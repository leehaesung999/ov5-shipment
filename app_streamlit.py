"""OV5 지정출고 자동매칭 — Streamlit 웹 UI.

흐름:
  1. 사이드바에 기준정보·옵션
  2. 재고 파일 업로드 → 즉시 Lock 재고 현황판
  3. 주문 파일 업로드 → 즉시 매칭 결과
  4. 결과 다운로드 (자동 로트 누적은 다운로드 클릭 시 수행)
"""
from __future__ import annotations

import io
import json
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

from core.stock_loader import load_locked_stock
from core.order_loader import load_orders
from core.master_loader import (
    load_shelf_life_months, load_fs_ms_items, load_lot_assignments,
)
from core import store


def migrate_fs_ms_to_lot() -> int:
    """기존 fs_ms_items.xlsx가 있으면 lot_assignments_auto.xlsx로 1회 자동 변환.
    이미 lot에 있는 (품목+유통기한)은 스킵. 변환 후 fs_ms 파일은 .bak 으로 이동.
    (클라우드/Supabase 모드에서는 레거시 파일이 없으므로 스킵)"""
    if store.use_supabase() or not FS_MS_PATH.exists():
        return 0
    fs_ms = load_fs_ms_items(str(FS_MS_PATH))
    if not fs_ms:
        return 0
    manual = load_lot_assignments(str(LOT_PATH))
    auto = load_lot_assignments(str(LOT_AUTO_PATH))
    seen = set(manual.keys()) | set(auto.keys())

    new_rows = []
    for code, rules in fs_ms.items():
        for ymd, tag in rules:
            if (code, ymd) in seen:
                continue
            seen.add((code, ymd))
            new_rows.append({
                "Item code": int(code), "Item": "",
                "유통기한": ymd if ymd else "", "카테고리": tag,
            })
    if not new_rows:
        FS_MS_PATH.rename(FS_MS_PATH.with_suffix(".bak.xlsx"))
        return 0

    if LOT_AUTO_PATH.exists():
        try:
            existing_df = pd.read_excel(str(LOT_AUTO_PATH))
        except Exception:
            existing_df = pd.DataFrame(columns=["Item code", "Item", "유통기한", "카테고리"])
    else:
        existing_df = pd.DataFrame(columns=["Item code", "Item", "유통기한", "카테고리"])
    merged = pd.concat([existing_df, pd.DataFrame(new_rows)], ignore_index=True)
    CONFIG_DIR.mkdir(exist_ok=True)
    merged.to_excel(str(LOT_AUTO_PATH), index=False, sheet_name="로트지정_자동")
    FS_MS_PATH.rename(FS_MS_PATH.with_suffix(".bak.xlsx"))
    return len(new_rows)
from core.shelf_life import calc_remaining_rate, calc_production_date, judge
from core.matcher import match, determine_lot_category
from core.writer import write_summary

CONFIG_DIR = ROOT / "config"
SETTINGS_PATH = CONFIG_DIR / "settings.json"
MASTER_CACHE = CONFIG_DIR / "master_info.xlsx"
FS_MS_PATH = CONFIG_DIR / "fs_ms_items.xlsx"
LOT_PATH = CONFIG_DIR / "lot_assignments.xlsx"
LOT_AUTO_PATH = CONFIG_DIR / "lot_assignments_auto.xlsx"
TMP_DIR = ROOT / "tmp_uploads"
OUT_DIR = ROOT / "output"

DEFAULT_SETTINGS = {
    "threshold": 0.70,
    "thresholds": {
        "대리점": 0.70, "FS": 0.70, "MS": 0.70,
        "급식": 0.70, "소재": 0.70, "온라인": 0.70,
    },
    "thresholds_ignore": ["급식"],  # 잔존율 무관 처리할 카테고리
    "ov_locations": ["OV5"],
    "nh_branch_keywords": {
        "포천": "농협포천", "평택": "농협평택", "횡성": "농협횡성",
        "군위": "농협군위", "장성": "농협장성", "경남": "농협경남",
    },
}


# ---------------- helpers ----------------
def load_settings() -> dict:
    return store.load_settings(DEFAULT_SETTINGS)


def save_settings(s: dict) -> None:
    store.save_settings(s)


def save_uploaded(file, dst: Path) -> Path:
    dst.parent.mkdir(exist_ok=True)
    dst.write_bytes(file.getvalue())
    return dst


# ---------------- 속도 개선: 무거운 엑셀 파싱 결과 캐시 ----------------
# 같은 파일/기준정보면 재실행(클릭)마다 다시 파싱하지 않고 캐시 재사용.
@st.cache_data(show_spinner=False)
def cached_locked_stock(data: bytes, ov_locations: tuple) -> pd.DataFrame:
    TMP_DIR.mkdir(exist_ok=True)
    p = TMP_DIR / "_cache_stock.xlsx"
    p.write_bytes(data)
    return load_locked_stock(str(p), list(ov_locations))


@st.cache_data(show_spinner=False)
def cached_orders(data: bytes) -> pd.DataFrame:
    TMP_DIR.mkdir(exist_ok=True)
    p = TMP_DIR / "_cache_orders.xlsx"
    p.write_bytes(data)
    return load_orders(str(p))


@st.cache_data(show_spinner=False)
def cached_shelf_map(_mtime: float) -> dict:
    # _mtime(기준정보 파일 수정시각)이 바뀌면 자동으로 캐시 갱신
    return load_shelf_life_months(str(MASTER_CACHE)) if MASTER_CACHE.exists() else {}


def _master_mtime() -> float:
    return MASTER_CACHE.stat().st_mtime if MASTER_CACHE.exists() else 0.0


def prune_auto_lots(stocks: pd.DataFrame) -> int:
    """오늘 재고에 없는 자동 lot 제거 (백엔드는 store가 처리)."""
    valid = set()
    for _, s in stocks.iterrows():
        try:
            code = str(int(s["제품코드"]))
            ymd = int(s["소비기한"]) if pd.notna(s["소비기한"]) else None
        except (ValueError, TypeError):
            continue
        valid.add((code, ymd))
    try:
        return store.prune_auto_lots(valid)
    except Exception as e:
        st.warning(f"⚠ 자동 lot 정리 실패(건너뜀): {e}")
        return 0


def append_auto_lots(df: pd.DataFrame) -> int:
    """매칭 결과 기반 자동 lot 누적 (기존에 없는 키만 추가)."""
    nh_cols = [c for c in df.columns if c.startswith("농협")]
    seen = set(store.get_manual_lots()) | set(store.get_auto_lots())

    new_map, names = {}, {}
    for _, r in df.iterrows():
        qty = r.get("매칭수량")
        if qty is None or pd.isna(qty) or qty == 0:
            continue
        try:
            item_id = str(int(r["Item ID"]))
            ymd_int = int(r["유통기한"]) if pd.notna(r.get("유통기한")) else None
        except (ValueError, TypeError):
            continue
        if (item_id, ymd_int) in seen:
            continue
        cat = determine_lot_category(r.to_dict(), nh_cols)
        cats = store.parse_cats(cat)
        if not cats:
            continue
        seen.add((item_id, ymd_int))
        new_map[(item_id, ymd_int)] = cats
        names[(item_id, ymd_int)] = r.get("Item") or ""

    if not new_map:
        return 0
    try:
        return store.upsert_auto_lots(new_map, names, overwrite=False)
    except Exception as e:
        st.warning(f"⚠ 자동 누적 실패: {e}")
        return 0


def persist_user_lots(user_lots: dict) -> int:
    """화면에서 직접 체크한 카테고리(다중 포함)를 자동 lot에 누적 저장.
    같은 (품목+유통기한)은 화면 선택값으로 갱신 → 다음 실행 때 그대로 체크됨.
    (재고에 없는 항목은 다음 매칭 시 prune_auto_lots가 자동 차감)
    """
    if not user_lots:
        return 0
    try:
        return store.upsert_auto_lots(user_lots, overwrite=True)
    except Exception as e:
        st.warning(f"⚠ 화면 체크 저장 실패: {e}")
        return 0


def build_stock_preview(stocks: pd.DataFrame, shelf_map: dict,
                        threshold: float, today: date) -> pd.DataFrame:
    """OV Lock 재고에 잔존율·판정 컬럼 미리 계산해 보여주기."""
    rows = []
    for _, s in stocks.iterrows():
        code = str(s["제품코드"])
        months = shelf_map.get(code, 24)
        rate = calc_remaining_rate(s["소비기한"], months, today)
        produced = calc_production_date(s["소비기한"], months)
        rows.append({
            "Item ID": code,
            "Item": s["제품명"],
            "유통기한": int(s["소비기한"]) if pd.notna(s["소비기한"]) else None,
            "Lock(Box)": int(s["Lock Qty(Box)"]),
            "현재고(Box)": int(s["현재고(Box)"]),
            "로케이션": s["Location"],
            "유통기한(월)": months,
            "제조일자": produced,
            "잔존율": round(rate, 4) if rate is not None else None,
            "판정": judge(rate, threshold),
        })
    return pd.DataFrame(rows)


def style_match_result(df: pd.DataFrame):
    """OK 녹색, NG 주황, 매칭된 행 노랑 — 행 단위 색상."""
    def row_color(r):
        mq = r.get("매칭수량")
        if pd.notna(mq) and mq and mq > 0:
            return ["background-color: #FFF7CC"] * len(r)
        v = r.get("판정")
        if v == "OK":
            return ["background-color: #E7F4E4"] * len(r)
        if v == "NG":
            return ["background-color: #FAD9C9"] * len(r)
        return [""] * len(r)
    return df.style.apply(row_color, axis=1)


# ---------------- main UI ----------------
st.set_page_config(page_title="OV5 지정출고 자동매칭", layout="wide")
st.title("OV5 지정출고 자동매칭")
st.caption("재고 파일 업로드 → 현황 확인 → 주문 파일 업로드 → 매칭 결과 → 다운로드")

settings = load_settings()
_migrated = migrate_fs_ms_to_lot()
if _migrated:
    st.toast(f"🔁 기존 FS/MS 양식 {_migrated}건을 로트 지정으로 통합했습니다", icon="✅")

# ===== 사이드바 =====
with st.sidebar:
    if store.use_supabase():
        st.caption("🟢 공유 모드 · Supabase (여러 사용자 데이터 공유)")
    else:
        st.caption("💾 로컬 모드 · 이 PC/서버 파일 저장")
    st.header("⚙️ 설정 / 기준정보")

    settings["threshold"] = st.number_input(
        "잔존율 기준 (전체 기본, fallback 용)",
        min_value=0.0, max_value=1.0, value=float(settings["threshold"]),
        step=0.05, format="%.2f")

    with st.expander("📐 카테고리별 잔존율 기준 (농협은 항상 무관)", expanded=True):
        prev_thr = dict(settings.get("thresholds") or {})
        prev_ignore = set(settings.get("thresholds_ignore") or [])
        thr = dict(prev_thr)
        ignore = set(prev_ignore)
        st.caption("💡 **무관** 체크 시 그 카테고리는 잔존율 검사 스킵 (NG 재고도 매칭). "
                    "농협은 항상 무관. 변경 즉시 저장됨.")
        for k in ("대리점", "FS", "MS", "급식", "소재", "온라인"):
            c1, c2 = st.columns([5, 2])
            # ① 체크박스를 먼저 평가 → is_ignored 결정 → 같은 rerun 내에서 disabled에 반영
            with c2:
                is_ignored = st.checkbox(
                    "무관", value=(k in ignore), key=f"ign_{k}",
                    help="체크 시 잔존율 검사 스킵")
            if is_ignored:
                ignore.add(k)
            else:
                ignore.discard(k)
            # ② 그 결과로 number_input disabled 결정
            label = f"{k}" + ("  (무관)" if is_ignored else "")
            with c1:
                val = st.number_input(
                    label, min_value=0.0, max_value=1.0,
                    value=float(thr.get(k, settings["threshold"])),
                    step=0.05, format="%.2f", key=f"thr_{k}",
                    disabled=is_ignored,
                    help=f"{k} 거래처 매칭 시 필요한 최소 잔존율"
                          + (" (현재 무관 — 체크 해제하면 적용)" if is_ignored else ""))
                thr[k] = val
        settings["thresholds"] = thr
        settings["thresholds_ignore"] = sorted(ignore)
        # 변경 즉시 디스크 저장 (새로고침/재실행 후에도 보존)
        if thr != prev_thr or set(ignore) != prev_ignore:
            save_settings(settings)

    settings["ov_locations"] = st.multiselect(
        "OV 로케이션 (Lock 재고 대상)",
        options=["OV1", "OV4", "OV5", "OV6"],
        default=settings.get("ov_locations", ["OV5"]))

    today = st.date_input("기준일자 (오늘)", value=date.today())

    if st.button("💾 설정 저장", use_container_width=True):
        save_settings(settings)
        st.success("저장됨")

    st.divider()
    st.subheader("📚 기준정보 (유통기한월)")
    if MASTER_CACHE.exists():
        n = len(cached_shelf_map(_master_mtime()))
        st.success(f"등록 {n}품목")
    else:
        st.warning("등록 안 됨 (24개월 기본값)")
    up = st.file_uploader("기준정보 xlsx 업로드", type=["xlsx"], key="master_up",
                            help="새 양식: 소비기한(월) 컬럼, AF=상세 / 하대 없음 — "
                                 "업로드 시 자동으로 컬럼 정규화 + 하대=배면×배단 생성")
    if up:
        raw_path = TMP_DIR / up.name
        save_uploaded(up, raw_path)
        try:
            xl = pd.ExcelFile(raw_path)
            # 'Item code' + (유통기한|소비기한)(월) 둘 다 있는 시트 자동 탐색
            df = pd.DataFrame()
            for sh in xl.sheet_names:
                tmp = pd.read_excel(raw_path, sheet_name=sh)
                if "Item code" in tmp.columns and (
                        "유통기한(월)" in tmp.columns or "소비기한(월)" in tmp.columns):
                    df = tmp
                    break
            if df.empty:
                st.error("기준정보 시트를 못 찾았습니다 — Item code · 유통기한(월) 또는 "
                          "소비기한(월) 컬럼이 필요합니다.")
            else:
                # 컬럼명 정규화: 소비기한(월) → 유통기한(월)
                if ("소비기한(월)" in df.columns
                        and "유통기한(월)" not in df.columns):
                    df = df.rename(columns={"소비기한(월)": "유통기한(월)"})
                # 하대 자동 생성: 배면 × 배단
                ng_msg = ""
                if "하대" not in df.columns:
                    if "배면" in df.columns and "배단" in df.columns:
                        df["하대"] = (pd.to_numeric(df["배면"], errors="coerce") *
                                       pd.to_numeric(df["배단"], errors="coerce"))
                        ng_msg = f" · 하대 자동생성 {int(df['하대'].notna().sum())}건"
                    else:
                        ng_msg = " · (배면/배단 없어 하대 미생성)"
                CONFIG_DIR.mkdir(exist_ok=True)
                df.to_excel(str(MASTER_CACHE), index=False, sheet_name="기준정보")
                st.success(f"✅ 기준정보 {len(df)}품목 등록 완료{ng_msg}. "
                            "새로고침하세요.")
        except Exception as e:
            st.error(f"업로드 처리 오류: {e}")

    st.divider()
    st.subheader("📌 로트 지정 거래처 (FS/MS 통합)")
    manual_n = len(store.get_manual_lots())
    auto_n = len(store.get_auto_lots())
    st.caption(f"수동 {manual_n}건 / 자동 {auto_n}건")
    st.caption("자동: 다운로드 시 누적, 다음 매칭 시 오늘 재고에 없으면 제거")
    up_lot = st.file_uploader("로트 지정 xlsx 업로드 (수동)",
                               type=["xlsx"], key="lot_up")
    if up_lot:
        try:
            _ldf = pd.read_excel(up_lot)
            store.set_manual_lots_from_df(_ldf)
            st.success(f"업로드 완료 ({len(_ldf)}행). 새로고침하세요.")
        except Exception as e:
            st.error(f"업로드 오류: {e}")
    lot_tmpl = pd.DataFrame({
        "Item code": [1010422, 2061502, 1015028, 1014854, 2032260, 2054307],
        "Item": ["진간장 금S 15L", "폰타나 학교급식", "맛간장 1.7L",
                 "유산균발효양조간장", "쓱쓱싹싹 깻잎", "소재용 ○○"],
        "유통기한": [20280309, 20270917, "", 20280116, "", ""],
        "카테고리": ["농협", "MS", "농협,대리점", "FS", "급식", "소재"],
    })
    buf2 = io.BytesIO()
    lot_tmpl.to_excel(buf2, index=False, sheet_name="로트지정")
    st.download_button("로트 지정 템플릿", buf2.getvalue(),
                        "로트지정_템플릿.xlsx", use_container_width=True)
    if auto_n > 0:
        if st.button("🗑️ 자동 누적 lot 전체 초기화"):
            try:
                store.clear_auto_lots()
                st.success("초기화 완료. 새로고침하세요.")
            except Exception as e:
                st.error(f"⚠ 초기화 실패: {e}")

# ===== 메인 — 단계별 =====
st.subheader("① 재고 파일 업로드")
stock_up = st.file_uploader(
    "로케이션별 재고조회 xlsx", type=["xlsx"], key="stock_up",
    help="OV 로케이션의 Lock 재고만 추출됩니다.")

if not stock_up:
    st.info("재고 파일을 업로드하면 현황판이 나타납니다.")
    st.stop()

# 재고 처리 (캐시: 같은 파일이면 재파싱 안 함)
try:
    stocks = cached_locked_stock(stock_up.getvalue(), tuple(settings["ov_locations"]))
except Exception as e:
    st.error(f"재고 로드 오류: {e}")
    st.stop()

if stocks.empty:
    st.warning(f"{', '.join(settings['ov_locations'])}에 Lock 걸린 재고가 없습니다.")
    st.stop()

shelf_map = cached_shelf_map(_master_mtime())
preview = build_stock_preview(stocks, shelf_map, settings["threshold"], today)

# 자동/수동 lot에서 이미 지정된 카테고리 가져오기
auto_lots_for_preview = store.get_auto_lots()
manual_lots_for_preview = store.get_manual_lots()


LOT_CAT_COLS = ["농협", "대리점", "FS", "MS", "급식", "소재"]


def _lot_key(row):
    code = str(row["Item ID"])
    ymd = int(row["유통기한"]) if pd.notna(row["유통기한"]) else None
    return (code, ymd)


def _initial_cats(row) -> list:
    key = _lot_key(row)
    return list(manual_lots_for_preview.get(key)
                or auto_lots_for_preview.get(key) or [])


def _initial_source(row):
    key = _lot_key(row)
    if key in manual_lots_for_preview:
        return "수동"
    if key in auto_lots_for_preview:
        return "자동(전일 누적)"
    return ""


# 카테고리별 체크박스 컬럼 (한 행에 여러 개 = 중복 지정)
for cat in LOT_CAT_COLS:
    preview[cat] = preview.apply(lambda r, c=cat: c in _initial_cats(r), axis=1)
preview["출처"] = preview.apply(_initial_source, axis=1)

_assigned_n = int(preview[LOT_CAT_COLS].any(axis=1).sum())

st.success(f"Lock 재고 **{len(preview)}건** 추출됨")
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("총 재고 행", len(preview))
col2.metric("총 박스", int(preview["Lock(Box)"].sum()))
col3.metric("OK (≥ 70%)", int((preview["판정"] == "OK").sum()))
col4.metric("NG (< 70%)", int((preview["판정"] == "NG").sum()))
col5.metric("이미 지정", _assigned_n)

st.markdown(
    "##### 📋 재고 현황판 (확인 + 거래처 조건 설정)\n"
    "각 행의 **카테고리 체크박스**(농협/대리점/FS/MS/급식/소재)를 켜면 그 재고는 "
    "**체크한 카테고리들**의 주문에만 매칭됩니다 — **여러 개 동시 체크 가능**. "
    "하나도 안 켜면 기본 우선순위(농협→나머지) 적용. "
    "**출처**가 '자동(전일 누적)'이면 어제 매칭 결과가 미리 채워진 것입니다.")

_preview_cfg = {
    "Item ID": st.column_config.TextColumn(disabled=True),
    "Item": st.column_config.TextColumn(disabled=True),
    "유통기한": st.column_config.NumberColumn(format="%d", disabled=True),
    "Lock(Box)": st.column_config.NumberColumn(disabled=True),
    "현재고(Box)": st.column_config.NumberColumn(disabled=True),
    "로케이션": st.column_config.TextColumn(disabled=True),
    "유통기한(월)": st.column_config.NumberColumn(format="%d", disabled=True),
    "제조일자": st.column_config.DateColumn(disabled=True),
    "잔존율": st.column_config.NumberColumn(format="%.4f", disabled=True),
    "판정": st.column_config.TextColumn(disabled=True),
    "출처": st.column_config.TextColumn(
        help="수동 / 자동(전일 누적) / 빈값", disabled=True),
}
for cat in LOT_CAT_COLS:
    _preview_cfg[cat] = st.column_config.CheckboxColumn(
        cat, help=f"{cat} 주문에 매칭 허용", default=False)

edited_preview = st.data_editor(
    preview, use_container_width=True, hide_index=True,
    column_config=_preview_cfg, key="stock_editor")

# 체크된 카테고리들을 세션 lot에 모으기 (이번 매칭에만 적용)
user_lots: dict = {}
for _, r in edited_preview.iterrows():
    cats = [c for c in LOT_CAT_COLS if bool(r.get(c))]
    if not cats:
        continue
    try:
        code = str(int(r["Item ID"]))
        ymd = int(r["유통기한"]) if pd.notna(r["유통기한"]) else None
    except (ValueError, TypeError):
        continue
    user_lots[(code, ymd)] = cats

st.divider()
st.subheader("② 주문 파일 업로드")
orders_up = st.file_uploader(
    "출고진행현황 xlsx", type=["xlsx"], key="orders_up",
    help="이 파일이 업로드되면 자동으로 매칭이 실행됩니다.")

if not orders_up:
    st.info("주문 파일을 업로드하면 매칭이 자동 실행됩니다.")
    st.stop()

try:
    orders = cached_orders(orders_up.getvalue())
except Exception as e:
    st.error(f"주문 로드 오류: {e}")
    st.stop()
st.caption(f"주문 {len(orders)}행 로드")

# ===== 매칭 실행 =====
with st.spinner("매칭 계산 중..."):
    removed = prune_auto_lots(stocks)
    fs_ms = load_fs_ms_items(str(FS_MS_PATH)) if FS_MS_PATH.exists() else {}
    manual_lots = store.get_manual_lots()
    auto_lots = store.get_auto_lots()
    # 우선순위: 화면 편집값 > 수동 > 자동
    lots = {**auto_lots, **manual_lots, **user_lots}

    # ignore된 카테고리는 threshold=0.0으로 변환해 잔존율 무관 처리
    ignore_set = set(settings.get("thresholds_ignore") or [])
    effective_thr = {k: (0.0 if k in ignore_set else v)
                      for k, v in (settings.get("thresholds") or {}).items()}
    result = match(
        stocks, orders,
        shelf_life_map=shelf_map, fs_ms_items=fs_ms,
        lot_assignments=lots, today=today,
        threshold=float(settings["threshold"]),
        thresholds=effective_thr,
        nh_branches=settings["nh_branch_keywords"],
    )

if removed:
    st.info(f"🧹 자동 lot 정리: {removed}건 (오늘 재고에 없는 항목 제거)")
if user_lots:
    st.info(f"🖋️ 화면에서 직접 지정한 카테고리 {len(user_lots)}건 적용됨")

# ===== 매칭 결과 =====
st.divider()
st.subheader("③ 매칭 결과")
matched_mask = result["매칭수량"].fillna(0) > 0
ok_n = int((result["판정"] == "OK").sum())
ng_n = int((result["판정"] == "NG").sum())
matched_n = int(matched_mask.sum())

m1, m2, m3, m4 = st.columns(4)
m1.metric("총 결과 행", len(result))
m2.metric("매칭 성공", matched_n)
m3.metric("OK", ok_n)
m4.metric("NG", ng_n)

only_matched = st.toggle("✅ 매칭된 항목만 보기 / 저장", value=True,
                          help="끄면 매칭 안 된 잔여 재고까지 모두 표시")
view = result[matched_mask].reset_index(drop=True) if only_matched else result

_meta_cols = {"Item ID","Item","로케이션","판정구분","거래처명","제조일자","판정"}
_float_cols = {"잔존율","남은율"}
_view_config = {
    "잔존율": st.column_config.NumberColumn(format="%.4f"),
    "남은율": st.column_config.NumberColumn(format="%.4f"),
}
for _c in view.columns:
    if _c in _view_config or _c in _meta_cols or _c in _float_cols:
        continue
    _view_config[_c] = st.column_config.NumberColumn(format="%d")
st.dataframe(
    style_match_result(view), use_container_width=True, hide_index=True,
    column_config=_view_config)

# ===== 다운로드 =====
st.divider()
st.subheader("④ 결과 다운로드")
prefix = "지정출고_매칭결과_매칭만" if only_matched else "지정출고_매칭결과"
ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M")
fname = f"{prefix}_{ts}.xlsx"

if st.button(f"📥 {fname} 생성 + 자동 로트 누적", type="primary", use_container_width=True):
    OUT_DIR.mkdir(exist_ok=True)
    saved_checks = persist_user_lots(user_lots)      # 화면 체크 선택을 누적 저장
    appended = append_auto_lots(view)                # 매칭 결과 기반 누적
    path = write_summary(view, str(OUT_DIR), prefix=prefix)
    st.success(f"저장됨: `{path}`")
    if saved_checks:
        st.success(f"☑ 화면 체크 {saved_checks}건 저장 (다음 실행 시 체크됨)")
    if appended:
        st.success(f"📌 자동 로트 누적: {appended}건")
    with open(path, "rb") as f:
        st.download_button(
            "💾 다운로드 (브라우저로 받기)",
            f.read(), fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
