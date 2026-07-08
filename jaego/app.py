"""
유통기한 재고 출고 모니터
- 로컬 실행: watchlist.json 사용
- 클라우드 배포: Supabase DB 사용 (st.secrets에 SUPABASE_URL, SUPABASE_KEY 설정)
"""

import streamlit as st
import pandas as pd
import io
import json
import re
from datetime import date
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


_XLSX_COLS = ["창고", "품목코드", "품목명", "등록 유통기한",
              "출고진행 유통기한", "최신 출고 유통기한", "상태", "비고", "담당자", "확인여부"]
_XLSX_FILL = {"red": "FFCCCC", "orange": "FFE0B2", "": "FFFFFF"}


def build_result_xlsx(df) -> bytes:
    """결과 DataFrame → 색상 입힌 엑셀 bytes (창고별로 정렬)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "유통기한 모니터"
    ws.append(_XLSX_COLS)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="305496")
        c.alignment = Alignment(horizontal="center", vertical="center")
    if df is not None and not df.empty:
        for _, r in df.sort_values(["창고", "상태"]).iterrows():
            ws.append([r.get(c, "") for c in _XLSX_COLS])
            fill = PatternFill("solid", fgColor=_XLSX_FILL.get(r.get("_color", ""), "FFFFFF"))
            for cell in ws[ws.max_row]:
                cell.fill = fill
    widths = [8, 11, 30, 13, 20, 16, 14, 40, 12, 8]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ─── Supabase 감지 ─────────────────────────────────────────────────────────
# 우선순위: JAEGO_SUPABASE_* (재고모니터 전용/기존 프로젝트) > SUPABASE_* (통합 공용)
def _sb_url():
    try:
        return st.secrets.get("JAEGO_SUPABASE_URL") or st.secrets.get("SUPABASE_URL")
    except Exception:
        return None


def _sb_key():
    try:
        return st.secrets.get("JAEGO_SUPABASE_KEY") or st.secrets.get("SUPABASE_KEY")
    except Exception:
        return None


def _supabase_configured() -> bool:
    return bool(_sb_url()) and bool(_sb_key())

USE_SUPABASE = _supabase_configured()

if USE_SUPABASE:
    from supabase import create_client

    @st.cache_resource
    def _sb():
        return create_client(_sb_url(), _sb_key())

# ─── 컬럼 인덱스 ──────────────────────────────────────────────────────────
COL_품목코드 = 4
COL_품목명   = 5
COL_유통기한 = 9
COL_출고예정 = 14
COL_LOCK    = 23

APP_DIR        = Path(__file__).parent
WATCHLIST_FILE = APP_DIR / "watchlist.json"


# ─── 워치리스트 CRUD ────────────────────────────────────────────────────────

WATCHLIST_TABLE_MISSING = False  # watchlist 테이블 미생성 감지 플래그


def load_watchlist() -> list[dict]:
    global WATCHLIST_TABLE_MISSING
    if USE_SUPABASE:
        try:
            resp = _sb().table("watchlist").select("*").order("id").execute()
            return resp.data or []
        except Exception:
            WATCHLIST_TABLE_MISSING = True   # 테이블 없음 → 빈 목록(안내는 UI에서)
            return []
    # 로컬 폴백
    if WATCHLIST_FILE.exists():
        with open(WATCHLIST_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def add_item(code: str, expiry: str) -> bool:
    """추가. 중복이면 False 반환."""
    wl = load_watchlist()
    if any(str(r["code"]) == code and str(r["expiry"]) == expiry for r in wl):
        return False
    if USE_SUPABASE:
        try:
            _sb().table("watchlist").insert({"code": code, "expiry": expiry}).execute()
        except Exception:
            return False
    else:
        wl.append({"code": code, "expiry": expiry})
        _save_local(wl)
    return True


def delete_item(row_id) -> None:
    if USE_SUPABASE:
        try:
            _sb().table("watchlist").delete().eq("id", row_id).execute()
        except Exception:
            pass
    else:
        wl = load_watchlist()
        wl = [r for r in wl if r.get("id") != row_id]
        _save_local(wl)


def _save_local(wl: list) -> None:
    # 로컬 전용 - id가 없으면 부여
    for i, r in enumerate(wl):
        if "id" not in r:
            r["id"] = i + 1
    with open(WATCHLIST_FILE, "w", encoding="utf-8") as f:
        json.dump(wl, f, ensure_ascii=False, indent=2)


# ─── 담당자 매핑 (품목코드 → 담당자) ─────────────────────────────────────────

def _fetch_담당자_supabase() -> dict:
    """통합 Supabase app_settings['inventory_담당자']에서 로드 (재고분석기와 공유)."""
    try:
        url = st.secrets.get("SUPABASE_URL")
        key = st.secrets.get("SUPABASE_KEY")
        if not (url and key):
            return {}
        from supabase import create_client
        cli = create_client(url, key)
        r = cli.table("app_settings").select("value").eq("key", "inventory_담당자").execute()
        raw = (r.data[0].get("value") or {}) if r.data else {}
        return {str(k).strip(): str(v).strip() for k, v in raw.items()}
    except Exception:
        return {}


def _parse_담당자_xlsx(file) -> dict:
    """업로드 담당자 엑셀 → {품목코드: 담당자} (1열 코드, 2열 담당자)."""
    try:
        d = pd.read_excel(file, dtype=str, header=0)
        if d.shape[1] < 2:
            return {}
        cc, nc = d.columns[0], d.columns[1]
        m = {}
        for _, r in d.iterrows():
            c = str(r[cc]).strip()
            n = str(r[nc]).strip()
            if c and c.lower() != "nan" and n and n.lower() != "nan":
                m[c] = n
        return m
    except Exception:
        return {}


# ─── 확인 체크 (영속) ────────────────────────────────────────────────────────

CHECKS_TABLE_MISSING = False
CHECKS_FILE = APP_DIR / "checks.json"


def load_checks() -> dict:
    """{key: True} — 확인 완료된 항목만."""
    global CHECKS_TABLE_MISSING
    if USE_SUPABASE:
        try:
            resp = _sb().table("jaego_checks").select("key,checked").execute()
            return {r["key"]: True for r in (resp.data or []) if r.get("checked")}
        except Exception:
            CHECKS_TABLE_MISSING = True
            return {}
    if CHECKS_FILE.exists():
        try:
            with open(CHECKS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def set_check(key: str, checked: bool) -> None:
    if USE_SUPABASE:
        try:
            if checked:
                _sb().table("jaego_checks").upsert(
                    {"key": key, "checked": True}, on_conflict="key").execute()
            else:
                _sb().table("jaego_checks").delete().eq("key", key).execute()
        except Exception:
            pass
    else:
        data = load_checks()
        if checked:
            data[key] = True
        else:
            data.pop(key, None)
        with open(CHECKS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)


# ─── 수량 파싱 ─────────────────────────────────────────────────────────────

def has_qty(val) -> bool:
    if pd.isna(val) or str(val).strip() in ("", "nan", "None", "NaN"):
        return False
    tokens = re.findall(r"[\d,]+", str(val))
    return any(float(t.replace(",", "")) > 0 for t in tokens)


def to_exp_str(raw) -> str | None:
    if pd.isna(raw) or str(raw).strip() in ("", "nan"):
        return None
    try:
        return str(int(float(str(raw).strip())))
    except Exception:
        return str(raw).strip()


# ─── 분석 로직 ─────────────────────────────────────────────────────────────

def analyze_item(df: pd.DataFrame, code: str, target_exp: str):
    """(품목명, status, color, note, 출고중_유통기한_목록) 반환."""
    mask    = df.iloc[:, COL_품목코드].astype(str).str.strip() == code.strip()
    item_df = df[mask]

    if item_df.empty:
        return "-", "미발견", "", "파일에서 품목을 찾을 수 없음", []

    품목명 = str(item_df.iloc[0, COL_품목명])

    exp_info: dict[str, dict] = {}
    for _, row in item_df.iterrows():
        exp = to_exp_str(row.iloc[COL_유통기한])
        if exp is None:
            continue
        if exp not in exp_info:
            exp_info[exp] = {"has_lock_free": False, "in_출고": False}
        if not has_qty(row.iloc[COL_LOCK]):
            exp_info[exp]["has_lock_free"] = True
        if has_qty(row.iloc[COL_출고예정]):
            exp_info[exp]["in_출고"] = True

    all_exps    = sorted(exp_info.keys())
    출고중_목록 = [e for e in all_exps if exp_info[e]["in_출고"]]

    if target_exp not in all_exps:
        return 품목명, "미발견", "", f"등록 유통기한({target_exp})이 파일에 없음", 출고중_목록

    idx         = all_exps.index(target_exp)
    before_exps = all_exps[:idx]
    after_exps  = all_exps[idx + 1:]
    t           = exp_info[target_exp]
    직전        = before_exps[-1] if before_exps else None
    이후_출고   = [e for e in after_exps if exp_info[e]["in_출고"]]

    if 이후_출고:
        return 품목명, "위험", "red", f"이후 유통기한 출고진행 중: {', '.join(이후_출고)}", 출고중_목록

    if t["in_출고"]:
        return 품목명, "주의(락 점검)", "orange", f"동일 유통기한({target_exp}) 출고진행 중", 출고중_목록

    if not t["has_lock_free"] and 직전 and exp_info[직전]["in_출고"]:
        return (
            품목명, "주의(락 점검)", "orange",
            f"직전 유통기한({직전}) 출고진행 중  /  락 없는 동일 유통기한 재고 없음",
            출고중_목록,
        )

    note = "락 없는 동일 유통기한 재고 있음" if t["has_lock_free"] else "출고진행 없음"
    return 품목명, "정상", "", note, 출고중_목록


# ─── UI ───────────────────────────────────────────────────────────────────

try:  # 단독 실행 시에만 (통합 Home.py에서 실행되면 무시)
    st.set_page_config(page_title="유통기한 재고 모니터", layout="wide", page_icon="📦")
except Exception:
    pass
st.title("📦 신규파우치 점검 프로그램")

# watchlist 테이블 미생성 감지 → 1회 안내 (테이블 만들면 사라짐)
if USE_SUPABASE:
    load_watchlist()  # 플래그 세팅용
    if WATCHLIST_TABLE_MISSING:
        st.warning(
            "⚠ `watchlist` 테이블이 없습니다. 방법 2가지:\n\n"
            "**A) 기존 재고모니터 데이터 그대로 쓰기** — 앱 Secrets에 기존 재고모니터 "
            "Supabase 값을 `JAEGO_SUPABASE_URL` / `JAEGO_SUPABASE_KEY` 로 추가 (이관 불필요)\n\n"
            "**B) 이 통합 Supabase에 새로 만들기** — SQL Editor에서 1회 실행:\n"
            "```sql\ncreate table watchlist (\n  id serial primary key,\n"
            "  code varchar(50) not null,\n  expiry varchar(8) not null,\n"
            "  unique(code, expiry)\n);\n```")

# ── 품목 등록 (숨김 토글) ──────────────────────────────────────────────────
col_cb, col_desc = st.columns([2, 5])
with col_cb:
    show_register = st.checkbox("⚙️ 품목 등록/삭제 열기", value=False)
with col_desc:
    st.markdown("**신규 파우치 품목코드 유통기한 등록**")

if show_register:
    with st.container(border=True):
        st.subheader("🔖 품목 등록 / 삭제")
        left, right = st.columns([1, 1])

        # 등록
        with left:
            with st.form("add_form", clear_on_submit=True):
                inp_code   = st.text_input("품목코드", placeholder="예: 2061438")
                inp_expiry = st.text_input("유통기한 (YYYYMMDD)", placeholder="예: 20270827")
                submitted  = st.form_submit_button("➕ 등록", use_container_width=True)

            if submitted:
                if inp_code and inp_expiry:
                    ok = add_item(inp_code.strip(), inp_expiry.strip())
                    if ok:
                        st.success(f"등록 완료: **{inp_code.strip()}** / `{inp_expiry.strip()}`")
                        st.rerun()
                    else:
                        st.warning("이미 등록된 항목입니다")
                else:
                    st.error("품목코드와 유통기한을 모두 입력해주세요")

        # 삭제
        with right:
            wl = load_watchlist()
            if not wl:
                st.info("등록된 품목이 없습니다")
            else:
                st.markdown("**등록 목록** (🗑 클릭 시 삭제)")
                for item in wl:
                    c1, c2 = st.columns([5, 1])
                    c1.markdown(f"**{item['code']}** &nbsp; `{item['expiry']}`")
                    if c2.button("🗑", key=f"del_{item['id']}"):
                        delete_item(item["id"])
                        st.rerun()

st.divider()

# ── 담당자 데이터 (분석결과에 품목 담당자 표시) ──────────────────────────────
with st.expander("👤 담당자 데이터 (분석결과에 품목 담당자 표시)"):
    st.caption("품목코드 → 담당자 매핑. 재고분석기에 등록된 담당자(Supabase)를 자동으로 불러오며, "
               "파일을 올리면 그 파일이 우선합니다. (1열: 품목코드, 2열: 담당자)")
    up_dam = st.file_uploader("담당자 로우데이터 xlsx", type=["xlsx"], key="jaego_dam")
    if up_dam is not None:
        st.session_state["_dam_map"] = _parse_담당자_xlsx(up_dam)
    dam_map = st.session_state.get("_dam_map") or _fetch_담당자_supabase()
    if dam_map:
        _src = "업로드 파일" if st.session_state.get("_dam_map") else "Supabase(재고분석기 공유)"
        st.success(f"담당자 {len(dam_map)}명 로드됨 · 출처: {_src}")
    else:
        st.info("담당자 데이터 없음 — '담당자' 열은 비어서 표시됩니다.")

# ── 파일 업로드 및 분석 ────────────────────────────────────────────────────
st.header("📂 파일 업로드")
uploaded = st.file_uploader(
    "로케이션별 재고조회 Excel 파일 (.xlsx)",
    type=["xlsx", "xls"],
)

if uploaded:
    wl = load_watchlist()
    if not wl:
        st.warning("⚠️ 등록된 품목이 없습니다. 위 '품목 등록/삭제 열기'를 체크하여 품목을 등록해주세요.")
    else:
        df = pd.read_excel(uploaded, dtype=str, header=0)
        # ── 창고 선택: 전체창고 파일도 창고별로 분리 (섞이지 않음) ──
        _whs_all = sorted({str(v).strip() for v in df.iloc[:, 0].dropna()
                           if str(v).strip()})
        # 기본 체크: IC930·IC920·IC906 (파일에 있는 것만)
        _default = [w for w in ("IC930", "IC920", "IC906") if w in _whs_all] or _whs_all
        if len(_whs_all) > 1:
            sel_whs = st.multiselect(
                "🏬 창고 선택 (창고별로 분리 표시 — 보고 싶은 창고만 체크, 섞이지 않음)",
                _whs_all, default=_default)
        else:
            sel_whs = _whs_all
        if not sel_whs:
            sel_whs = _whs_all
        with st.spinner("분석 중..."):
            rows = []
            for wh in sel_whs:
                df_wh = df[df.iloc[:, 0].astype(str).str.strip() == wh]
                for item in wl:
                    품목명, status, color, note, 출고중 = analyze_item(
                        df_wh, str(item["code"]), str(item["expiry"])
                    )
                    if status == "미발견":
                        continue  # 그 창고 파일에 품목/유통기한 없으면 제외
                    rows.append(
                        {
                            "창고":            wh,
                            "품목코드":         str(item["code"]),
                            "품목명":           품목명,
                            "등록 유통기한":     str(item["expiry"]),
                            "출고진행 유통기한":  ", ".join(출고중) if 출고중 else "-",
                            "최신 출고 유통기한": max(출고중) if 출고중 else "-",
                            "상태":            status,
                            "비고":            note,
                            "담당자":          dam_map.get(str(item["code"]).strip(), ""),
                            "_color":          color,
                        }
                    )

        result_df = pd.DataFrame(rows)
        if result_df.empty:
            st.info("선택한 창고에서 등록 품목을 찾지 못했습니다. 창고 체크를 확인하세요.")
            st.stop()
        BG = {"red": "#FFCCCC", "orange": "#FFE0B2", "": "#FFFFFF"}

        # 전체 요약
        c1, c2, c3 = st.columns(3)
        c1.metric("🔴 위험", int((result_df["_color"] == "red").sum()),
                  help="이후 유통기한 출고진행 중")
        c2.metric("🟠 주의", int((result_df["_color"] == "orange").sum()),
                  help="동일/직전 유통기한 출고진행 중")
        c3.metric("✅ 정상", int((result_df["_color"] == "").sum()),
                  help="락 없는 동일 유통기한 재고 있음")

        # ── 확인 체크 로드 (+ 엑셀용 확인여부 열) ──
        checks = load_checks()
        if USE_SUPABASE and CHECKS_TABLE_MISSING:
            st.warning(
                "⚠ `jaego_checks` 테이블이 없어 확인 체크가 저장되지 않습니다. "
                "SQL Editor에서 1회 실행:\n\n```sql\ncreate table jaego_checks (\n"
                "  key text primary key,\n  checked boolean not null default true,\n"
                "  updated timestamptz default now()\n);\n```")

        def _mk_key(r):
            return "|".join([str(r["창고"]), str(r["품목코드"]), str(r["등록 유통기한"]),
                             str(r["상태"]), str(r["출고진행 유통기한"])])
        result_df["확인여부"] = ["✔" if checks.get(_mk_key(r), False) else ""
                              for _, r in result_df.iterrows()]

        # 비정상만 보기 / 확인완료 숨기기 / 엑셀 다운로드
        _f1, _f2, _f3 = st.columns([2, 2, 2])
        with _f1:
            abn_only = st.toggle("⚠ 비정상(위험·주의)만 보기", value=False)
        with _f2:
            hide_checked = st.toggle("✔ 확인완료 숨기기", value=False,
                                     help="이미 확인(체크)한 항목을 목록에서 숨겨 반복 작업 방지")
        view_df = (result_df[result_df["_color"].isin(["red", "orange"])]
                   if abn_only else result_df)
        with _f3:
            st.download_button(
                "📥 결과 엑셀 다운로드",
                build_result_xlsx(view_df),
                f"유통기한모니터_{date.today():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        if view_df.empty:
            st.info("표시할 결과가 없습니다 (비정상 없음).")
        st.caption("✔ '확인' 열을 체크하면 저장되어, 다음에 다시 분석해도 유지됩니다 (동일 작업 반복 방지).")
        st.subheader("📊 분석 결과 (창고별)")

        _SEV = {"red": 0, "orange": 1, "": 2}
        _EMO = {"red": "🔴", "orange": "🟠", "": "✅"}
        for wh in sel_whs:
            wh_df = view_df[view_df["창고"] == wh].copy()
            if wh_df.empty:
                continue
            wh_df = wh_df.sort_values("_color", key=lambda s: s.map(_SEV),
                                      kind="stable").reset_index(drop=True)
            # 행별 key + 확인 seed (+ 확인완료 숨기기 필터)
            keys, seed, keep = [], [], []
            for _, r in wh_df.iterrows():
                k = _mk_key(r)
                chk = checks.get(k, False)
                if hide_checked and chk:
                    keep.append(False)
                    continue
                keep.append(True)
                keys.append(k)
                seed.append(chk)
            wh_df = wh_df[keep].reset_index(drop=True)
            if wh_df.empty:
                continue
            _r = int((wh_df["_color"] == "red").sum())
            _o = int((wh_df["_color"] == "orange").sum())
            _k = int((wh_df["_color"] == "").sum())
            st.markdown(f"#### 🏬 {wh}  —  🔴 위험 {_r} · 🟠 주의 {_o} · ✅ 정상 {_k}")
            disp = wh_df.drop(columns=["_color", "창고", "확인여부"]).copy()
            disp["상태"] = [f'{_EMO.get(c, "")} {s}'
                          for c, s in zip(wh_df["_color"], wh_df["상태"])]
            disp.insert(0, "확인", seed)
            _bg = wh_df["_color"].tolist()   # 행별 색 (disp 행 순서와 일치)

            def _row_style(row, bg=_bg):
                return [f"background-color: {BG.get(bg[row.name], '#FFFFFF')}"] * len(row)

            # Styler로 행 전체 배경색 유지 + 체크박스 편집 가능
            styled = disp.style.apply(_row_style, axis=1).map(
                lambda _: "font-weight: bold", subset=["등록 유통기한"])
            edited = st.data_editor(
                styled, key=f"ed_{wh}", hide_index=True, use_container_width=True,
                height=min(600, 60 + len(disp) * 38),
                column_config={"확인": st.column_config.CheckboxColumn(
                    "확인", help="확인 완료 시 체크 — 저장되어 다음 분석에도 유지", default=False)},
                disabled=[c for c in disp.columns if c != "확인"])
            # 변경분만 저장 후 새로고침(확인여부·숨기기 즉시 반영)
            changed = False
            for i, k in enumerate(keys):
                newv = bool(edited.iloc[i]["확인"])
                if newv != seed[i]:
                    set_check(k, newv)
                    changed = True
            if changed:
                st.rerun()

        st.markdown("""
---
**범례**

| 색상 | 의미 |
|------|------|
| 🔴 빨간 (위험) — 역순출고 우려 | 등록 유통기한보다 **이후** 유통기한이 출고 진행 중 |
| 🟠 주황 (주의) — 신규파우치 출고 임박 (락해제 점검 필요) | **동일** 유통기한 출고 진행 중  ·  또는  ·  락 없는 동일 유통기한 재고 없고 **직전** 유통기한 출고 진행 중 |
| ⬜ 색 없음 (정상) | 락 없는 동일 유통기한 재고 있음 |
        """)

        st.markdown("""
<div style="
    background-color: #FFF3CD;
    border: 3px solid #FF0000;
    border-radius: 8px;
    padding: 18px 24px;
    margin-top: 16px;
    text-align: center;
">
    <span style="font-size: 22px; font-weight: 900; color: #CC0000;">
        ⚠️ 락 해제 시 메일 발송 필수
    </span><br>
    <span style="font-size: 17px; font-weight: 700; color: #333;">
        수신 : 영업팀 전부 &nbsp;|&nbsp; 참조 : SCM
    </span>
</div>
""", unsafe_allow_html=True)
