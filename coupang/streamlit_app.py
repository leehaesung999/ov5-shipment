# -*- coding: utf-8 -*-
"""쿠팡 지정출고 자동매칭 — Streamlit 웹 UI.

흐름:
  1) 사이드바: 기능2 대상품목 관리
  2) 재고조회 업로드 → 즉시 OV5 락재고 + 대상품목 출고가능 현황 표시
  3) 출고진행현황 업로드 → 자동 매칭 (FEFO)
  4) 매칭 결과 표에서 거래처 셀 직접 수정
  5) Excel 다운로드
"""
from __future__ import annotations

import copy
import os
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd
import streamlit as st

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import 지정출고_자동매칭 as core  # noqa: E402
import cloud_store  # noqa: E402

TMP = HERE / "tmp_uploads"
TMP.mkdir(exist_ok=True)


# ---------- 속도 개선: 무거운 엑셀 파싱/매칭 결과 캐시 ----------
# 같은 파일이면 재실행(클릭)마다 재파싱하지 않고 캐시 재사용.
def _master_mtime() -> float:
    p = core.MASTER_ITEM_CACHE
    return os.path.getmtime(p) if os.path.isfile(p) else 0.0


@st.cache_data(show_spinner=False)
def cached_master_count(_mtime: float) -> int:
    return len(core._load_hadae_from_item(core.MASTER_ITEM_CACHE))


@st.cache_data(show_spinner=False)
def cached_analyze(inv_bytes: bytes, out_bytes: bytes, target_sig: tuple, _mmtime: float):
    ip = TMP / "_cache_inv.xlsx"
    ip.write_bytes(inv_bytes)
    op = TMP / "_cache_out.xlsx"
    op.write_bytes(out_bytes)
    # 대상품목은 target_sig(클라우드/로컬에서 읽은 값)로 복원 — 캐시 키와 일치
    targets = [{"code": c, "name": n} for c, n in target_sig]
    return core.analyze(str(ip), str(op), targets=targets)


try:  # 단독 실행 시에만 적용 (통합 Home.py에서 실행되면 이미 설정됨 → 무시)
    st.set_page_config(page_title="쿠팡 지정출고 자동매칭", layout="wide")
except Exception:
    pass
st.title("쿠팡 지정출고 자동매칭")
st.caption("재고 업로드 → 출고진행 업로드 → 매칭 결과 + 거래처 수정 → 엑셀 다운로드")


# ================== 사이드바 ==================
with st.sidebar:
    if cloud_store.use_supabase():
        st.caption("🟢 공유 모드 · Supabase (대상품목 공유)")
    else:
        st.caption("💾 로컬 모드 · 이 PC 파일")
    # ----- 마스터 정보 (하대·팔레트) -----
    st.header("📚 마스터 정보")
    n_master = cached_master_count(_master_mtime())
    if n_master:
        st.success(f"하대 마스터: **{n_master}품목** 캐시됨")
    else:
        st.warning("하대 마스터 미등록 (Item_*.xlsx 업로드)")
    up_master = st.file_uploader(
        "Item_*.xlsx 업로드로 갱신",
        type=["xlsx"], key="master_up",
        help="배면×배단 = 하대 자동 계산 후 저장.")
    if up_master:
        tmp = TMP / up_master.name
        tmp.write_bytes(up_master.getvalue())
        ok, n = core.update_master_cache(str(tmp))
        if ok:
            st.success(f"마스터 갱신 완료: {n}품목 — 새로고침하세요")
        else:
            st.error("갱신 실패")

    st.divider()
    # ----- 기능2 대상 품목 -----
    st.header("⚙️ 기능2 대상 품목")
    targets = cloud_store.load_targets()
    targets_df = pd.DataFrame(targets, columns=["code", "name"])
    edited_t = st.data_editor(
        targets_df, num_rows="dynamic", hide_index=True,
        column_config={
            "code": st.column_config.NumberColumn("품목코드", format="%d"),
            "name": st.column_config.TextColumn("품명 (비우면 마스터에서 자동)"),
        },
        key="targets_editor",
    )
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("💾 저장", use_container_width=True):
            item_names = core.load_item_names()
            items = []
            for _, row in edited_t.iterrows():
                if pd.isna(row.get("code")):
                    continue
                try:
                    code = int(row["code"])
                except Exception:
                    continue
                name = str(row.get("name") or "").strip()
                if not name:
                    name = item_names.get(code, "")
                items.append({"code": code, "name": name})
            cloud_store.save_targets(items)
            n_auto = sum(1 for r in edited_t.iterrows()
                          if not str(r[1].get("name") or "").strip()
                          and not pd.isna(r[1].get("code")))
            msg = f"저장: {len(items)}개"
            if n_auto:
                msg += f" (마스터에서 {n_auto}건 자동 채움)"
            st.success(msg)
    with col_b:
        if st.button("🔄 기본11", use_container_width=True,
                     help="기본 11개로 초기화"):
            cloud_store.save_targets(core.DEFAULT_TARGETS)
            st.success("초기화됨. 새로고침")

    st.divider()
    st.caption("팔레트 정보는 '~/Desktop/지정출고/' 폴더의 최신 "
               "'통합 지정출고 양식(쿠팡)' 파일에서 자동 인식")


# ============= Step 1: 재고조회 업로드 =============
st.subheader("① 재고조회 파일 업로드")
inv_up = st.file_uploader(
    "로케이션별 재고조회 xlsx", type=["xlsx"], key="inv_up")
if not inv_up:
    st.info("재고조회 파일을 업로드하세요.")
    st.stop()

inv_path = TMP / inv_up.name
inv_path.write_bytes(inv_up.getvalue())
st.success(f"✅ {inv_up.name} 업로드 완료")


# ============= Step 2: 출고진행현황 업로드 =============
st.divider()
st.subheader("② 출고진행현황 파일 업로드")
out_up = st.file_uploader(
    "출고진행현황 xlsx", type=["xlsx"], key="out_up",
    help="업로드 즉시 자동 매칭(FEFO)이 실행됩니다.")
if not out_up:
    st.info("출고진행현황을 업로드하면 자동 매칭됩니다.")
    st.stop()

out_path = TMP / out_up.name
out_path.write_bytes(out_up.getvalue())
targets = cloud_store.load_targets()

with st.spinner("매칭 중..."):
    try:
        _tsig = tuple((t["code"], t.get("name", "")) for t in targets)
        analysis = copy.deepcopy(cached_analyze(
            inv_up.getvalue(), out_up.getvalue(), _tsig, _master_mtime()))
    except Exception as e:
        st.error(f"매칭 오류: {e}")
        st.stop()

st_st = analysis["stats"]
m1, m2, m3, m4 = st.columns(4)
m1.metric("기능1 청크", st_st["f1_rows"])
m2.metric("기능2 청크", st_st["f2_rows"])
m3.metric("기능1 부족 품목", st_st["f1_shortage_items"])
m4.metric("기능2 부족 품목", st_st["f2_shortage_items"])


# ============= Step 3: 거래처 수정 =============
st.divider()
st.subheader("③ 매칭 결과 — 거래처 수정")
st.caption("거래처 칸을 클릭해서 다른 쿠팡 센터로 바꿀 수 있습니다. "
           "비우면 '잔여(미배정)' 처리.")

def rec_to_row(r, fn):
    row = {
        "rowid": r["rowid"],
        "품목코드": int(r["code"]),
        "품명": r["name"],
        "유통기한": r["exp"],
        "lot박스": float(r["lot_box"]),
    }
    if fn == 1:
        row["고정로케이션"] = r["fixloc"]
        row["팔레트"] = r.get("pallets", "")
    else:
        row["로케이션"] = r["locs_str"]
    row["하대"] = str(r.get("baemyeon", ""))
    row["거래처"] = r["cust"] or ""
    row["배정박스"] = float(r["qty"])
    row["상태"] = r["status"]
    return row


def render_tab(records, fn, key_prefix):
    """품목별 미니 data_editor — 드롭다운 옵션은 그 품목 주문 거래처로 한정.
    반환: {rowid: 선택거래처}
    """
    # 품목별 그룹
    by_code = {}
    for r in records:
        by_code.setdefault(r["code"], []).append(r)

    sel = {}
    sections = []

    for code in sorted(by_code):
        group = by_code[code]
        first = group[0]
        cust_orders = first["cust_orders"]

        # 모든 청크 sel 초기 세팅 (자동배정 cust 또는 빈문자열)
        for r in group:
            sel[r["rowid"]] = r["cust"] or ""

        if not cust_orders:           # 쿠팡 주문 없음 → 화면 X
            continue

        # 기능2: 선입 lot 단독 처리 가능 시 화면 X (export 와 일치)
        if fn == 2:
            exps_lots = [(r["exp"], r["lot_box"])
                         for r in group if r["exp"] != "-"]
            if exps_lots:
                earliest = min(e for e, _ in exps_lots)
                earliest_lot = next(b for e, b in exps_lots if e == earliest)
                if earliest_lot >= sum(cust_orders.values()):
                    continue

        matched = [r for r in group if r["cust"]]
        if not matched:
            continue
        sections.append((code, first, cust_orders, matched, len(group)))

    if not sections:
        st.info("매칭된 결과 없음")
        return sel

    for code, first, cust_orders, matched, total_lots in sections:
        opts = [""] + sorted(cust_orders.keys())
        header = (f"**{code} · {first['name']}** — "
                  f"매칭 {len(matched)}/{total_lots} lot, "
                  f"주문처 {len(cust_orders)}곳")
        with st.expander(header, expanded=True):
            df = pd.DataFrame([rec_to_row(r, fn) for r in matched])
            cfg = {
                "rowid": None,
                "품목코드": st.column_config.NumberColumn(format="%d", disabled=True),
                "품명": st.column_config.TextColumn(disabled=True),
                "유통기한": st.column_config.TextColumn(disabled=True),
                "lot박스": st.column_config.NumberColumn(format="%g", disabled=True),
                "하대": st.column_config.TextColumn(disabled=True),
                "배정박스": st.column_config.NumberColumn(format="%g", disabled=True),
                "상태": st.column_config.TextColumn(disabled=True),
                "거래처": st.column_config.SelectboxColumn(
                    options=opts, required=False,
                    help=f"이 품목 주문 {len(cust_orders)}곳"),
            }
            if fn == 1:
                cfg["고정로케이션"] = st.column_config.TextColumn(disabled=True)
                cfg["팔레트"] = st.column_config.TextColumn(disabled=True)
            else:
                cfg["로케이션"] = st.column_config.TextColumn(disabled=True)
            edited = st.data_editor(
                df, use_container_width=True, hide_index=True,
                column_config=cfg, key=f"{key_prefix}_ed_{code}")
            for _, row in edited.iterrows():
                sel[row["rowid"]] = (row["거래처"] or "")
    return sel


tab1, tab2 = st.tabs(
    [f"기능1 · OV5 락재고 ({len(analysis['f1'])}청크)",
     f"기능2 · 출고가능 재고 ({len(analysis['f2'])}청크)"])
with tab1:
    sel1 = render_tab(analysis["f1"], 1, "f1")
with tab2:
    sel2 = render_tab(analysis["f2"], 2, "f2")


# ============= Step 4: 다운로드 =============
st.divider()
st.subheader("④ 결과 다운로드")
short = analysis["date_short"]
fname = f"{short} 쿠팡 지정출고_자동.xlsx"

if st.button(f"📥 {fname} 생성 + 다운로드", type="primary",
             use_container_width=True):
    out_xlsx = TMP / fname
    try:
        core.export(analysis, sel1, sel2, str(out_xlsx))
    except Exception as e:
        st.error(f"저장 오류: {e}")
        st.stop()
    with open(out_xlsx, "rb") as f:
        data = f.read()
    st.success(f"생성 완료: {fname} ({len(data)/1024:.1f} KB)")
    st.download_button(
        "💾 브라우저로 다운로드",
        data, fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
