# -*- coding: utf-8 -*-
"""ATN/BNF 창고비교 — Streamlit 웹 (통합앱 페이지).

재고조회 업로드 → IC930(통합) vs 타 창고 유통기한 비교 → 정상/비정상 판정.
지정출고 유통기한을 붙여넣으면 '정상(지정출고)'로 재판정.
"""
from __future__ import annotations

import io
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import compare_core as cc  # noqa: E402

DEFAULT_MASTER = HERE / "default_master.xlsx"

try:
    st.set_page_config(page_title="창고비교 (ATN/BNF)", layout="wide")
except Exception:
    pass

st.title("🏬 ATN/BNF 창고비교")
st.caption("재고조회 업로드 → IC930(통합) vs 타 창고 유통기한 비교 → 정상/비정상 판정")


# ---------- 캐시 ----------
@st.cache_data(show_spinner=False)
def _pivot(inv_bytes: bytes, whs: tuple, min_wh: int) -> pd.DataFrame:
    p = HERE / "_inv_cache.xlsx"
    p.write_bytes(inv_bytes)
    df = cc.load_inventory(str(p))
    return cc.build_pivot(df, list(whs), min_wh)


@st.cache_data(show_spinner=False)
def _master(master_bytes: bytes | None) -> dict:
    if master_bytes:
        p = HERE / "_master_cache.xlsx"
        p.write_bytes(master_bytes)
        try:
            return cc.load_master(str(p))
        except Exception:
            return {}
    if DEFAULT_MASTER.exists():
        try:
            return cc.load_master(str(DEFAULT_MASTER))
        except Exception:
            return {}
    return {}


# ---------- 사이드바: 설정/마스터 (접기) ----------
with st.sidebar.expander("⚙️ 설정 / 물품정보 — 클릭해서 열기", expanded=False):
    wh_text = st.text_input("비교 창고 (쉼표, 첫 번째=통합기준 IC930)",
                            value=",".join(cc.DEFAULT_WAREHOUSES))
    warehouses = [w.strip() for w in wh_text.split(",") if w.strip()] or cc.DEFAULT_WAREHOUSES
    min_wh = st.number_input("최소 비교 창고 수", 1, len(warehouses),
                             value=min(cc.DEFAULT_MIN_WAREHOUSES, len(warehouses)))
    st.divider()
    st.caption("물품정보(소비기한월) — 잔존% 계산용. 미업로드 시 기본 마스터 사용.")
    up_master = st.file_uploader("물품정보 xlsx 업로드 (Item_*.xlsx)", type=["xlsx"], key="wh_master")
    master = _master(up_master.getvalue() if up_master else None)
    st.success(f"물품정보: {len(master)}품목") if master else st.warning("물품정보 없음 (잔존% 생략)")
    # 양식 다운로드
    tmpl = pd.DataFrame({"Item code": [1010422, 2032260], "소비기한(월)": [24, 12]})
    buf = io.BytesIO(); tmpl.to_excel(buf, index=False)
    st.download_button("📥 물품정보 양식", buf.getvalue(), "물품정보_양식.xlsx",
                       use_container_width=True)

st.info("판정 기준 · IC930(통합)이 타 창고보다 빠르면 **정상** · 타 창고가 더 빠른데 "
        "지정출고 내역 있으면 **정상(지정출고)** · 없으면 **비정상(점검)**", icon="📐")

# ---------- 입력 ----------
c1, c2 = st.columns([1, 1])
with c1:
    inv_up = st.file_uploader("① 로케이션별 재고조회 xlsx", type=["xlsx", "xls"], key="wh_inv",
                              help="ⓘ ORG IC 창고는 빈칸으로 조회")
with c2:
    pasted = st.text_area("② (선택) 지정출고 유통기한 붙여넣기 — 제품코드·유통기한 열 포함",
                          height=120, placeholder="엑셀에서 헤더 포함 복사 → 붙여넣기")

if not inv_up:
    st.info("재고조회 파일을 업로드하세요.")
    st.stop()

try:
    pivot = _pivot(inv_up.getvalue(), tuple(warehouses), int(min_wh)).copy()
except Exception as e:
    st.error(f"비교 오류: {e}")
    st.stop()

extra = {}
if pasted.strip():
    try:
        extra = cc.parse_pasted_data(pasted)
        st.success(f"지정출고 {len(extra)}품목 인식 — 판정에 반영")
    except Exception as e:
        st.warning(f"붙여넣기 파싱 실패: {e}")

# ---------- 판정 + 잔존% ----------
today = date.today()
pivot[cc.VERDICT_COL] = pivot.apply(lambda r: cc.verdict_row(r, warehouses, extra), axis=1)
pivot["_ratio"] = pivot.apply(lambda r: cc.row_min_ratio(r, warehouses, master, today), axis=1)
if master:
    pivot[cc.REMAINING_COL_NAME] = pivot["_ratio"].apply(
        lambda r: "" if r is None else (f"{int(round(r*100))}%" if r >= 0 else "만료"))
if extra:
    pivot[cc.EXTRA_COL_NAME] = pivot["품목코드"].map(lambda c: extra.get(cc._normalize_pivot_code(c), ""))

# ---------- 필터 ----------
fc1, fc2 = st.columns([1, 3])
with fc1:
    abnormal_only = st.toggle("⚠ 비정상만 보기", value=False)
with fc2:
    wh_filter = st.multiselect("창고 필터 (선택 창고에 재고 있는 품목만)", warehouses, default=[])

view = pivot
if abnormal_only:
    view = view[view[cc.VERDICT_COL] == "비정상"]
for wh in wh_filter:
    view = view[view[wh].astype(str).str.strip() != ""]

total = len(pivot)
abn = int((pivot[cc.VERDICT_COL] == "비정상").sum())
m1, m2, m3 = st.columns(3)
m1.metric("전체 품목", f"{total:,}")
m2.metric("표시", f"{len(view):,}")
m3.metric("비정상", f"{abn:,}")

# ---------- 표시 (색상) ----------
base = ["품목코드", "품목명"] + [w for w in warehouses if w in view.columns]
tail = [cc.VERDICT_COL]
if cc.REMAINING_COL_NAME in view.columns:
    tail.append(cc.REMAINING_COL_NAME)
if cc.EXTRA_COL_NAME in view.columns:
    tail.append(cc.EXTRA_COL_NAME)
disp = view[base + tail + ["_ratio"]].reset_index(drop=True)


def _style(row):
    v = row[cc.VERDICT_COL]
    bg = ""
    if v == "비정상":
        bg = "#ffd6d6"
    elif v == "정상(지정출고)":
        bg = "#fff7d6"
    else:
        r = row["_ratio"]
        if r is not None and r < cc.IMMINENT_THRESHOLD:
            bg = "#ffcc99"
    return [f"background-color: {bg}"] * len(row)


styler = disp.style.apply(_style, axis=1).hide(["_ratio"], axis="columns")
st.dataframe(styler, use_container_width=True, hide_index=True, height=560)

st.caption("🟥 비정상   🟨 정상(지정출고)   🟧 유통기한 임박(잔존<50%)")
