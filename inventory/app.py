# -*- coding: utf-8 -*-
"""통합센터 재고 분석기 — Streamlit 웹 (통합앱 페이지).

ERP 재고조회 업로드 → 이중적치 분석 / 재고지 생성 / 유통기한(OV5·OV6) 분석 → 다운로드.
기준정보(물품 마스터)는 레포에 포함된 data/기준정보.xlsx 사용(또는 업로드).
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import streamlit as st

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import core  # noqa: E402

DATA = HERE / "data"
DEFAULT_MASTER = DATA / "기준정보.xlsx"
TMP = HERE / "_tmp"
TMP.mkdir(exist_ok=True)

try:
    st.set_page_config(page_title="통합센터 재고 분석기", layout="wide")
except Exception:
    pass

st.title("🏭 통합센터 재고 분석기")
st.caption("ERP 재고조회 업로드 → 이중적치 / 재고지 / 유통기한 분석 → 엑셀 다운로드")


def _save(upload, name) -> str:
    p = TMP / name
    p.write_bytes(upload.getvalue())
    return str(p)


# ---------- 사이드바: 기준정보 (접기) ----------
with st.sidebar.expander("⚙️ 기준정보 / 옵션 — 클릭해서 열기", expanded=False):
    if DEFAULT_MASTER.exists():
        st.success("기준정보: 내장본 사용 중")
    else:
        st.warning("내장 기준정보 없음 — 업로드 필요")
    up_master = st.file_uploader("기준정보 xlsx 업로드(갱신)", type=["xlsx"], key="inv_master")
    master_path = _save(up_master, "_master.xlsx") if up_master else (
        str(DEFAULT_MASTER) if DEFAULT_MASTER.exists() else None)
    threshold = st.slider("유통기한 분석 잔존율 기준", 0.0, 1.0, 0.5, 0.05)
    today = st.date_input("기준일자", value=date.today())

# ---------- 입력 ----------
st.subheader("① ERP 재고조회 파일 업로드")
stock_up = st.file_uploader("로케이션별 재고조회 xlsx", type=["xlsx", "xls"], key="inv_stock")
if not stock_up:
    st.info("재고조회 파일을 업로드하세요.")
    st.stop()
stock_path = _save(stock_up, "_stock.xlsx")
if not master_path:
    st.error("기준정보가 없습니다. 사이드바에서 업로드하세요.")
    st.stop()

st.subheader("② 분석 선택 → 실행")
ACTIONS = {
    "이중적치 분석 (보충오류 점검)": ("analyze_and_save", "이중적치"),
    "재고지 — 1단 전체": ("edit_재고지_1단_전체", "재고지_1단전체"),
    "재고지 — 2~6단": ("edit_재고지_2_6단", "재고지_2_6단"),
    "유통기한 분석 — OV5 (잔존율 ≤ 기준)": ("analyze_ov5_expiry", "OV5유통기한"),
    "유통기한 분석 — OV6 (잔존율 ≤ 기준)": ("analyze_ov6_expiry", "OV6유통기한"),
}
choice = st.selectbox("분석 항목", list(ACTIONS.keys()))
fn_name, out_tag = ACTIONS[choice]

if st.button(f"▶ {choice} 실행", type="primary", use_container_width=True):
    logs: list[str] = []

    def log(*a):
        logs.append(" ".join(str(x) for x in a))

    out_path = TMP / f"{out_tag}_{date.today():%Y%m%d}.xlsx"
    fn = getattr(core, fn_name)
    with st.spinner("분석 중..."):
        try:
            if fn_name.endswith("_expiry"):
                fn(stock_path, master_path, str(out_path),
                   threshold=float(threshold), today=today, log=log)
            else:
                fn(stock_path, master_path, str(out_path), log=log)
        except Exception as e:
            st.error(f"오류: {e}")
            with st.expander("로그"):
                st.code("\n".join(logs) or str(e))
            st.stop()
    if logs:
        with st.expander("처리 로그", expanded=False):
            st.code("\n".join(logs))
    if out_path.exists():
        data = out_path.read_bytes()
        st.success(f"완료: {out_path.name} ({len(data)/1024:.0f} KB)")
        st.download_button("📥 결과 다운로드", data, out_path.name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    else:
        st.warning("결과 파일이 생성되지 않았습니다. 로그를 확인하세요.")

st.caption("ⓘ 담당자(공유여부 자동채움)는 개인정보라 웹엔 미포함 — 필요 시 사이드바 추가 없이 로컬판 사용")
