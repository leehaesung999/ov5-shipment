# -*- coding: utf-8 -*-
"""LOCK 재고 엑셀 변환기 — Streamlit 페이지 (통합앱).

조회1_xxx.xlsx (Lock 재고 원본) → 락.xlsx
  · sheet: 빈 출고진행열 삭제 + 필터(Location '-' 없음 & 사유≠불가_실사차이) + 정렬
  · Sheet1: (Location·ItemID·유통기한) 재고수량 합계 피벗 + 총합계
"""
from __future__ import annotations

import contextlib
import io
import sys
from datetime import date
from pathlib import Path

import streamlit as st

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import convert  # noqa: E402

TMP = HERE / "_tmp"
TMP.mkdir(exist_ok=True)
MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

try:
    st.set_page_config(page_title="LOCK 변환기", layout="wide")
except Exception:
    pass

st.title("🔒 LOCK 재고 변환기")
st.caption("조회1_xxx.xlsx (Lock 재고 원본) → 락.xlsx (필터·정렬 + 피벗)")
st.info("변환 규칙 · **sheet**: 빈 출고진행열 삭제 + 필터(Location '-' 없음 & 사유≠불가_실사차이) + 정렬 "
        "· **Sheet1**: (Location·ItemID·유통기한) 재고수량 합계 피벗 + 총합계", icon="📐")

up = st.file_uploader("Lock 재고 원본 xlsx 업로드", type=["xlsx"], key="lock_up")
if not up:
    st.info("Lock 재고 원본 파일을 업로드하세요.")
    st.stop()

src = TMP / "_src.xlsx"
src.write_bytes(up.getvalue())
dst = TMP / f"락_{date.today():%Y%m%d}.xlsx"

buf = io.StringIO()
try:
    with contextlib.redirect_stdout(buf):
        convert.transform(src, dst)
except Exception as e:
    st.error(f"변환 오류: {e}")
    st.stop()

if not dst.exists():
    st.error("변환 결과 파일이 생성되지 않았습니다.")
    st.stop()

st.success("✅ 변환 완료")
_stats = buf.getvalue().strip()
if _stats:
    st.code(_stats)
st.download_button(f"📥 {dst.name} 다운로드", dst.read_bytes(), dst.name,
                   mime=MIME, use_container_width=True)
