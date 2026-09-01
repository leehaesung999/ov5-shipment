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
# --- KST (Streamlit Cloud는 UTC 기본) ---
try:
    KST  # noqa: F821
except NameError:
    from datetime import datetime as _dt, timedelta as _td, timezone as _tz
    KST = _tz(_td(hours=9))
    def _now_kst():
        return _dt.now(KST)
    def _now_kst_naive():
        return _dt.now(KST).replace(tzinfo=None)
    def _today_kst():
        return _dt.now(KST).date()

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

st.title("🔒 LOCK 실사지 출력")
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from page_help import show_help  # noqa: E402
show_help({
    "목적": "Lock 재고 원본 엑셀(조회1_xxx.xlsx)을 실사용 락.xlsx 로 자동 변환.",
    "필요한 파일": "조회1_xxx.xlsx (Lock 재고 원본)",
    "사용 순서": "1. 원본 xlsx 업로드\n"
                 "2. 자동 처리: 빈 출고진행 열 삭제 + 필터(Location '-' 없음 & 사유≠불가_실사차이) + 정렬\n"
                 "3. 락.xlsx 다운로드",
})
st.caption("조회1_xxx.xlsx (Lock 재고 원본) → 락.xlsx (필터·정렬 + 피벗)")
st.info("변환 규칙 · **sheet**: 빈 출고진행열 삭제 + 필터(Location '-' 없음 & 사유≠불가_실사차이) + 정렬 "
        "· **Sheet1**: (Location·ItemID·유통기한) 재고수량 합계 피벗 + 총합계", icon="📐")

up = st.file_uploader("Lock 재고 원본 xlsx 업로드", type=["xlsx"], key="lock_up")
if not up:
    st.info("Lock 재고 원본 파일을 업로드하세요.")
    st.stop()

src = TMP / "_src.xlsx"
src.write_bytes(up.getvalue())
dst = TMP / f"락_{_today_kst():%Y%m%d}.xlsx"

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
                   mime=MIME, width='stretch')
