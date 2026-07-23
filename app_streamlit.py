# -*- coding: utf-8 -*-
"""통합 진입점 (기존 배포 앱의 Main file = app_streamlit.py 호환).

한 주소에서 OV5/쿠팡/재고모니터/창고비교/재고분석기 전환 + 외부앱 링크.
Home.py 와 동일한 멀티페이지 네비게이션.
"""
import sys
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))
import brand  # noqa: E402

st.set_page_config(page_title="샘표 물류 업무 시스템", layout="wide")
brand.apply()

ov5 = st.Page(
    "ov5_page.py", title="농협/대리점 지정출고", icon="🟩",
    url_path="ov5", default=True,
)
coupang = st.Page(
    "coupang/streamlit_app.py", title="쿠팡 지정출고", icon="🟦",
    url_path="coupang",
)
jaego = st.Page(
    "jaego/app.py", title="신규파우치 점검 프로그램", icon="📦",
    url_path="jaego",
)
warehouse = st.Page(
    "warehouse/app.py", title="분산재고 점검 프로그램", icon="🏬",
    url_path="warehouse",
)
sheets = st.Page(
    "inventory/sheets.py", title="실사지 출력", icon="📋",
    url_path="sheets",
)
checks = st.Page(
    "inventory/checks.py", title="점검 (이중적치·하프도달)", icon="🔎",
    url_path="checks",
)
lock = st.Page(
    "lock/app.py", title="LOCK 실사 프로그램", icon="🔒",
    url_path="lock",
)
export_label = st.Page(
    "export_label/app.py", title="수출 표식지", icon="🖨️",
    url_path="export",
)
paint = st.Page(
    "paint/app.py", title="창고 레이아웃 색칠", icon="🎨",
    url_path="paint",
)
abc = st.Page(
    "abc_analysis/app.py", title="ABC 재배치 분석", icon="📊",
    url_path="abc",
)

# 기본 사이드바 네비게이션은 숨기고(아래에서 커스텀 구성) 외부 링크를 최상단에 배치
nav = st.navigation(
    [ov5, coupang, jaego, warehouse, sheets, checks, lock, export_label, paint, abc],
    position="hidden",
)

# ── 사이드바 최상단: 외부 앱(PythonAnywhere) 바로가기 ──
st.sidebar.link_button("🔗 일일재고 (PythonAnywhere)",
                       "https://goal.pythonanywhere.com/",
                       width='stretch')
st.sidebar.link_button("🔗 입고계획 (PythonAnywhere)",
                       "https://goalkii.pythonanywhere.com/",
                       width='stretch')
st.sidebar.link_button("🔗 출하파트",
                       "https://sampyo-shipment-hkdmdboz3qwiysgcozwchd.streamlit.app/",
                       width='stretch')
st.sidebar.divider()

# ── 페이지 메뉴 (기본 네비게이션 대체) ──
for _p in (ov5, coupang, jaego, warehouse, sheets, checks, lock, export_label, paint, abc):
    st.sidebar.page_link(_p, width='stretch')
st.sidebar.divider()

nav.run()
