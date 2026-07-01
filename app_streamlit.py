# -*- coding: utf-8 -*-
"""통합 진입점 (기존 배포 앱의 Main file = app_streamlit.py 호환).

한 주소에서 OV5/쿠팡/재고모니터/창고비교/재고분석기 전환 + 외부앱 링크.
Home.py 와 동일한 멀티페이지 네비게이션.
"""
import streamlit as st

st.set_page_config(page_title="사내 도구 통합", layout="wide")

ov5 = st.Page(
    "ov5_page.py", title="OV5 / 농협 지정출고", icon="🟩",
    url_path="ov5", default=True,
)
coupang = st.Page(
    "coupang/streamlit_app.py", title="쿠팡 지정출고", icon="🟦",
    url_path="coupang",
)
jaego = st.Page(
    "jaego/app.py", title="유통기한 재고 모니터", icon="📦",
    url_path="jaego",
)
warehouse = st.Page(
    "warehouse/app.py", title="창고비교 (ATN/BNF)", icon="🏬",
    url_path="warehouse",
)
inventory = st.Page(
    "inventory/app.py", title="통합센터 재고 분석기", icon="🏭",
    url_path="inventory",
)
lock = st.Page(
    "lock/app.py", title="LOCK 변환기", icon="🔒",
    url_path="lock",
)

nav = st.navigation([ov5, coupang, jaego, warehouse, inventory, lock])

st.sidebar.divider()
st.sidebar.link_button("🔗 일일재고 (PythonAnywhere)",
                       "https://goal.pythonanywhere.com/",
                       use_container_width=True)

nav.run()
