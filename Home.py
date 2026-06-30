# -*- coding: utf-8 -*-
"""통합 — 한 주소에서 여러 사내 도구 전환 (멀티페이지).

배포: Streamlit Cloud Main file path = Home.py
사이드바 상단의 페이지 메뉴로 앱들을 오갑니다.
"""
import streamlit as st

st.set_page_config(page_title="사내 도구 통합", layout="wide")

ov5 = st.Page(
    "app_streamlit.py", title="OV5 / 농협 지정출고", icon="🟩",
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

nav = st.navigation([ov5, coupang, jaego, warehouse, inventory])

# 외부 앱(일일재고 — PythonAnywhere)으로 이동 링크
st.sidebar.divider()
st.sidebar.link_button("🔗 일일재고 (PythonAnywhere)",
                       "https://goal.pythonanywhere.com/",
                       use_container_width=True)

nav.run()
