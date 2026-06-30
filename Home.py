# -*- coding: utf-8 -*-
"""지정출고 통합 — 한 주소에서 OV5/농협 · 쿠팡 전환 (멀티페이지).

배포: Streamlit Cloud Main file path = Home.py
사이드바 상단의 페이지 메뉴로 두 앱을 오갑니다.
"""
import streamlit as st

st.set_page_config(page_title="지정출고 통합", layout="wide")

ov5 = st.Page(
    "app_streamlit.py",
    title="OV5 / 농협 지정출고",
    icon="🟩",
    default=True,
)
coupang = st.Page(
    "coupang/streamlit_app.py",
    title="쿠팡 지정출고",
    icon="🟦",
)

nav = st.navigation([ov5, coupang])
nav.run()
