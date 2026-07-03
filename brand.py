# -*- coding: utf-8 -*-
"""샘표 브랜드 테마 — 통합앱 전 페이지 공통 적용.

Home.py / app_streamlit.py 에서 set_page_config 직후 brand.apply()를 호출하면
사이드바 로고 + 폰트(Noto Sans KR) + 샘표 레드 포인트 + 상단 브랜드 바가
모든 페이지에 일관되게 적용된다.
"""
from __future__ import annotations

from pathlib import Path

import streamlit as st

RED = "#B61C21"       # 샘표 시그니처 레드
RED_DARK = "#8E161A"
GOLD = "#B39A78"      # 웜 골드/베이지 보조색
INK = "#2B2B2B"
ERP_URL = "https://sws.sempio.com/#SWS"
LOGO = str(Path(__file__).parent / "assets" / "sempio_logo.png")

_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700;800&display=swap');
html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stSidebar"] {{
    font-family: 'Noto Sans KR','Malgun Gothic','맑은 고딕','Nanum Gothic',sans-serif;
}}
/* 링크 = 샘표 레드 */
.stMarkdown a, a {{ color: {RED}; }}
.stMarkdown a:hover {{ color: {RED_DARK}; }}
/* 상단 브랜드 바 */
.sempio-bar {{
    display:flex; align-items:baseline; gap:12px; flex-wrap:wrap;
    border-left: 7px solid {RED};
    background: linear-gradient(90deg,#FBF6F0 0%,#FFFFFF 65%);
    padding: 11px 18px; margin: 2px 0 16px 0; border-radius: 8px;
    box-shadow: 0 1px 3px rgba(0,0,0,.06);
}}
.sempio-bar .brand {{ font-size: 21px; font-weight:800; color:{RED}; letter-spacing:-.4px; }}
.sempio-bar .sys {{ font-size: 19px; font-weight:800; color:{INK}; letter-spacing:-.3px; }}
.sempio-bar .sub {{ font-size: 12.5px; color:#9A8C7C; font-weight:600; margin-left:auto; }}
/* 사이드바 링크 hover / 선택 */
[data-testid="stSidebar"] a:hover {{ color: {RED} !important; }}
/* 탭 선택 강조 */
.stTabs [data-baseweb="tab-highlight"] {{ background-color: {RED} !important; }}
.stTabs [aria-selected="true"] {{ color: {RED} !important; }}
/* 익스팬더 헤더 hover */
details summary:hover {{ color: {RED}; }}
/* 메트릭 라벨 웜톤 */
[data-testid="stMetricValue"] {{ color: {INK}; }}
</style>
"""


def apply(system_name: str = "물류 업무 시스템",
          subtitle: str = "Sempio Logistics Workspace"):
    """테마 CSS + 사이드바 로고 + 상단 브랜드 바 적용."""
    st.markdown(_CSS, unsafe_allow_html=True)
    try:
        st.logo(LOGO, size="large", link=ERP_URL)
    except Exception:
        pass
    st.markdown(
        f'<div class="sempio-bar">'
        f'<span class="brand">샘표</span>'
        f'<span class="sys">{system_name}</span>'
        f'<span class="sub">{subtitle}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )
