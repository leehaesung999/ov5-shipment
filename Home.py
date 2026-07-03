# -*- coding: utf-8 -*-
"""통합 — 한 주소에서 여러 사내 도구 전환 (멀티페이지).

배포: Streamlit Cloud Main file path = Home.py
사이드바 상단의 페이지 메뉴로 앱들을 오갑니다.
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
export_label = st.Page(
    "export_label/app.py", title="수출 라벨 인쇄", icon="🖨️",
    url_path="export",
)
box_request = st.Page(
    "box_request/app.py", title="외박스 요청 리스트", icon="🗃️",
    url_path="box",
)

# 기본 사이드바 네비게이션은 숨기고(아래에서 커스텀 구성) 외부 링크를 최상단에 배치
nav = st.navigation(
    [ov5, coupang, jaego, warehouse, inventory, lock, export_label, box_request],
    position="hidden",
)

# ── 사이드바 최상단: 외부 앱(PythonAnywhere) 바로가기 ──
st.sidebar.link_button("🔗 일일재고 (PythonAnywhere)",
                       "https://goal.pythonanywhere.com/",
                       use_container_width=True)
st.sidebar.link_button("🔗 입고계획 (PythonAnywhere)",
                       "https://goalkii.pythonanywhere.com/",
                       use_container_width=True)
st.sidebar.divider()

# ── 페이지 메뉴 (기본 네비게이션 대체) ──
for _p in (ov5, coupang, jaego, warehouse, inventory, lock, export_label, box_request):
    st.sidebar.page_link(_p, use_container_width=True)
st.sidebar.divider()

nav.run()
