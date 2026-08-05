# -*- coding: utf-8 -*-
"""통합 — 한 주소에서 여러 사내 도구 전환 (멀티페이지).

배포: Streamlit Cloud Main file path = Home.py
사이드바 상단의 페이지 메뉴로 앱들을 오갑니다.
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
    "jaego/app.py", title="리뉴얼 재고 출고 시점 점검", icon="📦",
    url_path="jaego",
)
warehouse = st.Page(
    "warehouse/app.py", title="창고별 재고 유통기한 점검", icon="🏬",
    url_path="warehouse",
)
sheets = st.Page(
    "inventory/sheets.py", title="실사지 출력", icon="📋",
    url_path="sheets",
)
checks = st.Page(
    "inventory/checks.py", title="일 점검(이중적치, 유통기한)", icon="🔎",
    url_path="checks",
)
lock = st.Page(
    "lock/app.py", title="LOCK 실사지 출력", icon="🔒",
    url_path="lock",
)
export_label = st.Page(
    "export_label/app.py", title="수출 표식지 출력", icon="🖨️",
    url_path="export",
)
paint = st.Page(
    "paint/app.py", title="통합 기본로케이션 레이아웃", icon="🎨",
    url_path="paint",
)
abc = st.Page(
    "abc_analysis/app.py", title="ABC분석 로케이션 변경안", icon="📊",
    url_path="abc",
)
pallet = st.Page(
    "pallet/app.py", title="BNF(비네이버) 피킹지 표식지", icon="🧱",
    url_path="pallet",
)
janjon = st.Page(
    "janjon/app.py", title="거래처별 잔존율", icon="📅",
    url_path="janjon",
)
safety_stock = st.Page(
    "safety_stock/app.py", title="BNF(비네이버) 안전재고 산출", icon="📐",
    url_path="safety",
)
transfer = st.Page(
    "transfer/app.py", title="BNF(비네이버) 재고이동 계획", icon="🚚",
    url_path="transfer",
)

# 기본 사이드바 네비게이션은 숨기고(아래에서 커스텀 구성) 외부 링크를 최상단에 배치
nav = st.navigation(
    [ov5, coupang, checks, jaego, warehouse, sheets, lock, export_label, safety_stock, transfer, pallet, paint, abc, janjon],
    position="hidden",
)

# ── 사이드바 최상단: 외부 앱(PythonAnywhere) 바로가기 ──
st.sidebar.link_button("🔗 일일재고 (PythonAnywhere)",
                       "https://goal.pythonanywhere.com/",
                       width='stretch')
st.sidebar.link_button("🔗 입고계획 (PythonAnywhere)",
                       "https://goalkii.pythonanywhere.com/",
                       width='stretch')
st.sidebar.divider()

# ── 페이지 메뉴 (기본 네비게이션 대체) ──
for _p in (ov5, coupang, checks, jaego, warehouse, sheets, lock, export_label, safety_stock, transfer, pallet, paint, abc, janjon):
    st.sidebar.page_link(_p, width='stretch')
st.sidebar.divider()

nav.run()
