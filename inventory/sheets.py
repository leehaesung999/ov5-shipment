# -*- coding: utf-8 -*-
"""실사지 출력 — 일일·1단·토요일 실사지 + 2~6단 재고지 (4종)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent))
import inv_page  # noqa: E402
import streamlit as st  # noqa: E402
from page_help import show_help  # noqa: E402

try:
    st.set_page_config(page_title="통합센터 재고 분석기", layout="wide")
except Exception:
    pass
st.title("📋 실사지 출력")
st.caption("일일 · 토요일 · 1단 실사지 및 2~6단 재고지 생성 → 엑셀 다운로드")
show_help({
    "목적": "1단·2~6단 재고 실사지를 매일/주간/월간 용도로 자동 생성.",
    "필요한 파일": "① ERP 재고조회 xlsx (필수)  ② 일일입력 xlsx (차이수량 반영, 선택)  "
                   "③ 제품별리스트 xlsx (토요일 실사지용, 필수·컬리 선택)",
    "4가지 실사지": "**① 일일 재고실사지** — 지정로케이션 1단(-10), 매일 점검 (일일입력 있으면 차이수량·음영 반영)\n"
                    "**② 1단 재고실사지** — 마스터 등록 1단 전체 (빈 로케 포함, 월간 실사용)\n"
                    "**③ 토요일 실사지 (쿠팡+컬리)** — 두 리스트 합집합의 IC930 출하 품목만\n"
                    "**④ 재고지 (2~6단)** — 마스터 5개 단, 한 파일 5시트 (빈 로케 포함)",
    "사용 순서": "1. ① ERP 재고 파일 업로드\n"
                 "2. 원하는 실사지 카드로 이동 → 필요한 보조 파일 업로드\n"
                 "3. [▶] 버튼 클릭 → 결과 xlsx 자동 생성\n"
                 "4. [📥 다운로드] 클릭",
    "참고": "차이수량 있는 품목은 현재고 0이어도 자동 포함 (fixed_loc 매칭). 기준정보·지정로케이션은 사이드바에서 관리.",
}, expanded=False)
inv_page.render(
    ["일일실사", "토요일", "1단전체", "2_6단"],
    "📋 실사지 출력",
    "일일 · 토요일 · 1단 실사지 및 2~6단 재고지 생성 → 엑셀 다운로드",
    show_title=False,
)
