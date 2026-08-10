# -*- coding: utf-8 -*-
"""점검 — 이중적치 · OV5/OV6 하프도달 · 비Lock 유통기한 점검 (4종)."""
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
st.title("🔎 일 점검(이중적치, 유통기한)")
st.caption("이중적치 · OV5/OV6 하프도달 · 비Lock 유통기한 점검 → 화면 표시 + 엑셀 다운로드")
show_help({
    "목적": "재고 이상(이중적치)과 유통기한 임박(하프도달) 품목을 자동 검출.",
    "필요한 파일": "ERP 재고조회 xlsx",
    "4가지 점검": "**① 이중적치** — 같은 로케이션 잔량 혼재 (Lock·OV·정파렛트 제외)\n"
                  "**② OV5 하프도달** — OV5 로케이션 재고 중 잔존율 ≤ 50%\n"
                  "**③ OV6 하프도달** — OV6 로케이션 재고 중 잔존율 ≤ 50%\n"
                  "**④ 비Lock 유통기한** — Lock 무관 전체 재고 중 잔존율 ≤ 기준(50%)",
    "사용 순서": "1. ① ERP 재고 파일 업로드\n"
                 "2. 원하는 점검 카드 → [▶] 클릭\n"
                 "3. 결과가 화면 표로 표시 (엑셀 다운로드 병행)\n"
                 "4. 비Lock 점검: 각 행 [확인] 체크 → 다음 실행 시 자동 숨김 (누적 관리)",
    "판정 기준일": "하프도달(②③④) 판정은 **오늘이 아니라 실제 입고일(=다음 영업일)** 기준입니다. "
                  "오늘 점검한 재고는 익일(금요일이면 월요일) 거래처에 입고되므로, 그 날 시점의 잔존율로 봅니다. "
                  "결과의 **하프도달일** 열에 각 로트가 잔존율 50%가 되는 날짜가 표시됩니다.",
    "참고": "확인 체크는 Supabase에 누적 저장돼 재확인 방지. 담당자 매핑은 사이드바 관리.",
}, expanded=False)
inv_page.render(
    ["이중적치", "ov5", "ov6", "nonlock"],
    "",  # title은 위에서 이미 표시
    "",
    preview=True,
)
