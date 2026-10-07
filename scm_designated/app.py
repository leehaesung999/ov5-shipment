# -*- coding: utf-8 -*-
"""SCM 지정출고 주간작업 — SCM팀 주간 '지정출고 요약'을 실행용 작업지시로 변환.

흐름:
  1) '지정출고 요약' xlsx 업로드 → ①확인 / ②입고요청 / ③LOCK / ④해제 4그룹 분리
  2) (선택) '로케이션별 재고조회' xlsx 업로드 → ② 입고요청 품목에 재고(로케이션·LOT) 자동 조인
  3) 도달일자 기준 우선순위 + '이번 주 할 일' 하이라이트
  4) 완료 체크 + 작업지시 엑셀 다운로드

순수 로직은 logic.py (테스트 가능). 여기는 UI 전용.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import streamlit as st

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent))
import logic  # noqa: E402

try:
    st.set_page_config(page_title="SCM 지정출고 주간작업", layout="wide")
except Exception:
    pass

try:
    from page_help import show_help  # noqa: E402
    show_help({
        "목적": "SCM팀 주간 '지정출고 요약'을 물류팀이 바로 실행할 수 있는 작업지시서로 변환.",
        "필요한 파일": "① 지정출고 요약 xlsx (필수)   ② 로케이션별 재고조회 xlsx (선택 — ②입고요청 재고조인용)",
        "사용 순서": "1. 지정출고 요약 업로드 → 4개 할 일(①확인·②입고요청·③LOCK·④해제)로 자동 분리\n"
                     "2. 재고조회 업로드 → ②입고요청 품목에 '어디서 몇 박스 집어 보낼지' 자동 조인\n"
                     "3. 도달일자 기준 '이번 주 할 일'부터 처리 (②입고요청은 리드타임 때문에 최우선)\n"
                     "4. 완료 체크 후 작업지시 엑셀 다운로드",
        "참고": "②재고 매칭은 '요약서 제조일 LOT'만 집계. 해당 LOT 재고가 없으면 ⚠️ 표시.",
    })
except Exception:
    pass

st.title("🟨 SCM 지정출고 주간작업")
st.caption("요약서 업로드 → 4개 할 일 분리 → 재고 자동조인 → 우선순위 → 작업지시 다운로드")

c1, c2 = st.columns(2)
with c1:
    up_sum = st.file_uploader("① 지정출고 요약 (필수)", type=["xlsx"], key="scm_sum")
with c2:
    up_inv = st.file_uploader("② 로케이션별 재고조회 (선택)", type=["xlsx"], key="scm_inv")

if not up_sum:
    st.info("먼저 '지정출고 요약' 파일을 올려주세요.")
    st.stop()

try:
    summary = logic.parse_summary(up_sum.getvalue())
except Exception as e:
    st.error(f"요약서 파싱 실패: {e}")
    st.stop()

if summary.empty:
    st.warning("요약서에서 품목 행을 찾지 못했습니다.")
    st.stop()

inv = None
if up_inv:
    try:
        inv = logic.parse_inventory(up_inv.getvalue())
        whs = ", ".join(sorted(inv["창고"].dropna().unique()))
        st.caption(f"재고 파일 로드: {len(inv):,}행 · 창고: {whs}")
    except Exception as e:
        st.error(f"재고 파일 파싱 실패: {e}")

g1 = summary[summary["구분"].str.startswith("①")].sort_values("도달일자", na_position="last")
g2 = summary[summary["구분"].str.startswith("②")].sort_values("도달일자", na_position="last")
g3 = summary[summary["구분"].str.startswith("③")].sort_values("도달일자", na_position="last")
g4 = summary[summary["구분"].str.startswith("④")].sort_values("도달일자", na_position="last")

today = logic._today_kst()
this_week = summary[summary["도달일자"].apply(
    lambda d: d is not None and 0 <= (d - today).days <= 7)]
overdue = summary[summary["도달일자"].apply(
    lambda d: d is not None and (d - today).days < 0)]

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("①쿠팡확인", len(g1))
m2.metric("②입고요청", len(g2))
m3.metric("③LOCK", len(g3))
m4.metric("④해제", len(g4))
m5.metric("🟠이번주 / 🔴지남", f"{len(this_week)} / {len(overdue)}")

if up_inv is None and len(g2):
    st.warning("②입고요청 품목의 '어디서 몇 박스 집을지'를 보려면 **로케이션별 재고조회** 파일도 올려주세요.")

st.divider()

view2 = logic.join_move_stock(g2, inv) if len(g2) else pd.DataFrame()
view1 = logic.simple_view(g1) if len(g1) else pd.DataFrame()
view3 = logic.simple_view(g3) if len(g3) else pd.DataFrame()
view4 = logic.simple_view(g4) if len(g4) else pd.DataFrame()


def _editor(df, key):
    return st.data_editor(
        df, hide_index=True, width="stretch",
        disabled=[c for c in df.columns if c != "완료"], key=key,
    )


t2, t3, t4, t1 = st.tabs([
    f"② 입고요청(물품이동) · {len(g2)}",
    f"③ LOCK 설정 · {len(g3)}",
    f"④ LOCK 해제 · {len(g4)}",
    f"① 쿠팡 확인 · {len(g1)}",
])
with t2:
    st.markdown("**가장 급함** — 3PL 입고 리드타임(2~3일) 때문에 주초에 먼저 처리하세요.")
    _editor(view2, "ed2") if not view2.empty else st.info("②입고요청 항목이 없습니다.")
with t3:
    st.markdown("**WMS에서 지정출고 LOCK 설정** — 도달일자 임박 순.")
    _editor(view3, "ed3") if not view3.empty else st.info("③LOCK 항목이 없습니다.")
with t4:
    st.markdown("**WMS에서 지정출고 LOCK 해제**.")
    _editor(view4, "ed4") if not view4.empty else st.info("④해제 항목이 없습니다.")
with t1:
    st.markdown("**쿠팡 담당자에게 회신 요청 메일** 발송 대상.")
    _editor(view1, "ed1") if not view1.empty else st.info("①쿠팡 확인 항목이 없습니다.")

st.divider()
xls = logic.build_xlsx({
    "②입고요청(물품이동)": view2,
    "③LOCK설정": view3,
    "④LOCK해제": view4,
    "①쿠팡확인": view1,
})
st.download_button(
    "⬇️ 작업지시서 엑셀 다운로드",
    data=xls,
    file_name=f"지정출고_작업지시_{today.strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
