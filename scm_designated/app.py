# -*- coding: utf-8 -*-
"""SCM 지정출고 주간작업 — 상시 일정판.

핵심: SCM팀이 주 1회 보내는 '지정출고 요약'을 올려 **누적**하고, 매일 들어와
오늘/이번주 할 일과 종료(재고조정) 일정을 **한 타임라인**에서 확인·완료체크한다.

- 주 1회: 요약 업로드 → 누적 병합(완료건은 재진행 안 함)
- 매일: 접속 즉시 저장된 전체 일정이 날짜순으로 표시 (업로드 없이도)
- 작업/종료가 한 일정에 섞여 그날 할 일이 모두 보임
- 완료 체크는 영속(Supabase) → 다음날·다음주에도 유지

순수 로직은 logic.py, 영속 저장은 store.py.
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
import store  # noqa: E402

try:
    st.set_page_config(page_title="SCM 지정출고 주간작업", layout="wide")
except Exception:
    pass

try:
    from page_help import show_help  # noqa: E402
    show_help({
        "목적": "SCM 주간 '지정출고 요약'을 누적해, 매일 오늘 할 일·종료(재고조정) 일정을 한 눈에.",
        "필요한 파일": "지정출고 요약 xlsx (주 1회 업로드). 재고조회 xlsx는 ②입고요청 재고조인용(선택).",
        "일정 규칙": "② 보관창고 입고요청 → 당일(즉시)\n"
                     "① 쿠팡 / ③ LOCK 설정 → 하프(50%)면 도달일자, 그 외는 도달−2영업일\n"
                     "④ 지정출고 해제 → 도달일자\n"
                     "종료일자 있는 건 → 그날 '지정출고 종료 → 재고조정' 으로 함께 표시",
        "사용법": "1. (주1회) 요약 업로드 → 누적 반영  2. (매일) 일정 확인 + 완료 체크  3. 필요시 작업지시 엑셀 다운로드",
        "참고": "완료 체크한 건은 다음 주 업로드 때도 완료 유지(중복 진행 방지).",
    })
except Exception:
    pass

st.title("🟨 SCM 지정출고 주간작업 — 상시 일정판")
st.caption(f"저장: {store.backend_name()}  ·  작업+종료 한 타임라인, 날짜순. 주1회 업데이트 / 매일 확인.")

# ── 누적 task 로드 (상시) ──
if "scm_tasks" not in st.session_state:
    st.session_state.scm_tasks = store.load_tasks()

today = logic._today_kst()

# ── 주간 업데이트 (요약 업로드 → 누적 병합) ──
with st.expander("📥 주간 업데이트 — SCM '지정출고 요약' 업로드 (주 1회)",
                 expanded=(not st.session_state.scm_tasks)):
    up_sum = st.file_uploader("지정출고 요약 xlsx", type=["xlsx"], key="scm_sum")
    cc1, cc2 = st.columns([1, 3])
    if up_sum:
        if cc1.button("➕ 누적 반영", type="primary", width="stretch"):
            try:
                summ = logic.parse_summary(up_sum.getvalue())
                if summ.empty:
                    st.warning("요약서에서 품목 행을 찾지 못했습니다.")
                else:
                    new = logic.summary_to_tasks(summ, today=today)
                    merged, stat = logic.merge_tasks(st.session_state.scm_tasks, new)
                    st.session_state.scm_tasks = merged
                    ok = store.save_tasks(merged)
                    st.success(
                        f"반영됨 — 추가 {stat['추가']} · 갱신 {stat['갱신']} · 완료유지 {stat['완료유지']}"
                        + ("" if ok else "  (⚠️ 영속저장 미설정 — 세션에만 반영)"))
                    st.rerun()
            except Exception as e:
                st.error(f"파싱/반영 실패: {e}")
    with st.popover("🗑️ 전체 초기화"):
        st.caption("누적된 일정을 모두 지웁니다(되돌릴 수 없음).")
        if st.button("정말 초기화", type="secondary"):
            st.session_state.scm_tasks = []
            store.save_tasks([])
            st.rerun()

tasks = st.session_state.scm_tasks
if not tasks:
    st.info("아직 누적된 일정이 없습니다. 위 '주간 업데이트'에서 요약 파일을 올려 시작하세요.")
    st.stop()

# ── 요약(미완료 기준) ──
def _cnt(pred):
    return sum(1 for t in tasks if not t.get("완료") and pred(t))


wd_today = today.strftime("%Y%m%d")
overdue = _cnt(lambda t: (t.get("처리예정일") or "") and t["처리예정일"] < wd_today)
todo_today = _cnt(lambda t: t.get("처리예정일") == wd_today)
end_cnt = _cnt(lambda t: t.get("유형") == "종료")
done_cnt = sum(1 for t in tasks if t.get("완료"))

m1, m2, m3, m4 = st.columns(4)
m1.metric("🔴 지남(미완료)", overdue)
m2.metric("🟠 오늘", todo_today)
m3.metric("🔚 종료(재고조정)", end_cnt)
m4.metric("✅ 완료", done_cnt)

st.divider()

# ── 통합 일정 (상시, 날짜순) ──
show_done = st.checkbox("완료 포함 보기", value=False)
sch = logic.schedule_df(tasks, today=today, include_done=show_done)
if sch.empty:
    st.success("미완료 일정이 없습니다. 👍 (완료 포함 보기로 지난 건 확인)")
else:
    st.markdown("**📅 통합 일정** — 완료 체크하면 저장되어 다음에도 유지됩니다.")
    edited = st.data_editor(
        sch, hide_index=True, width="stretch",
        disabled=[c for c in sch.columns if c != "완료"],
        column_config={"_key": None},
        key="scm_sched_ed",
    )
    # 완료 동기화 → 저장
    done_map = {r["_key"]: bool(r["완료"]) for _, r in edited.iterrows()}
    changed = False
    for t in tasks:
        k = t.get("key")
        if k in done_map and bool(t.get("완료")) != done_map[k]:
            t["완료"] = done_map[k]
            changed = True
    if changed:
        store.save_tasks(tasks)
        st.session_state.scm_tasks = tasks
        st.rerun()

# ── (선택) ② 입고요청 재고 LOT 조인 ──
st.divider()
with st.expander("📦 ② 보관창고 입고요청 — 재고 LOT 조인 (로케이션별 재고조회 업로드 시)"):
    up_inv = st.file_uploader("로케이션별 재고조회 xlsx", type=["xlsx"], key="scm_inv")
    g2_tasks = [t for t in tasks if t.get("유형") == "작업"
                and str(t.get("구분", "")).startswith("②") and not t.get("완료")]
    if not g2_tasks:
        st.info("미완료 ②입고요청 항목이 없습니다.")
    elif up_inv is None:
        st.caption("재고조회 파일을 올리면 '어디서 몇 박스 집어 보낼지'를 보여줍니다.")
    else:
        try:
            inv = logic.parse_inventory(up_inv.getvalue())
            g2df = pd.DataFrame([{
                "품목코드": t["품목코드"], "품목명": t["품목명"], "보관창고": t["보관창고"],
                "제조일자": t["제조일자"], "요청Box": t["요청Box"],
                "도달일자": logic._parse_d8(t["도달일자"]), "종료일자": logic._parse_d8(t["종료일자"]),
                "처리사유": t["처리사유"],
            } for t in g2_tasks])
            view2 = logic.join_move_stock(g2df, inv)
            st.dataframe(view2, hide_index=True, width="stretch")
        except Exception as e:
            st.error(f"재고 조인 실패: {e}")

# ── 작업지시 엑셀 다운로드 (현재 미완료 일정) ──
st.divider()
dl = logic.schedule_df(tasks, today=today, include_done=False)
if not dl.empty:
    xls = logic.build_xlsx({"통합일정": dl.drop(columns=["_key"])})
    st.download_button(
        "⬇️ 작업지시(통합일정) 엑셀 다운로드",
        data=xls,
        file_name=f"지정출고_일정_{today.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
