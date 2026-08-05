# -*- coding: utf-8 -*-
"""안전재고 계산기 (Streamlit 페이지)

매월 하는 일: **그 달 CJ출고실적 1개만 업로드** → 히스토리에 누적 → 최근 12개월로 안전재고 재계산.
  · 비네이버(네이버·토스 제외)만 자동 추출, 행사no로 딜 식별(σ에서 제외)
  · 롤링 12개월이라 계절 변화(여름 급증 등) 자동 반영
  · 결과: 품목별 안전재고·Min·Max·파레트 → 엑셀 다운로드
"""
import io
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from safety_stock import core, store  # noqa: E402

KST = timezone(timedelta(hours=9))

st.title("📐 안전재고 계산기")
st.caption("매월 CJ출고실적 1개만 올리면 12개월 누적으로 안전재고를 재계산합니다. "
           "비네이버(네이버·토스 제외)·딜 제외·박스/파레트 기준.")

# ---------- 히스토리 로드 ----------
if "ss_history" not in st.session_state:
    h, src = store.load_history()
    st.session_state.ss_history = h
    st.session_state.ss_src = src
history = st.session_state.ss_history
master = store.load_master()

meta = store.history_meta()
c1, c2, c3 = st.columns(3)
dmax = max((d for it in history.values() for d in it.keys()), default="—")
c1.metric("누적 품목수", f"{len(history):,}")
c2.metric("최신 데이터", dmax)
c3.metric("저장소", "Supabase" if store.use_supabase() else "로컬/세션")

# ---------- 설정 ----------
with st.expander("⚙️ 설정 (리드타임·발주주기·서비스레벨)", expanded=False):
    cc = st.columns(4)
    lead = cc[0].number_input("리드타임(일)", 1, 30, core.DEFAULT_SETTINGS["lead_time"])
    cycle = cc[1].number_input("발주주기(일)", 1, 30, core.DEFAULT_SETTINGS["cycle"])
    zsel = cc[2].selectbox("서비스레벨", ["99% (2.33)", "98% (2.05)", "95% (1.65)", "99.5% (2.58)"])
    batch = cc[3].number_input("발주배치(일)", 1, 60, core.DEFAULT_SETTINGS["batch"])
    z = {"99% (2.33)": 2.33, "98% (2.05)": 2.05, "95% (1.65)": 1.65, "99.5% (2.58)": 2.58}[zsel]
    st.caption(f"노출기간 = 리드타임+발주주기 = **{lead+cycle}일**")
settings = {"lead_time": lead, "cycle": cycle, "z": z, "batch": batch}

st.divider()

# ---------- 월별 CJ출고실적 업로드 → 누적 ----------
st.subheader("1️⃣ 이번 달 실적 반영")
up = st.file_uploader("CJ출고실적 xlsx 업로드 (그 달 1개, 'raw' 시트 포함)", type=["xlsx"], key="cj")


def _find_cols(ws):
    """헤더행에서 필요한 열 인덱스 자동 탐색."""
    hdr = [ws.cell(1, j).value for j in range(1, ws.max_column + 1)]
    want = {"ship_date": ["출고일자"], "code": ["품목코드"], "qty": ["출고수량"],
            "channel": ["판매채널"], "event_no": ["행사no", "행사NO", "행사번호"]}
    col = {}
    for key, names in want.items():
        for i, h in enumerate(hdr):
            if h is not None and str(h).strip() in names:
                col[key] = i
                break
    return col


if up is not None:
    if st.button("➕ 이 실적을 히스토리에 반영", type="primary"):
        try:
            wb = openpyxl.load_workbook(io.BytesIO(up.getvalue()), data_only=True, read_only=True)
            ws = wb["raw"] if "raw" in wb.sheetnames else wb[wb.sheetnames[0]]
            col = _find_cols(ws)
            need = {"ship_date", "code", "qty", "channel", "event_no"}
            if not need.issubset(col):
                st.error(f"필요한 열을 못 찾았습니다: 누락 {need - set(col)}. (출고일자·품목코드·출고수량·판매채널·행사no)")
            else:
                rows = ws.iter_rows(min_row=2, values_only=True)
                chunk = core.extract_cj_month(rows, col)
                wb.close()
                before = dmax
                core.merge_history(history, chunk)
                core.trim_history(history, months=14)
                st.session_state.ss_history = history
                saved = store.save_history(history)
                new_dmax = max((d for it in history.values() for d in it.keys()), default="—")
                nqty = sum(v[0] for days in chunk.values() for v in days.values())
                st.success(f"반영 완료: {len(chunk)}품목, 비네이버 {nqty:,.0f}개. "
                           f"최신 {before} → {new_dmax}. "
                           + ("Supabase 저장됨" if saved else "세션에만 반영(로컬)"))
        except Exception as e:
            st.error(f"처리 실패: {e}")

# ---------- 재계산 ----------
st.subheader("2️⃣ 안전재고 재계산")
if st.button("🔄 최근 12개월로 재계산", type="primary"):
    rows, info = core.compute(history, master, settings)
    if not rows:
        st.warning("히스토리가 비어있습니다. 실적을 먼저 반영하세요.")
    else:
        df = pd.DataFrame(rows)
        tot = df["안전재고_EA"].sum()
        i1, i2, i3, i4 = st.columns(4)
        i1.metric("대상 품목", f"{info['품목수']:,}")
        i2.metric("안전재고 합계", f"{tot:,.0f} EA")
        i3.metric("계산 기간", f"{info['시작']} ~ {info['종료']}")
        i4.metric("노출기간", f"{info['노출기간']}일")
        st.dataframe(df, width="stretch", height=460)

        # 엑셀 다운로드
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            df.to_excel(xw, index=False, sheet_name="안전재고")
        dl, sv = st.columns([1, 1])
        dl.download_button("📥 안전재고 산출표 다운로드 (xlsx)", buf.getvalue(),
                           file_name=f"안전재고_산출_{datetime.now(KST):%y%m%d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width="stretch")
        if sv.button("✅ 이 기준을 재고이동계획에 적용", type="secondary", width="stretch"):
            ok = store.save_baseline(rows, settings)
            st.success("재고이동계획 기준으로 저장됨" + (" (Supabase)" if ok else " (로컬 시드)"))
        st.caption("이 표의 안전재고·Min·Max가 재고이동계획의 기준으로 쓰입니다. "
                   "딜(행사)은 제외된 평상시 기준이며, 딜 물량은 이동계획의 이벤트/행사일정으로 별도 반영합니다.")
