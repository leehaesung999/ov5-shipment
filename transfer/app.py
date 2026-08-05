# -*- coding: utf-8 -*-
"""재고이동계획 (Streamlit 페이지)

매일 하는 일: **재고입력(현재고) 업로드** → 이동_박스·파레트 자동 산출.
기준(안전재고·Min·Max)은 '안전재고 계산기'에서 저장한 값을 사용.
이동 = MIN(요청, 출고가능, 할당) 박스단위. 입고예정 차감·종료 제외·행사 프리쉽 반영.
"""
import io
import sys
from datetime import datetime, timezone, timedelta, date
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from transfer import core as T          # noqa: E402
from safety_stock import store          # noqa: E402

KST = timezone(timedelta(hours=9))

st.title("🚚 재고이동계획")
baseline, bmeta = store.load_baseline()
st.caption("기준(안전재고·Min·Max)은 '안전재고 계산기'에서 저장한 값 사용. "
           f"현재 기준: {bmeta.get('품목수','?')}품목 · 갱신 {bmeta.get('갱신','—')}")

if not baseline:
    st.warning("안전재고 기준이 없습니다. 먼저 '안전재고 계산기'에서 '재고이동계획에 적용'을 눌러 기준을 저장하세요.")
    st.stop()

lead = int((bmeta.get("설정") or {}).get("lead_time", 3))
plan_date = st.date_input("계획일자 (행사 프리쉽 기준)", value=datetime.now(KST).date())


# ---------- 업로드 파서 ----------
def _rows(file, sheet=None, min_row=2):
    # read_only=False: 일부 시스템 export가 read_only에서 행이 잘려 읽히는 문제 회피(파일 작음)
    wb = openpyxl.load_workbook(io.BytesIO(file.getvalue()), data_only=True)
    ws = wb[sheet] if (sheet and sheet in wb.sheetnames) else wb[wb.sheetnames[0]]
    for r in ws.iter_rows(min_row=min_row, values_only=True):
        yield r
    wb.close()


def _code(v):
    if v is None:
        return None
    if isinstance(v, float):
        return str(int(v))
    return str(v).strip()


def parse_stock(file):
    """재고입력: A=품목코드, B=현재고, C=할당(선택). 헤더 1행."""
    out = {}
    for r in _rows(file):
        c = _code(r[0])
        if not c or not (c.isdigit() or c.startswith("P")):
            continue
        cur = r[1] if len(r) > 1 and isinstance(r[1], (int, float)) else None
        alloc = r[2] if len(r) > 2 and isinstance(r[2], (int, float)) else None
        out[c] = {"cur": cur, "alloc": alloc}
    return out


def parse_avail(file):
    """출고가능재고 WMS export: C(3)=item id, AB(28)=출고가능(Box). 헤더 1행."""
    out = {}
    for r in _rows(file):
        if len(r) < 28:
            continue
        c = _code(r[2])
        box = r[27]
        if c and isinstance(box, (int, float)):
            out[c] = out.get(c, 0) + int(box)
    return out


def parse_incoming(file):
    """입고예정: A=코드, E(5)=낱개. 헤더 2행(그들 양식)."""
    out = {}
    for r in _rows(file, min_row=3):
        c = _code(r[0])
        ea = r[4] if len(r) > 4 and isinstance(r[4], (int, float)) else 0
        if c and ea:
            out[c] = out.get(c, 0) + ea
    return out


def parse_ended(file):
    """종료품목: A=CJ코드. 헤더 1행."""
    s = set()
    for r in _rows(file):
        c = _code(r[0])
        if c and (c.isdigit() or c.startswith("P")):
            s.add(c)
    return s


def parse_events_sched(file, plan_d, lead_days):
    """행사일정 export: S(19)=item_number, Q(17)=first_del_date, W(23)=tot_qty,
    J(10)=current_flag. 프리쉽일(=행사시작-리드타임)==계획일자 인 것만 합산."""
    out = {}
    for r in _rows(file):
        if len(r) < 23:
            continue
        if str(r[9]).strip() != "Y":     # current_flag
            continue
        fd = r[16]
        if not hasattr(fd, "date") and not isinstance(fd, date):
            continue
        fdd = fd.date() if hasattr(fd, "date") else fd
        preship = fdd - timedelta(days=lead_days)
        if preship != plan_d:
            continue
        c = _code(r[18])
        q = r[22]
        if c and isinstance(q, (int, float)):
            out[c] = out.get(c, 0) + q
    return out


# ---------- 입력 ----------
st.subheader("1️⃣ 입력 (재고 필수, 나머지 선택)")
up_stock = st.file_uploader("재고입력 xlsx (A:품목코드 B:현재고 C:할당(선택))", type=["xlsx"], key="t_stock")
with st.expander("추가 입력 (출고가능·입고예정·행사·종료)"):
    up_avail = st.file_uploader("출고가능재고 (WMS export)", type=["xlsx"], key="t_avail")
    up_inc = st.file_uploader("입고예정", type=["xlsx"], key="t_inc")
    up_sch = st.file_uploader("행사일정 (시스템 export)", type=["xlsx"], key="t_sch")
    up_end = st.file_uploader("종료품목 (A:CJ코드)", type=["xlsx"], key="t_end")

st.subheader("2️⃣ 수동 이벤트 (선택)")
ev_df = st.data_editor(pd.DataFrame({"품목코드": ["", "", ""], "추가수량(EA)": [None, None, None]}),
                       num_rows="dynamic", width="stretch", key="t_ev")

# ---------- 산출 ----------
if st.button("🚚 이동계획 산출", type="primary", disabled=up_stock is None):
    try:
        stock = parse_stock(up_stock)
        avail = parse_avail(up_avail) if up_avail else None
        incoming = parse_incoming(up_inc) if up_inc else {}
        ended = parse_ended(up_end) if up_end else set()
        events = {}
        if up_sch:
            events.update(parse_events_sched(up_sch, plan_date, lead))
        for _, row in ev_df.iterrows():          # 수동 이벤트 합산
            c = _code(row["품목코드"])
            q = row["추가수량(EA)"]
            if c and isinstance(q, (int, float)) and q:
                events[c] = events.get(c, 0) + q

        rows = T.compute_transfer(baseline, stock, avail, incoming, events, ended)
        summ = T.summarize(rows)
        df = pd.DataFrame(rows)

        m = st.columns(5)
        m[0].metric("이동 품목", f"{summ['이동품목수']:,}")
        m[1].metric("총 이동_박스", f"{summ['총이동_박스']:,}")
        m[2].metric("총 파레트", f"{summ['총파레트']}")
        m[3].metric("가용부족", f"{summ['가용부족']}")
        m[4].metric("종료제외/미입력", f"{summ['종료제외']}/{summ['미입력']}")

        show = df[df["사유"] != "미입력"] if len(stock) < len(baseline) else df
        st.dataframe(show, width="stretch", height=460)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            df.to_excel(xw, index=False, sheet_name="이동계획")
        st.download_button("📥 이동계획 다운로드 (xlsx)", buf.getvalue(),
                           file_name=f"재고이동계획_{plan_date:%y%m%d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width="stretch")
        st.caption("이동=MIN(요청,출고가능,할당) 박스. 파레트환산=이동박스÷하대박스수. "
                   "'미충족_박스'>0=가용부족. 행사는 프리쉽일 하루만 반영(중복 방지).")
    except Exception as e:
        st.error(f"산출 실패: {e}")
