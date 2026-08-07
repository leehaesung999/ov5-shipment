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
from master_hub import store as hub  # noqa: E402


def _apply_hub_master(master: dict, history: dict) -> dict:
    """공용 기준정보 허브의 입수·하대(배면×배단)·품명을 safety_stock 마스터에 덮어씌움.

    - 대상: 기존 마스터 코드 ∪ 히스토리 코드 (숫자 코드만).
    - 하대/입수: 비어있으면 채우고, 값이 다르면 허브값으로 교정(허브=ERP 원본, 정답).
    - 허브에 없는 코드(예: 8000xxx 특수코드)는 그대로 보존.
    반환: 갱신 통계.
    """
    hip, hha, hnm = hub.ipsu_map(), hub.hadae_map(), hub.name_map()
    if not hha and not hip:
        return {}
    codes = set(master) | {str(c) for c in history}
    st_ = {"하대교정": 0, "하대신규": 0, "입수교정": 0, "입수신규": 0, "신규품목": 0}
    for k in codes:
        if not str(k).isdigit():
            continue
        ci = int(k)
        h_plt, h_ip, h_nm = hha.get(ci), hip.get(ci), hnm.get(ci)
        if h_plt is None and h_ip is None:
            continue
        m = master.get(k)
        if m is None:
            master[k] = {"ip": int(h_ip) if h_ip else None,
                         "plt": int(h_plt) if h_plt else None,
                         "cat": "", "nm": h_nm or ""}
            st_["신규품목"] += 1
            continue
        if h_plt is not None:
            cur = m.get("plt")
            if cur in (None, 0, ""):
                m["plt"] = int(h_plt); st_["하대신규"] += 1
            elif int(cur) != int(h_plt):
                m["plt"] = int(h_plt); st_["하대교정"] += 1
        if h_ip is not None:
            cur = m.get("ip")
            if cur in (None, 0, ""):
                m["ip"] = int(h_ip); st_["입수신규"] += 1
            elif int(cur) != int(h_ip):
                m["ip"] = int(h_ip); st_["입수교정"] += 1
        if not m.get("nm") and h_nm:
            m["nm"] = h_nm
    return st_

KST = timezone(timedelta(hours=9))

st.title("📐 BNF(비네이버) 안전재고 산출")
st.caption("매월 CJ출고실적 1개만 올리면 12개월 누적으로 안전재고를 재계산합니다. "
           "비네이버(네이버·토스 제외)·딜 제외·박스/파레트 기준.")

# ---------- 히스토리 로드 ----------
if "ss_history" not in st.session_state:
    h, src = store.load_history()
    st.session_state.ss_history = h
    st.session_state.ss_src = src
history = st.session_state.ss_history
master = store.load_master()

# 공용 기준정보 허브의 입수·하대 자동 적용 (기본 ON) — 한 번 올리면 이 페이지도 최신 반영
use_hub = st.checkbox("🗂️ 공용 기준정보 허브의 입수·하대 자동 적용 (권장)", value=True,
                      help="공용 기준정보 관리에서 올린 ERP Item 마스터의 입수·하대(배면×배단)를 "
                           "이 계산에 반영합니다. 끄면 내장 마스터값을 사용합니다.")
if use_hub:
    _hstat = _apply_hub_master(master, history)
    if _hstat and sum(_hstat.values()):
        st.caption("🗂️ 허브 반영 — " + " · ".join(f"{k} {v}건" for k, v in _hstat.items() if v))

# 품목코드 승계(입수변경 등으로 코드 변경) — 구코드 이력·마스터를 신코드로 이전
remap = store.load_code_remap()
remap_log = core.apply_code_remap(history, master, remap)

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

# ---------- 품목코드 승계 (입수변경 등으로 코드가 바뀐 경우) ----------
with st.expander(f"🔀 품목코드 승계 (입수변경) — 적용 {len(remap)}건", expanded=False):
    st.caption("입수량이 바뀌며 품목코드가 변경된 경우, **구코드의 판매이력(낱개 기준)**을 "
               "신코드로 승계해 안전재고를 이어서 산출합니다. 낱개 상품이 동일할 때만 사용하세요.")
    if remap_log:
        st.dataframe(pd.DataFrame(remap_log), width="stretch", hide_index=True)
    else:
        st.info("등록된 승계 규칙이 없습니다.")
    with st.form("remap_add", clear_on_submit=True):
        rc = st.columns([1, 1, 1, 2])
        r_old = rc[0].text_input("구코드")
        r_new = rc[1].text_input("신코드")
        r_ip = rc[2].number_input("신입수", 1, 9999, 24)
        r_plt = rc[3].number_input("신하대박스수(선택,0=구코드유지)", 0, 9999, 0)
        if st.form_submit_button("➕ 승계규칙 추가", type="secondary"):
            if r_old.strip() and r_new.strip():
                ok = store.add_code_remap(r_old.strip(), r_new.strip(), r_ip,
                                          r_plt or None)
                st.success("추가됨" + (" (Supabase)" if ok else " (로컬은 시드파일로만 반영)")
                           + " — 재계산하면 반영됩니다.")
                st.rerun()
            else:
                st.error("구코드·신코드를 입력하세요.")

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
