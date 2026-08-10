# -*- coding: utf-8 -*-
"""월입고 집계 — WMS '입고 진행현황' 업로드 → 날짜별 건수·파레트 집계.

- 하대(배면×배단)는 **공용 기준정보 허브**에서 가져온다(별도 업로드 불필요).
- 파레트는 **날짜별**로 계산: (입고일자, 품목)별 ceil(입고박스수 / 하대)를 더함.
  → 같은 품목이 여러 날 나눠 들어와도 실제 물리 파레트 수와 일치.
- 건수는 (입고일자, Item code, 소비기한) 고유 조합 수를 날짜별로 집계.
"""
from __future__ import annotations

import io
import math
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent))  # 레포 루트 (master_hub, page_help)
from master_hub import store as hub  # noqa: E402
from page_help import show_help  # noqa: E402

try:
    st.set_page_config(page_title="월입고 집계", layout="wide")
except Exception:
    pass

# ---------- 컬럼 후보 ----------
DATE_CANDS = ["입고일자"]
ITEM_CANDS = ["Item code", "품목코드", "ITEMID"]
QTY_ACTUAL = ["입고수량(Box)", "입고Box수량"]   # 실제 입고(들어온) 박스
QTY_REQUEST = ["의뢰수량(Box)"]                 # 요청 박스
EXP_CANDS = ["소비기한", "유통기한"]
TYPE_CANDS = ["입고유형", "입고타입"]


def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [" ".join(str(c).replace("\n", " ").replace("\r", " ").split())
                  for c in df.columns]
    return df


def _pick(cols, cands):
    for c in cands:
        if c in cols:
            return c
    return None


def read_inbound(file, sheet_name) -> pd.DataFrame:
    """헤더 행을 자동 탐색해서 '입고일자'가 있는 행을 헤더로 읽는다."""
    raw = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=15)
    header_row = 0
    for i in range(len(raw)):
        vals = [" ".join(str(v).replace("\n", " ").split()) for v in raw.iloc[i].tolist()]
        if "입고일자" in vals:
            header_row = i
            break
    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
    return _norm_cols(df)


def _to_intcode(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def _fmt_date(v):
    """'20260701' / datetime → 'YYYY-MM-DD' 문자열."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    if len(s) == 8 and s.isdigit():
        return f"{s[:4]}-{s[4:6]}-{s[6:]}"
    return s


def compute(df: pd.DataFrame, hadae: dict, names: dict, qty_col: str):
    """날짜별 건수·파레트 집계.

    반환: (요약 by_date, 상세 detail, 미매핑 품목코드 리스트, 총건수, 총파레트)
    """
    cols = list(df.columns)
    date_col = _pick(cols, DATE_CANDS)
    item_col = _pick(cols, ITEM_CANDS)
    exp_col = _pick(cols, EXP_CANDS)
    if not date_col or not item_col:
        raise ValueError(f"'입고일자' 또는 품목코드 컬럼을 찾을 수 없습니다. 현재 컬럼: {cols}")
    if qty_col not in cols:
        raise ValueError(f"박스 수량 컬럼 '{qty_col}' 이 없습니다. 현재 컬럼: {cols}")

    work = df.copy()
    work["_날짜"] = work[date_col].map(_fmt_date)
    work = work[work["_날짜"].notna() & (work["_날짜"].astype(str).str.strip() != "")]
    work["_code"] = work[item_col].map(_to_intcode)
    work["_박스"] = pd.to_numeric(work[qty_col], errors="coerce").fillna(0)

    # 파레트: (날짜, 품목)별 박스 합 → ceil(합/하대)
    grp = (work.groupby(["_날짜", "_code"], dropna=False)["_박스"]
               .sum().reset_index())
    grp["하대"] = grp["_code"].map(hadae)
    grp["품명"] = grp["_code"].map(names)

    def _plt(r):
        q, h = r["_박스"], r["하대"]
        if pd.isna(h) or h == 0 or q <= 0:
            return 0
        return math.ceil(q / h)
    grp["파레트"] = grp.apply(_plt, axis=1)

    detail = grp.rename(columns={"_날짜": "입고일자", "_code": "품목코드",
                                 "_박스": "입고박스"})
    detail = detail[["입고일자", "품목코드", "품명", "입고박스", "하대", "파레트"]] \
        .sort_values(["입고일자", "품목코드"]).reset_index(drop=True)

    # 건수: (날짜, 품목, 소비기한) 고유 조합 수
    if exp_col:
        trip = work[["_날짜", "_code", exp_col]].copy()
        uniq = trip.drop_duplicates()
        cnt = uniq.groupby("_날짜").size()
    else:
        cnt = work.drop_duplicates(["_날짜", "_code"]).groupby("_날짜").size()

    plt_by_date = grp.groupby("_날짜")["파레트"].sum()
    by_date = pd.DataFrame({"건수": cnt, "파레트": plt_by_date}).fillna(0).astype(int)
    by_date = by_date.reset_index().rename(columns={"_날짜": "입고일자"}) \
        .sort_values("입고일자").reset_index(drop=True)

    unmatched = sorted({int(c) for c in grp[grp["하대"].isna()]["_code"].dropna().tolist()})
    return by_date, detail, unmatched, int(by_date["건수"].sum()), int(by_date["파레트"].sum())


def build_xlsx(by_date, detail, unmatched, qty_col) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        summary = pd.DataFrame({
            "항목": ["기간", "건수 합계", "파레트 합계", "사용 박스컬럼",
                     "미매핑 품목수(하대없음)", "처리 일시"],
            "값": [f"{by_date['입고일자'].min()} ~ {by_date['입고일자'].max()}" if len(by_date) else "-",
                   int(by_date["건수"].sum()) if len(by_date) else 0,
                   int(by_date["파레트"].sum()) if len(by_date) else 0,
                   qty_col, len(unmatched),
                   datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        })
        summary.to_excel(w, sheet_name="요약", index=False)
        by_date.to_excel(w, sheet_name="날짜별", index=False)
        detail.to_excel(w, sheet_name="날짜x품목_상세", index=False)
        if unmatched:
            pd.DataFrame({"미매핑_품목코드(하대없음)": unmatched}).to_excel(
                w, sheet_name="미매핑품목", index=False)
    return buf.getvalue()


# ======================= UI =======================
st.title("📥 월입고 집계")
st.caption("WMS 입고 진행현황 업로드 → 날짜별 건수·파레트 집계 (하대는 공용 기준정보 허브 사용)")
show_help({
    "목적": "한 기간의 입고 데이터를 올려 **날짜별 총 건수·총 파레트**를 집계.",
    "필요한 파일": "WMS **입고 진행현황** 다운로드 xlsx (입고일자·Item code·입고수량(Box)·소비기한 포함)",
    "하대 기준": "배면×배단=하대는 **공용 기준정보 관리**의 마스터에서 자동으로 가져옵니다(별도 업로드 불필요).",
    "파레트 계산": "**날짜별**로 (입고일자,품목) 박스합 ÷ 하대를 올림해서 더합니다 → 실제 들어온 파레트 수.",
    "사용 순서": "1. 입고 진행현황 파일 업로드\n"
                 "2. 박스 기준(실입고/의뢰) 선택\n"
                 "3. 날짜별 건수·파레트 표 확인 → 엑셀 다운로드",
    "참고": "하대가 마스터에 없는 품목은 파레트 0으로 집계되며 '미매핑품목'에 표시됩니다.",
}, expanded=False)

hadae = hub.hadae_map()
names = hub.name_map()
if hadae:
    st.success(f"공용 기준정보 허브: 하대 등록 {len(hadae):,}품목")
else:
    st.warning("공용 기준정보 허브에 하대 정보가 없습니다 — **공용 기준정보 관리**에서 마스터를 먼저 등록하세요. "
               "(지금 실행하면 파레트가 전부 0으로 나옵니다)")

up = st.file_uploader("입고 진행현황 xlsx 업로드", type=["xlsx", "xls"], key="wolipgo_up")
if not up:
    st.info("입고 진행현황 파일을 업로드하세요.")
    st.stop()

# 시트 선택
try:
    xls = pd.ExcelFile(up)
    sheets = xls.sheet_names
except Exception as e:
    st.error(f"파일을 읽을 수 없습니다: {e}")
    st.stop()
default_sheet = "sheet" if "sheet" in sheets else sheets[0]
sheet = st.selectbox("시트 선택", sheets, index=sheets.index(default_sheet))

df = read_inbound(up, sheet)
qty_choice = st.radio(
    "박스 수량 기준", ["실입고 (입고수량(Box))", "의뢰 (의뢰수량(Box))"],
    horizontal=True,
    help="'들어온' 파레트는 실입고 기준이 맞습니다. 요청 기준으로 보려면 의뢰 선택.")
qty_col = _pick(list(df.columns), QTY_ACTUAL if qty_choice.startswith("실입고") else QTY_REQUEST)
if qty_col is None:  # 선택한 기준 컬럼이 없으면 반대 후보로 폴백
    qty_col = _pick(list(df.columns), QTY_ACTUAL + QTY_REQUEST)

try:
    by_date, detail, unmatched, tot_cnt, tot_plt = compute(df, hadae, names, qty_col)
except Exception as e:
    st.error(str(e))
    st.stop()

c1, c2, c3 = st.columns(3)
c1.metric("총 건수", f"{tot_cnt:,}")
c2.metric("총 파레트", f"{tot_plt:,}")
c3.metric("기간", f"{by_date['입고일자'].min()} ~ {by_date['입고일자'].max()}" if len(by_date) else "-")
if unmatched:
    st.warning(f"하대 미매핑 {len(unmatched)}품목 → 파레트 0 처리 (엑셀 '미매핑품목' 시트 참고). "
               f"예: {', '.join(map(str, unmatched[:10]))}{' …' if len(unmatched) > 10 else ''}")

st.subheader("날짜별 집계")
st.dataframe(by_date, hide_index=True, width='stretch')

with st.expander("날짜×품목 상세 보기"):
    st.dataframe(detail, hide_index=True, width='stretch')

st.download_button(
    "📥 집계 결과 Excel 다운로드",
    build_xlsx(by_date, detail, unmatched, qty_col),
    file_name=f"월입고집계_{datetime.now():%Y%m%d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    width='stretch')
