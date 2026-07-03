# -*- coding: utf-8 -*-
"""외박스 요청 리스트 — Streamlit 페이지 (통합앱).

외박스 요청 리스트 xlsx 업로드 → 정리된 '복사 가능한 표' 생성.
  · 요청 리스트 시트(코드·품명·수량·업체 헤더)에서 품목 수집
  · 품목코드로 마스터(Sheet3: 품목코드↔ERP코드·업체)에서 입고업체·ERP코드 자동조회
  · 표에서 직접 수정/행추가, [요청] 체크박스로 요청 표시
  · 표를 그대로 Ctrl+C 복사해서 메일/ERP 등 다른 곳에 붙여넣기
"""
from __future__ import annotations

import pandas as pd
import streamlit as st

try:
    st.set_page_config(page_title="외박스 요청 리스트", layout="wide")
except Exception:
    pass

st.title("🗃️ 외박스 요청 리스트")
st.caption("요청 리스트 xlsx → 입고업체·ERP코드 자동조회 → 복사 가능한 표")

COLS = ["요청", "코드", "품명", "수량", "입고업체", "ERP코드", "비고"]


def norm(s):
    return " ".join(str(s).replace("\n", " ").split())


def clean_num(v):
    s = norm(v)
    if s.endswith(".0"):
        s = s[:-2]
    return "" if s.lower() == "nan" else s


def load_requests(xls):
    """헤더에 '품명'+'수량'이 있는 시트에서 코드·품명·수량·업체·ERP·비고 수집."""
    items = []
    for sn in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=sn, header=None)
        hdr, cols = None, []
        for i in range(min(8, len(raw))):
            row = [norm(x) for x in raw.iloc[i].tolist()]
            if "품명" in row and "수량" in row:
                hdr, cols = i, row
                break
        if hdr is None:
            continue

        def idx(opts):
            for j, c in enumerate(cols):
                if c in opts:
                    return j
            return None

        ji = {k: idx(v) for k, v in {
            "코드": ["코드", "품목코드", "Item code"],
            "품명": ["품명"],
            "수량": ["수량"],
            "업체": ["업체", "공급업체", "Supplier"],
            "erp": ["ERP", "ERP코드"],
            "비고": ["비고"],
        }.items()}
        for _, r in raw.iloc[hdr + 1:].iterrows():
            name = clean_num(r.iloc[ji["품명"]]) if (ji["품명"] is not None and pd.notna(r.iloc[ji["품명"]])) else ""
            if not name:
                continue

            def get(key):
                j = ji[key]
                return clean_num(r.iloc[j]) if (j is not None and pd.notna(r.iloc[j])) else ""

            items.append({"코드": get("코드"), "품명": name, "수량": get("수량"),
                          "업체": get("업체"), "erp": get("erp"), "비고": get("비고")})
    return items


def load_master(xls):
    """Sheet3형(품목코드+업체 헤더) → {품목코드: (ERP코드, 업체)}."""
    m = {}
    for sn in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=sn, header=None)
        for i in range(min(6, len(raw))):
            row = [norm(x) for x in raw.iloc[i].tolist()]
            if "품목코드" in row and "업체" in row:
                jcode = row.index("품목코드")
                jerp = next((k for k in range(jcode + 1, len(row)) if row[k] in ("ERP", "ERP코드")), None)
                jsup = len(row) - 1 - row[::-1].index("업체")
                for _, r in raw.iloc[i + 1:].iterrows():
                    code = clean_num(r.iloc[jcode]) if pd.notna(r.iloc[jcode]) else ""
                    if not code:
                        continue
                    erp = clean_num(r.iloc[jerp]) if (jerp is not None and pd.notna(r.iloc[jerp])) else ""
                    sup = clean_num(r.iloc[jsup]) if pd.notna(r.iloc[jsup]) else ""
                    m[code] = (erp, sup)
                return m
    return m


up = st.file_uploader("외박스 요청 리스트 xlsx 업로드", type=["xlsx"], key="box_up")
if not up:
    st.info("요청 리스트 파일을 업로드하세요. (Sheet3 같은 마스터가 있으면 ERP·업체 자동조회)")
    st.stop()

try:
    xls = pd.ExcelFile(up)
except Exception as e:
    st.error(f"엑셀 읽기 오류: {e}")
    st.stop()

items = load_requests(xls)
if not items:
    st.error("요청 리스트를 인식하지 못했습니다. (헤더에 '품명'과 '수량'이 있는 시트 필요)")
    st.stop()
master = load_master(xls)

auto = st.checkbox("코드로 입고업체·ERP코드 자동조회 (마스터 있을 때)", value=True)

rows, filled = [], 0
for it in items:
    erp, sup = it["erp"], it["업체"]
    if auto and it["코드"] in master:
        m_erp, m_sup = master[it["코드"]]
        if not erp and m_erp:
            erp = m_erp
            filled += 1
        if not sup and m_sup:
            sup = m_sup
    rows.append({"요청": False, "코드": it["코드"], "품명": it["품명"], "수량": it["수량"],
                 "입고업체": sup, "ERP코드": erp, "비고": it["비고"]})

st.success(f"품목 {len(rows)}건 · 마스터 {len(master)}건" +
           (f" · 자동조회로 ERP {filled}건 채움" if auto else ""))
st.markdown("#### 요청 표 (셀 선택 후 **Ctrl+C**로 복사 · 직접 수정·행 추가 가능)")

edited = st.data_editor(
    pd.DataFrame(rows, columns=COLS),
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key="box_editor",
    column_config={
        "요청": st.column_config.CheckboxColumn("요청", help="요청한 항목 체크", default=False),
        "품명": st.column_config.TextColumn(width="large"),
        "코드": st.column_config.TextColumn(help="품목코드 (마스터 조회 키)"),
        "ERP코드": st.column_config.TextColumn(),
        "입고업체": st.column_config.TextColumn(),
    },
)

req = edited[edited["요청"] == True] if "요청" in edited else edited.iloc[0:0]  # noqa: E712
st.markdown(f"**요청 표시: {len(req)}건**")
if len(req):
    with st.expander("✅ 요청 표시된 항목만 보기 (복사용)", expanded=True):
        st.dataframe(req[["코드", "품명", "수량", "입고업체", "ERP코드", "비고"]],
                     use_container_width=True, hide_index=True)
