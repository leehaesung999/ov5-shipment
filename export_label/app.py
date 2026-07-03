# -*- coding: utf-8 -*-
"""수출 라벨(피킹분) 자동 인쇄 — Streamlit 페이지 (통합앱).

'복사본 수출.xlsm'의 자동출력() 매크로를 웹으로 옮긴 것.
  · 상세_리스트(inventory,order).xlsx 업로드
  · 라벨 텍스트 = 'Shipping Instruction'
  · 출력장수   = ROUNDUP(P/L환산) + 여유분(기본 2)
  · [🖨️ 인쇄] 버튼 → 브라우저 인쇄창(프린터 인쇄 또는 PDF 저장)
"""
from __future__ import annotations

import html
import json
import math

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

try:
    st.set_page_config(page_title="수출 라벨 인쇄", layout="wide")
except Exception:
    pass

st.title("🖨️ 수출 라벨(피킹분) 인쇄")
st.caption("상세_리스트(inventory,order).xlsx → 출고건별 라벨을 출력장수만큼 자동 인쇄")
st.info("규칙 · **라벨** = Shipping Instruction · **출력장수** = ROUNDUP(P/L환산) + 여유분  "
        "→ 아래 표에서 수정 가능, [🖨️ 인쇄] 누르면 인쇄창이 바로 열립니다.", icon="📐")


def find_col(cols, *cands):
    low = {str(c).strip().lower(): c for c in cols}
    for cand in cands:
        for k, orig in low.items():
            if cand.lower() in k:
                return orig
    return None


up = st.file_uploader("상세_리스트 xlsx 업로드", type=["xlsx"], key="exlabel_up")
if not up:
    st.info("상세_리스트(inventory,order) 파일을 업로드하세요.")
    st.stop()

try:
    df = pd.read_excel(up, sheet_name=0)
except Exception as e:
    st.error(f"엑셀 읽기 오류: {e}")
    st.stop()
df.columns = [" ".join(str(c).replace("\n", " ").split()) for c in df.columns]

c_label = find_col(df.columns, "shipping instruction", "shipping", "orderid")
c_pl = find_col(df.columns, "p/l", "pl환산", "파레트")
c_qty = find_col(df.columns, "수량")
if c_label is None or c_pl is None:
    st.error(f"필수 컬럼을 찾지 못했습니다. (라벨={c_label}, P/L환산={c_pl})\n"
             f"현재 컬럼: {list(df.columns)}")
    st.stop()

base = st.number_input("출력장수 여유분 (파레트 올림 + N장)", min_value=0, max_value=20, value=2, step=1)

# 원본 → 표 데이터 구성 (라벨 빈 행 제외)
rows = []
for _, r in df.iterrows():
    label = str(r[c_label]).strip() if pd.notna(r[c_label]) else ""
    if not label:
        continue
    pl = pd.to_numeric(r[c_pl], errors="coerce")
    pl = 0.0 if pd.isna(pl) else float(pl)
    qty = "" if (c_qty is None or pd.isna(r[c_qty])) else str(r[c_qty]).strip()
    copies = max(0, math.ceil(pl) + int(base))
    rows.append({"라벨(출고지시)": label, "출력장수": copies, "파레트": round(pl, 2), "수량": qty})

if not rows:
    st.warning("라벨로 쓸 'Shipping Instruction' 값이 있는 행이 없습니다.")
    st.stop()

st.markdown("#### 출력 목록 (수정 가능 · 행 삭제·추가 가능)")
edited = st.data_editor(
    pd.DataFrame(rows),
    num_rows="dynamic",
    use_container_width=True,
    key="exlabel_editor",
    column_config={
        "라벨(출고지시)": st.column_config.TextColumn(width="large"),
        "출력장수": st.column_config.NumberColumn(min_value=0, max_value=999, step=1),
        "파레트": st.column_config.NumberColumn(disabled=True),
        "수량": st.column_config.TextColumn(),
    },
)

# 유효 행만 (라벨 있고 출력장수>=1)
final = []
for _, r in edited.iterrows():
    label = str(r["라벨(출고지시)"]).strip() if pd.notna(r["라벨(출고지시)"]) else ""
    copies = int(r["출력장수"]) if pd.notna(r["출력장수"]) else 0
    if not label or copies < 1:
        continue
    final.append({"label": label, "copies": copies,
                  "pallet": r.get("파레트", ""), "qty": r.get("수량", "")})

total = sum(r["copies"] for r in final)
st.markdown(f"**출고건 {len(final)}건 · 총 인쇄 {total}장**")
if total == 0:
    st.stop()


def build_html(items, auto_print):
    pages, idx = [], 0
    grand = sum(i["copies"] for i in items)
    for it in items:
        for k in range(it["copies"]):
            idx += 1
            qty = html.escape(str(it["qty"]))
            pages.append(
                f'<div class="page"><div class="top">'
                f'<span>📦 수출 피킹분</span>'
                f'<span>파레트 {html.escape(str(it["pallet"]))} · 수량 {qty}</span></div>'
                f'<div class="mid"><div class="txt">{html.escape(it["label"])}</div></div>'
                f'<div class="bot"><span>이 출고건 {k+1} / {it["copies"]} 장</span>'
                f'<span>전체 {idx} / {grand}</span></div></div>'
            )
    onload = ' onload="setTimeout(function(){window.print();},250)"' if auto_print else ""
    return (
        '<!doctype html><html lang="ko"><head><meta charset="utf-8"><style>'
        '@page{size:A4;margin:12mm;}*{box-sizing:border-box;}'
        "body{font-family:'Malgun Gothic','맑은 고딕',sans-serif;margin:0;}"
        '.page{page-break-after:always;height:265mm;display:flex;flex-direction:column;}'
        '.page:last-child{page-break-after:auto;}'
        '.top{font-size:20px;color:#222;border-bottom:3px solid #111;padding-bottom:8px;'
        'display:flex;justify-content:space-between;align-items:flex-end;}'
        '.mid{flex:1;display:flex;align-items:center;justify-content:center;text-align:center;padding:10px;}'
        '.mid .txt{font-size:46px;font-weight:800;line-height:1.35;word-break:keep-all;}'
        '.bot{font-size:17px;color:#555;border-top:2px solid #bbb;padding-top:8px;'
        'display:flex;justify-content:space-between;}'
        f'</style></head><body{onload}>' + "".join(pages) + "</body></html>"
    )


print_html = build_html(final, auto_print=True)
comp = (
    '<div style="padding:6px 0;">'
    f'<button id="pbtn" style="font-size:18px;padding:12px 28px;background:#ff4b4b;'
    'color:#fff;border:none;border-radius:8px;cursor:pointer;font-weight:700;">'
    f'🖨️ 라벨 인쇄 ({total}장)</button>'
    '<span style="margin-left:12px;color:#666;">버튼을 누르면 인쇄창이 열립니다 '
    '(프린터로 인쇄하거나 대상=PDF로 저장)</span></div>'
    f'<script>const HTML={json.dumps(print_html)};'
    "document.getElementById('pbtn').addEventListener('click',function(){"
    "var w=window.open('','_blank');"
    "if(!w){alert('팝업이 차단되었습니다. 팝업 허용 후 다시 시도하세요.');return;}"
    "w.document.open();w.document.write(HTML);w.document.close();w.focus();});</script>"
)
components.html(comp, height=80)

with st.expander("🔍 인쇄 미리보기 (스크롤)"):
    components.html(build_html(final, auto_print=False), height=560, scrolling=True)
