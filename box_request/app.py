# -*- coding: utf-8 -*-
"""외박스 요청 리스트 — Streamlit 페이지 (통합앱).

여러 사용자가 지속적으로 함께 추가·수정하는 '공유 요청 리스트'.
  · Supabase(또는 로컬 파일)에 저장되어 새로고침·다른 사용자에게 유지
  · 표에서 직접 추가/수정/삭제, [요청] 체크박스로 요청 표시
  · 엑셀에서 항목 일괄 추가 + 품목코드로 마스터(ERP코드·입고업체) 자동조회
  · 표를 그대로 Ctrl+C 복사해 메일/ERP 등에 붙여넣기
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import streamlit as st

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import store  # noqa: E402

try:
    st.set_page_config(page_title="외박스 요청 리스트", layout="wide")
except Exception:
    pass

COLS = ["요청", "코드", "품명", "수량", "입고업체", "ERP코드", "비고"]

st.title("🗃️ 외박스 요청 리스트")
st.caption(f"저장 위치: {store.backend_name()} · 여러 사용자가 함께 추가·수정")


def norm(s):
    return " ".join(str(s).replace("\n", " ").split())


def clean_num(v):
    s = norm(v)
    if s.endswith(".0"):
        s = s[:-2]
    return "" if s.lower() == "nan" else s


def parse_requests(xls):
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
            return next((j for j, c in enumerate(cols) if c in opts), None)

        ji = {"코드": idx(["코드", "품목코드", "Item code"]), "품명": idx(["품명"]),
              "수량": idx(["수량"]), "업체": idx(["업체", "공급업체", "Supplier"]),
              "erp": idx(["ERP", "ERP코드"]), "비고": idx(["비고"])}
        for _, r in raw.iloc[hdr + 1:].iterrows():
            j = ji["품명"]
            name = clean_num(r.iloc[j]) if (j is not None and pd.notna(r.iloc[j])) else ""
            if not name:
                continue

            def get(key):
                jj = ji[key]
                return clean_num(r.iloc[jj]) if (jj is not None and pd.notna(r.iloc[jj])) else ""

            items.append({"코드": get("코드"), "품명": name, "수량": get("수량"),
                          "업체": get("업체"), "erp": get("erp"), "비고": get("비고")})
    return items


def parse_master(xls):
    """Sheet3형(품목코드+업체 헤더) → {품목코드: [ERP코드, 업체]}."""
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
                    m[code] = [erp, sup]
                return m
    return m


def empty_df():
    return pd.DataFrame(columns=COLS)


def rows_to_df(rows):
    if not rows:
        return empty_df()
    df = pd.DataFrame(rows)
    for c in COLS:
        if c not in df.columns:
            df[c] = "" if c != "요청" else False
    df["요청"] = df["요청"].fillna(False).astype(bool)
    return df[COLS]


# ── 공유 데이터 로드 / 컨트롤 ───────────────────────────────
c1, c2, c3 = st.columns([1, 1, 5])
refresh = c1.button("🔄 새로고침", use_container_width=True, help="공유 저장소에서 최신 리스트 다시 불러오기")
save = c2.button("💾 저장", type="primary", use_container_width=True, help="현재 표를 공유 저장소에 저장")

if refresh or "box_df" not in st.session_state:
    st.session_state.box_df = rows_to_df(store.load_rows())

st.markdown("#### 요청 표 (직접 추가·수정 · 셀 선택 후 **Ctrl+C** 복사)")
edited = st.data_editor(
    st.session_state.box_df,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key="box_editor",
    column_config={
        "요청": st.column_config.CheckboxColumn("요청", help="요청한 항목 체크", default=False),
        "품명": st.column_config.TextColumn(width="large"),
        "코드": st.column_config.TextColumn(help="품목코드 (마스터 조회 키)"),
    },
)

if save:
    df = edited.copy()
    df["요청"] = df["요청"].fillna(False).astype(bool)
    df = df.fillna("")
    df = df[df["품명"].astype(str).str.strip() != ""]
    store.save_rows(df[COLS].to_dict("records"))
    st.session_state.box_df = df[COLS].reset_index(drop=True)
    st.success(f"저장 완료 · {len(df)}건")

st.caption("여럿이 동시에 쓸 때는 큰 수정 전에 **새로고침**으로 최신본을 먼저 받는 걸 권장합니다 "
           "(저장은 마지막에 저장한 내용으로 덮어씁니다).")

# ── 요청 표시 항목만 (복사용) ───────────────────────────────
req = edited[edited["요청"].fillna(False) == True] if "요청" in edited else edited.iloc[0:0]  # noqa: E712
if len(req):
    with st.expander(f"✅ 요청 표시된 항목만 보기 ({len(req)}건 · 복사용)", expanded=False):
        st.dataframe(req[["코드", "품명", "수량", "입고업체", "ERP코드", "비고"]],
                     use_container_width=True, hide_index=True)

# ── 엑셀에서 일괄 추가 / 마스터 업데이트 ────────────────────
with st.expander("📥 엑셀에서 항목 추가 · 마스터(ERP·업체) 업데이트"):
    up = st.file_uploader("요청 리스트 / 마스터 xlsx", type=["xlsx"], key="box_import")
    if up:
        try:
            xls = pd.ExcelFile(up)
        except Exception as e:
            st.error(f"엑셀 읽기 오류: {e}")
            st.stop()
        new_master = parse_master(xls)
        new_items = parse_requests(xls)
        m1, m2 = st.columns(2)
        with m1:
            st.write(f"마스터 감지: **{len(new_master)}건**")
            if new_master and st.button("🔁 이 파일로 마스터 갱신·저장"):
                store.save_master(new_master)
                st.success(f"마스터 저장 완료 · {len(new_master)}건")
        with m2:
            st.write(f"요청 품목 감지: **{len(new_items)}건**")
            auto = st.checkbox("추가 시 코드로 ERP·업체 자동조회", value=True)
            if new_items and st.button("➕ 현재 리스트에 추가"):
                master = store.load_master()
                add = []
                for it in new_items:
                    erp, sup = it["erp"], it["업체"]
                    if auto and it["코드"] in master:
                        m_erp, m_sup = (master[it["코드"]] + ["", ""])[:2]
                        erp = erp or m_erp
                        sup = sup or m_sup
                    add.append({"요청": False, "코드": it["코드"], "품명": it["품명"],
                                "수량": it["수량"], "입고업체": sup, "ERP코드": erp, "비고": it["비고"]})
                merged = pd.concat([st.session_state.box_df, pd.DataFrame(add)], ignore_index=True)
                store.save_rows(merged[COLS].to_dict("records"))
                st.session_state.box_df = merged[COLS]
                st.success(f"{len(add)}건 추가·저장 완료 (총 {len(merged)}건). 위 표에 반영됩니다.")
                st.rerun()
