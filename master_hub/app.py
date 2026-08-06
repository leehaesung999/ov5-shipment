# -*- coding: utf-8 -*-
"""공용 기준정보 관리 — 한 번 업로드하면 여러 페이지가 공용 사용.

- Item 마스터 (ERP Item_*.xlsx): 하대(배면×배단)·품명·입수·유통기한
- 고정로케이션 매핑 (편집본_*.xlsx): Item code → 고정로케이션
"""
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from master_hub import store          # noqa: E402

st.title("🗂️ 공용 기준정보 관리")
st.caption("여기서 한 번 업데이트하면 이 값을 쓰는 페이지들이 **공용으로** 사용합니다. "
           "(현재 연동: 초도물량 차량 배차 — 이후 다른 페이지도 순차 연동 예정)")

if not store.use_supabase():
    st.warning("⚠️ Supabase 미설정(로컬 모드) — 업로드분은 이 앱 폴더의 시드파일에만 저장됩니다. "
               "라이브(클라우드)에서 공용 저장되려면 SUPABASE_URL/KEY 시크릿이 필요합니다.")

item_data, item_meta = store.load_item()
loc_data, loc_meta = store.load_loc()

c1, c2 = st.columns(2)

# ---------- Item 마스터 ----------
with c1:
    st.subheader("① Item 마스터")
    if item_meta:
        hae = item_meta.get("하대보유", "?")
        st.success(f"등록됨: **{item_meta.get('품목수','?')}품목** "
                   f"(하대 보유 {hae}) · 갱신 {item_meta.get('갱신','—')}")
    else:
        st.info("아직 등록된 Item 마스터가 없습니다.")
    up_item = st.file_uploader("ERP Item_*.xlsx 업로드 (A=Item code, D=입수, AD=배면, AE=배단)",
                               type=["xlsx"], key="hub_item")
    if up_item is not None and st.button("Item 마스터 저장", type="primary", key="save_item"):
        try:
            n, ok = store.save_item(up_item.getvalue())
            st.success(f"저장 완료: {n}품목" + (" (공용/Supabase)" if ok else " (로컬 시드)"))
            st.rerun()
        except Exception as e:
            st.error(f"저장 실패: {e}")

# ---------- 고정로케이션 ----------
with c2:
    st.subheader("② 고정로케이션 매핑")
    if loc_meta:
        st.success(f"등록됨: **{loc_meta.get('품목수','?')}품목** · 갱신 {loc_meta.get('갱신','—')}")
    else:
        st.info("아직 등록된 고정로케이션 매핑이 없습니다.")
    up_loc = st.file_uploader("편집본_*.xlsx 업로드 (C=로케이션ID, D=보관타입, G=Item code)",
                              type=["xlsx"], key="hub_loc")
    if up_loc is not None and st.button("고정로케이션 저장", type="primary", key="save_loc"):
        try:
            n, ok = store.save_loc(up_loc.getvalue())
            st.success(f"저장 완료: {n}품목" + (" (공용/Supabase)" if ok else " (로컬 시드)"))
            st.rerun()
        except Exception as e:
            st.error(f"저장 실패: {e}")

st.divider()

# ---------- 미리보기 ----------
with st.expander("미리보기 (상위 20행)"):
    if item_data:
        st.markdown("**Item 마스터**")
        rows = [{"코드": k, "품명": v.get("nm"), "입수": v.get("ip"),
                 "배면": v.get("bm"), "배단": v.get("bd"), "하대": v.get("hadae"),
                 "유통기한(월)": v.get("shelf")}
                for k, v in list(item_data.items())[:20]]
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
    if loc_data:
        st.markdown("**고정로케이션 매핑**")
        lr = [{"코드": k, "고정로케이션": v} for k, v in list(loc_data.items())[:20]]
        st.dataframe(pd.DataFrame(lr), width="stretch", hide_index=True)
