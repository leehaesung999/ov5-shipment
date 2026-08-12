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

# ---------- 🌟 모니터링 재고표로 한 번에 업데이트 ----------
st.subheader("🌟 모니터링 재고표로 한 번에 업데이트 (.xlsb)")
st.caption("`재고현황_NewForm(모니터링 재고표).xlsb`의 **Item기준정보** 시트로 "
           "**Item마스터(하대·품명·입수·유통기한월) + 담당자 + 품목등록일/Flag**를 한 번에 갱신합니다. "
           "쿠팡·OV5·실사지용 원본은 자동 합성됩니다.")
up_mon = st.file_uploader("모니터링 재고표 .xlsb 업로드", type=["xlsb"], key="hub_mon")
if up_mon is not None and st.button("📥 모니터링으로 일괄 갱신", type="primary", key="save_mon"):
    try:
        with st.spinner("파싱 중 (.xlsb는 다소 느립니다)…"):
            r = store.save_monitoring(up_mon.getvalue())
        tail = " (공용/Supabase)" if (r["item_ok"] or r["dam_ok"] or r["flag_ok"]) else " (로컬 시드)"
        st.success(f"일괄 갱신 완료 — Item {r['item']:,}개 갱신(전체 {r.get('item_total', r['item']):,}) · "
                   f"담당자 {r['담당자']:,} · 등록일/Flag {r['flag']:,}{tail}")
        st.rerun()
    except Exception as e:
        st.error(f"갱신 실패: {e}")
st.divider()

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

# ---------- ③ 물품담당자 ----------
st.subheader("③ 물품담당자 (개인정보 · 공유)")
_dam = store.load_damdangja()
if _dam:
    st.success(f"등록됨: **{len(_dam):,}명 매핑** — 실사지/점검·재고출고시점과 공유")
else:
    st.info("아직 등록된 물품담당자가 없습니다." +
            ("" if store.use_supabase() else " (로컬 모드는 Supabase 필요)"))
up_dam = st.file_uploader("물품담당자 xlsx 업로드 (코드·담당자 열)", type=["xlsx"], key="hub_dam")
if up_dam is not None and st.button("물품담당자 저장", type="primary", key="save_dam"):
    try:
        n, ok = store.save_damdangja(up_dam.getvalue())
        if ok:
            st.success(f"저장 완료: {n:,}명 (공용/Supabase) — 실사지·재고출고시점에 자동 반영")
        else:
            st.warning(f"파싱 {n:,}명 — 저장 실패(로컬은 Supabase 미설정)")
        st.rerun()
    except Exception as e:
        st.error(f"저장 실패: {e}")
st.caption("담당자는 개인정보라 저장소(Supabase)에만 보관하고 파일/시드로 남기지 않습니다.")

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

st.divider()

# ---------- 🔎 품목 담당자·등록일 조회 (검색/정렬) ----------
st.subheader("🔎 품목 담당자·등록일 조회")
_flags = store.load_itemflags()
if not _flags:
    st.info("모니터링 재고표를 먼저 업로드하면 담당자·품목등록일·Flag를 검색/정렬할 수 있습니다.")
else:
    fdf = pd.DataFrame([{"코드": k, "품명": v.get("nm"), "담당자": v.get("담당자"),
                         "품목등록일": v.get("등록일"), "Flag": v.get("flag")}
                        for k, v in _flags.items()])
    fc = st.columns([2, 1, 1, 1])
    q = fc[0].text_input("검색 (코드·품명·담당자)", key="flag_q")
    dam_opts = sorted([x for x in fdf["담당자"].dropna().unique() if str(x).strip()])
    flag_opts = sorted([x for x in fdf["Flag"].dropna().unique() if str(x).strip()])
    dam_sel = fc[1].multiselect("담당자", dam_opts, key="flag_dam")
    flag_sel = fc[2].multiselect("Flag", flag_opts, key="flag_flag")
    sort_col = fc[3].selectbox("정렬", ["담당자", "품목등록일", "Flag", "코드", "품명"], key="flag_sort")

    view = fdf
    if q:
        ql = q.strip().lower()
        view = view[view.apply(
            lambda r: ql in str(r["코드"]).lower() or ql in str(r["품명"]).lower()
            or ql in str(r["담당자"]).lower(), axis=1)]
    if dam_sel:
        view = view[view["담당자"].isin(dam_sel)]
    if flag_sel:
        view = view[view["Flag"].isin(flag_sel)]
    view = view.sort_values(sort_col, na_position="last", kind="stable")

    _flagcnt = ", ".join(f"{k}:{int((fdf['Flag'] == k).sum())}" for k in flag_opts)
    st.caption(f"{len(view):,} / {len(fdf):,}품목  ·  담당자 {len(dam_opts)}명  ·  Flag {_flagcnt}")
    st.dataframe(view, width="stretch", hide_index=True, height=460)

    import io as _io
    _buf = _io.BytesIO()
    with pd.ExcelWriter(_buf, engine="openpyxl") as _w:
        view.to_excel(_w, index=False, sheet_name="담당자_등록일")
    st.download_button("📥 조회결과 다운로드 (xlsx)", _buf.getvalue(),
                       "담당자_품목등록일.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
