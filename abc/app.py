# -*- coding: utf-8 -*-
"""ABC 재배치 분석 — Streamlit 페이지.

로컬 ABC 재배치 분석기(tkinter)의 웹 이식본.
데이터는 Supabase(`app_settings` 테이블 + base64) 영구 저장.
"""
from __future__ import annotations

import base64
import io
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import core  # noqa: E402

TMP = HERE / "_tmp"
TMP.mkdir(exist_ok=True)
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

MASTER_META_KEY = "abc_master_meta"
MASTER_B64_KEY = "abc_master_b64"
INDEX_KEY = "abc_monthly_index"


def _month_key(ym: str) -> str:
    return f"abc_monthly_b64_{ym.replace('-', '_')}"


# ---------- Supabase ----------
try:
    _SB_OK = bool(st.secrets.get("SUPABASE_URL")) and bool(st.secrets.get("SUPABASE_KEY"))
except Exception:
    _SB_OK = False


@st.cache_resource(show_spinner=False)
def _sb():
    from supabase import create_client
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])


def _fetch(key):
    if not _SB_OK:
        return None
    try:
        r = _sb().table("app_settings").select("value").eq("key", key).execute()
        return r.data[0]["value"] if r.data else None
    except Exception:
        return None


def _upsert(key, value):
    if not _SB_OK:
        return False
    try:
        _sb().table("app_settings").upsert(
            {"key": key, "value": value}, on_conflict="key").execute()
        return True
    except Exception as e:
        st.error(f"Supabase 저장 실패: {e}")
        return False


def _delete(key):
    if not _SB_OK:
        return False
    try:
        _sb().table("app_settings").delete().eq("key", key).execute()
        return True
    except Exception:
        return False


def _fetch_master_meta():
    return _fetch(MASTER_META_KEY)


@st.cache_data(show_spinner=False)
def _fetch_master_bytes(version):
    v = _fetch(MASTER_B64_KEY)
    return base64.b64decode(v) if v else None


def _store_master(data: bytes, count: int):
    ok = _upsert(MASTER_B64_KEY, base64.b64encode(data).decode())
    if ok:
        _upsert(MASTER_META_KEY, {
            "updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "n": int(count),
        })
    _fetch_master_bytes.clear()
    return ok


def _fetch_index() -> list:
    idx = _fetch(INDEX_KEY) or []
    return sorted(set(idx))


def _store_index(index_list):
    return _upsert(INDEX_KEY, sorted(set(index_list)))


@st.cache_data(show_spinner=False)
def _fetch_month_bytes(ym: str, version):
    v = _fetch(_month_key(ym))
    return base64.b64decode(v) if v else None


def _store_month(ym: str, data: bytes):
    ok = _upsert(_month_key(ym), base64.b64encode(data).decode())
    if ok:
        idx = _fetch_index()
        if ym not in idx:
            idx.append(ym)
            _store_index(idx)
    _fetch_month_bytes.clear()
    return ok


def _delete_month(ym: str):
    _delete(_month_key(ym))
    idx = [x for x in _fetch_index() if x != ym]
    _store_index(idx)
    _fetch_month_bytes.clear()


# ---------- 헬퍼 ----------
def _guess_ym(filename: str) -> str:
    m = re.search(r"(20\d{2})[-_]?(\d{2})", filename)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return ""


def _load_workbook_count(data: bytes) -> int:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
        n = sum(1 for _ in wb.active.iter_rows(min_row=2, values_only=True)) or 0
        wb.close()
        return n
    except Exception:
        return 0


# ---------- UI ----------
st.title("📊 ABC 재배치 분석")

if not _SB_OK:
    st.error("Supabase 자격증명이 설정되지 않았습니다. `st.secrets`에 SUPABASE_URL/KEY를 등록하세요.")
    st.stop()


# ────────── ① 마스터 관리 ──────────
st.header("① 로케 마스터")
meta = _fetch_master_meta()
c1, c2 = st.columns([3, 2])
with c1:
    if meta:
        st.success(f"등록됨 · 마지막 업데이트 **{meta.get('updated','?')}** · {meta.get('n',0):,}품목")
    else:
        st.warning("아직 등록되지 않았습니다.")
with c2:
    if meta:
        b = _fetch_master_bytes(meta.get("updated", ""))
        if b:
            st.download_button("📥 현재 마스터 다운로드", b, file_name="location_master.xlsx",
                               mime=MIME_XLSX, width='stretch')

up_master = st.file_uploader("교체할 마스터 xlsx 업로드", type=["xlsx"], key="upl_master")
if up_master:
    data = up_master.getvalue()
    n = _load_workbook_count(data)
    if st.button(f"✓ 마스터 등록 ({n:,}행)", type="primary", width='stretch'):
        if _store_master(data, n):
            st.success("등록 완료")
            st.rerun()

st.divider()

# ────────── ② 월별 데이터 등록 ──────────
st.header("② 월별 출고 데이터")

up_month = st.file_uploader("월별 xlsx 업로드 (예: 2026-06.xlsx)", type=["xlsx"], key="upl_month")
if up_month:
    default_ym = _guess_ym(up_month.name)
    ym = st.text_input("기준월 (YYYY-MM)", value=default_ym, max_chars=7, key="ym_input")
    valid = bool(re.match(r"^20\d{2}-\d{2}$", ym or ""))
    if not valid:
        st.warning("YYYY-MM 형식으로 입력하세요.")
    if valid and st.button(f"✓ {ym} 로 등록/교체", type="primary", width='stretch'):
        if _store_month(ym, up_month.getvalue()):
            st.success(f"{ym} 저장 완료")
            st.rerun()

# 등록된 월 리스트
index = _fetch_index()
if index:
    st.caption(f"등록된 월: **{len(index)}개** ({index[0]} ~ {index[-1]})")
    with st.expander("등록된 월 목록 (삭제 가능)"):
        del_target = st.selectbox("삭제할 월 선택", ["(선택 안함)"] + index[::-1], key="del_sel")
        if del_target != "(선택 안함)":
            if st.button(f"🗑 {del_target} 삭제", key="del_btn"):
                _delete_month(del_target)
                st.success(f"{del_target} 삭제됨")
                st.rerun()
        # 목록 표
        st.dataframe(pd.DataFrame({"기준월": index}), width='stretch',
                     height=min(300, 80 + len(index) * 32), hide_index=True)
else:
    st.info("아직 등록된 월이 없습니다.")

st.divider()

# ────────── ③ 분석 실행 ──────────
st.header("③ ABC 재분류 + 히트맵 생성")

if not (meta and index):
    st.warning("마스터와 최소 1개월 이상 데이터가 필요합니다.")
    st.stop()

ref_ym = st.selectbox("기준월 선택", index[::-1], index=0, key="ref_sel")

if st.button("▶ 분석 실행", type="primary", width='stretch'):
    log_lines = []
    def log(m):
        log_lines.append(m)

    try:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        sess = TMP / f"session_{ts}"
        (sess / "monthly").mkdir(parents=True)
        (sess / "output").mkdir()

        with st.spinner("Supabase에서 데이터 다운로드..."):
            # 마스터
            master_path = sess / "location_master.xlsx"
            master_path.write_bytes(_fetch_master_bytes(meta.get("updated", "")))
            # 월별 (기준월까지만 사용해도 core가 알아서 처리)
            for ym in index:
                b = _fetch_month_bytes(ym, meta.get("updated", ""))
                if b:
                    (sess / "monthly" / f"{ym}.xlsx").write_bytes(b)

        with st.spinner(f"분석 실행 중 (기준월 {ref_ym})..."):
            result = core.run_analysis(
                str(sess / "monthly"),
                str(master_path),
                str(sess / "output"),
                log=log,
            )

        st.success("분석 완료")
        with st.expander("실행 로그", expanded=False):
            st.code("\n".join(log_lines))

        # 결과 파일
        out_files = sorted((sess / "output").glob("*.xlsx"))
        if not out_files:
            st.error("결과 파일이 생성되지 않았습니다.")
        else:
            for f in out_files:
                st.download_button(f"📥 {f.name} 다운로드",
                                   data=f.read_bytes(), file_name=f.name,
                                   mime=MIME_XLSX, width='stretch',
                                   key=f"dl_{f.name}")
            # 요약 시트 미리보기
            try:
                xls = pd.ExcelFile(out_files[0])
                if "요약" in xls.sheet_names:
                    st.subheader("요약")
                    st.dataframe(xls.parse("요약"), width='stretch', hide_index=True)
                if "ABC_리스트" in xls.sheet_names:
                    st.subheader("ABC_리스트 (상위 30)")
                    st.dataframe(xls.parse("ABC_리스트").head(30),
                                 width='stretch', hide_index=True)
            except Exception:
                pass
    except Exception as e:
        st.error(f"오류: {e}")
        st.code(traceback.format_exc())
        with st.expander("실행 로그"):
            st.code("\n".join(log_lines))
