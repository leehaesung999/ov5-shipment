# -*- coding: utf-8 -*-
"""통합센터 재고 분석기 — Streamlit 웹 (통합앱 페이지).

ERP 재고조회 업로드 → 이중적치 분석 / 재고지 생성 / 유통기한(OV5·OV6) 분석 → 다운로드.
기준정보(물품 마스터)는 레포에 포함된 data/기준정보.xlsx 사용(또는 업로드).
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import streamlit as st

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import core  # noqa: E402

DATA = HERE / "data"
DEFAULT_MASTER = DATA / "기준정보.xlsx"
TMP = HERE / "_tmp"
TMP.mkdir(exist_ok=True)

try:
    st.set_page_config(page_title="통합센터 재고 분석기", layout="wide")
except Exception:
    pass

st.title("🏭 통합센터 재고 분석기")
st.caption("ERP 재고조회 업로드 → 이중적치 / 재고지 / 유통기한 분석 → 엑셀 다운로드")


def _save(upload, name) -> str:
    p = TMP / name
    p.write_bytes(upload.getvalue())
    return str(p)


# ---------- 담당자(개인정보) — 공개 레포 대신 Supabase(비공개)에서 ----------
try:
    _SB_OK = bool(st.secrets.get("SUPABASE_URL")) and bool(st.secrets.get("SUPABASE_KEY"))
except Exception:
    _SB_OK = False


@st.cache_resource(show_spinner=False)
def _sb():
    from supabase import create_client
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])


def _fetch_담당자() -> dict:
    if not _SB_OK:
        return {}
    try:
        r = _sb().table("app_settings").select("value").eq("key", "inventory_담당자").execute()
        return (r.data[0].get("value") or {}) if r.data else {}
    except Exception:
        return {}


def _setup_담당자(up_file) -> tuple:
    """담당자 매핑을 core.담당자_PATH 에 준비. 우선순위: 세션 업로드 > Supabase."""
    import openpyxl
    src = "업로드" if up_file else ("Supabase" if _SB_OK else None)
    if up_file:
        p = TMP / "물품담당자.xlsx"
        p.write_bytes(up_file.getvalue())
        core.담당자_PATH = p
        try:
            return src, len(core.load_담당자(str(p)))
        except Exception:
            return src, 0
    m = _fetch_담당자()
    if m:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["코드", "담당자"])
        for c, n in m.items():
            ws.append([c, n])
        p = TMP / "물품담당자.xlsx"
        wb.save(str(p))
        core.담당자_PATH = p
        return src, len(m)
    return None, 0


# ---------- 사이드바: 기준정보 (접기) ----------
with st.sidebar.expander("⚙️ 기준정보 / 옵션 — 클릭해서 열기", expanded=False):
    if DEFAULT_MASTER.exists():
        st.success("기준정보: 내장본 사용 중")
    else:
        st.warning("내장 기준정보 없음 — 업로드 필요")
    up_master = st.file_uploader("기준정보 xlsx 업로드(갱신)", type=["xlsx"], key="inv_master")
    master_path = _save(up_master, "_master.xlsx") if up_master else (
        str(DEFAULT_MASTER) if DEFAULT_MASTER.exists() else None)
    st.divider()
    st.caption("물품담당자(개인정보) — 유통기한 분석의 '공유여부' 자동채움용. 공개 레포 대신 Supabase 저장.")
    up_dam = st.file_uploader("물품담당자 xlsx 업로드(갱신)", type=["xlsx"], key="inv_dam")
    _dam_src, _dam_n = _setup_담당자(up_dam)
    if _dam_n:
        st.success(f"담당자 {_dam_n}명 적용 ({_dam_src})")
    else:
        st.warning("담당자 미적용 — 공유여부 자동채움 생략")
    st.divider()
    threshold = st.slider("유통기한 분석 잔존율 기준", 0.0, 1.0, 0.5, 0.05)
    today = st.date_input("기준일자", value=date.today())

# ---------- 입력 ----------
st.subheader("① ERP 재고조회 파일 업로드")
stock_up = st.file_uploader("로케이션별 재고조회 xlsx", type=["xlsx", "xls"], key="inv_stock")
if not stock_up:
    st.info("재고조회 파일을 업로드하세요.")
    st.stop()
stock_path = _save(stock_up, "_stock.xlsx")
if not master_path:
    st.error("기준정보가 없습니다. 사이드바에서 업로드하세요.")
    st.stop()

st.subheader("② 분석 선택 → 실행")
ACTIONS = {
    "이중적치 분석 (보충오류 점검)": ("analyze_and_save", "이중적치"),
    "재고지 — 1단 전체": ("edit_재고지_1단_전체", "재고지_1단전체"),
    "재고지 — 2~6단": ("edit_재고지_2_6단", "재고지_2_6단"),
    "유통기한 분석 — OV5 (잔존율 ≤ 기준)": ("analyze_ov5_expiry", "OV5유통기한"),
    "유통기한 분석 — OV6 (잔존율 ≤ 기준)": ("analyze_ov6_expiry", "OV6유통기한"),
}
choice = st.selectbox("분석 항목", list(ACTIONS.keys()))
fn_name, out_tag = ACTIONS[choice]

if st.button(f"▶ {choice} 실행", type="primary", use_container_width=True):
    logs: list[str] = []

    def log(*a):
        logs.append(" ".join(str(x) for x in a))

    out_path = TMP / f"{out_tag}_{date.today():%Y%m%d}.xlsx"
    fn = getattr(core, fn_name)
    with st.spinner("분석 중..."):
        try:
            if fn_name.endswith("_expiry"):
                fn(stock_path, master_path, str(out_path),
                   threshold=float(threshold), today=today, log=log)
            else:
                fn(stock_path, master_path, str(out_path), log=log)
        except Exception as e:
            st.error(f"오류: {e}")
            with st.expander("로그"):
                st.code("\n".join(logs) or str(e))
            st.stop()
    if logs:
        with st.expander("처리 로그", expanded=False):
            st.code("\n".join(logs))
    if out_path.exists():
        data = out_path.read_bytes()
        st.success(f"완료: {out_path.name} ({len(data)/1024:.0f} KB)")
        st.download_button("📥 결과 다운로드", data, out_path.name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    else:
        st.warning("결과 파일이 생성되지 않았습니다. 로그를 확인하세요.")

st.caption("ⓘ 담당자(공유여부 자동채움)는 Supabase(비공개)에 저장 — 결과 파일에 반영됩니다. 사이드바에서 갱신 가능.")
