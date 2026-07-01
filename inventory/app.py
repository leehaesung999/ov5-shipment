# -*- coding: utf-8 -*-
"""통합센터 재고 분석기 — Streamlit 웹 (통합앱 페이지).

ERP 재고조회 업로드 → 이중적치 분석 / 재고지 생성 / 유통기한(OV5·OV6) 분석 → 다운로드.
기준정보(물품 마스터)는 레포에 포함된 data/기준정보.xlsx 사용(또는 업로드).
"""
from __future__ import annotations

import base64
import sys
from datetime import date, datetime
from pathlib import Path

import streamlit as st

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import inv_core as core  # noqa: E402  (OV5의 core 패키지와 이름 충돌 방지)

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


# ---------- 기준정보(마스터) — Supabase 영구 저장 + 마지막 업데이트 날짜 ----------
MASTER_META_KEY = "inventory_master_meta"
MASTER_B64_KEY = "inventory_master_b64"


def _fetch_master_meta():
    if not _SB_OK:
        return None
    try:
        r = _sb().table("app_settings").select("value").eq("key", MASTER_META_KEY).execute()
        return r.data[0].get("value") if r.data else None
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def _fetch_master_bytes(version: str):
    try:
        r = _sb().table("app_settings").select("value").eq("key", MASTER_B64_KEY).execute()
        if r.data:
            return base64.b64decode(r.data[0]["value"])
    except Exception:
        pass
    return None


def _store_master(data: bytes, count: int) -> bool:
    if not _SB_OK:
        return False
    try:
        meta = {"updated": datetime.now().strftime("%Y-%m-%d %H:%M"), "n": count}
        _sb().table("app_settings").upsert(
            {"key": MASTER_META_KEY, "value": meta}, on_conflict="key").execute()
        _sb().table("app_settings").upsert(
            {"key": MASTER_B64_KEY, "value": base64.b64encode(data).decode()},
            on_conflict="key").execute()
        return True
    except Exception:
        return False


# ---------- 사이드바: 기준정보 (접기) ----------
with st.sidebar.expander("⚙️ 기준정보 / 옵션 — 클릭해서 열기", expanded=False):
    up_master = st.file_uploader("기준정보 xlsx 업로드(갱신)", type=["xlsx"], key="inv_master")
    _mmeta = None
    if up_master:
        master_path = _save(up_master, "_master.xlsx")
        try:
            _mn = len(core.load_master(master_path))
        except Exception:
            _mn = 0
        if _store_master(up_master.getvalue(), _mn):
            st.success(f"기준정보 {_mn}품목 업데이트 저장됨 (모두에게 공유·영구)")
            _fetch_master_bytes.clear()  # 캐시 무효화
        else:
            st.info(f"기준정보 {_mn}품목 (이 세션에만 적용 — 로컬/비공유 모드)")
        _mmeta = _fetch_master_meta()
    else:
        _mmeta = _fetch_master_meta()
        if _mmeta:
            _mb = _fetch_master_bytes(str(_mmeta.get("updated", "")))
            if _mb:
                _p = TMP / "_master_sb.xlsx"
                _p.write_bytes(_mb)
                master_path = str(_p)
            else:
                master_path = str(DEFAULT_MASTER) if DEFAULT_MASTER.exists() else None
        else:
            master_path = str(DEFAULT_MASTER) if DEFAULT_MASTER.exists() else None
    # 마지막 업데이트 날짜 표시
    if _mmeta:
        st.info(f"📅 기준정보 마지막 업데이트: **{_mmeta.get('updated', '?')}** "
                f"({_mmeta.get('n', '?')}품목)")
    elif DEFAULT_MASTER.exists():
        _dm = datetime.fromtimestamp(DEFAULT_MASTER.stat().st_mtime)
        st.caption(f"기준정보: 내장 기본본 ({_dm:%Y-%m-%d} 기준) — 업로드하면 날짜가 갱신됩니다")
    else:
        st.warning("기준정보 없음 — 업로드 필요")
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
    st.divider()
    st.caption("실사지 보조 입력 (선택)")
    up_daily = st.file_uploader("일일입력 xlsx (차이수량 반영)", type=["xlsx"], key="inv_daily")
    up_prod = st.file_uploader("제품별리스트 xlsx (토요일 실사지용)", type=["xlsx"], key="inv_prod")

# ---------- 입력 ----------
st.subheader("① ERP 재고조회 파일 업로드")
stock_up = st.file_uploader("로케이션별 재고조회 xlsx", type=["xlsx", "xls"], key="inv_stock")
if not stock_up:
    st.info("재고조회 파일을 업로드하세요.")
    st.session_state.pop("inv_results", None)  # 파일 없으면 이전 결과 제거
    st.session_state.pop("inv_stock_sig", None)
    st.stop()
stock_path = _save(stock_up, "_stock.xlsx")
# 재고파일이 바뀌면(삭제 후 재업로드 포함) 이전 분석 다운로드 버튼 초기화 — 혼동 방지
_stock_sig = f"{stock_up.name}|{stock_up.size}"
if st.session_state.get("inv_stock_sig") != _stock_sig:
    st.session_state.inv_results = {}
    st.session_state.inv_stock_sig = _stock_sig
if not master_path:
    st.error("기준정보가 없습니다. 사이드바에서 업로드하세요.")
    st.stop()

st.subheader("② 분석 실행 — 원하는 기능 버튼 클릭 (각각 독립 실행)")
MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
LOC_LIST = str(DATA / "지정로케이션.xlsx")

# (버튼라벨, 설명 hint, key, 결과파일 태그)
ACTIONS = [
    ("🔀 이중적치 분석", "같은 로케이션에 잔량 혼재 검출 (Lock·OV·정파렛트 제외)", "이중적치", "이중적치"),
    ("📋 일일 재고실사지", "지정 로케이션 일일 실사지 · 일일입력 있으면 차이수량 반영(대시보드 다운)", "일일실사", "일일재고실사지"),
    ("📋 1단 재고실사지", "1단 전체 로케이션 실사지", "1단전체", "1단재고실사지"),
    ("📅 토요일 실사지 (쿠팡 출고)", "출하Inv=IC930 품목만 · 제품별리스트 업로드 필요(출하체크·대시보드 다운)", "토요일", "토요일실사지"),
    ("📋 재고지 (2~6단)", "고단(2-6단) 재고지", "2_6단", "재고지_2_6단"),
    ("⏳ OV5 하프도달 점검", "OV5 재고 중 잔존율 ≤ 기준(50%) 유통기한 임박", "ov5", "OV5유통기한"),
    ("⏳ OV6 하프도달 점검", "OV6 재고 중 잔존율 ≤ 기준(50%) 유통기한 임박", "ov6", "OV6유통기한"),
    ("⏳ 비Lock 유통기한 점검", "Lock 무관 전체 재고 중 잔존율 ≤ 기준 임박", "nonlock", "비Lock유통기한"),
]

if "inv_results" not in st.session_state:
    st.session_state.inv_results = {}


def run_action(key, out_tag):
    logs: list[str] = []

    def log(*a):
        logs.append(" ".join(str(x) for x in a))

    op = TMP / f"{out_tag}_{date.today():%Y%m%d}.xlsx"
    thr = float(threshold)
    # 유통기한 함수는 내부에서 datetime과 빼기 → date를 datetime으로 변환
    _today = datetime.combine(today, datetime.min.time()) if isinstance(today, date) else today
    diff_qty = None
    try:
        if up_daily:
            diff_qty = core.load_차이수량_from_일일입력(_save(up_daily, "_daily.xlsx"))
            log(f"일일입력 차이수량 {len(diff_qty):,}건 반영")
        if key == "이중적치":
            core.analyze_and_save(stock_path, master_path, str(op), log=log)
        elif key == "일일실사":
            core.edit_재고지_1단(stock_path, master_path, LOC_LIST, str(op),
                                 diff_qty=diff_qty, log=log)
        elif key == "1단전체":
            core.edit_재고지_1단_전체(stock_path, master_path, str(op), log=log)
        elif key == "토요일":
            if not up_prod:
                return {"error": "제품별리스트 파일을 사이드바에서 업로드하세요.", "logs": logs}
            code_set = core.load_품목코드_from_제품별리스트(_save(up_prod, "_prod.xlsx"))
            log(f"출하 대상 품목 {len(code_set):,}건")
            core.edit_재고지_1단(stock_path, master_path, LOC_LIST, str(op),
                                 code_filter=code_set, diff_qty=diff_qty, log=log)
        elif key == "2_6단":
            core.edit_재고지_2_6단(stock_path, master_path, str(op), log=log)
        elif key == "ov5":
            core.analyze_ov5_expiry(stock_path, master_path, str(op),
                                    threshold=thr, today=_today, log=log)
        elif key == "ov6":
            core.analyze_ov6_expiry(stock_path, master_path, str(op),
                                    threshold=thr, today=_today, log=log)
        elif key == "nonlock":
            core.analyze_nonlock_expiry(stock_path, master_path, str(op),
                                        threshold=thr, today=_today, log=log)
    except Exception as e:
        return {"error": str(e), "logs": logs}
    if not op.exists():
        return {"error": "결과 파일이 생성되지 않았습니다.", "logs": logs}
    return {"data": op.read_bytes(), "name": op.name, "logs": logs}


for label, hint, key, out_tag in ACTIONS:
    c1, c2 = st.columns([3, 2])
    with c1:
        clicked = st.button(f"▶ {label}", key=f"btn_{key}", use_container_width=True)
        st.caption(hint)
    with c2:
        if clicked:
            with st.spinner(f"{label} 실행 중..."):
                st.session_state.inv_results[key] = run_action(key, out_tag)
        res = st.session_state.inv_results.get(key)
        if res:
            if res.get("error"):
                st.error(res["error"])
            else:
                st.download_button(
                    f"📥 {res['name']}", res["data"], res["name"],
                    mime=MIME, key=f"dl_{key}", use_container_width=True)

st.caption("ⓘ 각 버튼은 독립 실행 — 원하는 것만 눌러 결과를 받으세요. "
           "담당자(공유여부 자동채움)는 Supabase(비공개)에 저장돼 결과에 반영됩니다.")
