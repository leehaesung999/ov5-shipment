# -*- coding: utf-8 -*-
"""통합센터 재고 분석 엔진 — '실사지 출력'/'점검' 페이지 공용.

두 페이지(sheets.py / checks.py)가 이 모듈의 render(action_keys, title, caption)를
각자 기능 subset으로 호출한다. 기준정보·담당자는 Supabase(app_settings)에서 공유.
"""
from __future__ import annotations

import base64
import io
import json
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import inv_core as core  # noqa: E402  (OV5의 core 패키지와 이름 충돌 방지)

DATA = HERE / "data"
DEFAULT_MASTER = DATA / "기준정보.xlsx"
TMP = HERE / "_tmp"
TMP.mkdir(exist_ok=True)
MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
LOC_DEFAULT = str(DATA / "지정로케이션.xlsx")  # 내장 기본본 (Supabase 갱신 없을 때 사용)


def _save(upload, name) -> str:
    p = TMP / name
    p.write_bytes(upload.getvalue())
    return str(p)


# ---------- 확인 체크 (영속·누적) — 기존 app_settings(jsonb)에 저장 ----------
NONLOCK_CHECK_KEY = "inventory_nonlock_checks"
_CHECK_FILE = HERE / "_checks.json"  # 로컬 폴백


def _load_checks(settings_key) -> dict:
    """{rowkey: True} — 확인 완료 항목."""
    if _SB_OK:
        try:
            r = _sb().table("app_settings").select("value").eq("key", settings_key).execute()
            v = (r.data[0].get("value") or {}) if r.data else {}
            return {k: True for k, val in v.items() if val}
        except Exception:
            pass
    try:
        if _CHECK_FILE.exists():
            return json.loads(_CHECK_FILE.read_text(encoding="utf-8")).get(settings_key, {})
    except Exception:
        pass
    return {}


def _save_checks(settings_key, mapping: dict) -> None:
    if _SB_OK:
        try:
            _sb().table("app_settings").upsert(
                {"key": settings_key, "value": mapping}, on_conflict="key").execute()
            return
        except Exception:
            pass
    try:
        alld = {}
        if _CHECK_FILE.exists():
            alld = json.loads(_CHECK_FILE.read_text(encoding="utf-8"))
        alld[settings_key] = mapping
        _CHECK_FILE.write_text(json.dumps(alld, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def _show_nonlock_checks(rows):
    """비Lock 유통기한 결과를 확인 체크칸과 함께 표시. 체크는 누적 저장 → 다음엔 재확인 불필요."""
    if not rows:
        st.info("결과 데이터가 없습니다 (비Lock 임박 없음).")
        return

    def _mkkey(r):  # (품목코드, 유통기한) 조합 = 합산 단위, 실행마다 안정적
        return f'{r.get("품목코드", "")}|{r.get("유통기한_원본", "")}'

    saved = _load_checks(NONLOCK_CHECK_KEY)
    keys_all = [_mkkey(r) for r in rows]
    done_n = sum(1 for k in keys_all if saved.get(k))
    hide = st.toggle(f"✔ 확인완료({done_n}) 숨기기 — 다음엔 재확인 안 함",
                     value=False, key="nl_hide")

    view, vkeys = [], []
    for r, k in zip(rows, keys_all):
        if hide and saved.get(k):
            continue
        view.append(r)
        vkeys.append(k)
    if not view:
        st.success("표시할 미확인 항목이 없습니다. (모두 확인 완료) 🎉")
        return

    disp = pd.DataFrame([{
        "확인": saved.get(k, False),
        "품목코드": r.get("품목코드", ""),
        "품목명": r.get("품목명", ""),
        "유통기한": r.get("유통기한"),
        "수량": r.get("수량"),
        "잔존일수": r.get("잔존일수"),
        "잔존개월": r.get("잔존개월"),
        "잔존율(%)": round((r.get("잔존율") or 0) * 100, 1),
        "잔존율40도달일": r.get("잔존율40_도달일"),
        "공유여부": r.get("공유여부"),
    } for r, k in zip(view, vkeys)])

    st.caption(f"총 {len(disp):,}행 · '확인' 체크는 저장되어 다음 분석에도 누적됩니다")
    edited = st.data_editor(
        disp, key="nl_editor", hide_index=True, width='stretch',
        height=min(600, 80 + len(disp) * 35),
        column_config={"확인": st.column_config.CheckboxColumn(
            "확인", help="확인 완료 시 체크 → 저장·누적, 다음엔 재확인 불필요", default=False)},
        disabled=[c for c in disp.columns if c != "확인"])

    # 변경분만 저장 (강제 rerun 없음 → 무한루프 방지)
    changed = False
    for i, k in enumerate(vkeys):
        newv = bool(edited.iloc[i]["확인"])
        if newv != saved.get(k, False):
            if newv:
                saved[k] = True
            else:
                saved.pop(k, None)
            changed = True
    if changed:
        _save_checks(NONLOCK_CHECK_KEY, saved)


def _show_xlsx(data: bytes):
    """결과 xlsx bytes를 화면에 표로 표시 (시트별). 다운로드와 별개."""
    try:
        xls = pd.ExcelFile(io.BytesIO(data))
    except Exception as e:
        st.warning(f"미리보기 표시 실패: {e} (엑셀 다운로드는 정상입니다)")
        return
    multi = len(xls.sheet_names) > 1
    for sh in xls.sheet_names:
        try:
            d = xls.parse(sh)
        except Exception:
            continue
        if multi:
            st.markdown(f"**[{sh}]**  ·  {len(d):,}행")
        if d is None or d.empty:
            st.info("결과 데이터가 없습니다 (해당 없음).")
            continue
        st.caption(f"총 {len(d):,}행")
        st.dataframe(d, width='stretch',
                     height=min(560, 80 + len(d) * 35))


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
LOC_META_KEY = "inventory_loc_meta"
LOC_B64_KEY = "inventory_loc_b64"


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


# ---------- 지정로케이션 — Supabase 영구 저장 (기준정보와 동일 방식) ----------
def _fetch_loc_meta():
    if not _SB_OK:
        return None
    try:
        r = _sb().table("app_settings").select("value").eq("key", LOC_META_KEY).execute()
        return r.data[0].get("value") if r.data else None
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def _fetch_loc_bytes(version: str):
    try:
        r = _sb().table("app_settings").select("value").eq("key", LOC_B64_KEY).execute()
        if r.data:
            return base64.b64decode(r.data[0]["value"])
    except Exception:
        pass
    return None


def _store_loc(data: bytes, count: int) -> bool:
    if not _SB_OK:
        return False
    try:
        meta = {"updated": datetime.now().strftime("%Y-%m-%d %H:%M"), "n": count}
        _sb().table("app_settings").upsert(
            {"key": LOC_META_KEY, "value": meta}, on_conflict="key").execute()
        _sb().table("app_settings").upsert(
            {"key": LOC_B64_KEY, "value": base64.b64encode(data).decode()},
            on_conflict="key").execute()
        return True
    except Exception:
        return False


# (버튼라벨, 설명 hint, key, 결과파일 태그) — 전체 8개. 페이지별로 subset만 노출.
ALL_ACTIONS = [
    ("🔀 이중적치 분석", "같은 로케이션에 잔량 혼재 검출 (Lock·OV·정파렛트 제외)", "이중적치", "이중적치"),
    ("📋 일일 재고실사지", "지정 로케이션 일일 실사지 · 일일입력 있으면 차이수량 반영(대시보드 다운)", "일일실사", "일일재고실사지"),
    ("📋 1단 재고실사지", "1단 전체 로케이션 실사지", "1단전체", "1단재고실사지"),
    ("📅 토요일 실사지 (쿠팡+컬리)", "쿠팡 필수 + 컬리 선택 · 두 제품별리스트의 출하Inv=IC930 품목 합집합", "토요일", "토요일실사지"),
    ("📋 재고지 (2~6단)", "고단(2-6단) 재고지", "2_6단", "재고지_2_6단"),
    ("⏳ OV5 하프도달 점검", "OV5 재고 중 잔존율 ≤ 기준(50%) 유통기한 임박", "ov5", "OV5유통기한"),
    ("⏳ OV6 하프도달 점검", "OV6 재고 중 잔존율 ≤ 기준(50%) 유통기한 임박", "ov6", "OV6유통기한"),
    ("⏳ 비Lock 유통기한 점검", "Lock 무관 전체 재고 중 잔존율 ≤ 기준 임박", "nonlock", "비Lock유통기한"),
]


def render(action_keys, title: str, caption: str, preview: bool = False):
    """action_keys(집합)에 해당하는 기능만 노출하는 재고 분석 페이지 렌더.

    preview=True 이면 실행 결과를 엑셀 다운로드와 함께 화면에 표로 표시한다.
    """
    try:
        st.set_page_config(page_title="통합센터 재고 분석기", layout="wide")
    except Exception:
        pass

    st.title(title)
    st.caption(caption)

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
        st.caption("지정로케이션 — 일일/토요일 실사지 대상 위치(1단). 바뀌면 업로드하면 저장·공유됩니다.")
        up_loc = st.file_uploader("지정로케이션 xlsx 업로드(갱신)", type=["xlsx"], key="inv_loc")
        _lmeta = None
        if up_loc:
            _loc_tmp = _save(up_loc, "_loc.xlsx")
            try:
                _ln = len(core.load_location_list(_loc_tmp))
            except Exception:
                _ln = 0
            if _store_loc(up_loc.getvalue(), _ln):
                st.success(f"지정로케이션 {_ln}건 저장됨 (모두에게 공유·영구)")
                _fetch_loc_bytes.clear()  # 캐시 무효화
            else:
                st.info(f"지정로케이션 {_ln}건 (이 세션에만 적용 — 로컬/비공유 모드)")
            loc_list_path = _loc_tmp
            _lmeta = _fetch_loc_meta()
        else:
            _lmeta = _fetch_loc_meta()
            if _lmeta:
                _lb = _fetch_loc_bytes(str(_lmeta.get("updated", "")))
                if _lb:
                    _lp = TMP / "_loc_sb.xlsx"
                    _lp.write_bytes(_lb)
                    loc_list_path = str(_lp)
                else:
                    loc_list_path = LOC_DEFAULT if Path(LOC_DEFAULT).exists() else None
            else:
                loc_list_path = LOC_DEFAULT if Path(LOC_DEFAULT).exists() else None
        if _lmeta:
            st.info(f"📅 지정로케이션 마지막 업데이트: **{_lmeta.get('updated', '?')}** "
                    f"({_lmeta.get('n', '?')}건)")
        elif Path(LOC_DEFAULT).exists():
            st.caption("지정로케이션: 내장 기본본 — 업로드하면 날짜가 갱신됩니다")
        else:
            st.warning("지정로케이션 없음 — 업로드 필요")
        st.divider()
        threshold = st.slider("유통기한 분석 잔존율 기준", 0.0, 1.0, 0.5, 0.05)
        today = st.date_input("기준일자", value=date.today())

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

    if "inv_results" not in st.session_state:
        st.session_state.inv_results = {}

    def run_action(key, out_tag, up_daily=None, up_prod=None, up_prod2=None):
        logs: list[str] = []

        def log(*a):
            logs.append(" ".join(str(x) for x in a))

        op = TMP / f"{out_tag}_{date.today():%Y%m%d}.xlsx"
        thr = float(threshold)
        # 유통기한 함수는 내부에서 datetime과 빼기 → date를 datetime으로 변환
        _today = datetime.combine(today, datetime.min.time()) if isinstance(today, date) else today
        diff_qty = None
        extra = {}
        try:
            if key in ("일일실사", "토요일") and not loc_list_path:
                return {"error": "지정로케이션 파일이 없습니다. 사이드바 '지정로케이션'에서 업로드하세요.",
                        "logs": logs}
            if up_daily:
                diff_qty = core.load_차이수량_from_일일입력(_save(up_daily, "_daily.xlsx"))
                log(f"일일입력 차이수량 {len(diff_qty):,}건 반영")
            if key == "이중적치":
                core.analyze_and_save(stock_path, master_path, str(op), log=log)
            elif key == "일일실사":
                core.edit_재고지_1단(stock_path, master_path, loc_list_path, str(op),
                                     diff_qty=diff_qty, log=log)
            elif key == "1단전체":
                core.edit_재고지_1단_전체(stock_path, master_path, str(op), log=log)
            elif key == "토요일":
                if not up_prod:
                    return {"error": "쿠팡 제품별리스트 파일을 이 버튼 바로 아래에서 업로드하세요.", "logs": logs}
                code_set = core.load_품목코드_from_제품별리스트(_save(up_prod, "_prod.xlsx"))
                log(f"쿠팡 출하 대상 품목 {len(code_set):,}건")
                if up_prod2:
                    code_set2 = core.load_품목코드_from_제품별리스트(_save(up_prod2, "_prod2.xlsx"))
                    log(f"컬리 출하 대상 품목 {len(code_set2):,}건")
                    before = len(code_set)
                    code_set = code_set | code_set2
                    log(f"합집합: {len(code_set):,}건 (쿠팡 {before} + 컬리 {len(code_set2)} − 중복 {before + len(code_set2) - len(code_set)})")
                core.edit_재고지_1단(stock_path, master_path, loc_list_path, str(op),
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
                _stat = core.analyze_nonlock_expiry(stock_path, master_path, str(op),
                                                    threshold=thr, today=_today, log=log)
                if isinstance(_stat, dict):
                    extra["rows"] = _stat.get("rows")
        except Exception as e:
            return {"error": str(e), "logs": logs}
        if not op.exists():
            return {"error": "결과 파일이 생성되지 않았습니다.", "logs": logs}
        return {"data": op.read_bytes(), "name": op.name, "logs": logs, **extra}

    # 기능별 추가 입력 (해당 버튼 바로 아래에서 업로드) — (kind, 라벨, 필수여부)
    AUX = {
        "일일실사": [("daily", "📎 일일입력 xlsx — 차이수량 반영 (선택)", False)],
        "토요일":   [("prod", "📎 쿠팡 제품별리스트 xlsx — 출하 대상 품목 (필수)", True),
                     ("prod2", "📎 컬리 제품별리스트 xlsx — 출하 대상 품목 (선택, 있으면 합집합)", False),
                     ("daily", "📎 일일입력 xlsx — 차이수량 반영 (선택)", False)],
    }
    _amap = {a[2]: a for a in ALL_ACTIONS}
    actions = [_amap[k] for k in action_keys if k in _amap]  # 전달된 순서 그대로 표시
    for label, hint, key, out_tag in actions:
        c1, c2 = st.columns([3, 2])
        with c1:
            clicked = st.button(f"▶ {label}", key=f"btn_{key}", width='stretch')
            st.caption(hint)
            aux = {}
            for kind, lbl, required in AUX.get(key, []):
                f = st.file_uploader(lbl, type=["xlsx"], key=f"aux_{key}_{kind}")
                aux[kind] = f
                if required and not f:
                    st.caption("⚠ 이 파일을 올려야 실행됩니다")
        with c2:
            if clicked:
                with st.spinner(f"{label} 실행 중..."):
                    st.session_state.inv_results[key] = run_action(
                        key, out_tag, up_daily=aux.get("daily"),
                        up_prod=aux.get("prod"), up_prod2=aux.get("prod2"))
            res = st.session_state.inv_results.get(key)
            if res:
                if res.get("error"):
                    st.error(res["error"])
                else:
                    st.download_button(
                        f"📥 {res['name']}", res["data"], res["name"],
                        mime=MIME, key=f"dl_{key}", width='stretch')
        # 결과 화면 미리보기 (점검 페이지) — 엑셀 다운로드는 위에 그대로 유지
        if preview:
            res = st.session_state.inv_results.get(key)
            if res and not res.get("error") and res.get("data"):
                with st.expander(f"📄 {label} — 결과 화면 보기", expanded=True):
                    if key == "nonlock" and res.get("rows") is not None:
                        _show_nonlock_checks(res["rows"])
                    else:
                        _show_xlsx(res["data"])

    st.caption("ⓘ 각 버튼은 독립 실행 — 원하는 것만 눌러 결과를 받으세요. "
               "담당자(공유여부 자동채움)는 Supabase(비공개)에 저장돼 결과에 반영됩니다.")
