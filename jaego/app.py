"""
유통기한 재고 출고 모니터
- 로컬 실행: watchlist.json 사용
- 클라우드 배포: Supabase DB 사용 (st.secrets에 SUPABASE_URL, SUPABASE_KEY 설정)
"""

import streamlit as st
import pandas as pd
import json
import re
from pathlib import Path

# ─── Supabase 감지 ─────────────────────────────────────────────────────────
# 우선순위: JAEGO_SUPABASE_* (재고모니터 전용/기존 프로젝트) > SUPABASE_* (통합 공용)
def _sb_url():
    try:
        return st.secrets.get("JAEGO_SUPABASE_URL") or st.secrets.get("SUPABASE_URL")
    except Exception:
        return None


def _sb_key():
    try:
        return st.secrets.get("JAEGO_SUPABASE_KEY") or st.secrets.get("SUPABASE_KEY")
    except Exception:
        return None


def _supabase_configured() -> bool:
    return bool(_sb_url()) and bool(_sb_key())

USE_SUPABASE = _supabase_configured()

if USE_SUPABASE:
    from supabase import create_client

    @st.cache_resource
    def _sb():
        return create_client(_sb_url(), _sb_key())

# ─── 컬럼 인덱스 ──────────────────────────────────────────────────────────
COL_품목코드 = 4
COL_품목명   = 5
COL_유통기한 = 9
COL_출고예정 = 14
COL_LOCK    = 23

APP_DIR        = Path(__file__).parent
WATCHLIST_FILE = APP_DIR / "watchlist.json"


# ─── 워치리스트 CRUD ────────────────────────────────────────────────────────

WATCHLIST_TABLE_MISSING = False  # watchlist 테이블 미생성 감지 플래그


def load_watchlist() -> list[dict]:
    global WATCHLIST_TABLE_MISSING
    if USE_SUPABASE:
        try:
            resp = _sb().table("watchlist").select("*").order("id").execute()
            return resp.data or []
        except Exception:
            WATCHLIST_TABLE_MISSING = True   # 테이블 없음 → 빈 목록(안내는 UI에서)
            return []
    # 로컬 폴백
    if WATCHLIST_FILE.exists():
        with open(WATCHLIST_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def add_item(code: str, expiry: str) -> bool:
    """추가. 중복이면 False 반환."""
    wl = load_watchlist()
    if any(str(r["code"]) == code and str(r["expiry"]) == expiry for r in wl):
        return False
    if USE_SUPABASE:
        try:
            _sb().table("watchlist").insert({"code": code, "expiry": expiry}).execute()
        except Exception:
            return False
    else:
        wl.append({"code": code, "expiry": expiry})
        _save_local(wl)
    return True


def delete_item(row_id) -> None:
    if USE_SUPABASE:
        try:
            _sb().table("watchlist").delete().eq("id", row_id).execute()
        except Exception:
            pass
    else:
        wl = load_watchlist()
        wl = [r for r in wl if r.get("id") != row_id]
        _save_local(wl)


def _save_local(wl: list) -> None:
    # 로컬 전용 - id가 없으면 부여
    for i, r in enumerate(wl):
        if "id" not in r:
            r["id"] = i + 1
    with open(WATCHLIST_FILE, "w", encoding="utf-8") as f:
        json.dump(wl, f, ensure_ascii=False, indent=2)


# ─── 수량 파싱 ─────────────────────────────────────────────────────────────

def has_qty(val) -> bool:
    if pd.isna(val) or str(val).strip() in ("", "nan", "None", "NaN"):
        return False
    tokens = re.findall(r"[\d,]+", str(val))
    return any(float(t.replace(",", "")) > 0 for t in tokens)


def to_exp_str(raw) -> str | None:
    if pd.isna(raw) or str(raw).strip() in ("", "nan"):
        return None
    try:
        return str(int(float(str(raw).strip())))
    except Exception:
        return str(raw).strip()


# ─── 분석 로직 ─────────────────────────────────────────────────────────────

def analyze_item(df: pd.DataFrame, code: str, target_exp: str):
    """(품목명, status, color, note, 출고중_유통기한_목록) 반환."""
    mask    = df.iloc[:, COL_품목코드].astype(str).str.strip() == code.strip()
    item_df = df[mask]

    if item_df.empty:
        return "-", "미발견", "", "파일에서 품목을 찾을 수 없음", []

    품목명 = str(item_df.iloc[0, COL_품목명])

    exp_info: dict[str, dict] = {}
    for _, row in item_df.iterrows():
        exp = to_exp_str(row.iloc[COL_유통기한])
        if exp is None:
            continue
        if exp not in exp_info:
            exp_info[exp] = {"has_lock_free": False, "in_출고": False}
        if not has_qty(row.iloc[COL_LOCK]):
            exp_info[exp]["has_lock_free"] = True
        if has_qty(row.iloc[COL_출고예정]):
            exp_info[exp]["in_출고"] = True

    all_exps    = sorted(exp_info.keys())
    출고중_목록 = [e for e in all_exps if exp_info[e]["in_출고"]]

    if target_exp not in all_exps:
        return 품목명, "미발견", "", f"등록 유통기한({target_exp})이 파일에 없음", 출고중_목록

    idx         = all_exps.index(target_exp)
    before_exps = all_exps[:idx]
    after_exps  = all_exps[idx + 1:]
    t           = exp_info[target_exp]
    직전        = before_exps[-1] if before_exps else None
    이후_출고   = [e for e in after_exps if exp_info[e]["in_출고"]]

    if 이후_출고:
        return 품목명, "위험", "red", f"이후 유통기한 출고진행 중: {', '.join(이후_출고)}", 출고중_목록

    if t["in_출고"]:
        return 품목명, "주의(락 점검)", "orange", f"동일 유통기한({target_exp}) 출고진행 중", 출고중_목록

    if not t["has_lock_free"] and 직전 and exp_info[직전]["in_출고"]:
        return (
            품목명, "주의(락 점검)", "orange",
            f"직전 유통기한({직전}) 출고진행 중  /  락 없는 동일 유통기한 재고 없음",
            출고중_목록,
        )

    note = "락 없는 동일 유통기한 재고 있음" if t["has_lock_free"] else "출고진행 없음"
    return 품목명, "정상", "", note, 출고중_목록


# ─── UI ───────────────────────────────────────────────────────────────────

try:  # 단독 실행 시에만 (통합 Home.py에서 실행되면 무시)
    st.set_page_config(page_title="유통기한 재고 모니터", layout="wide", page_icon="📦")
except Exception:
    pass
st.title("📦 유통기한 재고 출고 모니터")

# watchlist 테이블 미생성 감지 → 1회 안내 (테이블 만들면 사라짐)
if USE_SUPABASE:
    load_watchlist()  # 플래그 세팅용
    if WATCHLIST_TABLE_MISSING:
        st.warning(
            "⚠ `watchlist` 테이블이 없습니다. 방법 2가지:\n\n"
            "**A) 기존 재고모니터 데이터 그대로 쓰기** — 앱 Secrets에 기존 재고모니터 "
            "Supabase 값을 `JAEGO_SUPABASE_URL` / `JAEGO_SUPABASE_KEY` 로 추가 (이관 불필요)\n\n"
            "**B) 이 통합 Supabase에 새로 만들기** — SQL Editor에서 1회 실행:\n"
            "```sql\ncreate table watchlist (\n  id serial primary key,\n"
            "  code varchar(50) not null,\n  expiry varchar(8) not null,\n"
            "  unique(code, expiry)\n);\n```")

# ── 품목 등록 (숨김 토글) ──────────────────────────────────────────────────
col_cb, col_desc = st.columns([2, 5])
with col_cb:
    show_register = st.checkbox("⚙️ 품목 등록/삭제 열기", value=False)
with col_desc:
    st.markdown("**신규 파우치 품목코드 유통기한 등록**")

if show_register:
    with st.container(border=True):
        st.subheader("🔖 품목 등록 / 삭제")
        left, right = st.columns([1, 1])

        # 등록
        with left:
            with st.form("add_form", clear_on_submit=True):
                inp_code   = st.text_input("품목코드", placeholder="예: 2061438")
                inp_expiry = st.text_input("유통기한 (YYYYMMDD)", placeholder="예: 20270827")
                submitted  = st.form_submit_button("➕ 등록", use_container_width=True)

            if submitted:
                if inp_code and inp_expiry:
                    ok = add_item(inp_code.strip(), inp_expiry.strip())
                    if ok:
                        st.success(f"등록 완료: **{inp_code.strip()}** / `{inp_expiry.strip()}`")
                        st.rerun()
                    else:
                        st.warning("이미 등록된 항목입니다")
                else:
                    st.error("품목코드와 유통기한을 모두 입력해주세요")

        # 삭제
        with right:
            wl = load_watchlist()
            if not wl:
                st.info("등록된 품목이 없습니다")
            else:
                st.markdown("**등록 목록** (🗑 클릭 시 삭제)")
                for item in wl:
                    c1, c2 = st.columns([5, 1])
                    c1.markdown(f"**{item['code']}** &nbsp; `{item['expiry']}`")
                    if c2.button("🗑", key=f"del_{item['id']}"):
                        delete_item(item["id"])
                        st.rerun()

st.divider()

# ── 파일 업로드 및 분석 ────────────────────────────────────────────────────
st.header("📂 파일 업로드")
uploaded = st.file_uploader(
    "로케이션별 재고조회 Excel 파일 (.xlsx)",
    type=["xlsx", "xls"],
)

if uploaded:
    wl = load_watchlist()
    if not wl:
        st.warning("⚠️ 등록된 품목이 없습니다. 위 '품목 등록/삭제 열기'를 체크하여 품목을 등록해주세요.")
    else:
        df = pd.read_excel(uploaded, dtype=str, header=0)
        # ── 창고 선택: 전체창고 파일도 창고별로 분리 (섞이지 않음) ──
        _whs_all = sorted({str(v).strip() for v in df.iloc[:, 0].dropna()
                           if str(v).strip()})
        # 기본 체크: IC930·IC920·IC906 (파일에 있는 것만)
        _default = [w for w in ("IC930", "IC920", "IC906") if w in _whs_all] or _whs_all
        if len(_whs_all) > 1:
            sel_whs = st.multiselect(
                "🏬 창고 선택 (창고별로 분리 표시 — 보고 싶은 창고만 체크, 섞이지 않음)",
                _whs_all, default=_default)
        else:
            sel_whs = _whs_all
        if not sel_whs:
            sel_whs = _whs_all
        with st.spinner("분석 중..."):
            rows = []
            for wh in sel_whs:
                df_wh = df[df.iloc[:, 0].astype(str).str.strip() == wh]
                for item in wl:
                    품목명, status, color, note, 출고중 = analyze_item(
                        df_wh, str(item["code"]), str(item["expiry"])
                    )
                    rows.append(
                        {
                            "창고":            wh,
                            "품목코드":         str(item["code"]),
                            "품목명":           품목명,
                            "등록 유통기한":     str(item["expiry"]),
                            "출고진행 유통기한":  ", ".join(출고중) if 출고중 else "-",
                            "최신 출고 유통기한": max(출고중) if 출고중 else "-",
                            "상태":            status,
                            "비고":            note,
                            "_color":          color,
                        }
                    )

        result_df  = pd.DataFrame(rows)
        display_df = result_df.drop(columns=["_color"])
        BG = {"red": "#FFCCCC", "orange": "#FFE0B2", "": "#FFFFFF"}

        def row_style(row):
            bg = BG.get(result_df.at[row.name, "_color"], "#FFFFFF")
            return [f"background-color: {bg}"] * len(row)

        styled = display_df.style.apply(row_style, axis=1).map(
            lambda _: "font-weight: bold", subset=["등록 유통기한"]
        )

        # 요약 카운트
        cnt_red    = (result_df["_color"] == "red").sum()
        cnt_orange = (result_df["_color"] == "orange").sum()
        cnt_ok     = (result_df["_color"] == "").sum()

        c1, c2, c3 = st.columns(3)
        c1.metric("🔴 위험", cnt_red,    help="이후 유통기한 출고진행 중")
        c2.metric("🟠 주의", cnt_orange, help="동일/직전 유통기한 출고진행 중")
        c3.metric("✅ 정상", cnt_ok,     help="락 없는 동일 유통기한 재고 있음")

        st.subheader("📊 분석 결과")
        st.dataframe(styled, use_container_width=True, height=min(600, 60 + len(rows) * 38))

        st.markdown("""
---
**범례**

| 색상 | 의미 |
|------|------|
| 🔴 빨간 (위험) — 역순출고 우려 | 등록 유통기한보다 **이후** 유통기한이 출고 진행 중 |
| 🟠 주황 (주의) — 신규파우치 출고 임박 (락해제 점검 필요) | **동일** 유통기한 출고 진행 중  ·  또는  ·  락 없는 동일 유통기한 재고 없고 **직전** 유통기한 출고 진행 중 |
| ⬜ 색 없음 (정상) | 락 없는 동일 유통기한 재고 있음 |
        """)

        st.markdown("""
<div style="
    background-color: #FFF3CD;
    border: 3px solid #FF0000;
    border-radius: 8px;
    padding: 18px 24px;
    margin-top: 16px;
    text-align: center;
">
    <span style="font-size: 22px; font-weight: 900; color: #CC0000;">
        ⚠️ 락 해제 시 메일 발송 필수
    </span><br>
    <span style="font-size: 17px; font-weight: 700; color: #333;">
        수신 : 영업팀 전부 &nbsp;|&nbsp; 참조 : SCM
    </span>
</div>
""", unsafe_allow_html=True)
