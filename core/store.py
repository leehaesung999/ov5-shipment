# -*- coding: utf-8 -*-
"""저장 백엔드 추상화 — Supabase(클라우드 공유) 또는 로컬 파일.

st.secrets 에 SUPABASE_URL / SUPABASE_KEY 가 있으면 Supabase, 없으면
기존과 동일하게 config/*.xlsx · settings.json 을 사용한다(데스크톱 app.py 와 호환).

다루는 영구 데이터
  · lot_assignments : 로트 지정 거래처 (source='manual' 업로드 / 'auto' 누적·화면체크)
  · settings        : 잔존율 기준·OV 로케이션 등

master_info(유통기한월·하대)는 파일 기반 유지(레포 기본 + 세션 업로드).
"""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

try:
    import streamlit as st
except Exception:  # streamlit 밖에서 import 될 때
    st = None

ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = ROOT / "config"
SETTINGS_PATH = CONFIG_DIR / "settings.json"
LOT_PATH = CONFIG_DIR / "lot_assignments.xlsx"            # 수동
LOT_AUTO_PATH = CONFIG_DIR / "lot_assignments_auto.xlsx"  # 자동 누적

LOT_CATEGORIES = ("농협", "대리점", "FS", "MS", "급식", "소재")
SETTINGS_KEY = "ov5"


# ───────────────────────── 백엔드 감지 ─────────────────────────
def _secret(name):
    try:
        return st.secrets.get(name) if st is not None else None
    except Exception:
        return None


def use_supabase() -> bool:
    return bool(_secret("SUPABASE_URL")) and bool(_secret("SUPABASE_KEY"))


def backend_name() -> str:
    return "Supabase (클라우드 공유)" if use_supabase() else "로컬 파일"


if st is not None:
    @st.cache_resource(show_spinner=False)
    def _sb():
        from supabase import create_client
        return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
else:  # pragma: no cover
    def _sb():
        from supabase import create_client
        import os
        return create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_KEY"])


# ───────────────────────── 공용 파서 ─────────────────────────
def parse_cats(raw) -> list:
    """'농협,대리점' / '농협 대리점' → ['농협','대리점'] (유효 카테고리만)."""
    if raw is None:
        return []
    parts = str(raw).replace("/", ",").replace(" ", ",").split(",")
    out = []
    for p in parts:
        p = p.strip()
        if p in LOT_CATEGORIES and p not in out:
            out.append(p)
    return out


def _norm_key(code, ymd):
    code_s = str(int(code)) if str(code).strip() not in ("", "nan", "None") else None
    try:
        ymd_i = int(str(ymd).split(".")[0]) if ymd not in (None, "", "nan") and pd.notna(ymd) else None
    except (ValueError, TypeError):
        ymd_i = None
    return code_s, ymd_i


def _expiry_text(ymd) -> str:
    return str(ymd) if ymd not in (None, "", 0) else ""


# ───────────────────────── lot 읽기 ─────────────────────────
def _read_lots_local(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        df = pd.read_excel(str(path))
    except Exception:
        return {}
    if df.empty or "Item code" not in df.columns:
        return {}
    cat_col = next((c for c in ("카테고리", "구분", "Category") if c in df.columns), None)
    if cat_col is None:
        return {}
    exp_col = next((c for c in ("유통기한", "소비기한") if c in df.columns), None)
    out: dict = {}
    for _, r in df.iterrows():
        code, ymd = _norm_key(r.get("Item code"), r.get(exp_col) if exp_col else None)
        if code is None:
            continue
        cats = parse_cats(r.get(cat_col))
        if cats:
            out[(code, ymd)] = cats
    return out


def _read_lots_supabase(source: str) -> dict:
    resp = _sb().table("lot_assignments").select("*").eq("source", source).execute()
    out: dict = {}
    for row in (resp.data or []):
        code, ymd = _norm_key(row.get("item_code"), row.get("expiry") or None)
        if code is None:
            continue
        cats = parse_cats(row.get("category"))
        if cats:
            out[(code, ymd)] = cats
    return out


def get_manual_lots() -> dict:
    return _read_lots_supabase("manual") if use_supabase() else _read_lots_local(LOT_PATH)


def get_auto_lots() -> dict:
    return _read_lots_supabase("auto") if use_supabase() else _read_lots_local(LOT_AUTO_PATH)


# ───────────────────────── lot 쓰기 ─────────────────────────
def _map_to_rows(m: dict, names: dict | None = None) -> list[dict]:
    rows = []
    for (code, ymd), cats in m.items():
        cat_str = ",".join(cats) if isinstance(cats, (list, tuple)) else str(cats)
        rows.append({
            "Item code": int(code),
            "Item": (names or {}).get((code, ymd), ""),
            "유통기한": ymd if ymd else "",
            "카테고리": cat_str,
        })
    return rows


def _write_local(path: Path, m: dict, names: dict | None = None):
    CONFIG_DIR.mkdir(exist_ok=True)
    rows = _map_to_rows(m, names)
    sheet = "로트지정_자동" if path == LOT_AUTO_PATH else "로트지정"
    if not rows:
        path.unlink(missing_ok=True)
        return
    pd.DataFrame(rows).to_excel(str(path), index=False, sheet_name=sheet)


def _replace_supabase(source: str, m: dict, names: dict | None = None):
    sb = _sb()
    sb.table("lot_assignments").delete().eq("source", source).execute()
    rows = []
    for (code, ymd), cats in m.items():
        cat_str = ",".join(cats) if isinstance(cats, (list, tuple)) else str(cats)
        rows.append({
            "item_code": str(int(code)),
            "expiry": _expiry_text(ymd),
            "category": cat_str,
            "item_name": (names or {}).get((code, ymd), "") or "",
            "source": source,
        })
    if rows:
        sb.table("lot_assignments").insert(rows).execute()


def replace_auto_lots(m: dict, names: dict | None = None):
    """자동 lot 전체를 m 으로 교체."""
    if use_supabase():
        _replace_supabase("auto", m, names)
    else:
        _write_local(LOT_AUTO_PATH, m, names)


def upsert_auto_lots(new_map: dict, names: dict | None = None,
                     overwrite: bool = True) -> int:
    """자동 lot 에 new_map 병합. overwrite=True면 같은 키를 화면값으로 갱신,
    False면 기존에 없는 키만 추가. 반환: 반영 건수."""
    if not new_map:
        return 0
    cur = get_auto_lots()
    merged = dict(cur)
    n = 0
    for key, cats in new_map.items():
        if overwrite:
            merged[key] = list(cats)
            n += 1
        elif key not in merged:
            merged[key] = list(cats)
            n += 1
    if n:
        replace_auto_lots(merged, names)
    return n


def prune_auto_lots(valid_keys: set) -> int:
    """자동 lot 중 valid_keys 에 없는 항목 제거. 반환: 제거 건수."""
    cur = get_auto_lots()
    if not cur:
        return 0
    kept = {k: v for k, v in cur.items() if k in valid_keys}
    removed = len(cur) - len(kept)
    if removed:
        replace_auto_lots(kept)
    return removed


def clear_auto_lots():
    if use_supabase():
        _sb().table("lot_assignments").delete().eq("source", "auto").execute()
    else:
        LOT_AUTO_PATH.unlink(missing_ok=True)


def set_manual_lots_from_df(df: pd.DataFrame):
    """수동 lot 업로드(xlsx 내용)를 저장. 로컬은 파일로, Supabase는 source='manual' 교체."""
    if not use_supabase():
        CONFIG_DIR.mkdir(exist_ok=True)
        df.to_excel(str(LOT_PATH), index=False, sheet_name="로트지정")
        return
    # Supabase: df → map → 교체
    cat_col = next((c for c in ("카테고리", "구분", "Category") if c in df.columns), None)
    exp_col = next((c for c in ("유통기한", "소비기한") if c in df.columns), None)
    name_col = "Item" if "Item" in df.columns else None
    m, names = {}, {}
    if cat_col and "Item code" in df.columns:
        for _, r in df.iterrows():
            code, ymd = _norm_key(r.get("Item code"), r.get(exp_col) if exp_col else None)
            if code is None:
                continue
            cats = parse_cats(r.get(cat_col))
            if cats:
                m[(code, ymd)] = cats
                if name_col:
                    names[(code, ymd)] = r.get(name_col) or ""
    _replace_supabase("manual", m, names)


# ───────────────────────── settings ─────────────────────────
def load_settings(defaults: dict) -> dict:
    if use_supabase():
        try:
            resp = _sb().table("app_settings").select("value").eq("key", SETTINGS_KEY).execute()
            if resp.data:
                return {**defaults, **(resp.data[0].get("value") or {})}
        except Exception:
            pass
        return defaults.copy()
    # 로컬
    if SETTINGS_PATH.exists():
        try:
            return {**defaults, **json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))}
        except Exception:
            pass
    return defaults.copy()


def save_settings(s: dict):
    if use_supabase():
        _sb().table("app_settings").upsert(
            {"key": SETTINGS_KEY, "value": s}, on_conflict="key").execute()
        return
    CONFIG_DIR.mkdir(exist_ok=True)
    SETTINGS_PATH.write_text(json.dumps(s, ensure_ascii=False, indent=2), encoding="utf-8")
