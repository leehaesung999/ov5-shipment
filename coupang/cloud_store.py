# -*- coding: utf-8 -*-
"""쿠팡 지정출고 — 대상품목(기능2) 공유 저장.

st.secrets에 SUPABASE_URL/KEY가 있으면 Supabase(app_settings 테이블, key='coupang_targets')에
대상품목 목록을 저장·공유한다. 없으면 기존 로컬 대상품목.json 사용.
※ OV5와 같은 Supabase 프로젝트·app_settings 테이블을 재사용 (추가 SQL 불필요).
"""
from __future__ import annotations

import 지정출고_자동매칭 as core  # noqa: E402

try:
    import streamlit as st
except Exception:
    st = None

KEY = "coupang_targets"


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


def _clean(items):
    out = [{"code": int(t["code"]), "name": t.get("name", "")} for t in items]
    out.sort(key=lambda t: t["code"])
    return out


def load_targets():
    if not use_supabase():
        return core.load_targets()
    try:
        r = _sb().table("app_settings").select("value").eq("key", KEY).execute()
        if r.data:
            items = r.data[0].get("value") or []
            clean = _clean(items)
            if clean:
                return clean
    except Exception:
        pass
    # 클라우드에 아직 없으면 기본값(또는 로컬 파일)으로 시드
    defaults = core.load_targets()
    try:
        save_targets(defaults)
    except Exception:
        pass
    return _clean(defaults)


def save_targets(items):
    clean = _clean(items)
    if use_supabase():
        _sb().table("app_settings").upsert(
            {"key": KEY, "value": clean}, on_conflict="key").execute()
        return clean
    return core.save_targets(items)
