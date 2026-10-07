# -*- coding: utf-8 -*-
"""SCM 지정출고 주간작업 — 누적 task 영속 저장 (Supabase app_settings).

key='scm_designated_tasks' 에 전체 task 리스트를 JSON으로 저장/로드.
OV5와 같은 Supabase 프로젝트·app_settings 테이블 재사용(추가 SQL 불필요).
st.secrets 에 SUPABASE_URL/KEY 없으면 세션 임시(영속 안 됨).
"""
from __future__ import annotations

import json

try:
    import streamlit as st
except Exception:
    st = None

KEY = "scm_designated_tasks"


def _secret(name):
    try:
        return st.secrets.get(name) if st is not None else None
    except Exception:
        return None


def use_supabase() -> bool:
    return bool(_secret("SUPABASE_URL")) and bool(_secret("SUPABASE_KEY"))


def backend_name() -> str:
    return "☁️ Supabase(클라우드 공유·영속)" if use_supabase() else "⚠️ 세션 임시(영속 안 됨 — 새로고침 시 사라짐)"


if st is not None:
    @st.cache_resource(show_spinner=False)
    def _sb():
        from supabase import create_client
        return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])


def load_tasks() -> list:
    """저장된 누적 task 리스트. 없거나 미설정이면 []."""
    if not use_supabase():
        return []
    try:
        r = _sb().table("app_settings").select("value").eq("key", KEY).execute()
        if r.data:
            v = r.data[0].get("value")
            if isinstance(v, str):
                v = json.loads(v)
            return v or []
    except Exception:
        pass
    return []


def save_tasks(tasks) -> bool:
    """누적 task 리스트 저장(upsert)."""
    if not use_supabase():
        return False
    try:
        _sb().table("app_settings").upsert(
            {"key": KEY, "value": tasks}, on_conflict="key").execute()
        return True
    except Exception:
        return False
