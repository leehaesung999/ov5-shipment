# -*- coding: utf-8 -*-
"""외박스 요청 리스트 — 공유 저장(지속·다중 사용자).

st.secrets에 SUPABASE_URL/KEY가 있으면 Supabase(app_settings 테이블)에
  · key='box_requests' : 요청 리스트(행 목록)
  · key='box_master'   : 품목코드→[ERP코드, 업체] 매핑
을 저장·공유한다. 없으면 이 서버의 로컬 JSON 파일로 폴백.
※ OV5/쿠팡과 같은 Supabase 프로젝트·app_settings 테이블 재사용 (추가 SQL 불필요).
"""
from __future__ import annotations

import json
from pathlib import Path

try:
    import streamlit as st
except Exception:
    st = None

HERE = Path(__file__).parent
LIST_KEY = "box_requests"
MASTER_KEY = "box_master"
LOCAL_LIST = HERE / "_box_data.json"
LOCAL_MASTER = HERE / "_box_master.json"


def _secret(name):
    try:
        return st.secrets.get(name) if st is not None else None
    except Exception:
        return None


def use_supabase() -> bool:
    return bool(_secret("SUPABASE_URL")) and bool(_secret("SUPABASE_KEY"))


def backend_name() -> str:
    return "🟢 공유 모드 · Supabase (여러 사용자 공유)" if use_supabase() \
        else "💾 로컬 파일 (이 서버에만 저장)"


if st is not None:
    @st.cache_resource(show_spinner=False)
    def _sb():
        from supabase import create_client
        return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])


def _load(key, local_path, default):
    if use_supabase():
        try:
            r = _sb().table("app_settings").select("value").eq("key", key).execute()
            if r.data:
                v = r.data[0].get("value")
                return v if v is not None else default
        except Exception:
            pass
        return default
    try:
        if local_path.exists():
            return json.loads(local_path.read_text(encoding="utf-8"))
    except Exception:
        pass
    return default


def _save(key, local_path, value):
    if use_supabase():
        _sb().table("app_settings").upsert(
            {"key": key, "value": value}, on_conflict="key").execute()
        return
    local_path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def load_rows():
    v = _load(LIST_KEY, LOCAL_LIST, [])
    return v if isinstance(v, list) else []


def save_rows(rows):
    _save(LIST_KEY, LOCAL_LIST, rows)


def load_master():
    v = _load(MASTER_KEY, LOCAL_MASTER, {})
    return v if isinstance(v, dict) else {}


def save_master(m):
    _save(MASTER_KEY, LOCAL_MASTER, m)
