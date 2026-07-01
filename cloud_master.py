# -*- coding: utf-8 -*-
"""마스터(기준정보) Supabase 영구저장 + 마지막 업데이트 날짜 — 여러 페이지 공용.

app_settings 테이블에 <prefix>_master_meta(날짜·품목수) / <prefix>_master_b64(파일) 저장.
Supabase 미설정(로컬)이면 no-op → 파일 mtime 기반 표시.
"""
from __future__ import annotations

import base64
import os
from datetime import datetime
from pathlib import Path

try:
    import streamlit as st
except Exception:
    st = None


def _secret(n):
    try:
        return st.secrets.get(n) if st is not None else None
    except Exception:
        return None


def use_supabase() -> bool:
    return bool(_secret("SUPABASE_URL")) and bool(_secret("SUPABASE_KEY"))


if st is not None:
    @st.cache_resource(show_spinner=False)
    def _sb():
        from supabase import create_client
        return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

    @st.cache_data(show_spinner=False)
    def _fetch_bytes(prefix: str, version: str):
        try:
            r = _sb().table("app_settings").select("value").eq(
                "key", f"{prefix}_master_b64").execute()
            if r.data:
                return base64.b64decode(r.data[0]["value"])
        except Exception:
            pass
        return None


def fetch_meta(prefix: str):
    if not use_supabase():
        return None
    try:
        r = _sb().table("app_settings").select("value").eq(
            "key", f"{prefix}_master_meta").execute()
        return r.data[0].get("value") if r.data else None
    except Exception:
        return None


def store(prefix: str, data: bytes, count: int) -> bool:
    """마스터 파일을 Supabase에 영구 저장하고 업데이트 시각 기록."""
    if not use_supabase():
        return False
    try:
        meta = {"updated": datetime.now().strftime("%Y-%m-%d %H:%M"), "n": int(count)}
        _sb().table("app_settings").upsert(
            {"key": f"{prefix}_master_meta", "value": meta}, on_conflict="key").execute()
        _sb().table("app_settings").upsert(
            {"key": f"{prefix}_master_b64", "value": base64.b64encode(data).decode()},
            on_conflict="key").execute()
        try:
            _fetch_bytes.clear()
        except Exception:
            pass
        return True
    except Exception:
        return False


def restore_to(prefix: str, dest_path) -> dict | None:
    """Supabase에 저장된 마스터가 있으면 dest_path에 1회 써넣고 meta 반환.
    같은 버전은 세션당 1번만 기록(불필요한 재쓰기·캐시무효화 방지)."""
    meta = fetch_meta(prefix)
    if not meta:
        return None
    version = str(meta.get("updated", ""))
    skey = f"_mrestored_{prefix}"
    if st is not None and st.session_state.get(skey) == version:
        return meta
    b = _fetch_bytes(prefix, version) if st is not None else None
    if b:
        try:
            Path(dest_path).write_bytes(b)
            if st is not None:
                st.session_state[skey] = version
        except Exception:
            pass
    return meta


def show_date(prefix: str, default_file=None):
    """마지막 업데이트 날짜 배지 표시. (Supabase meta 우선, 없으면 파일 mtime)"""
    if st is None:
        return None
    meta = fetch_meta(prefix)
    if meta:
        st.info(f"📅 기준정보 마지막 업데이트: **{meta.get('updated', '?')}** "
                f"({meta.get('n', '?')}품목)")
        return meta
    if default_file and os.path.isfile(default_file):
        mt = datetime.fromtimestamp(os.path.getmtime(default_file))
        st.caption(f"기준정보: 내장 기본본 ({mt:%Y-%m-%d} 기준) — 업로드하면 날짜 갱신")
    return None
