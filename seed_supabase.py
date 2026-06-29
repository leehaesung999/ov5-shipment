# -*- coding: utf-8 -*-
"""현재 로컬 config/lot_assignments*.xlsx 를 Supabase로 1회 이관(시드).

배포 직후, 기존에 쌓아둔 수동/자동 로트 지정을 클라우드로 옮길 때 한 번만 실행.
사용법(로컬에서):
    set SUPABASE_URL=https://xxxx.supabase.co     (PowerShell: $env:SUPABASE_URL="...")
    set SUPABASE_KEY=eyJ...
    python seed_supabase.py
"""
import os
import sys

if not (os.environ.get("SUPABASE_URL") and os.environ.get("SUPABASE_KEY")):
    print("환경변수 SUPABASE_URL / SUPABASE_KEY 를 먼저 설정하세요.")
    sys.exit(1)

from supabase import create_client  # noqa: E402
from core import store  # noqa: E402

sb = create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_KEY"])


def seed(source: str, lots: dict):
    sb.table("lot_assignments").delete().eq("source", source).execute()
    rows = []
    for (code, ymd), cats in lots.items():
        rows.append({
            "item_code": str(int(code)),
            "expiry": str(ymd) if ymd else "",
            "category": ",".join(cats),
            "source": source,
        })
    if rows:
        sb.table("lot_assignments").insert(rows).execute()
    print(f"  {source}: {len(rows)}건 이관")


# 로컬 파일에서 직접 읽기 (use_supabase 분기와 무관하게 파일 강제)
manual = store._read_lots_local(store.LOT_PATH)
auto = store._read_lots_local(store.LOT_AUTO_PATH)
print("로컬 로트 → Supabase 이관 시작")
seed("manual", manual)
seed("auto", auto)

# settings 도 이관
import json  # noqa: E402
if store.SETTINGS_PATH.exists():
    s = json.loads(store.SETTINGS_PATH.read_text(encoding="utf-8"))
    sb.table("app_settings").upsert(
        {"key": store.SETTINGS_KEY, "value": s}, on_conflict="key").execute()
    print("  settings 이관 완료")

print("✅ 시드 완료")
