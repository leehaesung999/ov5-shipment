# -*- coding: utf-8 -*-
"""ABC 재배치 분석기 초기 시드 (로컬 1회 실행).

로컬 abc-updater/data/ 안의 28개월 파일 + 마스터를 Supabase app_settings 에 base64 저장.
자격증명은 designated-shipment/.streamlit/secrets.toml 에서 읽음.
"""
from __future__ import annotations

import base64
import re
import sys
import tomllib
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).parent.parent
SECRETS = HERE / ".streamlit" / "secrets.toml"
SRC = Path("c:/Users/sempio/Desktop/gpters-21th-temp-main/my/projects/abc-updater/data")
MONTHLY = SRC / "monthly_data"
MASTER = SRC / "location_master.xlsx"


def load_secrets():
    with open(SECRETS, "rb") as f:
        return tomllib.load(f)


def main():
    secrets = load_secrets()
    from supabase import create_client
    sb = create_client(secrets["SUPABASE_URL"], secrets["SUPABASE_KEY"])

    def upsert(key, value):
        sb.table("app_settings").upsert(
            {"key": key, "value": value}, on_conflict="key").execute()

    # 마스터
    if not MASTER.exists():
        print(f"ERROR: 마스터 없음 — {MASTER}")
        return 1
    import openpyxl
    wb = openpyxl.load_workbook(MASTER, data_only=True, read_only=True)
    n_master = sum(1 for _ in wb.active.iter_rows(min_row=2, values_only=True))
    wb.close()
    upsert("abc_master_b64", base64.b64encode(MASTER.read_bytes()).decode())
    upsert("abc_master_meta", {
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "n": n_master,
    })
    print(f"[master] {n_master:,}행 업로드")

    # 월별
    files = sorted(MONTHLY.glob("????-??.xlsx"))
    if not files:
        print(f"ERROR: 월별 파일 없음 — {MONTHLY}")
        return 1
    index = []
    for f in files:
        m = re.match(r"(\d{4}-\d{2})\.xlsx$", f.name)
        if not m:
            continue
        ym = m.group(1)
        key = f"abc_monthly_b64_{ym.replace('-', '_')}"
        upsert(key, base64.b64encode(f.read_bytes()).decode())
        index.append(ym)
        print(f"[month] {ym} ({f.stat().st_size:,} B) 업로드")

    upsert("abc_monthly_index", sorted(set(index)))
    print(f"[index] {len(index)}개월 등록")

    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
