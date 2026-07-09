# -*- coding: utf-8 -*-
"""점검 — 이중적치 · OV5/OV6 하프도달 · 비Lock 유통기한 점검 (4종)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import inv_page  # noqa: E402

inv_page.render(
    ["이중적치", "ov5", "ov6", "nonlock"],
    "🔎 점검 · 이중적치 / 하프도달 / 유통기한",
    "이중적치 · OV5/OV6 하프도달 · 비Lock 유통기한 점검 → 화면 표시 + 엑셀 다운로드",
    preview=True,
)
