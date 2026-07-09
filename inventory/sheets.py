# -*- coding: utf-8 -*-
"""실사지 출력 — 일일·1단·토요일 실사지 + 2~6단 재고지 (4종)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import inv_page  # noqa: E402

inv_page.render(
    ["일일실사", "토요일", "1단전체", "2_6단"],
    "📋 실사지 출력",
    "일일 · 토요일 · 1단 실사지 및 2~6단 재고지 생성 → 엑셀 다운로드",
)
