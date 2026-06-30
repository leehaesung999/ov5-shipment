# -*- coding: utf-8 -*-
"""ATN/BNF 창고비교 — 순수 로직 (tkinter 의존 제거, Streamlit/웹용).

원본 compare_warehouses.py 의 비교·판정·잔존율·붙여넣기 파싱 로직을 그대로 이식.
"""
from __future__ import annotations

import re
from datetime import date, datetime

import pandas as pd

DEFAULT_WAREHOUSES = ["IC930", "IC920", "IC906"]
DEFAULT_MIN_WAREHOUSES = 2
IMMINENT_THRESHOLD = 0.5
DAYS_PER_MONTH = 30
EXTRA_COL_NAME = "지정출고 유통기한"
REMAINING_COL_NAME = "잔존(%)"
VERDICT_COL = "판정"

COLUMNS = {
    0: "Inventory", 1: "Inventory명", 2: "고정로케이션", 3: "Location",
    4: "품목코드", 5: "품목명", 6: "UOM", 7: "입수", 8: "제조일", 9: "유통기한",
    10: "현재고", 11: "현재고_Box", 12: "현재고_Ea", 13: "출고예정", 14: "출고진행",
    15: "출고진행_Box", 16: "출고진행_Ea", 17: "출고가능", 18: "출고가능_Box",
    19: "출고가능_Ea", 20: "Location이동", 21: "Location이동_Box", 22: "Location이동_Ea",
    23: "Lock_Qty", 24: "Lock_Qty_Box", 25: "Lock_Qty_Ea",
}


def load_inventory(path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0, header=0, dtype={9: str})
    if df.shape[1] < len(COLUMNS):
        raise ValueError(f"엑셀 컬럼 수가 예상보다 적습니다: {df.shape[1]}개 (예상 {len(COLUMNS)}개)")
    df = df.iloc[:, : len(COLUMNS)].copy()
    df.columns = [COLUMNS[i] for i in range(len(COLUMNS))]
    return df


def has_shippable_stock(row: pd.Series) -> bool:
    box = pd.to_numeric(row["출고가능_Box"], errors="coerce")
    ea = pd.to_numeric(row["출고가능_Ea"], errors="coerce")
    return (pd.notna(box) and box > 0) or (pd.notna(ea) and ea > 0)


def parse_yyyymmdd(value) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if len(s) == 8 and s.isdigit():
        return f"{s[:4]}-{s[4:6]}-{s[6:]}"
    return s


def _normalize_pivot_code(value) -> str:
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def build_pivot(df: pd.DataFrame, warehouses: list, min_warehouses: int) -> pd.DataFrame:
    df = df[df["Inventory"].isin(warehouses)].copy()
    if df.empty:
        raise ValueError(f"지정한 창고({', '.join(warehouses)}) 데이터가 없습니다.")
    df["품목코드"] = df["품목코드"].map(_normalize_pivot_code)
    df = df[df.apply(has_shippable_stock, axis=1)].copy()
    if df.empty:
        raise ValueError("지정한 창고들에 출고가능재고가 있는 행이 없습니다.")
    # 주의: 점은 반드시 이스케이프(\.0$). ".0$"는 "20260710"→"202607"처럼 끝2자리를 날린다(원본 버그)
    df["유통기한_str"] = df["유통기한"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df = df[df["유통기한_str"].str.match(r"^\d{8}$", na=False)].copy()
    grouped = (
        df.groupby(["품목코드", "품목명", "Inventory"], as_index=False)
        .agg(가장빠른유통기한=("유통기한_str", "min"))
    )
    pivot = grouped.pivot_table(
        index=["품목코드", "품목명"], columns="Inventory",
        values="가장빠른유통기한", aggfunc="min",
    ).reset_index()
    for wh in warehouses:
        if wh not in pivot.columns:
            pivot[wh] = pd.NA
    pivot = pivot[["품목코드", "품목명", *warehouses]]
    wcount = pivot[warehouses].notna().sum(axis=1)
    pivot = pivot[wcount >= min_warehouses].copy()
    pivot["최빠른_전체"] = pivot[warehouses].min(axis=1, skipna=True)
    pivot = pivot.sort_values(["최빠른_전체", "품목코드"]).drop(columns=["최빠른_전체"])
    for wh in warehouses:
        pivot[wh] = pivot[wh].apply(parse_yyyymmdd)
    return pivot.reset_index(drop=True)


def load_master(path) -> dict:
    """물품정보 엑셀 → {품목코드: 유통기한_월}"""
    df = pd.read_excel(path, sheet_name=0, header=0, dtype=str)
    code_col = next((c for c in ("Item code", "품목코드", "제품코드") if c in df.columns),
                    df.columns[0])
    month_col = next((c for c in ("소비기한(월)", "유통기한(월)") if c in df.columns), None)
    if month_col is None:
        for c in df.columns:
            if isinstance(c, str) and ("소비기한" in c or "유통기한" in c) and "월" in c:
                month_col = c
                break
    if month_col is None:
        raise ValueError("'소비기한(월)' 또는 '유통기한(월)' 컬럼을 찾을 수 없습니다.")
    result = {}
    for code, months in zip(df[code_col], df[month_col]):
        if not isinstance(code, str):
            continue
        c = code.strip()
        if c.endswith(".0"):
            c = c[:-2]
        if not c:
            continue
        try:
            m = int(float(months))
        except (TypeError, ValueError):
            continue
        if m > 0:
            result[c] = m
    return result


def calc_remaining_ratio(expiry_str: str, months: int, today: date):
    if not expiry_str or len(expiry_str) != 10:
        return None
    try:
        exp = datetime.strptime(expiry_str, "%Y-%m-%d").date()
    except ValueError:
        return None
    total = months * DAYS_PER_MONTH
    if total <= 0:
        return None
    return (exp - today).days / total


# ---------- 붙여넣기 파서 ----------
def _normalize_code(s: str) -> str:
    s = s.strip()
    if s.endswith(".0"):
        s = s[:-2]
    nc = s.replace(",", "")
    return nc if nc.isdigit() else s


def _normalize_date(s: str) -> str:
    s = s.strip()
    if s.endswith(".0"):
        s = s[:-2]
    if len(s) == 8 and s.isdigit():
        return f"{s[:4]}-{s[4:6]}-{s[6:]}"
    if re.match(r"^\d{4}[-/.]\d{2}[-/.]\d{2}$", s):
        return s.replace("/", "-").replace(".", "-")
    return ""


def _is_date_like(s: str) -> bool:
    return bool(_normalize_date(s))


def _is_code_like(s: str) -> bool:
    s = s.replace(",", "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.isdigit() and 5 <= len(s) <= 10


def _detect_by_header(header):
    code_idx = date_idx = None
    kws = ("제품코드", "품목코드", "상품코드", "SKU")
    for i, cell in enumerate(header):
        c = cell.strip()
        if not c:
            continue
        if code_idx is None and any(k in c for k in kws):
            code_idx = i
        if date_idx is None and "유통" in c:
            date_idx = i
    if code_idx is None:
        for i, cell in enumerate(header):
            if "코드" in cell.strip():
                code_idx = i
                break
    return code_idx, date_idx


def _detect_by_content(rows):
    if not rows:
        return None, None
    n = max(len(r) for r in rows)
    code_idx = date_idx = None
    bc = bd = 0
    for col in range(n):
        vals = [r[col].strip() if col < len(r) else "" for r in rows]
        ne = [v for v in vals if v]
        if not ne:
            continue
        cs = sum(1 for v in ne if _is_code_like(v))
        ds = sum(1 for v in ne if _is_date_like(v))
        if cs > bc and cs >= max(1, len(ne) // 2):
            bc, code_idx = cs, col
        if ds > bd and ds >= max(1, len(ne) // 2):
            bd, date_idx = ds, col
    return code_idx, date_idx


def parse_pasted_data(raw: str) -> dict:
    """엑셀 복사 TSV/CSV → {제품코드: 가장빠른 유통기한}."""
    lines = [ln for ln in raw.replace("\r", "").split("\n") if ln.strip()]
    if not lines:
        raise ValueError("붙여넣은 내용이 비어있습니다.")
    sep = "\t" if "\t" in lines[0] else ("," if "," in lines[0] else "\t")
    rows = [ln.split(sep) for ln in lines]
    code_idx, date_idx = _detect_by_header(rows[0])
    data_rows = rows[1:] if (code_idx is not None or date_idx is not None) else rows
    if code_idx is None or date_idx is None:
        c2, d2 = _detect_by_content(rows)
        code_idx = code_idx if code_idx is not None else c2
        date_idx = date_idx if date_idx is not None else d2
        if code_idx is not None or date_idx is not None:
            data_rows = rows
    if code_idx is None or date_idx is None:
        raise ValueError("제품코드 또는 유통기한 컬럼을 찾을 수 없습니다. 헤더 포함해 복사하세요.")
    mapping = {}
    for row in data_rows:
        if max(code_idx, date_idx) >= len(row):
            continue
        code = _normalize_code(row[code_idx])
        d = _normalize_date(row[date_idx])
        if not _is_code_like(code) or not d:
            continue
        prev = mapping.get(code)
        if prev is None or d < prev:
            mapping[code] = d
    if not mapping:
        raise ValueError("매칭 가능한 제품코드/유통기한이 없습니다.")
    return mapping


def _date_cell(v) -> str:
    try:
        if v is None or pd.isna(v):
            return ""
    except Exception:
        pass
    s = str(v).strip()
    return s if re.match(r"^\d{4}-\d{2}-\d{2}$", s) else ""


def verdict_row(row, warehouses, extra: dict) -> str:
    code = _normalize_pivot_code(row["품목코드"])
    ic930 = _date_cell(row.get("IC930"))
    others = [_date_cell(row.get(wh)) for wh in warehouses if wh != "IC930"]
    others = [d for d in others if d]
    if not ic930:
        return "930없음"
    if not others:
        return "-"
    if min(others) < ic930:
        return "정상(지정출고)" if code in extra else "비정상"
    return "정상"


def row_min_ratio(row, warehouses, master: dict, today: date):
    if not master:
        return None
    code = _normalize_pivot_code(row["품목코드"])
    months = master.get(code)
    if not months:
        return None
    ratios = []
    for wh in warehouses:
        d = _date_cell(row.get(wh))
        if d:
            r = calc_remaining_ratio(d, months, today)
            if r is not None:
                ratios.append(r)
    return min(ratios) if ratios else None
