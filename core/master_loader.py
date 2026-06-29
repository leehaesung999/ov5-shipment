import os
from typing import Optional
import pandas as pd

DEFAULT_MONTHS = 24


def _read_sheet_with_columns(path: str, required_cols: list[str]) -> pd.DataFrame:
    """필수 컬럼이 모두 있는 첫 시트를 찾아 반환.

    pd.ExcelFile은 with로 열고 반드시 닫는다 — 안 닫으면 (long-running Streamlit에서)
    파일 핸들이 새어 해당 엑셀 파일이 계속 잠긴다(삭제/덮어쓰기 불가).
    """
    with pd.ExcelFile(path) as xl:
        for name in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=name)
            if all(c in df.columns for c in required_cols):
                return df
    return pd.DataFrame()


def load_shelf_life_months(master_xlsx: str) -> dict:
    """반환: {item_code: 유통기한월(int)}.
    컬럼명은 '유통기한(월)' 또는 '소비기한(월)' 둘 다 인식 (둘 다 같은 의미).
    """
    if not master_xlsx or not os.path.exists(master_xlsx):
        return {}
    df = pd.DataFrame()
    for col in ("유통기한(월)", "소비기한(월)"):
        try:
            tmp = _read_sheet_with_columns(master_xlsx, ["Item code", col])
        except Exception:
            continue
        if not tmp.empty:
            df = tmp.rename(columns={col: "유통기한(월)"})
            break
    if df.empty:
        return {}
    df = df[["Item code", "유통기한(월)"]].dropna()
    df["Item code"] = df["Item code"].astype("Int64").astype(str)
    df["유통기한(월)"] = pd.to_numeric(df["유통기한(월)"], errors="coerce")
    df = df.dropna()
    return dict(zip(df["Item code"], df["유통기한(월)"].astype(int)))


def load_fs_ms_items(fs_ms_xlsx: str) -> dict:
    """반환: {item_code: [(consume_ymd|None, 'FS'|'MS'), ...]}
    - 유통기한 컬럼이 비어 있으면 None → 해당 품목 전체에 적용
    - 값이 있으면 그 유통기한(YYYYMMDD)과 일치하는 재고만 매칭
    """
    if not fs_ms_xlsx or not os.path.exists(fs_ms_xlsx):
        return {}
    try:
        df = _read_sheet_with_columns(fs_ms_xlsx, ["Item code", "구분"])
    except Exception:
        return {}
    if df.empty:
        return {}

    df = df.dropna(subset=["Item code", "구분"]).copy()
    df["Item code"] = df["Item code"].astype("Int64").astype(str)
    df["구분"] = df["구분"].astype(str).str.upper().str.strip()

    expiry_col = next((c for c in ("유통기한", "소비기한") if c in df.columns), None)
    out: dict[str, list[tuple]] = {}
    for _, r in df.iterrows():
        ymd = None
        if expiry_col is not None:
            val = r[expiry_col]
            if pd.notna(val):
                try:
                    ymd = int(str(val).split(".")[0])
                except (ValueError, TypeError):
                    ymd = None
        out.setdefault(r["Item code"], []).append((ymd, r["구분"]))
    return out


LOT_CATEGORIES = ("농협", "대리점", "FS", "MS", "급식", "소재")


def _parse_categories(raw) -> list:
    """'농협,대리점' / '농협 대리점' 등 → ['농협','대리점'] (유효 카테고리만)."""
    if raw is None:
        return []
    parts = str(raw).replace("/", ",").replace(" ", ",").split(",")
    cats = []
    for p in parts:
        p = p.strip()
        if p in LOT_CATEGORIES and p not in cats:
            cats.append(p)
    return cats


def load_lot_assignments(xlsx_path: str) -> dict:
    """반환: {(item_code, consume_ymd|None): ['농협', '대리점', ...]}
    (품목+유통기한)별로 매칭 가능한 카테고리를 1개 이상 지정 (쉼표 구분 다중).
    유통기한 비우면 그 품목 전체에 적용. 같은 lot 중복 시 마지막 값.
    """
    if not xlsx_path or not os.path.exists(xlsx_path):
        return {}
    cat_aliases = ["카테고리", "구분", "Category"]
    df = pd.DataFrame()
    for alias in cat_aliases:
        try:
            tmp = _read_sheet_with_columns(xlsx_path, ["Item code", alias])
        except Exception:
            continue
        if not tmp.empty:
            df = tmp.rename(columns={alias: "카테고리"})
            break
    if df.empty:
        return {}

    expiry_col = next((c for c in ("유통기한", "소비기한") if c in df.columns), None)
    df = df.dropna(subset=["Item code", "카테고리"]).copy()
    df["Item code"] = df["Item code"].astype("Int64").astype(str)

    out: dict = {}
    for _, r in df.iterrows():
        cats = _parse_categories(r["카테고리"])
        if not cats:
            continue
        ymd = None
        if expiry_col is not None:
            val = r[expiry_col]
            if pd.notna(val):
                try:
                    ymd = int(str(val).split(".")[0])
                except (ValueError, TypeError):
                    ymd = None
        out[(r["Item code"], ymd)] = cats
    return out


def lot_lookup(assignments: dict, item_code: str, consume_ymd) -> list:
    """(item_code, 유통기한)에 지정된 카테고리 리스트. 정확 일치 우선, None(전체) 폴백.
    없으면 빈 리스트."""
    if not assignments:
        return []
    try:
        consume_int = int(str(consume_ymd).split(".")[0]) if consume_ymd is not None else None
    except (ValueError, TypeError):
        consume_int = None
    code = str(item_code)
    val = assignments.get((code, consume_int)) or assignments.get((code, None)) or []
    return list(val) if isinstance(val, (list, tuple)) else [val]


def customer_in_category(customer_code: str, customer_name: str,
                         category: str, nh_keywords) -> bool:
    """주문 거래처가 지정 카테고리에 속하는지 판정."""
    name = customer_name or ""
    code = (customer_code or "").upper()
    if category == "농협":
        # 농협 + 매핑된 지점 키워드 중 하나라도 포함 (제주 등 미매핑 자동 제외)
        if not ("[NH]" in name or "농협경제지주" in name):
            return False
        return any(kw in name for kw in nh_keywords)
    if category == "대리점":
        return code.startswith("GT") or "GT" in name
    if category == "FS":
        return "[FS]" in name
    if category == "MS":
        return "[MS]" in name
    if category == "급식":
        return "급식" in name
    if category == "소재":
        return "[소재]" in name or "소재" in name
    return False


def fs_ms_lookup(fs_ms_items: dict, item_code: str, consume_ymd) -> Optional[str]:
    """재고의 (item_code, 소비기한)에 맞는 FS/MS 구분을 반환. 없으면 None.
    유통기한 None 등록 = 그 품목 전체에 적용 (fallback).
    """
    rules = fs_ms_items.get(str(item_code)) if fs_ms_items else None
    if not rules:
        return None
    try:
        consume_int = int(str(consume_ymd).split(".")[0]) if consume_ymd is not None else None
    except (ValueError, TypeError):
        consume_int = None

    for ymd, tag in rules:
        if ymd is not None and ymd == consume_int:
            return tag
    for ymd, tag in rules:
        if ymd is None:
            return tag
    return None


