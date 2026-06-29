from typing import Iterable
import pandas as pd

REQUIRED_COLS = ["Location", "제품코드", "제품명", "UOM", "소비기한",
                 "현재고(Box)", "Lock Qty(Box)"]


def load_locked_stock(xlsx_path: str, ov_locations: Iterable[str] = ("OV5",)) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path)
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"재고 파일에 필수 컬럼 누락: {missing}")

    ov_set = {str(x).strip().upper() for x in ov_locations}
    mask_loc = df["Location"].astype(str).str.strip().str.upper().isin(ov_set)
    mask_lock = df["Lock Qty(Box)"].fillna(0).astype(float) > 0
    filtered = df[mask_loc & mask_lock].copy()

    filtered["제품코드"] = filtered["제품코드"].astype("Int64").astype(str)
    filtered["Lock Qty(Box)"] = filtered["Lock Qty(Box)"].astype(int)
    filtered["현재고(Box)"] = filtered["현재고(Box)"].fillna(0).astype(int)
    return filtered[REQUIRED_COLS].reset_index(drop=True)
