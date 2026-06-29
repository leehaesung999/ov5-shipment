import pandas as pd

REQUIRED = ["item code", "Customer code", "Customer", "출고지정수량_Box"]


def load_orders(xlsx_path: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path)
    if df.empty:
        return pd.DataFrame(columns=["item_code", "customer_code", "customer", "qty_box"])
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"주문 파일에 필수 컬럼 누락: {missing}")

    out = df[REQUIRED].copy()
    out.columns = ["item_code", "customer_code", "customer", "qty_box"]
    out["item_code"] = out["item_code"].astype("Int64").astype(str)
    out["customer_code"] = out["customer_code"].astype(str).str.strip()
    out["customer"] = out["customer"].astype(str).str.strip()
    out["qty_box"] = pd.to_numeric(out["qty_box"], errors="coerce").fillna(0).astype(int)
    out = out[out["qty_box"] > 0]
    return (out.groupby(["item_code", "customer_code", "customer"], as_index=False)["qty_box"]
              .sum())
