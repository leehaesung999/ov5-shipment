"""제공된 3개 파일로 end-to-end 검증."""
import sys
from datetime import date
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).parent))

from core.stock_loader import load_locked_stock
from core.order_loader import load_orders
from core.master_loader import load_shelf_life_months, load_fs_ms_items
from core.matcher import match
from core.writer import write_summary

STOCK = r"C:\Users\sempio\Downloads\로케이션별 재고조회_20260515102050.xlsx"
ORDERS = r"C:\Users\sempio\Downloads\출고진행현황_전체탭_20260515121254.xlsx"
TEMPLATE = r"C:\Users\sempio\Desktop\새 폴더\지정출고\0514 통합 지정출고 양식(농협)_수식추가.xlsx"
TODAY = date(2026, 5, 15)


def main():
    print("=== 재고 로드 ===")
    stocks = load_locked_stock(STOCK, ov_locations=["OV5"])
    print(f"OV5 Lock 재고: {len(stocks)}건")
    print(stocks.to_string())
    print()

    print("=== 주문 로드 ===")
    orders = load_orders(ORDERS)
    print(f"주문 행: {len(orders)}")
    print()

    print("=== 기준정보 로드 (현재 통합 양식의 기준정보 시트 활용) ===")
    shelf_map = load_shelf_life_months(TEMPLATE)
    print(f"품목 등록 수: {len(shelf_map)}")
    print()

    print("=== 매칭 (잔존율 70%) ===")
    result = match(stocks, orders, shelf_life_map=shelf_map,
                   fs_ms_items={}, today=TODAY, threshold=0.70)
    print(result[["Item ID", "Item", "잔존율", "판정", "판정구분"]].to_string())
    print()

    print("=== 진간장 금S 15L 검증 ===")
    target = result[result["Item ID"] == "1010422"]
    print(target.to_string())

    print("\n=== 결과 xlsx 저장 ===")
    out_dir = str(Path(__file__).parent / "test_output")
    path = write_summary(result, out_dir)
    print(f"저장 완료: {path}")


if __name__ == "__main__":
    main()
