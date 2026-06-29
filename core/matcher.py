import re
from datetime import date
from typing import Optional
import pandas as pd

from .shelf_life import (
    calc_remaining_rate, calc_production_date, judge,
    DEFAULT_MONTHS, DEFAULT_THRESHOLD,
)
from .classifier import classify_order, DEFAULT_NH_BRANCHES
from .master_loader import fs_ms_lookup, lot_lookup, customer_in_category

DEFAULT_NH_COLS = ["농협포천", "농협평택", "농협횡성", "농협군위", "농협장성", "농협경남"]
NH_COLS = DEFAULT_NH_COLS  # backward compat
NH_BRANCH_PATTERN = re.compile(r"\[NH\].*?\((.+?)물류센터\)")


def auto_detect_nh_branches(orders: pd.DataFrame,
                            excluded: list = None) -> dict:
    """주문 거래처명에서 [NH]…(○○물류센터) 패턴을 추출. excluded 키워드는 제외."""
    if orders is None or orders.empty or "customer" not in orders.columns:
        return {}
    excluded_set = {k.strip() for k in (excluded or []) if k}
    detected: dict[str, str] = {}
    for name in orders["customer"].dropna().unique():
        m = NH_BRANCH_PATTERN.search(str(name))
        if m:
            kw = m.group(1).strip()
            if kw not in excluded_set:
                detected[kw] = f"농협{kw}"
    return detected


def build_nh_cols(nh_branches: dict) -> list:
    """매핑값(컬럼명) unique 유지 + 기본 순서 보존."""
    seen: list = []
    for col in nh_branches.values():
        if col not in seen:
            seen.append(col)
    return seen


BASE_LEFT = ["Item ID", "Item", "유통기한", "재고수량(BOX)", "로케이션",
             "잔존율", "판정구분", "매칭수량", "거래처명"]
BASE_RIGHT = ["잔여량", "유통기한(월)", "제조일자", "남은율", "판정"]
META_COLS = set(BASE_LEFT) | set(BASE_RIGHT)


def determine_lot_category(row_dict: dict, nh_cols: list) -> Optional[str]:
    """매칭된 행에서 단일 카테고리를 추출. 복수 또는 모호하면 None."""
    used: set = set()
    nh_total = sum(row_dict.get(c, 0) or 0 for c in nh_cols)
    if nh_total > 0:
        used.add("농협")
    for col, qty in row_dict.items():
        if not qty or col in META_COLS or col in nh_cols:
            continue
        if "[MS]" in col:
            used.add("MS")
        elif "[FS]" in col:
            used.add("FS")
        elif "급식" in col:
            used.add("급식")
        else:
            used.add("대리점")  # GT 또는 기타
    return next(iter(used)) if len(used) == 1 else None


def build_output_cols(nh_cols: list, other_cols: list = None) -> list:
    other_cols = other_cols or []
    return [*BASE_LEFT, *nh_cols, *other_cols, *BASE_RIGHT]


OUTPUT_COLS = build_output_cols(NH_COLS)  # backward compat


def _build_order_pool(orders: pd.DataFrame, nh_branches: dict) -> dict:
    """item_code → list of order dict with mutable 'remaining'.
    농협 6개(또는 settings 매핑) 키워드에 매칭 안 되는 농협 거래처는 자동 제외.
    거래처명에 '(잔여)' / '(미배정)' 포함된 주문도 매칭 풀에서 제외.
    """
    nh_keywords = list(nh_branches.keys())
    pool: dict[str, list[dict]] = {}
    for _, row in orders.iterrows():
        cust = str(row["customer"] or "")
        # 잔여/미배정 거래처는 매칭에서 제외
        if "(잔여)" in cust or "(미배정)" in cust or "잔여)" in cust or "미배정)" in cust:
            continue
        is_nh = "[NH]" in cust or "농협경제지주" in cust
        if is_nh and not any(kw in cust for kw in nh_keywords):
            continue  # 매핑되지 않은 농협 지점(제주 등)은 매칭 제외
        cat, label = classify_order(row["customer_code"], row["customer"], nh_branches)
        pool.setdefault(row["item_code"], []).append({
            "category": cat,
            "label": label,
            "customer": row["customer"],
            "customer_code": row["customer_code"],
            "remaining": int(row["qty_box"]),
        })
    return pool


def _order_priority(od: dict) -> int:
    """거래처명에 '급식' 포함된 주문은 매칭 우선순위 뒤로."""
    return 1 if "급식" in (od.get("customer") or "") else 0


def _allocate(stock_qty: int, candidates: list[dict],
              row_buckets: dict, customer_labels: list[str],
              full_only: bool = False) -> int:
    """full_only=True면 stock_qty가 그 주문 잔량 이상일 때만 전량 매칭 (분할 X)."""
    candidates = sorted(candidates, key=_order_priority)
    used = 0
    for od in candidates:
        if stock_qty <= 0:
            break
        if full_only:
            if stock_qty < od["remaining"]:
                continue  # 전량 못 보내면 패스
            take = od["remaining"]
        else:
            take = min(stock_qty, od["remaining"])
        if take <= 0:
            continue
        if od["category"] == "nh":
            bucket_key = od["label"]
            if bucket_key and bucket_key in row_buckets:
                row_buckets[bucket_key] += take
        else:
            # 비농협(GT/FS/MS): 거래처명 그대로 컬럼명
            bucket_key = od["customer"]
            row_buckets[bucket_key] = row_buckets.get(bucket_key, 0) + take
        od["remaining"] -= take
        stock_qty -= take
        used += take
        customer_labels.append(f"{od['customer']}({take})")
    return used


def match(stocks: pd.DataFrame, orders: pd.DataFrame, *,
          shelf_life_map: dict, fs_ms_items: dict = None,  # noqa: ARG001
          lot_assignments: dict = None,
          today: Optional[date] = None,
          threshold: float = DEFAULT_THRESHOLD,
          thresholds: dict = None,
          nh_branches: dict = None,
          excluded_nh_keywords: list = None) -> pd.DataFrame:  # noqa: ARG001
    # fs_ms_items / excluded_nh_keywords 는 backward compat
    _ = (fs_ms_items, excluded_nh_keywords)
    lot_assignments = lot_assignments or {}
    today = today or date.today()
    nh_branches = dict(nh_branches or DEFAULT_NH_BRANCHES)
    nh_cols = build_nh_cols(nh_branches)
    pool = _build_order_pool(orders, nh_branches)

    # 카테고리별 잔존율 기준 (없으면 단일 threshold fallback)
    thresholds = dict(thresholds or {})
    THR = {
        "대리점": thresholds.get("대리점", threshold),
        "FS":     thresholds.get("FS",     threshold),
        "MS":     thresholds.get("MS",     threshold),
        "급식":   thresholds.get("급식",   threshold),
        "소재":   thresholds.get("소재",   threshold),
        "온라인": thresholds.get("온라인", threshold),
    }

    rows = []
    for _, s in stocks.iterrows():
        code = str(s["제품코드"])
        months = shelf_life_map.get(code, DEFAULT_MONTHS)
        rate = calc_remaining_rate(s["소비기한"], months, today)
        produced = calc_production_date(s["소비기한"], months)
        verdict = judge(rate, threshold)
        qty = int(s["Lock Qty(Box)"])

        row = {c: 0 for c in nh_cols}
        labels: list[str] = []
        cands = pool.get(code, [])
        matched_total = 0
        lot_category = lot_lookup(lot_assignments, code, s["소비기한"])

        # 카테고리별 잔존율 통과 여부
        def passes(thr_key: str) -> bool:
            thr = THR[thr_key]
            if rate is None:
                return False
            return rate >= thr

        if lot_category:
            # 로트 지정 시:
            #   ① 농협 우선 매칭 (체크 무관, 잔존율 무관) — 최우선
            #   ② 체크된 카테고리(농협 제외) 잔존율 OK시 매칭, 체크 안 된 카테고리 차단
            nh_kw = list(nh_branches.keys())

            # ① 농협
            nh = [c for c in cands if c["category"] == "nh" and c["remaining"] > 0]
            used = _allocate(qty, nh, row, labels)
            qty -= used; matched_total += used

            # ② 체크된 비농협 카테고리만
            other_cats = [c for c in lot_category if c != "농협"]
            if qty > 0 and other_cats:
                target = []
                for c in cands:
                    if c["remaining"] <= 0 or c["category"] == "nh":
                        continue
                    hit_cat = None
                    for cat in other_cats:
                        if customer_in_category(c.get("customer_code", ""),
                                                c["customer"], cat, nh_kw):
                            hit_cat = cat
                            break
                    if hit_cat is None:
                        continue
                    if hit_cat in ("대리점", "FS", "MS", "급식", "소재"):
                        if not passes(hit_cat):
                            continue
                    target.append(c)
                used = _allocate(qty, target, row, labels)
                qty -= used; matched_total += used

            category = "로트:" + "/".join(lot_category)
        else:
            # 기본 룰: 농협 우선 → 5개 카테고리 각자 threshold → 온라인큐브 전량
            nh = [c for c in cands if c["category"] == "nh" and c["remaining"] > 0]
            used = _allocate(qty, nh, row, labels)
            qty -= used; matched_total += used

            # 5개 카테고리 (GT/FS/MS/소재/급식) 각자 threshold
            non_nh_buckets = [
                ("gt_dealer", None, "대리점"),
                ("fs_ms",     "FS", "FS"),
                ("fs_ms",     "MS", "MS"),
                ("material",  None, "소재"),
                ("school",    None, "급식"),
            ]
            for cat_filter, label_filter, thr_key in non_nh_buckets:
                if qty <= 0:
                    break
                if not passes(thr_key):
                    continue
                target = [c for c in cands
                          if c["category"] == cat_filter
                          and (label_filter is None or c["label"] == label_filter)
                          and c["remaining"] > 0]
                used = _allocate(qty, target, row, labels)
                qty -= used; matched_total += used

            # 온라인큐브 — 전량 매칭만 (분할 X)
            if qty > 0 and passes("온라인"):
                qube = [c for c in cands
                        if c["category"] == "online_qube" and c["remaining"] > 0]
                used = _allocate(qty, qube, row, labels, full_only=True)
                qty -= used; matched_total += used

            category = _summarize_category(cands, None, verdict)

        matched_sum = matched_total

        rows.append({
            "Item ID": code,
            "Item": s["제품명"],
            "유통기한": int(s["소비기한"]) if pd.notna(s["소비기한"]) else None,
            "재고수량(BOX)": int(s["Lock Qty(Box)"]),
            "로케이션": s["Location"],
            "잔존율": round(rate, 6) if rate is not None else None,
            "판정구분": category,
            "매칭수량": matched_sum if matched_sum else None,
            "거래처명": ", ".join(labels) if labels else None,
            **row,
            "잔여량": qty if qty > 0 else None,
            "유통기한(월)": months,
            "제조일자": produced,
            "남은율": round(rate, 6) if rate is not None else None,
            "판정": verdict,
        })
    # 매칭된 비농협 거래처 모아 동적 컬럼 정의
    used_others: list = []
    for r in rows:
        for k in r.keys():
            if k in BASE_LEFT or k in BASE_RIGHT or k in nh_cols:
                continue
            if k not in used_others:
                used_others.append(k)
    other_cols = sorted(used_others)
    output_cols = build_output_cols(nh_cols, other_cols)

    df = pd.DataFrame(rows)
    for c in output_cols:
        if c not in df.columns:
            df[c] = 0
    df = df[output_cols]

    # 소수점 없이 정수로 표시되도록 nullable Int64 적용
    FLOAT_COLS = {"잔존율", "남은율"}
    NON_NUMERIC = {"Item ID", "Item", "로케이션", "판정구분",
                   "거래처명", "제조일자", "판정"}
    for c in df.columns:
        if c in FLOAT_COLS or c in NON_NUMERIC:
            continue
        try:
            df[c] = df[c].astype("Int64")
        except (ValueError, TypeError):
            pass

    # 품목코드 → 유통기한 순 정렬
    df = df.sort_values(["Item ID", "유통기한"], kind="stable",
                         na_position="last", ignore_index=True)
    return df


def _summarize_category(cands: list[dict], fs_ms_tag: Optional[str], verdict: str) -> str:
    if fs_ms_tag:
        return fs_ms_tag.lower()  # 'ms' or 'fs' — 농협 매칭은 별도로 같이 일어남
    has_nh = any(c["category"] == "nh" for c in cands)
    has_gt = any(c["category"] == "gt_dealer" for c in cands)
    if has_gt and verdict == "OK":
        return "농협/대리점" if has_nh else "대리점"
    if has_nh:
        return "농협만"
    return ""
