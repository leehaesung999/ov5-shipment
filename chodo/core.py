# -*- coding: utf-8 -*-
"""BNF 초도입고 → 파레트 혼적 규칙 → 차량 분할 로직.

파레트 구성 규칙 (사용자 정의):
  각 품목: 요청박스 // 하대 = 꽉찬 파레트, 나머지(잔바리)를 아래로 분류.
    - 잔바리 ≤ 5박스   : 한 파레트에 최대 tier1_max(기본10) 품목 묶음
    - 잔바리 6~24박스  : 한 파레트에 최대 tier2_max(기본3)  품목 묶음
    - 잔바리 25박스+   : 적재율(박스/하대) 합이 tier3_ratio(기본0.8) 이하가
                         되도록 최대 tier3_max(기본2) 품목 페어링
  소형/중형 개수 묶음은 물리적 안전한도(적재율 합 ≤ pallet_cap, 기본1.2)를 넘지 않음.

차량 분할: 완성된 파레트들을 용차 한도(voncha_capacity, 기본16파레트)씩 순서대로 적재.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class PalletItem:
    제품코드: int
    제품명: str
    박스: int
    하대: int
    로케이션: str = ""      # 고정로케이션(C열, 홈 위치) — 피킹 동선 정렬용

    @property
    def 적재율(self) -> float:
        return self.박스 / self.하대 if self.하대 else 0.0


@dataclass
class Pallet:
    kind: str                       # 'full' | 'small' | 'mid' | 'big'
    items: list[PalletItem] = field(default_factory=list)

    @property
    def ratio_sum(self) -> float:
        return sum(i.적재율 for i in self.items)


@dataclass
class Truck:
    label: str
    capacity: int
    loc: str = ""            # 출고지 표시명 (예: '930', '재고부족')
    pallets: list[Pallet] = field(default_factory=list)


# ----------------- 출고 로케이션 배정 -----------------

STOCK_SHORT = "재고부족"


def loc_display(inv: str) -> str:
    """IC930 → '930' 처럼 접미 숫자만. 그 외는 원문."""
    if isinstance(inv, str) and inv.upper().startswith("IC"):
        return inv[2:]
    return str(inv)


def _inv_box(avail: dict, inv: str) -> float:
    d = avail.get(inv)
    return d["box"] if isinstance(d, dict) else float(d or 0)


def _inv_loc(avail: dict, inv: str) -> str:
    d = avail.get(inv)
    return (d.get("loc") if isinstance(d, dict) else None) or ""


def assign_locations(
    request_rows: list[dict],
    inv_avail: dict[int, dict[str, dict]],
    priority: list[str],
) -> tuple[dict[str, list[dict]], list[str]]:
    """각 주문 품목을 우선순위 로케이션에 배정.

    규칙: priority 순서로 훑어 출고가능Box ≥ 요청박스 인 첫 Inventory를 출고지로.
          어느 우선순위도 전량 커버 못하면 STOCK_SHORT('재고부족') 그룹 + 경고.

    Returns:
      groups: {출고지표시명: [ {제품코드, 요청박스, 출고지, 출고Inventory, 출고Location} ]}
              키 순서는 priority(표시명) → 재고부족.
      warnings: 재고부족 품목 경고 리스트.
    """
    groups: dict[str, list[dict]] = {}
    # 표시명 순서 확보
    order = [loc_display(p) for p in priority] + [STOCK_SHORT]
    for k in order:
        groups[k] = []

    warnings: list[str] = []
    for r in request_rows:
        code = r["제품코드"]
        qty = r["요청박스"]
        avail = inv_avail.get(code, {})
        chosen_inv = None
        for inv in priority:
            if _inv_box(avail, inv) >= qty:
                chosen_inv = inv
                break
        if chosen_inv is None:
            disp = STOCK_SHORT
            have = sum(_inv_box(avail, inv) for inv in avail)
            note = "재고 0" if have <= 0 else f"우선순위 3곳 부족(합 {int(have)}박스)"
            warnings.append(f"재고부족: 코드 {code} (요청 {qty}박스) — {note}")
            row2 = {**r, "출고지": disp, "출고Inventory": None, "출고Location": ""}
        else:
            disp = loc_display(chosen_inv)
            row2 = {**r, "출고지": disp, "출고Inventory": chosen_inv,
                    "출고Location": _inv_loc(avail, chosen_inv)}
        groups.setdefault(disp, []).append(row2)

    # 빈 그룹 제거 (순서 유지)
    groups = {k: v for k, v in groups.items() if v}
    return groups, warnings


# ----------------- 파레트 구성 -----------------

def _pack_ffd(items: list[PalletItem], max_count: int, cap: float, kind: str) -> list[Pallet]:
    """First-Fit-Decreasing: 적재율 큰 순으로, 개수·적재율 한도를 지키며 최소 파레트에 담음."""
    items = sorted(items, key=lambda x: -x.적재율)
    pallets: list[Pallet] = []
    for it in items:
        placed = False
        for p in pallets:
            if len(p.items) < max_count and p.ratio_sum + it.적재율 <= cap + 1e-9:
                p.items.append(it)
                placed = True
                break
        if not placed:
            pallets.append(Pallet(kind=kind, items=[it]))
    return pallets


def _loc_key(it: PalletItem):
    # 로케이션 없는 항목은 뒤로. 같은 로케이션 내에서는 코드 순.
    return (it.로케이션 or "￿", it.제품코드)


def _pack_by_location(items: list[PalletItem], max_count: int, cap: float, kind: str) -> list[Pallet]:
    """창고 로케이션(피킹 동선) 순으로 정렬 후, 인접한 품목끼리 순차로 파레트에 담음(Next-Fit).

    개수 한도·적재율합 한도를 넘으면 새 파레트로. 로케이션 인접성이 최대한 보존됨.
    """
    items = sorted(items, key=_loc_key)
    pallets: list[Pallet] = []
    cur: Pallet | None = None
    for it in items:
        if (cur is None or len(cur.items) >= max_count
                or cur.ratio_sum + it.적재율 > cap + 1e-9):
            cur = Pallet(kind=kind, items=[])
            pallets.append(cur)
        cur.items.append(it)
    return pallets


def build_pallets(
    request_rows: list[dict],
    master: dict[int, dict],
    tier1_box: int = 5, tier1_max: int = 10,
    tier2_box: int = 24, tier2_max: int = 3,
    tier3_max: int = 2, tier3_ratio: float = 0.8,
    pallet_cap: float = 1.2,
    group_mode: str = "location",
    big_pairing: bool = True,
) -> tuple[list[Pallet], list[str]]:
    """요청 목록 → 파레트 리스트 + 경고.

    group_mode: 소형/중형 묶음 기준.
      'location' — 창고 로케이션(피킹 동선) 순으로 인접 품목끼리 (기본)
      'quantity' — 적재율 순 최소 파레트(FFD)
    big_pairing: 대형(25박스+) 잔바리 처리.
      True  — 적재율합 ≤ tier3_ratio 로 최대 tier3_max 품목 페어링 (기본)
      False — 페어링 없이 1품목=1파레트 (25박스 이상은 무조건 별도 파레트)
    """
    warnings: list[str] = []
    full_pallets: list[Pallet] = []
    t1: list[PalletItem] = []
    t2: list[PalletItem] = []
    t3: list[PalletItem] = []

    for r in request_rows:
        code = r["제품코드"]
        qty = r["요청박스"]
        m = master.get(code)
        if not m:
            warnings.append(f"기준정보 누락: 코드 {code} (요청 {qty}박스) — 제외됨")
            continue
        hadae = m.get("하대")
        if not hadae or hadae <= 0:
            warnings.append(f"하대 없음: 코드 {code} ({m.get('제품명')}) — 제외됨")
            continue
        name = m.get("제품명") or ""
        loc = r.get("출고Location", "")

        full = qty // hadae
        rem = qty % hadae
        for _ in range(full):
            full_pallets.append(Pallet(
                kind="full",
                items=[PalletItem(code, name, hadae, hadae, loc)],
            ))
        if rem > 0:
            pi = PalletItem(code, name, rem, hadae, loc)
            if rem <= tier1_box:
                t1.append(pi)
            elif rem <= tier2_box:
                t2.append(pi)
            else:
                t3.append(pi)

    pack = _pack_by_location if group_mode == "location" else _pack_ffd
    small = pack(t1, max_count=tier1_max, cap=pallet_cap, kind="small")
    mid = pack(t2, max_count=tier2_max, cap=pallet_cap, kind="mid")
    # 대형: 페어링 켜짐이면 적재율합 ≤ tier3_ratio(0.8)·최대 tier3_max(2) FFD,
    #       꺼짐이면 1품목=1파레트(각자 별도). group_mode=location이면 로케이션 순 유지.
    if big_pairing:
        big = _pack_ffd(t3, max_count=tier3_max, cap=tier3_ratio, kind="big")
    else:
        t3_sorted = sorted(t3, key=_loc_key) if group_mode == "location" \
            else sorted(t3, key=lambda x: -x.적재율)
        big = [Pallet(kind="big", items=[pi]) for pi in t3_sorted]

    if group_mode == "location":
        full_pallets.sort(key=lambda p: _loc_key(p.items[0]))

    pallets = full_pallets + big + mid + small
    return pallets, warnings


# ----------------- 차량 분할 -----------------

def assign_trucks(pallets: list[Pallet], voncha_capacity: int = 16, loc: str = "") -> list[Truck]:
    """완성된 파레트를 용차 한도씩 순서대로 나눔. loc가 있으면 라벨에 접두."""
    trucks: list[Truck] = []
    if voncha_capacity <= 0:
        voncha_capacity = 16
    for i in range(0, len(pallets), voncha_capacity):
        chunk = pallets[i:i + voncha_capacity]
        n = len(trucks) + 1
        trucks.append(Truck(label=f"{n}호차", capacity=voncha_capacity, loc=loc, pallets=chunk))
    return trucks


def dispatch_by_location(
    request_rows: list[dict],
    master: dict[int, dict],
    inv_avail: dict[int, dict[str, float]],
    priority: list[str],
    voncha_capacity: int = 16,
    **pallet_kwargs,
) -> tuple[list[Truck], list[Pallet], list[str], dict]:
    """전체 파이프라인: 로케이션 배정 → 로케이션별 파레트 구성 → 로케이션별 차량 분할.

    Returns: (all_trucks, all_pallets, warnings, groups_meta, shortages)
      groups_meta: {출고지: {'품목수','박스','파레트','차량수'}}
      shortages: 재고부족(우선순위 3곳 커버 불가) 품목 상세 리스트.
    """
    groups, loc_warns = assign_locations(request_rows, inv_avail, priority)
    all_trucks: list[Truck] = []
    all_pallets: list[Pallet] = []
    warnings: list[str] = list(loc_warns)
    groups_meta: dict[str, dict] = {}

    # 재고부족 상세 (제품명·로케이션별 재고 포함)
    prio_disp = [loc_display(p) for p in priority]
    shortages: list[dict] = []
    for r in groups.get(STOCK_SHORT, []):
        code = r["제품코드"]
        avail = inv_avail.get(code, {})
        other = sum(_inv_box(avail, inv) for inv in avail if inv not in priority)
        shortages.append({
            "제품코드": code,
            "제품명": (master.get(code) or {}).get("제품명", ""),
            "요청박스": r["요청박스"],
            "우선순위재고": {d: int(_inv_box(avail, inv)) for d, inv in zip(prio_disp, priority)},
            "기타재고": int(other),
        })

    for loc, rows in groups.items():
        pallets, warns = build_pallets(rows, master, **pallet_kwargs)
        warnings.extend(warns)
        trucks = assign_trucks(pallets, voncha_capacity=voncha_capacity, loc=loc)
        all_pallets.extend(pallets)
        all_trucks.extend(trucks)
        groups_meta[loc] = {
            "품목수": len(rows),
            "박스": sum(r["요청박스"] for r in rows),
            "파레트": len(pallets),
            "차량수": len(trucks),
        }
    return all_trucks, all_pallets, warnings, groups_meta, shortages


# ----------------- 요약 / 내보내기 -----------------

def summarize(request_rows: list[dict], pallets: list[Pallet],
              trucks: list[Truck], voncha_capacity: int,
              groups_meta: dict | None = None) -> dict:
    kinds = {"full": 0, "small": 0, "mid": 0, "big": 0}
    for p in pallets:
        kinds[p.kind] = kinds.get(p.kind, 0) + 1
    return {
        "품목수": len(request_rows),
        "총박스": sum(r["요청박스"] for r in request_rows),
        "총파레트": len(pallets),
        "full": kinds["full"],
        "small": kinds["small"],
        "mid": kinds["mid"],
        "big": kinds["big"],
        "차량수": len(trucks),
        "용차한도": voncha_capacity,
        "groups": groups_meta or {},
    }


def trucks_to_export(trucks: list[Truck]) -> list[dict]:
    out = []
    for t in trucks:
        out.append({
            "label": t.label,
            "capacity": t.capacity,
            "loc": t.loc,
            "pallets": [
                {
                    "kind": p.kind,
                    "items": [
                        {
                            "제품코드": i.제품코드,
                            "제품명": i.제품명,
                            "로케이션": i.로케이션,
                            "박스": i.박스,
                            "하대": i.하대,
                            "적재율": i.적재율,
                        }
                        for i in p.items
                    ],
                }
                for p in t.pallets
            ],
        })
    return out
