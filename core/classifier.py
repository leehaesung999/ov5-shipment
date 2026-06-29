from typing import Optional

DEFAULT_NH_BRANCHES = {
    "포천": "농협포천",
    "평택": "농협평택",
    "횡성": "농협횡성",
    "군위": "농협군위",
    "장성": "농협장성",
    "경남": "농협경남",
}


def classify_order(customer_code: str, customer_name: str,
                   nh_branches: dict = None) -> tuple[str, Optional[str]]:
    """주문 거래처만 보고 분류. (category, label)
    category ∈ {'nh', 'fs_ms', 'gt_dealer', 'other'}
    label   = 농협지점명 또는 'FS'/'MS' 또는 None
    """
    nh_branches = nh_branches or DEFAULT_NH_BRANCHES
    name = customer_name or ""
    code = (customer_code or "").upper()

    # ① 농협
    if "[NH]" in name or "농협경제지주" in name:
        for keyword, branch in nh_branches.items():
            if keyword in name:
                return "nh", branch
        return "nh", None

    # ② 온라인 큐브온커머스 — 예외 매칭 대상 (다른 [온라인] 은 매칭 X)
    if "[온라인]큐브온커머스" in name or "큐브온커머스" in name:
        return "online_qube", "온라인큐브"

    # ③ 다른 [온라인] 거래처 → 매칭 제외
    if "[온라인]" in name:
        return "other", None

    # ④ MS / FS
    if "[MS]" in name:
        return "fs_ms", "MS"
    if "[FS]" in name:
        return "fs_ms", "FS"

    # ⑤ 소재
    if "[소재]" in name or "소재" in name:
        return "material", "소재"

    # ⑥ GT 대리점
    if code.startswith("GT") or "GT" in name:
        return "gt_dealer", "대리점"

    # ⑦ 급식 (다른 카테고리 분류 안 된 일반 급식 거래처)
    if "급식" in name:
        return "school", "급식"

    return "other", None
