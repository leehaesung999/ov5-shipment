# -*- coding: utf-8 -*-
"""안전재고 히스토리/마스터 저장소.

히스토리(품목×일자×[총,딜])는 Supabase app_settings 에 gzip+base64 블롭으로 누적.
  key = 'ss_history_b64', 메타 = 'ss_history_meta'(종료일·품목수·갱신시각)
Supabase 미설정(로컬)이면 번들 시드(data/history_seed.json.gz) 사용 + 세션 유지.
마스터는 번들 data/master.json (신제품 생기면 화면에서 갱신).
"""
from __future__ import annotations
import base64
import gzip
import json
from datetime import datetime, timezone, timedelta
from pathlib import Path

try:
    import streamlit as st
except Exception:
    st = None

KST = timezone(timedelta(hours=9))
DATA = Path(__file__).parent / "data"
SEED = DATA / "history_seed.json.gz"
MASTER = DATA / "master.json"
HKEY = "ss_history_b64"
MKEY = "ss_history_meta"


def _secret(n):
    try:
        return st.secrets.get(n) if st is not None else None
    except Exception:
        return None


def use_supabase() -> bool:
    return bool(_secret("SUPABASE_URL")) and bool(_secret("SUPABASE_KEY"))


def _sb():
    from supabase import create_client
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])


def _encode(history: dict) -> str:
    raw = json.dumps(history, ensure_ascii=False).encode("utf-8")
    return base64.b64encode(gzip.compress(raw, 6)).decode("ascii")


def _decode(b64: str) -> dict:
    return json.loads(gzip.decompress(base64.b64decode(b64)).decode("utf-8"))


def load_seed() -> dict:
    if SEED.exists():
        with gzip.open(SEED, "rt", encoding="utf-8") as f:
            return json.load(f)
    return {}


def load_history() -> tuple:
    """(history, source) 반환. Supabase 우선, 없으면 번들 시드."""
    if use_supabase():
        try:
            r = _sb().table("app_settings").select("value").eq("key", HKEY).execute()
            if r.data and r.data[0].get("value"):
                return _decode(r.data[0]["value"]), "Supabase"
        except Exception:
            pass
        # Supabase 비었으면 시드로 초기화
        h = load_seed()
        return h, "시드(초기)"
    return load_seed(), "시드(로컬)"


def save_history(history: dict) -> bool:
    """Supabase에 히스토리 저장 + 메타 기록."""
    if not use_supabase():
        return False
    try:
        b64 = _encode(history)
        dmax = max((d for it in history.values() for d in it.keys()), default=None)
        meta = {"종료일": dmax, "품목수": len(history),
                "갱신": datetime.now(KST).strftime("%Y-%m-%d %H:%M")}
        _sb().table("app_settings").upsert({"key": HKEY, "value": b64}).execute()
        _sb().table("app_settings").upsert({"key": MKEY, "value": meta}).execute()
        return True
    except Exception:
        return False


def history_meta() -> dict | None:
    if not use_supabase():
        return None
    try:
        r = _sb().table("app_settings").select("value").eq("key", MKEY).execute()
        return r.data[0].get("value") if r.data else None
    except Exception:
        return None


# ---------- 반영 완료된 행사(header_id) 목록 — 신규만 반영 위한 중복방지 ----------
EKEY = "ss_events_reflected"


def load_reflected_events() -> set:
    """이미 반영 완료한 행사 header_id 집합. Supabase 우선, 없으면 빈 set."""
    if use_supabase():
        try:
            r = _sb().table("app_settings").select("value").eq("key", EKEY).execute()
            if r.data and r.data[0].get("value") is not None:
                return set(str(x) for x in (r.data[0]["value"] or []))
        except Exception:
            pass
    return set()


def add_reflected_events(header_ids) -> bool:
    """header_id들을 반영완료 목록에 추가 저장."""
    ids = load_reflected_events() | set(str(x) for x in header_ids)
    if not use_supabase():
        return False
    try:
        _sb().table("app_settings").upsert(
            {"key": EKEY, "value": sorted(ids)}).execute()
        return True
    except Exception:
        return False


# ---------- 품목코드 승계표 (입수변경 등으로 코드가 바뀐 경우) ----------
RKEY = "ss_code_remap"
REMAP_SEED = DATA / "code_remap_seed.json"


def load_code_remap() -> list:
    """번들 시드 ∪ Supabase 승계규칙(구코드→신코드). (old,new)로 중복제거,
    Supabase 항목이 시드를 덮어쓴다(입수 수정 반영)."""
    seed = []
    if REMAP_SEED.exists():
        try:
            seed = json.loads(REMAP_SEED.read_text(encoding="utf-8"))
        except Exception:
            seed = []
    extra = []
    if use_supabase():
        try:
            r = _sb().table("app_settings").select("value").eq("key", RKEY).execute()
            if r.data and r.data[0].get("value"):
                extra = r.data[0]["value"] or []
        except Exception:
            pass
    merged = {}
    for r in list(seed) + list(extra):
        old = str(r.get("old", "")).strip()
        new = str(r.get("new", "")).strip()
        if old and new:
            merged[(old, new)] = {"old": old, "new": new,
                                  "ip": r.get("ip"), "plt": r.get("plt")}
    return list(merged.values())


def add_code_remap(old, new, ip=None, plt=None) -> bool:
    """승계규칙 1건을 Supabase에 추가(동일 old→new 있으면 갱신)."""
    cur = []
    if use_supabase():
        try:
            r = _sb().table("app_settings").select("value").eq("key", RKEY).execute()
            if r.data and r.data[0].get("value"):
                cur = r.data[0]["value"] or []
        except Exception:
            pass
    cur = [c for c in cur if not (str(c.get("old")) == str(old)
                                  and str(c.get("new")) == str(new))]
    cur.append({"old": str(old), "new": str(new),
                "ip": int(ip) if ip else None, "plt": int(plt) if plt else None})
    if not use_supabase():
        return False
    try:
        _sb().table("app_settings").upsert({"key": RKEY, "value": cur}).execute()
        return True
    except Exception:
        return False


# ---------- 신규품목 초기기준 (유사품 기반) ----------
NKEY = "ss_new_items"
NEWITEMS_SEED = DATA / "new_items_seed.json"


def load_new_items() -> list:
    """번들 시드 ∪ Supabase 신규품목 등록. code로 중복제거(Supabase 우선)."""
    seed = []
    if NEWITEMS_SEED.exists():
        try:
            seed = json.loads(NEWITEMS_SEED.read_text(encoding="utf-8"))
        except Exception:
            seed = []
    extra = []
    if use_supabase():
        try:
            r = _sb().table("app_settings").select("value").eq("key", NKEY).execute()
            if r.data and r.data[0].get("value"):
                extra = r.data[0]["value"] or []
        except Exception:
            pass
    merged = {}
    for r in list(seed) + list(extra):
        c = str(r.get("code", "")).strip()
        if c and str(r.get("analog", "")).strip():
            merged[c] = {"code": c, "nm": r.get("nm", ""),
                         "ip": r.get("ip"), "plt": r.get("plt"),
                         "analog": str(r.get("analog")).strip(),
                         "factor": r.get("factor") or 1.0}
    return list(merged.values())


def add_new_item(code, analog, ip=None, plt=None, nm="", factor=1.0) -> bool:
    """신규품목 초기기준 1건 Supabase 저장(동일 code 있으면 갱신)."""
    cur = []
    if use_supabase():
        try:
            r = _sb().table("app_settings").select("value").eq("key", NKEY).execute()
            if r.data and r.data[0].get("value"):
                cur = r.data[0]["value"] or []
        except Exception:
            pass
    cur = [c for c in cur if str(c.get("code")) != str(code)]
    cur.append({"code": str(code), "analog": str(analog), "nm": nm,
                "ip": int(ip) if ip else None, "plt": int(plt) if plt else None,
                "factor": float(factor) if factor else 1.0})
    if not use_supabase():
        return False
    try:
        _sb().table("app_settings").upsert({"key": NKEY, "value": cur}).execute()
        return True
    except Exception:
        return False


def remove_new_item(code) -> bool:
    """신규품목 등록 해제(자립 후 정리 등)."""
    if not use_supabase():
        return False
    try:
        r = _sb().table("app_settings").select("value").eq("key", NKEY).execute()
        cur = (r.data[0].get("value") or []) if r.data else []
        cur = [c for c in cur if str(c.get("code")) != str(code)]
        _sb().table("app_settings").upsert({"key": NKEY, "value": cur}).execute()
        return True
    except Exception:
        return False


# ---------- BNF 거래처(행사 CHANNEL) 화이트리스트 ----------
CHKEY = "bnf_channels"
BNF_CH_SEED = DATA / "bnf_channels_seed.json"


def load_bnf_channels() -> list:
    """행사 필터용 BNF 거래처 키워드(부분일치). 번들 시드 ∪ Supabase(있으면 대체)."""
    seed = []
    if BNF_CH_SEED.exists():
        try:
            seed = json.loads(BNF_CH_SEED.read_text(encoding="utf-8"))
        except Exception:
            seed = []
    if use_supabase():
        try:
            r = _sb().table("app_settings").select("value").eq("key", CHKEY).execute()
            if r.data and r.data[0].get("value") is not None:
                return [str(x).strip() for x in (r.data[0]["value"] or []) if str(x).strip()]
        except Exception:
            pass
    return [str(x).strip() for x in seed if str(x).strip()]


def save_bnf_channels(channels) -> bool:
    clean = [str(x).strip() for x in channels if str(x).strip()]
    if not use_supabase():
        return False
    try:
        _sb().table("app_settings").upsert({"key": CHKEY, "value": clean}).execute()
        return True
    except Exception:
        return False


# ---------- 상습 초과 품목 행사수량 보정 배수 ----------
EUKEY = "event_uplift"
EU_SEED = DATA / "event_uplift_seed.json"


def load_event_uplift() -> dict:
    """{품목코드: 배수} — 상습 초과 품목의 행사수량 상향. 번들시드 ∪ Supabase(우선)."""
    seed = {}
    if EU_SEED.exists():
        try:
            seed = {str(k): float(v) for k, v in json.loads(EU_SEED.read_text(encoding="utf-8")).items()}
        except Exception:
            seed = {}
    if use_supabase():
        try:
            r = _sb().table("app_settings").select("value").eq("key", EUKEY).execute()
            if r.data and r.data[0].get("value") is not None:
                return {str(k): float(v) for k, v in (r.data[0]["value"] or {}).items()}
        except Exception:
            pass
    return seed


def save_event_uplift(uplift: dict) -> bool:
    clean = {str(k).strip(): float(v) for k, v in uplift.items() if str(k).strip() and float(v) > 0}
    if not use_supabase():
        return False
    try:
        _sb().table("app_settings").upsert({"key": EUKEY, "value": clean}).execute()
        return True
    except Exception:
        return False


# ---------- BNF 할당 누적사용량 (기간 총량 제한) ----------
AUKEY = "bnf_alloc_used"


def load_alloc_used() -> dict:
    """{품목코드: {"used": 누적이동EA, "end": 할당종료일ISO}}. 종료일 바뀌면 새 기간(리셋)."""
    if use_supabase():
        try:
            r = _sb().table("app_settings").select("value").eq("key", AUKEY).execute()
            if r.data and r.data[0].get("value") is not None:
                return r.data[0]["value"] or {}
        except Exception:
            pass
    return {}


def save_alloc_used(used: dict) -> bool:
    if not use_supabase():
        return False
    try:
        _sb().table("app_settings").upsert({"key": AUKEY, "value": used}).execute()
        return True
    except Exception:
        return False


def load_master() -> dict:
    if MASTER.exists():
        with open(MASTER, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_master(master: dict):
    with open(MASTER, "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False)


# ---------- 안전재고 기준 (계산기 → 이동계획 공유) ----------
BKEY = "ss_baseline_b64"     # 품목별 안전재고·Min·Max·입수·하대·일평균
BMETA = "ss_baseline_meta"
BASELINE = DATA / "baseline_seed.json.gz"


def _encode_obj(obj) -> str:
    return base64.b64encode(gzip.compress(
        json.dumps(obj, ensure_ascii=False).encode("utf-8"), 6)).decode("ascii")


def save_baseline(rows: list, settings: dict) -> bool:
    """계산 결과(rows)에서 이동계획용 기준만 뽑아 저장(Supabase + 번들 시드)."""
    base = {r["품목코드"]: {
        "nm": r.get("품목명", ""), "ip": r["입수"], "plt": r.get("하대박스수"),
        "rate": r["일평균(딜제외)"], "ss": r["안전재고_EA"],
        "mn": r["Min(발주점,EA)"], "mx": r["Max(목표재고,EA)"],
        "months": r.get("months"),   # 계절(월별) Min/Max·SS 프로파일
        "note": r.get("비고", ""),
    } for r in rows}
    meta = {"품목수": len(base),
            "갱신": datetime.now(KST).strftime("%Y-%m-%d %H:%M"),
            "설정": settings}
    payload = {"base": base, "meta": meta}
    # 번들 시드도 갱신(로컬/최초 배포용)
    try:
        with gzip.open(BASELINE, "wt", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
    except Exception:
        pass
    if not use_supabase():
        return False
    try:
        _sb().table("app_settings").upsert({"key": BKEY, "value": _encode_obj(base)}).execute()
        _sb().table("app_settings").upsert({"key": BMETA, "value": meta}).execute()
        return True
    except Exception:
        return False


def load_baseline() -> tuple:
    """(base dict, meta) — Supabase 우선, 없으면 번들 시드."""
    if use_supabase():
        try:
            r = _sb().table("app_settings").select("value").eq("key", BKEY).execute()
            m = _sb().table("app_settings").select("value").eq("key", BMETA).execute()
            if r.data and r.data[0].get("value"):
                base = _decode(r.data[0]["value"])
                meta = m.data[0].get("value") if m.data else {}
                return base, meta
        except Exception:
            pass
    if BASELINE.exists():
        with gzip.open(BASELINE, "rt", encoding="utf-8") as f:
            p = json.load(f)
        return p.get("base", {}), p.get("meta", {})
    return {}, {}
