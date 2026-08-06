# -*- coding: utf-8 -*-
"""안전재고 계산 코어 (순수 로직, Streamlit 무관).

핵심 개념
  · 안전재고 = z × σ(노출기간 이동합계) → 박스단위 올림   (박스가 기준단위)
  · Min(발주점) = 일평균 × 리드타임 + 안전재고
  · Max(목표재고) = (Min + 일평균 × 발주배치) → 박스단위 올림
  · 노출기간 = 리드타임 + 발주주기
  · 딜(행사) 물량은 사전 별도공급 전제 → σ 계산에서 제외

데이터 모델 (누적 히스토리)
  history = { 품목코드(str): { "YYYY-MM-DD": [총출고, 딜출고] } }
    · 비딜(평상시) = 총출고 − 딜출고  ← σ는 이 값으로 계산
    · CJ출고실적: 딜 = 행사no 있는 출고(실측)
    · 초기 시드(일별판매건수): 딜 = winsorize 초과분(중앙값×배수)
  최근 12개월만 롤링 사용 → 계절 변화 자동 반영(여름 급증 등)

master = { 품목코드(str): {"ip":입수, "plt":하대박스수, "cat":카테고리, "nm":품목명} }
"""
from __future__ import annotations
import math
import statistics as st
from datetime import date, timedelta


# ---------- 설정 기본값 ----------
DEFAULT_SETTINGS = {
    "lead_time": 3,      # 리드타임(일)
    "cycle": 1,          # 발주주기(일)
    "z": 2.33,           # 서비스레벨 계수 (2.33=99%)
    "batch": 15,         # 발주배치(일)
    "window_months": 12, # 롤링 기간(개월)
}
PROMO_MULT = 4          # 시드용 딜 판정(중앙값×배수) — CJ실적엔 미사용
DIL_ABS_MIN = 30


def _sd(v):
    return st.stdev(v) if len(v) > 1 else 0.0


def _rolling_dates(history, months):
    """히스토리 전체에서 최근 N개월 날짜 리스트(오름차순) 산출."""
    all_d = set()
    for it in history.values():
        all_d.update(it.keys())
    if not all_d:
        return []
    dmax = max(date.fromisoformat(d) for d in all_d)
    cutoff = dmax - timedelta(days=int(months * 30.44))
    return sorted(d for d in all_d if date.fromisoformat(d) >= cutoff)


def nondeal_series(hist_item, dates):
    """품목의 날짜순 비딜(평상시) 일수요 리스트."""
    out = []
    for d in dates:
        v = hist_item.get(d)
        if v is None:
            out.append(0.0)
        else:
            tot, deal = (v[0] or 0), (v[1] or 0)
            out.append(max(0.0, tot - deal))
    return out


def compute(history, master, settings=None, uplift=None):
    """안전재고 산출표(품목별 dict 리스트) 반환.

    uplift: {품목코드: 계수} 결품보정(선택). 없으면 미적용.
    """
    s = {**DEFAULT_SETTINGS, **(settings or {})}
    lead, cycle, z, batch = s["lead_time"], s["cycle"], s["z"], s["batch"]
    exp = int(lead + cycle)                    # 노출기간
    uplift = uplift or {}

    dates = _rolling_dates(history, s["window_months"])
    ndays = len(dates)
    rows = []
    for code, hist_item in history.items():
        ser = nondeal_series(hist_item, dates)
        if sum(ser) <= 0:
            continue
        m = master.get(str(code), {})
        ip = m.get("ip") or 1
        no_ip = not m.get("ip")
        plt = m.get("plt")
        rate = sum(ser) / ndays                # 비딜 일평균
        u = uplift.get(str(code), 1.0)
        win = [sum(ser[i:i + exp]) for i in range(len(ser) - exp + 1)]
        sigma = _sd(win) * u
        ss = max(ip, math.ceil(z * sigma / ip) * ip)   # 안전재고(박스올림)
        ss_box = ss // ip
        mn = math.ceil(rate * lead + ss)               # Min(발주점, EA)
        mx = max(math.ceil((mn + rate * batch) / ip),
                 math.ceil((mn + 1) / ip)) * ip         # Max(박스배수)
        mx_box = mx // ip
        deal_days = sum(1 for d in dates if (hist_item.get(d) or [0, 0])[1] > 0)
        rows.append({
            "품목코드": str(code),
            "품목명": m.get("nm", ""),
            "카테고리": m.get("cat", ""),
            "입수": ip,
            "하대박스수": plt,
            "일평균(딜제외)": round(rate, 1),
            "딜발생일": deal_days,
            "안전재고_박스": ss_box,
            "안전재고_EA": ss,
            "Min(발주점,EA)": mn,
            "Max(목표재고,EA)": mx,
            "초기이관_박스": mx_box,
            "초기이관_EA": mx,
            "이관_파레트": round(mx_box / plt, 2) if plt else None,
            "비고": " / ".join(x for x in [
                ("입수미상(1로가정)" if no_ip else ""),
                ("하대박스없음" if not plt else ""),
                (f"결품보정×{round(u,2)}" if u > 1.0 else ""),
            ] if x),
        })
    rows.sort(key=lambda r: -r["초기이관_EA"])
    return rows, {"기간일수": ndays,
                  "시작": dates[0] if dates else None,
                  "종료": dates[-1] if dates else None,
                  "품목수": len(rows),
                  "노출기간": exp}


# ---------- 월별 CJ출고실적 → 히스토리 조각 ----------
NAVER_CHANNELS = ("NFA",)   # + '이벤트_네이버' 접두어


def is_naver(ch):
    return ch == "NFA" or (isinstance(ch, str) and ch.startswith("이벤트_네이버"))


def extract_cj_month(rows, col):
    """CJ출고실적 raw 행 iterable → {품목코드: {날짜: [총, 딜]}}.
    비네이버(네이버·토스 제외)만. col=열인덱스 dict(0-based):
      ship_date, code, qty, channel, event_no
    """
    out = {}
    for r in rows:
        ch = r[col["channel"]]
        if is_naver(ch) or ch == "토스쇼핑":
            continue
        sd = r[col["ship_date"]]
        code = r[col["code"]]
        qty = r[col["qty"]]
        if sd is None or code is None or not isinstance(qty, (int, float)) or qty <= 0:
            continue
        # 날짜 정규화
        if hasattr(sd, "strftime"):
            dstr = sd.strftime("%Y-%m-%d")
        else:
            dstr = str(sd)[:10]
        code = str(code).strip()
        ev = r[col["event_no"]]
        deal = qty if ev not in (None, "", 0) else 0
        d = out.setdefault(code, {}).setdefault(dstr, [0.0, 0.0])
        d[0] += qty
        d[1] += deal
    return out


def apply_code_remap(history, master, remap):
    """입수량 변경 등으로 품목코드가 바뀐 경우 구코드 이력·마스터를 신코드로 승계.

    remap = [{"old": 구코드, "new": 신코드, "ip": 신입수(선택), "plt": 신하대박스수(선택)}]
    전제: 낱개(EA) 상품은 동일하고 박스 입수만 바뀐 것 → EA 수요이력을 그대로 승계.
      · 이력: 구코드 날짜별 [총,딜] 을 신코드로 이전(같은 날짜는 합산). 구코드는 제거.
      · 마스터: 구코드 마스터(품목명·하대·카테고리)를 신코드로 복사 + 입수/하대 갱신.
    매 재계산마다 호출(멱등): 이미 승계돼 구코드가 없으면 이력이전은 건너뛰고
    마스터 입수갱신만 재적용된다. 반환: 적용 로그.
    """
    log = []
    for r in (remap or []):
        old = str(r.get("old", "")).strip()
        new = str(r.get("new", "")).strip()
        if not old or not new or old == new:
            continue
        # 1) 이력 승계 (구코드 → 신코드, 날짜별 합산)
        oh = history.pop(old, None)
        if oh:
            nh = history.setdefault(new, {})
            for d, v in oh.items():
                tot, deal = (v[0] or 0), (v[1] or 0)
                if d in nh:
                    nh[d] = [nh[d][0] + tot, nh[d][1] + deal]
                else:
                    nh[d] = [tot, deal]
        # 2) 마스터 승계 + 입수/하대 갱신 (구코드 없어도 매번 재적용)
        base_m = master.get(new) or master.get(old) or {}
        nm = dict(base_m)
        if r.get("ip"):
            nm["ip"] = int(r["ip"])
        if r.get("plt"):
            nm["plt"] = int(r["plt"])
        if nm:
            master[new] = nm
        master.pop(old, None)
        log.append({"구코드": old, "신코드": new, "신입수": nm.get("ip"),
                    "이력승계": bool(oh), "승계일수": len(oh) if oh else 0})
    return log


def merge_history(history, chunk):
    """chunk를 history에 병합(같은 날짜는 덮어쓰기 = 재업로드 대비)."""
    for code, days in chunk.items():
        h = history.setdefault(code, {})
        for dstr, val in days.items():
            h[dstr] = [round(val[0], 3), round(val[1], 3)]
    return history


def trim_history(history, months=14):
    """최근 N개월(계산은 12, 보관 여유 14)만 남기고 오래된 날짜 제거."""
    all_d = set()
    for it in history.values():
        all_d.update(it.keys())
    if not all_d:
        return history
    dmax = max(date.fromisoformat(d) for d in all_d)
    cutoff = dmax - timedelta(days=int(months * 30.44))
    for code in list(history.keys()):
        it = history[code]
        for d in list(it.keys()):
            if date.fromisoformat(d) < cutoff:
                del it[d]
        if not it:
            del history[code]
    return history
