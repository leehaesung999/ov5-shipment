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
    "new_item_grad_days": 30,  # 신규품목 자립(자기 통계 전환) 임계 판매일수
    # 신규(이력 짧은)&활발 품목: 이력 없는 달 SS=0 대신 최소 N박스 바닥값.
    #   실적이 쌓이면(span↑) 자동으로 바닥값 해제 → 자기 통계로 전환.
    "min_ss_box": 10,       # 신규품목 최소 안전재고(박스)
    "new_span_days": 183,   # 시작~최신 span 이 이하면 '신규(이력짧음)'
    "new_active_days": 45,  # 최근 이 일수내 출고 있으면 '활발'
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


def _new_thin_floor(hist_item, data_max, new_span, new_active, min_ss_ea):
    """신규(이력 짧음)&활발 품목이면 최소 안전재고(EA) 반환, 아니면 0.
    · 시작(최초 출고일)이 최신일 기준 new_span 이내(=최근 시작, 이력 짧음)
    · 최근 new_active 이내 출고(=아직 활발, 단종 아님)
    실적이 쌓여 span 이 커지면 자동으로 0(바닥값 해제) → 자기 통계 전환."""
    if not hist_item or data_max is None or min_ss_ea <= 0:
        return 0
    ks = list(hist_item.keys())
    first = date.fromisoformat(min(ks))
    last = date.fromisoformat(max(ks))
    started_recent = (data_max - first).days <= new_span
    still_active = (data_max - last).days <= new_active
    return min_ss_ea if (started_recent and still_active) else 0


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


def _item_stats(hist_item, dates, ndays, exp):
    """품목 이력에서 (일평균, σ, 활성판매일수) 반환. 창내 수요 0이면 None."""
    ser = nondeal_series(hist_item, dates)
    total = sum(ser)
    if total <= 0:
        return None
    rate = total / ndays
    win = [sum(ser[i:i + exp]) for i in range(len(ser) - exp + 1)]
    sigma = _sd(win)
    active = sum(1 for d in dates if d in hist_item)
    return rate, sigma, active


def _ss_row(code, ip, plt, nm, cat, rate, sigma, z, lead, batch,
            deal_days=0, note_extra="", floor_ea=0):
    """일평균·σ 로 안전재고 산출 행 1건 생성(박스단위 정렬).
    floor_ea: 신규(이력짧음)&활발 품목 최소 안전재고(EA). 계산값이 작아도 이 값 이상."""
    no_ip = not ip
    ip = ip or 1
    ss = max(ip, floor_ea or 0, math.ceil(z * sigma / ip) * ip)   # 안전재고(박스올림, 최소 floor)
    ss_box = ss // ip
    mn = math.ceil(rate * lead + ss)                   # Min(발주점, EA)
    mx = max(math.ceil((mn + rate * batch) / ip),
             math.ceil((mn + 1) / ip)) * ip             # Max(박스배수)
    mx_box = mx // ip
    return {
        "품목코드": str(code),
        "품목명": nm or "",
        "카테고리": cat or "",
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
            note_extra,
        ] if x),
    }


def _exposure_windows(dv, exp):
    """(date, value) 정렬 리스트에서 달력상 연속한 exp일 창합계만 생성.
    달력 간격 >7일(연 경계 등)이면 창을 잇지 않음(잘못된 합 방지)."""
    wins = []
    n = len(dv)
    for i in range(n):
        seg = [dv[i][1]]
        ok = True
        for k in range(1, exp):
            if i + k >= n or (dv[i + k][0] - dv[i + k - 1][0]).days > 7:
                ok = False
                break
            seg.append(dv[i + k][1])
        if ok and len(seg) == exp:
            wins.append(sum(seg))
    return wins


def _monthly_profile(hist_item, dates, ip, z, lead, batch, exp, floor_ea=0):
    """달력월(1~12)별 {rate, ss, mn, mx} 프로파일. 그 달 날들로 직접 산출.
    비수기(수요0)는 0. 같은 달이 여러 해 있으면 함께 풀링(달력 연속만 창 구성).
    floor_ea>0(신규&활발 품목): 이력 없는 달은 SS=Max=floor(최소 유지), 있는 달은 max(계산,floor)."""
    ip = ip or 1
    fl = floor_ea or 0
    by_m = {}
    for d in dates:
        by_m.setdefault(int(d[5:7]), []).append(d)
    prof = {}
    for m in range(1, 13):
        dd = by_m.get(m)
        tot = 0.0
        if dd:
            ser = [max(0.0, (hist_item[d][0] - hist_item[d][1])) if d in hist_item else 0.0
                   for d in dd]
            tot = sum(ser)
        if not dd or tot <= 0:                 # 이력 없는 달
            prof[m] = ({"rate": 0, "ss": fl, "mn": fl, "mx": fl} if fl
                       else {"rate": 0, "ss": 0, "mn": 0, "mx": 0})
            continue
        rate = tot / len(dd)
        dv = [(date.fromisoformat(dd[i]), ser[i]) for i in range(len(dd))]
        win = _exposure_windows(dv, exp)
        sigma = _sd(win) if win else 0.0
        ss = max(ip, fl, math.ceil(z * sigma / ip) * ip)
        mn = math.ceil(rate * lead + ss)
        mx = max(math.ceil((mn + rate * batch) / ip), math.ceil((mn + 1) / ip)) * ip
        prof[m] = {"rate": round(rate, 1), "ss": ss, "mn": mn, "mx": mx}
    return prof


def compute(history, master, settings=None, uplift=None, new_items=None):
    """안전재고 산출표(품목별 dict 리스트) 반환.

    uplift: {품목코드: 계수} 결품보정(선택). 없으면 미적용.
    new_items: 신규품목 초기기준(유사품 기반) 리스트
       [{"code","nm","ip","plt","analog"(유사품목코드),"factor"(계수)}].
       판매일수가 임계(new_item_grad_days) 미만인 신규품목은 유사품 수요패턴을
       계수배해 초기 Min/Max 산출. 이력이 충분히 쌓이면 자기 통계로 자동 전환.
    """
    s = {**DEFAULT_SETTINGS, **(settings or {})}
    lead, cycle, z, batch = s["lead_time"], s["cycle"], s["z"], s["batch"]
    exp = int(lead + cycle)                    # 노출기간
    grad = int(s.get("new_item_grad_days", 30))
    uplift = uplift or {}

    # 신규품목 등록 정리 + 마스터에 입수/하대 주입(자립 시 박스정렬 위해)
    new_reg = {str(r["code"]): r for r in (new_items or [])
               if r.get("code") and r.get("analog")}
    for scode, reg in new_reg.items():
        m = dict(master.get(scode) or {})
        if reg.get("ip"):
            m["ip"] = int(reg["ip"])
        if reg.get("plt"):
            m["plt"] = int(reg["plt"])
        if reg.get("nm") and not m.get("nm"):
            m["nm"] = reg["nm"]
        master[scode] = m

    dates = _rolling_dates(history, s["window_months"])
    ndays = len(dates)
    data_max = date.fromisoformat(dates[-1]) if dates else None
    min_ss_box = int(s.get("min_ss_box", 0))
    new_span = int(s.get("new_span_days", 183))
    new_active = int(s.get("new_active_days", 45))
    floor_by = {}                               # {코드: 최소SS EA} 월별에도 동일 적용
    rows = []
    n_grad = 0
    n_floor = 0
    for code, hist_item in history.items():
        scode = str(code)
        stx = _item_stats(hist_item, dates, ndays, exp)
        if stx is None:
            continue
        rate, sigma, active = stx
        is_new = scode in new_reg
        if is_new and active < grad:
            continue                            # 아직 자립 전 → 아래서 유사품 시드
        m = master.get(scode, {})
        u = uplift.get(scode, 1.0)
        ip_i = m.get("ip") or 1
        floor_ea = _new_thin_floor(hist_item, data_max, new_span, new_active,
                                   min_ss_box * ip_i)
        floor_by[scode] = floor_ea
        deal_days = sum(1 for d in dates if (hist_item.get(d) or [0, 0])[1] > 0)
        rows.append(_ss_row(scode, m.get("ip"), m.get("plt"), m.get("nm"),
                            m.get("cat"), rate, sigma * u, z, lead, batch,
                            deal_days,
                            note_extra=(f"결품보정×{round(u,2)}" if u > 1.0 else "")
                            + (" / 신규(자립)" if is_new else "")
                            + (f" / 신규최소SS {min_ss_box}박스" if floor_ea else ""),
                            floor_ea=floor_ea))
        if is_new:
            n_grad += 1
        if floor_ea:
            n_floor += 1

    # 유사품 기반 신규품목 시드 (자립 전인 것)
    n_seed = 0
    no_analog = []
    for scode, reg in new_reg.items():
        hist_item = history.get(scode, {})
        stx = _item_stats(hist_item, dates, ndays, exp)
        if stx is not None and stx[2] >= grad:
            continue                            # 이미 자립 → 위에서 처리됨
        an = history.get(str(reg["analog"]))
        an_stx = _item_stats(an, dates, ndays, exp) if an else None
        if an_stx is None:
            no_analog.append(scode)
            continue
        f = float(reg.get("factor") or 1.0)
        m = master.get(scode, {})
        rows.append(_ss_row(scode, m.get("ip"), m.get("plt"),
                            m.get("nm") or reg.get("nm"), m.get("cat"),
                            an_stx[0] * f, an_stx[1] * f, z, lead, batch, 0,
                            note_extra=f"신규(유사 {reg['analog']}×{round(f,2)})"))
        n_seed += 1

    # 월별(계절) 프로파일 부착 — 이력 있으면 월별 산출, 없으면(신규 시드) 연값 평탄화
    for r in rows:
        hi = history.get(r["품목코드"])
        if hi:
            r["months"] = _monthly_profile(hi, dates, r["입수"], z, lead, batch, exp,
                                           floor_ea=floor_by.get(r["품목코드"], 0))
        else:
            v = {"rate": r["일평균(딜제외)"], "ss": r["안전재고_EA"],
                 "mn": r["Min(발주점,EA)"], "mx": r["Max(목표재고,EA)"]}
            r["months"] = {m: dict(v) for m in range(1, 13)}

    rows.sort(key=lambda r: -r["초기이관_EA"])
    return rows, {"기간일수": ndays,
                  "시작": dates[0] if dates else None,
                  "종료": dates[-1] if dates else None,
                  "품목수": len(rows),
                  "노출기간": exp,
                  "신규시드": n_seed,
                  "신규자립": n_grad,
                  "신규최소SS": n_floor,
                  "최소SS박스": min_ss_box,
                  "신규_유사품없음": no_analog}


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
