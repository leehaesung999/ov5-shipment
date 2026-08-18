# -*- coding: utf-8 -*-
"""재고이동계획 코어 (순수 로직).

이동_박스 = MIN( 요청, 출고가능, 할당 ) 을 박스단위로. 종료품목은 0.
  · 요청_박스   = ROUNDUP( MAX(정상보충 + 이벤트 - 입고예정, 0) / 입수 )
  · 정상보충    = 현재고 ≤ Min 이면 (Max - 현재고), 아니면 0
  · 이동가능_박스 = MIN( 출고가능_박스, 내림(할당/입수) )
  · 파레트환산   = 이동_박스 / 하대박스수

입력(앱이 업로드→dict로 정리해서 넘김):
  baseline = {코드: {ip, plt, rate, ss, mn, mx, nm, note}}   (안전재고 계산기 산출)
  stock    = {코드: {"cur": 현재고EA, "alloc": 할당EA|None}}   (재고입력)
  avail    = {코드: 출고가능_박스}  또는 None(=출고가능 미적용)
  incoming = {코드: 입고예정EA}
  events   = {코드: 이벤트추가EA}   (수동 + 행사 프리쉽 반영분)
  ended    = set(코드)              (종료품목)
"""
from __future__ import annotations
import math

INF = 9e15


def _season(b, plan_month):
    """계획월의 계절 프로파일(mn,mx,ss) 반환. months 없으면 연 고정값 폴백.
    months 키는 JSON 저장으로 문자열('1'~'12')일 수 있어 양쪽 조회."""
    if plan_month:
        mp = b.get("months")
        if mp:
            mm = mp.get(str(plan_month)) or mp.get(plan_month)
            if mm:
                return mm.get("mn"), mm.get("mx"), mm.get("ss")
    return b.get("mn"), b.get("mx"), b.get("ss")


def compute_transfer(baseline, stock, avail=None, incoming=None,
                     events=None, ended=None, plan_month=None):
    stock = stock or {}
    incoming = incoming or {}
    events = events or {}
    ended = ended or set()
    rows = []
    for code, b in baseline.items():
        ip = b.get("ip") or 1
        plt = b.get("plt")
        mn, mx, ss_m = _season(b, plan_month)   # 계절(계획월) 기준
        s = stock.get(code, {})
        cur = s.get("cur")
        alloc = s.get("alloc")
        evt = events.get(code, 0) or 0
        inc = incoming.get(code, 0) or 0
        is_end = code in ended

        if cur is None:                     # 재고 미입력
            rows.append(_row(code, b, ip, plt, cur, mn, mx, None, evt, inc,
                             None, alloc, None, 0, 0, "미입력"))
            continue

        정상보충 = max(mx - cur, 0) if cur <= mn else 0
        요청_박스 = math.ceil(max(정상보충 + evt - inc, 0) / ip)

        # 출고가능_박스: avail None=무제한 / dict에 없으면 0
        if avail is None:
            av_box = INF
        else:
            av_box = avail.get(code, 0)
        al_box = INF if alloc is None else math.floor(alloc / ip)
        이동가능_박스 = min(av_box, al_box)

        if is_end:
            이동_박스 = 0
        else:
            이동_박스 = min(요청_박스, 이동가능_박스)
        미충족_박스 = 0 if is_end else max(요청_박스 - 이동_박스, 0)

        # 사유
        if is_end:
            사유 = "종료(제외)"
        elif 이동_박스 < 요청_박스:
            사유 = "할당제한" if al_box <= av_box else "출고가능제한"
        elif 정상보충 > 0 and evt > 0:
            사유 = "발주점미달+이벤트"
        elif evt > 0:
            사유 = "이벤트"
        elif 정상보충 > 0:
            사유 = "발주점미달"
        else:
            사유 = "충분"

        rows.append(_row(code, b, ip, plt, cur, mn, mx, 요청_박스, evt, inc,
                         (None if av_box >= INF else av_box), alloc,
                         이동가능_박스, 이동_박스, 미충족_박스, 사유, ss_eff=ss_m))
    # 정렬: 종료 최하 → 미충족 → 이벤트 → 발주점미달 → 나머지 → 미입력 최하
    order = {"할당제한": 0, "출고가능제한": 0, "발주점미달+이벤트": 1, "발주점미달": 2,
             "이벤트": 3, "충분": 8, "종료(제외)": 9, "미입력": 9}
    rows.sort(key=lambda r: (order.get(r["사유"], 5),
                             -(r["★이동_박스"] if isinstance(r["★이동_박스"], (int, float)) else 0)))
    return rows


def _row(code, b, ip, plt, cur, mn, mx, req, evt, inc, av, alloc, movable,
         move_box, short_box, reason, ss_eff=None):
    move_ea = (move_box * ip) if move_box else 0
    # 소진경고: 현재고가 안전재고(SS) 밑으로 = 예상보다 빨리 소진(행사 초과 등).
    #   보충은 매일 Max까지 따라가지만, 리드타임(수일) 동안은 SS가 완충. SS를
    #   이미 까먹었다면 리드타임 내 결품 위험 → 사람이 확인하라는 신호(공급 로직은 불변).
    #   ss_eff = 계획월의 계절 SS(있으면), 없으면 연 고정 SS.
    ss = ss_eff if ss_eff is not None else (b.get("ss") or 0)
    if cur is None or reason == "종료(제외)":
        warn = ""
    elif cur <= 0:
        warn = "🔴 결품(현재고 0이하)"
    elif ss and cur <= ss:
        warn = f"🟠 안전재고({int(ss)}) 이하 — 소진 빠름(행사 초과 등 점검)"
    else:
        warn = ""
    return {
        "품목코드": code,
        "품목명": b.get("nm", ""),
        "입수": ip,
        "현재고": cur,
        "Min": mn,
        "Max": mx,
        "이벤트": evt or "",
        "입고예정": inc or "",
        "요청_박스": req if req is not None else "",
        "출고가능_박스": av if av is not None else "",
        "할당(EA)": alloc if alloc is not None else "",
        "★이동_박스": move_box if cur is not None else "",
        "이동_EA": move_ea if cur is not None else "",
        "미충족_박스": short_box if cur is not None else "",
        "하대박스수": plt,
        "파레트환산": round(move_box / plt, 2) if (plt and move_box) else (0 if plt else None),
        "사유": reason,
        "소진경고": warn,
        "비고": b.get("note", ""),
    }


# ---------- 창고 배정 (한 창고에서만, 창고 우선순위, 부족시 경고) ----------
# 배정 우선순위: IC930 → IC100 → IC920 (930부터, 없으면 100, 그다음 920)
WH_PRIORITY = {"IC930": 0, "IC100": 1, "IC920": 2}
WH_SOURCES = ("IC930", "IC100", "IC920")


def allocate_warehouse(rows, loc_inv, sources=WH_SOURCES):
    """이동_박스>0 각 품목에 출고 창고 1개 배정.
    loc_inv = {코드: {창고: {"avail": 출고가능박스, "exp": 최단소비기한(비교가능값)}}}
    규칙: 단독으로 이동량 채우는 창고 중 우선순위 IC930→IC100→IC920 로 선택(유통기한 무관).
          예) 930이 채울 수 있으면 100/920에 더 빠른 재고 있어도 930. 없으면 배정=분할필요(경고).
    각 row에 컬럼 추가: 배정창고 / 배정재고(Box) / 최단유통기한 / 창고재고표기 / 창고경고
    """
    prio = {w: WH_PRIORITY.get(w, 9) for w in sources}
    for r in rows:
        mb = r.get("★이동_박스")
        if not isinstance(mb, (int, float)) or mb <= 0:
            r["배정창고"] = ""; r["배정재고(Box)"] = ""; r["최단유통기한"] = ""
            r["창고재고"] = ""; r["창고경고"] = ""
            continue
        whs = (loc_inv.get(r["품목코드"]) or {})
        # 소스 창고만
        cand = {w: whs[w] for w in sources if w in whs and (whs[w].get("avail") or 0) > 0}
        note = " ".join(f"{w}:{int(cand[w]['avail'])}" for w in sources if w in cand) or "재고없음"
        # 단독 가능 창고 → 창고 우선순위(IC930→IC100→IC920)로 선택. 유통기한 무관.
        solo = [w for w in cand if (cand[w].get("avail") or 0) >= mb]
        if solo:
            solo.sort(key=lambda w: prio.get(w, 9))
            best = solo[0]
            r["배정창고"] = best
            r["배정재고(Box)"] = int(cand[best]["avail"])
            r["최단유통기한"] = _fmt_exp(cand[best].get("exp"))
            r["창고재고"] = note
            r["창고경고"] = ""
        else:
            total = sum((cand[w].get("avail") or 0) for w in cand)
            r["배정창고"] = "분할필요" if cand else "재고없음"
            r["배정재고(Box)"] = int(total)
            r["최단유통기한"] = ""
            r["창고재고"] = note
            r["창고경고"] = (f"⚠️ 단독가능 창고없음(이동 {int(mb)}박스 > 각 창고재고) — 분할검토"
                            if cand else f"⚠️ 소스창고에 재고없음(이동 {int(mb)}박스)")
    return rows


def _fmt_exp(v):
    if v is None:
        return ""
    s = str(int(v)) if isinstance(v, (int, float)) else str(v)
    return f"{s[:4]}-{s[4:6]}-{s[6:8]}" if len(s) == 8 and s.isdigit() else s


def summarize(rows):
    mv = [r for r in rows if isinstance(r["★이동_박스"], (int, float)) and r["★이동_박스"] > 0]
    return {
        "이동품목수": len(mv),
        "총이동_박스": sum(r["★이동_박스"] for r in mv),
        "총이동_EA": sum(r["이동_EA"] for r in mv),
        "총파레트": round(sum((r["파레트환산"] or 0) for r in mv), 1),
        "미입력": sum(1 for r in rows if r["사유"] == "미입력"),
        "가용부족": sum(1 for r in rows if r["사유"] in ("할당제한", "출고가능제한")),
        "종료제외": sum(1 for r in rows if r["사유"] == "종료(제외)"),
        "분할필요": sum(1 for r in rows if r.get("배정창고") == "분할필요"),
        "창고재고없음": sum(1 for r in rows if r.get("배정창고") == "재고없음"),
        "결품": sum(1 for r in rows if str(r.get("소진경고", "")).startswith("🔴")),
        "안전재고이하": sum(1 for r in rows if str(r.get("소진경고", "")).startswith("🟠")),
    }
