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


def compute_transfer(baseline, stock, avail=None, incoming=None,
                     events=None, ended=None):
    stock = stock or {}
    incoming = incoming or {}
    events = events or {}
    ended = ended or set()
    rows = []
    for code, b in baseline.items():
        ip = b.get("ip") or 1
        plt = b.get("plt")
        mn, mx = b.get("mn"), b.get("mx")
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
                         이동가능_박스, 이동_박스, 미충족_박스, 사유))
    # 정렬: 종료 최하 → 미충족 → 이벤트 → 발주점미달 → 나머지 → 미입력 최하
    order = {"할당제한": 0, "출고가능제한": 0, "발주점미달+이벤트": 1, "발주점미달": 2,
             "이벤트": 3, "충분": 8, "종료(제외)": 9, "미입력": 9}
    rows.sort(key=lambda r: (order.get(r["사유"], 5),
                             -(r["★이동_박스"] if isinstance(r["★이동_박스"], (int, float)) else 0)))
    return rows


def _row(code, b, ip, plt, cur, mn, mx, req, evt, inc, av, alloc, movable,
         move_box, short_box, reason):
    move_ea = (move_box * ip) if move_box else 0
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
        "비고": b.get("note", ""),
    }


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
    }
