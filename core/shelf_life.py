from datetime import date, datetime
from typing import Optional

DAYS_PER_MONTH = 30.4
DEFAULT_THRESHOLD = 0.70
DEFAULT_MONTHS = 24


def parse_ymd(value) -> Optional[date]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip().split(".")[0]
    if not s or s.lower() == "nan":
        return None
    try:
        if len(s) == 8 and s.isdigit():
            return date(int(s[:4]), int(s[4:6]), int(s[6:8]))
        return datetime.fromisoformat(s.replace("/", "-")).date()
    except (ValueError, TypeError):
        return None


def calc_remaining_rate(consume_ymd, months: int, today: Optional[date] = None) -> Optional[float]:
    end = parse_ymd(consume_ymd)
    if end is None or not months:
        return None
    today = today or date.today()
    return (end - today).days / (months * DAYS_PER_MONTH)


def calc_production_date(consume_ymd, months: int) -> Optional[date]:
    end = parse_ymd(consume_ymd)
    if end is None or not months:
        return None
    y, m = end.year, end.month - months
    while m <= 0:
        m += 12
        y -= 1
    day = min(end.day, 28)
    return date(y, m, day)


def judge(rate: Optional[float], threshold: float = DEFAULT_THRESHOLD) -> str:
    if rate is None:
        return ""
    return "OK" if rate >= threshold else "NG"
