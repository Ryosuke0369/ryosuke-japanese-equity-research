"""
jp_calendar.py — 日本の営業日カレンダー（祝日・振替休日・国民の休日対応）
決算カレンダー推定の「±3営業日」計算に使用。外部依存なし。
"""
from datetime import date, timedelta


def _nth_monday(year: int, month: int, n: int) -> date:
    d = date(year, month, 1)
    offset = (0 - d.weekday()) % 7  # 月曜=0
    return d + timedelta(days=offset + 7 * (n - 1))


def _equinox_day(year: int, spring: bool) -> int:
    """春分/秋分の日（2000-2099年の近似式）"""
    if spring:
        return int(20.8431 + 0.242194 * (year - 1980)) - int((year - 1980) / 4)
    return int(23.2488 + 0.242194 * (year - 1980)) - int((year - 1980) / 4)


def national_holidays(year: int) -> set:
    """指定年の祝日（振替休日・国民の休日込み）を返す"""
    h = set()
    h.add(date(year, 1, 1))                       # 元日
    h.add(_nth_monday(year, 1, 2))                # 成人の日
    h.add(date(year, 2, 11))                      # 建国記念の日
    h.add(date(year, 2, 23))                      # 天皇誕生日
    h.add(date(year, 3, _equinox_day(year, True)))   # 春分の日
    h.add(date(year, 4, 29))                      # 昭和の日
    h.add(date(year, 5, 3))                       # 憲法記念日
    h.add(date(year, 5, 4))                       # みどりの日
    h.add(date(year, 5, 5))                       # こどもの日
    h.add(_nth_monday(year, 7, 3))                # 海の日
    h.add(date(year, 8, 11))                      # 山の日
    h.add(_nth_monday(year, 9, 3))                # 敬老の日
    h.add(date(year, 9, _equinox_day(year, False)))  # 秋分の日
    h.add(_nth_monday(year, 10, 2))               # スポーツの日
    h.add(date(year, 11, 3))                      # 文化の日
    h.add(date(year, 11, 23))                     # 勤労感謝の日

    # 国民の休日（祝日に挟まれた平日）
    sorted_h = sorted(h)
    for a, b in zip(sorted_h, sorted_h[1:]):
        if (b - a).days == 2:
            mid = a + timedelta(days=1)
            if mid.weekday() < 5 and mid not in h:
                h.add(mid)

    # 振替休日（日曜の祝日→翌平日）
    extra = set()
    for d in sorted(h):
        if d.weekday() == 6:  # 日曜
            sub = d + timedelta(days=1)
            while sub in h or sub in extra:
                sub += timedelta(days=1)
            extra.add(sub)
    h |= extra
    return h


_HOLIDAY_CACHE = {}


def _holidays(year: int) -> set:
    if year not in _HOLIDAY_CACHE:
        _HOLIDAY_CACHE[year] = national_holidays(year)
    return _HOLIDAY_CACHE[year]


def is_business_day(d: date) -> bool:
    return d.weekday() < 5 and d not in _holidays(d.year)


def add_business_days(d: date, n: int) -> date:
    """n営業日後（負値で前）。起点日は含まない"""
    step = 1 if n >= 0 else -1
    cur, remaining = d, abs(n)
    while remaining > 0:
        cur += timedelta(days=step)
        if is_business_day(cur):
            remaining -= 1
    return cur


def business_days_between(start: date, end: date) -> int:
    """start〜end間の営業日数（end自身は含む、startは含まない）。end<startなら負"""
    if end < start:
        return -business_days_between(end, start)
    cnt, cur = 0, start
    while cur < end:
        cur += timedelta(days=1)
        if is_business_day(cur):
            cnt += 1
    return cnt


def snap_to_business_day(d: date, prefer: str = "backward") -> date:
    """非営業日を最寄りの営業日に寄せる"""
    if is_business_day(d):
        return d
    fwd = d
    while not is_business_day(fwd):
        fwd += timedelta(days=1)
    bwd = d
    while not is_business_day(bwd):
        bwd -= timedelta(days=1)
    return fwd if prefer == "forward" else bwd
