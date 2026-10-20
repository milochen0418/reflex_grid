"""Date and time functions."""

from __future__ import annotations

import math
from typing import List, Optional

from ..core.date import DateSystem, DateTime, datetime_from_serial, days_in_month, days_in_month_in, is_leap, serial_from_ymd, serial_in_system
from ..core.parse import parse_input
from ..core.value import Array, CellError, Value, XlError, scalar, to_number
from .registry import A, Arg, Ctx, S, fn
from .util import as_array, has, num, opt_bool, opt_num, text, to_int

D = "Date & Time"
MAX_SERIAL = 2_958_465.0
ND = (True, True, False)
NDI = (True, True, True, False)


def _serial_arg(args, i: int) -> float:
    n = num(args, i)
    if not 0.0 <= n < MAX_SERIAL + 1.0:
        raise XlError(CellError.NUM)
    return n


def _day_arg(args, i: int) -> int:
    return int(math.floor(_serial_arg(args, i)))


def _dt(sys: DateSystem, serial: float) -> DateTime:
    d = datetime_from_serial(sys, serial)
    if d is None:
        raise XlError(CellError.NUM)
    return d


def _ymd(sys: DateSystem, y: int, m: int, d: int) -> int:
    s = serial_from_ymd(sys, y, m, d)
    if s is None:
        raise XlError(CellError.NUM)
    return int(s)


def weekday0(sys: DateSystem, serial: int) -> int:
    """Weekday of a serial, 0 = Sunday."""
    return (serial + 6) % 7 if sys is DateSystem.D1900 else (serial + 5) % 7


@fn("DATE", 3, 3, D, S, "DATE(year, month, day)", "Returns the serial number of a particular date.")
def _date(args, ctx):
    y, m, d = to_int(num(args, 0)), to_int(num(args, 1)), to_int(num(args, 2))
    if not 0 <= y < 10000 or abs(m) > 1_000_000 or abs(d) > 100_000_000:
        raise XlError(CellError.NUM)
    return float(_ymd(ctx.date_system, y, m, d))


@fn("TIME", 3, 3, D, S, "TIME(hour, minute, second)", "Returns the serial number of a particular time.")
def _time(args, ctx):
    h, m, s = math.trunc(num(args, 0)), math.trunc(num(args, 1)), math.trunc(num(args, 2))
    if h > 32767 or m > 32767 or s > 32767:
        raise XlError(CellError.NUM)
    total = h * 3600 + m * 60 + s
    if total < 0:
        raise XlError(CellError.NUM)
    return (total % 86400) / 86400.0


def _part(args, ctx, f) -> float:
    s = _serial_arg(args, 0)
    return float(f(_dt(ctx.date_system, math.floor(s))))


@fn("YEAR", 1, 1, D, S, "YEAR(serial_number)", "Converts a serial number to a year.")
def _year(args, ctx):
    return _part(args, ctx, lambda d: d.year)


@fn("MONTH", 1, 1, D, S, "MONTH(serial_number)", "Converts a serial number to a month.")
def _month(args, ctx):
    return _part(args, ctx, lambda d: d.month)


@fn("DAY", 1, 1, D, S, "DAY(serial_number)", "Converts a serial number to a day of the month.")
def _day(args, ctx):
    return _part(args, ctx, lambda d: d.day)


def _day_seconds(args) -> int:
    s = _serial_arg(args, 0)
    secs = int(math.floor((s - math.floor(s)) * 86400.0 + 0.5))
    return 0 if secs >= 86400 else secs


@fn("HOUR", 1, 1, D, S, "HOUR(serial_number)", "Converts a serial number to an hour.")
def _hour(args, ctx):
    return float(_day_seconds(args) // 3600)


@fn("MINUTE", 1, 1, D, S, "MINUTE(serial_number)", "Converts a serial number to a minute.")
def _minute(args, ctx):
    return float(_day_seconds(args) // 60 % 60)


@fn("SECOND", 1, 1, D, S, "SECOND(serial_number)", "Converts a serial number to a second.")
def _second(args, ctx):
    return float(_day_seconds(args) % 60)


@fn("NOW", 0, 0, D, S, "NOW()", "Returns the serial number of the current date and time.", volatile=True)
def _now(args, ctx):
    return serial_in_system(ctx.date_system, ctx.now_serial())


@fn("TODAY", 0, 0, D, S, "TODAY()", "Returns the serial number of today's date.", volatile=True)
def _today(args, ctx):
    return float(math.floor(serial_in_system(ctx.date_system, ctx.now_serial())))


@fn("WEEKDAY", 1, 2, D, S, "WEEKDAY(serial_number, [return_type])", "Converts a serial number to a day of the week.")
def _weekday(args, ctx):
    s = _day_arg(args, 0)
    t = to_int(opt_num(args, 1, 1.0)) if has(args, 1) else 1
    wd = weekday0(ctx.date_system, s)
    if t in (1, 17):
        r = wd + 1
    elif t in (2, 11):
        r = (wd + 6) % 7 + 1
    elif t == 3:
        r = (wd + 6) % 7
    elif 12 <= t <= 16:
        r = (wd + 7 - (t - 10)) % 7 + 1
    else:
        raise XlError(CellError.NUM)
    return float(r)


def _isoweek(sys: DateSystem, s: int) -> int:
    iso_wd = (weekday0(sys, s) + 6) % 7 + 1
    thursday = s - iso_wd + 4
    y = 1900 if thursday < 1 else _dt(sys, float(thursday)).year
    jan1 = _ymd(sys, y, 1, 1)
    return (thursday - jan1) // 7 + 1


@fn("WEEKNUM", 1, 2, D, S, "WEEKNUM(serial_number, [return_type])", "Converts a serial number to a number representing where the week falls numerically with a year.")
def _weeknum(args, ctx):
    sys = ctx.date_system
    s = _day_arg(args, 0)
    t = to_int(opt_num(args, 1, 1.0)) if has(args, 1) else 1
    if t in (1, 17):
        start = 0
    elif t in (2, 11):
        start = 1
    elif 12 <= t <= 16:
        start = t - 10
    elif t == 21:
        return float(_isoweek(sys, s))
    else:
        raise XlError(CellError.NUM)
    y = _dt(sys, float(s)).year
    jan1 = serial_from_ymd(sys, y, 1, 1)
    jan1 = int(jan1) if jan1 is not None else min(s, 1)
    offset = (weekday0(sys, jan1) + 7 - start) % 7
    return float((s - jan1 + offset) // 7 + 1)


@fn("ISOWEEKNUM", 1, 1, D, S, "ISOWEEKNUM(date)", "Returns the number of the ISO week number of the year for a given date.")
def _isoweeknum(args, ctx):
    return float(_isoweek(ctx.date_system, _day_arg(args, 0)))


def _months_arg(args, i: int) -> int:
    m = to_int(num(args, i))
    if abs(m) > 200_000:
        raise XlError(CellError.NUM)
    return m


@fn("EDATE", 2, 2, D, S, "EDATE(start_date, months)", "Returns the serial number of the date that is the indicated number of months before or after the start date.")
def _edate(args, ctx):
    sys = ctx.date_system
    d = _dt(sys, float(_day_arg(args, 0)))
    m = _months_arg(args, 1)
    total = d.year * 12 + d.month - 1 + m
    y, mo = total // 12, total % 12 + 1
    if not 1900 <= y <= 9999:
        raise XlError(CellError.NUM)
    # Excel's EDATE and EOMONTH keep day 0 (an empty date) and give February 1900 its 28 days.
    dd = min(d.day, days_in_month(y, mo))
    return float(_ymd(sys, y, mo, dd))


@fn("EOMONTH", 2, 2, D, S, "EOMONTH(start_date, months)", "Returns the serial number of the last day of the month before or after a specified number of months.")
def _eomonth(args, ctx):
    sys = ctx.date_system
    d = _dt(sys, float(_day_arg(args, 0)))
    m = _months_arg(args, 1)
    total = d.year * 12 + d.month - 1 + m
    y, mo = total // 12, total % 12 + 1
    if not 1900 <= y <= 9999:
        raise XlError(CellError.NUM)
    return float(_ymd(sys, y, mo, days_in_month(y, mo)))


@fn("DATEDIF", 3, 3, D, S, "DATEDIF(start_date, end_date, unit)", "Counts whole years, months or days between two dates. unit: Y, M, D, MD, YM or YD.")
def _datedif(args, ctx):
    sys = ctx.date_system
    s, e = _day_arg(args, 0), _day_arg(args, 1)
    unit = text(args, 2).upper()
    if s > e:
        raise XlError(CellError.NUM)
    d1, d2 = _dt(sys, float(s)), _dt(sys, float(e))
    months = (d2.year - d1.year) * 12 + d2.month - d1.month - (1 if d2.day < d1.day else 0)
    if unit == "D":
        r = e - s
    elif unit == "M":
        r = months
    elif unit == "Y":
        r = months // 12
    elif unit == "YM":
        r = months % 12
    elif unit == "MD":
        if d2.day >= d1.day:
            r = d2.day - d1.day
        else:
            py, pm = (d2.year - 1, 12) if d2.month == 1 else (d2.year, d2.month - 1)
            r = days_in_month(py, pm) - d1.day + d2.day
    elif unit == "YD":
        later = (d1.month, d1.day) > (d2.month, d2.day)
        y = d2.year - 1 if later else d2.year
        dd = min(d1.day, max(days_in_month(y, d1.month), 1))
        r = e - _ymd(sys, y, d1.month, dd)
    else:
        raise XlError(CellError.NUM)
    return float(r)


def _parsed_text(args, ctx):
    v = scalar(args[0].value)
    if not isinstance(v, str):
        raise XlError(v if isinstance(v, CellError) else CellError.VALUE)
    p = parse_input(v, ctx.date_system)
    if isinstance(p.value, float) and p.format is not None and any(ch in p.format for ch in "dyh"):
        return p.value, p.format
    raise XlError(CellError.VALUE)


@fn("DATEVALUE", 1, 1, D, S, "DATEVALUE(date_text)", "Converts a date in the form of text to a serial number.")
def _datevalue(args, ctx):
    n, f = _parsed_text(args, ctx)
    if not any(ch in f for ch in "dy"):
        raise XlError(CellError.VALUE)
    return float(math.floor(n))


@fn("TIMEVALUE", 1, 1, D, S, "TIMEVALUE(time_text)", "Converts a time in the form of text to a serial number.")
def _timevalue(args, ctx):
    n, _ = _parsed_text(args, ctx)
    return n - math.floor(n)


@fn("DAYS", 2, 2, D, S, "DAYS(end_date, start_date)", "Returns the number of days between two dates.")
def _days(args, ctx):
    return float(_day_arg(args, 0) - _day_arg(args, 1))


def _last_of_feb(d: DateTime) -> bool:
    return d.month == 2 and d.day == days_in_month(d.year, 2)


def days360_core(d1: DateTime, d2: DateTime, european: bool) -> int:
    a, b = d1.day, d2.day
    if european:
        a, b = min(a, 30), min(b, 30)
    else:
        if a == 31 or _last_of_feb(d1):
            a = 30
        if b == 31 and a >= 30:
            b = 30
    return (d2.year - d1.year) * 360 + (d2.month - d1.month) * 30 + b - a


@fn("DAYS360", 2, 3, D, S, "DAYS360(start_date, end_date, [method])", "Calculates the number of days between two dates based on a 360-day year.")
def _days360(args, ctx):
    sys = ctx.date_system
    d1, d2 = _dt(sys, float(_day_arg(args, 0))), _dt(sys, float(_day_arg(args, 1)))
    return float(days360_core(d1, d2, opt_bool(args, 2, False)))


def _weekend_mask(a: Optional[Arg]) -> List[bool]:
    default = [True, False, False, False, False, False, True]
    if a is None:
        return default
    v = scalar(a.value)
    if v is None and not a.from_ref:
        return default
    if isinstance(v, str):
        if len(v) != 7 or any(c not in "01" for c in v) or all(c == "1" for c in v):
            raise XlError(CellError.VALUE)
        m = [False] * 7
        for i, c in enumerate(v):
            m[(i + 1) % 7] = c == "1"
        return m
    if isinstance(v, CellError):
        raise XlError(v)
    k = to_int(to_number(v))
    m = [False] * 7
    if 1 <= k <= 7:
        m[(k + 5) % 7] = True
        m[(k + 6) % 7] = True
    elif 11 <= k <= 17:
        m[k - 11] = True
    else:
        raise XlError(CellError.NUM)
    return m


def _holidays(a: Optional[Arg]) -> List[int]:
    out = set()
    if a is not None:
        for x in as_array(a.value).data:
            if x is None:
                continue
            if isinstance(x, CellError):
                raise XlError(x)
            n = to_number(x)
            if not 0 <= n < MAX_SERIAL + 1:
                raise XlError(CellError.NUM)
            out.add(int(math.floor(n)))
    return sorted(out)


def _count_workdays(sys, a: int, b: int, mask: List[bool]) -> int:
    per_week = sum(1 for w in mask if not w)
    total = b - a + 1
    weeks = total // 7
    n = weeks * per_week
    d = a + weeks * 7
    while d <= b:
        if not mask[weekday0(sys, d)]:
            n += 1
        d += 1
    return n


def _networkdays_core(args, ctx, mask, hol):
    sys = ctx.date_system
    s, e = _day_arg(args, 0), _day_arg(args, 1)
    hs = _holidays(hol)
    lo, hi, sign = (s, e, 1) if s <= e else (e, s, -1)
    n = _count_workdays(sys, lo, hi, mask)
    n -= sum(1 for h in hs if lo <= h <= hi and not mask[weekday0(sys, h)])
    return float(n * sign)


@fn("NETWORKDAYS", 2, 3, D, ND, "NETWORKDAYS(start_date, end_date, [holidays])", "Returns the number of whole workdays between two dates.")
def _networkdays(args, ctx):
    return _networkdays_core(args, ctx, _weekend_mask(None), args[2] if len(args) > 2 else None)


@fn("NETWORKDAYS.INTL", 2, 4, D, NDI, "NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])", "Returns the number of whole workdays between two dates using parameters to indicate which and how many days are weekend days.")
def _networkdays_intl(args, ctx):
    mask = _weekend_mask(args[2] if len(args) > 2 else None)
    return _networkdays_core(args, ctx, mask, args[3] if len(args) > 3 else None)


def _advance(sys, start: int, n: int, d: int, mask: List[bool], per_week: int) -> int:
    weeks = (n - 1) // per_week
    cur = start + d * 7 * weeks
    rem = n - weeks * per_week
    while rem > 0:
        cur += d
        if not mask[weekday0(sys, cur)]:
            rem -= 1
    return cur


def _workday_core(args, ctx, mask, hol):
    sys = ctx.date_system
    start = _day_arg(args, 0)
    days = math.trunc(num(args, 1))
    if abs(days) > 1e7:
        raise XlError(CellError.NUM)
    hs = _holidays(hol)
    per_week = sum(1 for w in mask if not w)
    if per_week == 0:
        raise XlError(CellError.VALUE)
    if days == 0:
        return float(start)
    d = 1 if days > 0 else -1
    cur = start
    rem = abs(days)
    while rem > 0:
        nxt = _advance(sys, cur, rem, d, mask, per_week)
        lo, hi = (cur + 1, nxt) if d > 0 else (nxt, cur - 1)
        rem = sum(1 for h in hs if lo <= h <= hi and not mask[weekday0(sys, h)])
        cur = nxt
        if not 0 <= cur <= MAX_SERIAL:
            raise XlError(CellError.NUM)
    return float(cur)


@fn("WORKDAY", 2, 3, D, ND, "WORKDAY(start_date, days, [holidays])", "Returns the serial number of the date before or after a specified number of workdays.")
def _workday(args, ctx):
    return _workday_core(args, ctx, _weekend_mask(None), args[2] if len(args) > 2 else None)


@fn("WORKDAY.INTL", 2, 4, D, NDI, "WORKDAY.INTL(start_date, days, [weekend], [holidays])", "Returns the serial number of the date before or after a specified number of workdays using parameters to indicate which and how many days are weekend days.")
def _workday_intl(args, ctx):
    mask = _weekend_mask(args[2] if len(args) > 2 else None)
    return _workday_core(args, ctx, mask, args[3] if len(args) > 3 else None)


def _feb29_between(sys, s: int, e: int, d1: DateTime, d2: DateTime) -> bool:
    for y in range(d1.year, d2.year + 1):
        if is_leap(y):
            f = serial_from_ymd(sys, y, 2, 29)
            if f is not None and s <= f <= e:
                return True
    return False


def yearfrac_core(sys: DateSystem, s: int, e: int, basis: int) -> float:
    if s > e:
        s, e = e, s
    d1, d2 = _dt(sys, float(s)), _dt(sys, float(e))
    days = float(e - s)
    if basis == 0:
        a, b = d1.day, d2.day
        if a == 31 and b == 31:
            a = b = 30
        elif a == 31:
            a = 30
        elif a == 30 and b == 31:
            b = 30
        elif _last_of_feb(d1) and _last_of_feb(d2):
            a = b = 30
        elif _last_of_feb(d1):
            a = 30
        return ((d2.year - d1.year) * 360 + (d2.month - d1.month) * 30 + b - a) / 360.0
    if basis == 1:
        within = d1.year == d2.year or (d2.year == d1.year + 1 and (d1.month > d2.month or (d1.month == d2.month and d1.day >= d2.day)))
        if within:
            leap = (d1.year == d2.year and is_leap(d1.year)) or _feb29_between(sys, s, e, d1, d2) or (d2.month == 2 and d2.day == 29)
            return days / (366.0 if leap else 365.0)
        years = float(d2.year - d1.year + 1)
        a = serial_from_ymd(sys, d2.year + 1, 1, 1)
        b = serial_from_ymd(sys, d1.year, 1, 1)
        span = (a if a is not None else e + 1) - (b if b is not None else s)
        return days / (span / years)
    if basis == 2:
        return days / 360.0
    if basis == 3:
        return days / 365.0
    if basis == 4:
        return days360_core(d1, d2, True) / 360.0
    raise XlError(CellError.NUM)


@fn("YEARFRAC", 2, 3, D, S, "YEARFRAC(start_date, end_date, [basis])", "Returns the year fraction representing the number of whole days between start_date and end_date.")
def _yearfrac(args, ctx):
    s, e = _day_arg(args, 0), _day_arg(args, 1)
    basis = to_int(num(args, 2)) if has(args, 2) else 0
    return yearfrac_core(ctx.date_system, s, e, basis)
