"""Parsing what people type into a cell (and numeric text in formulas).

Recognises numbers with thousands separators, percentages, currency, scientific notation,
fractions (``1 1/2``), dates (``3/4/2026``, ``2026-03-04``, ``4-Mar-2026``, ``Mar 4, 2026``),
times (``10:30``, ``10:30:15 PM``) and date-times. Locale: en-US.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from functools import lru_cache
from typing import Optional

from .date import MONTHS, DateSystem, current_year, days_in_month_in, serial_from_ymd, time_fraction
from .value import CellError, Value, number


@dataclass
class Parsed:
    value: Value
    # The number format Excel would apply automatically.
    format: Optional[str] = None


def parse_input(s: str, sys: DateSystem = DateSystem.D1900) -> Parsed:
    """Interprets typed cell input (not formulas). Text that isn't a number, date, boolean or
    error stays text. A leading apostrophe forces text."""
    if s == "":
        return Parsed(None)
    if s.startswith("'"):
        return Parsed(s[1:])
    t = s.strip()
    u = t.upper()
    if u == "TRUE":
        return Parsed(True)
    if u == "FALSE":
        return Parsed(False)
    e = CellError.parse(t)
    if e is not None:
        return Parsed(e)
    r = _parse_numeric(t, sys)
    if r is not None:
        return Parsed(number(r[0]), r[1])
    return Parsed(s)


@lru_cache(maxsize=8192)
def parse_number_text(s: str) -> Optional[float]:
    """Numeric text as arithmetic coerces it ("1,000", "50%", "$3", "1/2/2020", "10:00")."""
    r = _parse_numeric(s.strip(), DateSystem.D1900)
    return None if r is None else r[0]


def _parse_numeric(t: str, sys: DateSystem) -> Optional[tuple[float, Optional[str]]]:
    if not t:
        return None
    # Parenthesised negatives: (100) = -100.
    if t.startswith("(") and t.endswith(")") and len(t) >= 2:
        inner = _parse_numeric(t[1:-1].strip(), sys)
        return None if inner is None else (-inner[0], inner[1])
    r = _parse_plain(t)
    if r is not None:
        return r
    f = _parse_fraction(t)
    if f is not None:
        return f, "# ?/?"
    return _parse_datetime(t, sys)


_DIGITS_RE = re.compile(r"^[0-9.eE+\-]+$")
_FLOAT_RE = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")


def _float(s: str) -> Optional[float]:
    """Rust-style ``str::parse::<f64>`` for plain decimal text (no underscores, no inf/nan)."""
    if not _FLOAT_RE.match(s):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _parse_plain(t: str) -> Optional[tuple[float, Optional[str]]]:
    """Signs, currency, thousands separators, percent, exponent."""
    s = t
    neg = currency = percent = False
    if s.startswith("-"):
        neg = True
        s = s[1:].lstrip()
    elif s.startswith("+"):
        s = s[1:].lstrip()
    if s.startswith("$"):
        currency = True
        s = s[1:].lstrip()
        if not neg and s.startswith("-"):
            neg = True
            s = s[1:]
    if s.endswith("%"):
        percent = True
        s = s[:-1].rstrip()
    if not s:
        return None
    has_sep = "," in s
    if has_sep:
        # Thousands separators must be in groups of three before the decimal point.
        if "." in s:
            int_part, frac = s.split(".", 1)
        else:
            int_part, frac = s, None
        groups = int_part.split(",")
        first = groups[0]
        if not first or len(first) > 3 or any(len(g) != 3 for g in groups[1:]):
            return None
        digits = "".join(groups)
        if frac is not None:
            digits += "." + frac
    else:
        digits = s
    if not _DIGITS_RE.match(digits):
        return None
    if not (digits[0].isdigit() or digits[0] == "."):
        return None
    try:
        n = float(_fifteen_digits(digits))
    except ValueError:
        return None
    if not math.isfinite(n):
        return None
    sci = "e" in digits or "E" in digits
    if percent:
        n /= 100.0
    if neg:
        n = -n
    if currency:
        # Excel's built-in currency formats (5..8): negatives in red, in parentheses.
        fmt = "$#,##0.00_);[Red]($#,##0.00)" if "." in digits else "$#,##0_);[Red]($#,##0)"
    elif percent:
        fmt = "0.00%" if "." in digits else "0%"
    elif sci:
        fmt = "0.00E+00"
    elif has_sep:
        fmt = "#,##0.00" if "." in digits else "#,##0"
    else:
        fmt = None
    return n, fmt


def _fifteen_digits(digits: str) -> str:
    """Excel keeps 15 significant digits of a typed number; the rest become zeros
    (123456789012345678 is stored as 123456789012345000)."""
    mant, _, exp = digits.replace("E", "e").partition("e")
    sig = 0
    out = []
    for ch in mant:
        if ch.isdigit():
            if sig or ch != "0":
                sig += 1
            out.append(ch if sig <= 15 else "0")
        else:
            out.append(ch)
    return "".join(out) + ("e" + exp if exp else "")


def _parse_fraction(t: str) -> Optional[float]:
    """``1 1/2``, ``0 3/4`` (a bare ``1/2`` is a date)."""
    if " " not in t:
        return None
    whole, frac = t.split(" ", 1)
    neg = whole.startswith("-")
    w = _float(whole.lstrip("-"))
    if w is None:
        return None
    frac = frac.strip()
    if "/" not in frac:
        return None
    a, b = frac.split("/", 1)
    fa, fb = _float(a), _float(b)
    if fa is None or fb is None:
        return None
    if fb == 0.0 or not math.isfinite(w):
        return None
    v = w + fa / fb
    return -v if neg else v


def month_from_name(s: str) -> Optional[int]:
    l = s.lower()
    if len(l) < 3:
        return None
    for i, m in enumerate(MONTHS):
        if m.lower().startswith(l):
            return i + 1
    return None


def _parse_time(t: str) -> Optional[tuple[float, bool, bool]]:
    """Time part: ``h:mm``, ``h:mm:ss``, ``h:mm:ss.000``, ``h AM``, with optional AM/PM."""
    lower = t.strip().lower()
    ampm: Optional[bool] = None
    if lower.endswith("am"):
        body, ampm = lower[:-2].strip(), False
    elif lower.endswith("a"):
        body, ampm = lower[:-1].strip(), False
    elif lower.endswith("pm"):
        body, ampm = lower[:-2].strip(), True
    elif lower.endswith("p"):
        body, ampm = lower[:-1].strip(), True
    else:
        body = lower
    parts = body.split(":")
    if len(parts) == 1 and ampm is None:
        return None
    if len(parts) > 3 or any(p == "" for p in parts):
        return None
    h = _float(parts[0])
    m = _float(parts[1]) if len(parts) > 1 else 0.0
    s = _float(parts[2]) if len(parts) > 2 else 0.0
    if h is None or m is None or s is None:
        return None
    if not all(math.isfinite(x) for x in (h, m, s)):
        return None
    if m >= 60 or s >= 60 or h < 0 or m < 0 or s < 0:
        return None
    if ampm is not None:
        if not 0 <= h <= 12:
            return None
        if h == 12:
            h = 0
        if ampm:
            h += 12
    elif h >= 10000:
        return None
    return time_fraction(h, m, s), len(parts) == 3, ampm is not None


def _num(s: str) -> Optional[int]:
    s = s.strip()
    if s[:1] in "+-":
        body = s[1:]
    else:
        body = s
    if not body.isdigit() or not body.isascii():
        return None
    try:
        return int(s)
    except ValueError:
        return None


def _two_digit_year(y: int, text: str) -> int:
    if len(text.strip()) <= 2:
        return 2000 + y if y < 30 else 1900 + y
    return y


def _valid_ymd(sys: DateSystem, y: int, m: int, d: int) -> Optional[float]:
    # Dates before 1900 stay text (DATE() would add 1900 to such a year).
    if not 1 <= m <= 12 or d < 1 or y < 1900:
        return None
    if d > days_in_month_in(sys, y, m):
        return None
    return serial_from_ymd(sys, y, m, d)


_SEP_RE = re.compile(r"[/\-.]")


def _parse_date(t: str, sys: DateSystem) -> Optional[tuple[float, str]]:
    t = t.strip().rstrip(",")
    parts = _SEP_RE.split(t)
    if len(parts) == 3:
        a, b, c = (p.strip() for p in parts)
        if len(a) == 4:
            y, m, d = _num(a), _num(b), _num(c)
            if y is not None and m is not None and d is not None:
                s = _valid_ymd(sys, y, m, d)
                return None if s is None else (s, "m/d/yyyy")
        m, d, y = _num(a), _num(b), _num(c)
        if m is not None and d is not None and y is not None:
            y = _two_digit_year(y, c)
            s = _valid_ymd(sys, y, m, d)
            return None if s is None else (s, "m/d/yyyy")
        # 4-Mar-2026
        d, mo, y = _num(a), month_from_name(b), _num(c)
        if d is not None and mo is not None and y is not None:
            y = _two_digit_year(y, c)
            s = _valid_ymd(sys, y, mo, d)
            return None if s is None else (s, "d-mmm-yy")
    if len(parts) == 2:
        a, b = parts[0].strip(), parts[1].strip()
        m, d = _num(a), _num(b)
        if m is not None and d is not None:
            if 1 <= m <= 12 and 1 <= d <= 31:
                s = _valid_ymd(sys, current_year(), m, d)
                return None if s is None else (s, "d-mmm")
            # m/yyyy
            if 1 <= m <= 12 and len(b) == 4:
                s = _valid_ymd(sys, d, m, 1)
                return None if s is None else (s, "mmm-yy")
        d, mo = _num(a), month_from_name(b)
        if d is not None and mo is not None:
            s = _valid_ymd(sys, current_year(), mo, d)
            return None if s is None else (s, "d-mmm")
        mo, y = month_from_name(a), _num(b)
        if mo is not None and y is not None:
            y = _two_digit_year(y, b)
            s = _valid_ymd(sys, y, mo, 1)
            return None if s is None else (s, "mmm-yy")
    # "Mar 4, 2026" / "March 4 2026" / "4 March 2026"
    words = [w for w in re.split(r"[ ,]", t) if w]
    if len(words) == 2:
        # "Jan 2024": the first of the month.
        mo, y = month_from_name(words[0]), _num(words[1])
        if mo is not None and y is not None and len(words[1]) == 4:
            s = _valid_ymd(sys, y, mo, 1)
            return None if s is None else (s, "mmm-yy")
    if len(words) == 3:
        mo, d, y = month_from_name(words[0]), _num(words[1]), _num(words[2])
        if mo is not None and d is not None and y is not None:
            s = _valid_ymd(sys, y, mo, d)
            return None if s is None else (s, "d-mmm-yy")
        d, mo, y = _num(words[0]), month_from_name(words[1]), _num(words[2])
        if d is not None and mo is not None and y is not None:
            s = _valid_ymd(sys, y, mo, d)
            return None if s is None else (s, "d-mmm-yy")
    return None


def _parse_datetime(t: str, sys: DateSystem) -> Optional[tuple[float, Optional[str]]]:
    tm = _parse_time(t)
    if tm is not None:
        tf, secs, ampm = tm
        if tf >= 1.0:
            # 24 hours or more: elapsed hours, so 25:00 shows as 25:00:00, not 1:00.
            f = "[h]:mm:ss"
        elif not secs and not ampm:
            f = "h:mm"
        elif secs and not ampm:
            f = "h:mm:ss"
        elif not secs and ampm:
            f = "h:mm AM/PM"
        else:
            f = "h:mm:ss AM/PM"
        return tf, f
    d = _parse_date(t, sys)
    if d is not None:
        return d[0], d[1]
    # Date + time: split at a space before a time-looking tail.
    for i, ch in enumerate(t):
        if ch == " ":
            a, b = t[:i], t[i + 1:]
            da, tb = _parse_date(a, sys), _parse_time(b)
            if da is not None and tb is not None:
                return da[0] + tb[0], "m/d/yyyy h:mm"
    return None
