"""Rendering a value with a parsed section."""

from __future__ import annotations

import math
from typing import List, Optional, Tuple

from ..core.date import MONTHS, WEEKDAYS, DateSystem, datetime_from_serial, max_serial
from .decimal import decimal_digits, fixed, fixed_from_digits, format_general_fit
from .parse import Section, is_calendar, literal_text


class Out:
    """Text being built plus the first repeat-fill position; ``zero`` is set when every digit a
    number section shows is 0 (a negative number then shows no minus sign)."""

    __slots__ = ("text", "fill", "zero")

    def __init__(self) -> None:
        self.text = ""
        self.fill: Optional[Tuple[str, int]] = None
        self.zero = False

    def push(self, s: str) -> None:
        self.text += s

    def set_fill(self, c: str) -> None:
        if self.fill is None:
            self.fill = (c, len(self.text))

    def prepend(self, s: str) -> None:
        self.text = s + self.text
        if self.fill is not None:
            self.fill = (self.fill[0], self.fill[1] + len(s))


def render_text(sec: Section, text: str) -> Out:
    out = Out()
    for t in sec.toks:
        k = t[0]
        if k == "at":
            out.push(text)
        elif k == "fill":
            out.set_fill(t[1])
        else:
            out.push(literal_text(t))
    return out


def render_general(sec: Section, v: float, width: int) -> Out:
    out = Out()
    for t in sec.toks:
        k = t[0]
        if k == "general":
            g = format_general_fit(v, width)
            out.push(g if g is not None else "#" * min(max(width, 1), 11))
        elif k == "fill":
            out.set_fill(t[1])
        else:
            out.push(literal_text(t))
    return out


def _pad(n: int, w: int) -> str:
    if w >= 2:
        return str(n).zfill(min(w, 20))
    return str(n)


def render_date(sec: Section, v: float, sys: DateSystem) -> Optional[Out]:
    """Renders a date/time. ``None`` when the serial is outside the supported range."""
    if not math.isfinite(v) or not 0.0 <= v < max_serial(sys) + 1.0:
        return None
    sub_digits = 0
    for t in sec.toks:
        if t[0] == "date" and t[1] == "sub":
            sub_digits = max(sub_digits, t[2])
    sub_digits = min(sub_digits, 3)
    unit = 10 ** (3 - sub_digits)
    day = math.floor(v)
    ms = int(math.floor(((v - day) * 86_400_000.0) / unit + 0.5)) * unit
    if ms >= 86_400_000:
        ms -= 86_400_000
        day += 1
    total_ms = day * 86_400_000 + ms
    needs_date = any(t[0] == "date" and is_calendar(t[1]) for t in sec.toks)
    if needs_date:
        dt = datetime_from_serial(sys, float(day))
        if dt is None:
            return None
    else:
        dt = datetime_from_serial(sys, 0.0)
    ampm = any(t[0] == "date" and t[1] in ("ampm", "ap") for t in sec.toks)
    hour = ms // 3_600_000
    minute = (ms // 60_000) % 60
    second = (ms // 1000) % 60
    milli = ms % 1000
    month_name = MONTHS[dt.month - 1] if 1 <= dt.month <= 12 else ""
    day_name = WEEKDAYS[dt.weekday] if 0 <= dt.weekday <= 6 else ""
    out = Out()
    for t in sec.toks:
        k = t[0]
        if k == "date":
            part, n = t[1], t[2]
            if part == "y":
                s = f"{dt.year % 100:02d}" if n <= 2 else f"{dt.year:04d}"
            elif part == "mo":
                if n in (1, 2):
                    s = _pad(dt.month, n)
                elif n == 3:
                    s = month_name[:3]
                elif n == 5:
                    s = month_name[:1]
                else:
                    s = month_name
            elif part == "d":
                if n in (1, 2):
                    s = _pad(dt.day, n)
                elif n == 3:
                    s = day_name[:3]
                else:
                    s = day_name
            elif part == "h":
                h = hour
                if ampm:
                    h = hour % 12
                    if h == 0:
                        h = 12
                s = _pad(h, n)
            elif part == "mi":
                s = _pad(minute, n)
            elif part == "s":
                s = _pad(second, n)
            elif part == "sub":
                s = f".{milli:03d}"[: 1 + min(n, 3)]
                if n > 3:
                    s += "0" * (n - 3)
            elif part == "eh":
                s = _pad(total_ms // 3_600_000, n)
            elif part == "em":
                s = _pad(total_ms // 60_000, n)
            elif part == "es":
                s = _pad(total_ms // 1000, n)
            elif part == "ampm":
                s = "PM" if hour >= 12 else "AM"
            else:  # ap
                a, p = n
                s = p if hour >= 12 else a
            out.push(s)
        elif k == "fill":
            out.set_fill(t[1])
        else:
            out.push(literal_text(t))
    return out


def _fill_int(toks: List[tuple], idxs: List[int], digits: str, thousands: bool, force_zero: bool, pieces: List[str]) -> None:
    """Assigns integer digits right-to-left to the placeholders at ``idxs``; the leftmost
    placeholder takes all remaining digits. ``force_zero`` shows ``0`` in the units placeholder
    when there are no digits. A blank ``?`` pads with a space, and so does the thousands separator
    it would carry."""
    ds = digits
    rem = len(ds)
    pos = 0
    count = len(idxs)
    for k, ti in enumerate(reversed(idxs)):
        leftmost = k + 1 == count
        t = toks[ti] if ti < len(toks) else None
        ph = t[1] if t is not None and t[0] == "digit" else "0"
        rev: List[str] = []

        def push_digit(c: str) -> None:
            nonlocal pos
            if thousands and pos > 0 and pos % 3 == 0:
                rev.append(",")
            rev.append(c)
            pos += 1

        take = rem if leftmost else min(rem, 1)
        if take == 0:
            if ph == "0":
                push_digit("0")
            elif force_zero and k == 0:
                push_digit("0")
            elif ph == "?":
                if thousands and pos > 0 and pos % 3 == 0:
                    rev.append(" ")
                rev.append(" ")
                pos += 1
        else:
            for _ in range(take):
                rem -= 1
                push_digit(ds[rem] if 0 <= rem < len(ds) else "0")
        if ti < len(pieces):
            pieces[ti] = "".join(reversed(rev))


def _with_commas(digits: str, thousands: bool) -> str:
    if not thousands:
        return digits
    n = len(digits)
    out = []
    for i, c in enumerate(digits):
        if i > 0 and (n - i) % 3 == 0:
            out.append(",")
        out.append(c)
    return "".join(out)


def _fill_frac(toks: List[tuple], idxs: List[int], frac: str, pieces: List[str]) -> None:
    """Fraction-digit placeholders: trailing zeros show as ``0``, nothing (``#``) or a space (``?``)."""
    last_sig = None
    for i in range(len(frac) - 1, -1, -1):
        if frac[i] != "0":
            last_sig = i
            break
    for i, ti in enumerate(idxs):
        t = toks[ti]
        ph = t[1] if t[0] == "digit" else "0"
        if last_sig is not None and i <= last_sig:
            s = frac[i] if i < len(frac) else "0"
        elif ph == "0":
            s = "0"
        elif ph == "?":
            s = " "
        else:
            s = ""
        pieces[ti] = s


def _digit_indices(toks: List[tuple], lo: int, hi: int) -> List[int]:
    return [i for i in range(lo, min(hi, len(toks))) if toks[i][0] == "digit"]


def render_number(sec: Section, v: float) -> Out:
    """Renders a non-negative number through a number section."""
    x = abs(v)
    # In a scientific section a % or a trailing comma doesn't scale the number (Excel).
    if sec.percent > 0 and not sec.exp:
        x *= 100.0 ** min(sec.percent, 20)
    if sec.scale > 0 and not sec.exp:
        x /= 1000.0 ** min(sec.scale, 110)
    if not math.isfinite(x):
        x = 1.7976931348623157e308
    toks = sec.toks
    n = len(toks)
    pieces = [""] * n
    before = [""] * (n + 1)
    lits: List[Optional[str]] = []
    zero = False
    if sec.fraction:
        lits = [None] * n
        _render_fraction(sec, x, pieces, lits)
    else:
        point = next((i for i, t in enumerate(toks) if t[0] == "point"), None)
        exp = next((i for i, t in enumerate(toks) if t[0] == "exp"), None)
        int_end = point if point is not None else (exp if exp is not None else n)
        frac_end = exp if exp is not None else n
        int_idx = _digit_indices(toks, 0, int_end)
        frac_idx = _digit_indices(toks, point + 1, frac_end) if point is not None else []
        if exp is not None:
            exp_idx = _digit_indices(toks, exp + 1, n)
            step = max(len(int_idx), 1)
            if x == 0.0:
                # Zero shows a 0 in each integer placeholder of an all-# mantissa (Excel).
                hashes = not any(toks[i] == ("digit", "0") for i in int_idx)
                e, int_s, frac_s = 0, "0" * len(int_idx) if hashes else "", "0" * len(frac_idx)
            else:
                digits, xe = decimal_digits(x)
                e = xe - (xe % step)
                int_s, frac_s = fixed_from_digits(digits, xe - e, len(frac_idx))
                if len(int_s) > step:
                    e += step
                    int_s, frac_s = fixed_from_digits(digits, xe - e, len(frac_idx))
            t = toks[exp]
            upper, plus = t[1], t[2]
            if not int_s and all(ch == "0" for ch in frac_s):
                e = 0
            head = "E" if upper else "e"
            sign = "-" if e < 0 else ("+" if plus else "")
            width = min(sum(1 for i in exp_idx if toks[i] == ("digit", "0")), 20)
            edigits = str(abs(e)).zfill(width)
            if exp_idx:
                # The sign goes before the exponent digits, after a literal that follows the E.
                pieces[exp] = head
                pieces[exp_idx[0]] = sign + edigits
            else:
                pieces[exp] = head + sign + edigits
        else:
            int_s, frac_s = fixed(x, len(frac_idx))
        zero = exp is None and not any(c in "123456789" for c in int_s + frac_s)
        _fill_int(toks, int_idx, int_s, sec.thousands, False, pieces)
        if not int_idx and int_s:
            p = point if point is not None else exp
            if p is not None:
                before[p] = _with_commas(int_s, sec.thousands)
        _fill_frac(toks, frac_idx, frac_s, pieces)
        if point is not None:
            pieces[point] = "."
    out = Out()
    out.zero = zero
    for i, t in enumerate(toks):
        out.push(before[i])
        k = t[0]
        if k == "lit":
            lt = lits[i] if i < len(lits) else None
            out.push(lt if lt is not None else t[1])
        elif k == "percent":
            out.push("%")
        elif k == "fill":
            out.set_fill(t[1])
        elif k in ("at", "general", "date", "comma"):
            pass
        else:
            out.push(pieces[i])
    return out


def approximate(x: float, max_den: int) -> Tuple[int, int]:
    """Best rational approximation of ``x`` with denominator at most ``max_den`` (continued
    fractions)."""
    if x != x or x <= 0.0 or not math.isfinite(x):
        return 0, 1
    if x >= 1e15:
        return int(math.floor(x + 0.5)), 1
    p0, q0, p1, q1 = 0, 1, 1, 0
    r = x
    for _ in range(64):
        a_f = math.floor(r)
        if a_f > 1e15:
            break
        a = int(a_f)
        q2 = a * q1 + q0
        if q2 > max_den:
            break
        p2 = a * p1 + p0
        p0, q0, p1, q1 = p1, q1, p2, q2
        frac = r - a_f
        if frac < 1e-12:
            break
        r = 1.0 / frac
    if q1 == 0:
        return int(math.floor(x + 0.5)), 1
    k = (max_den - q0) // q1
    pb, qb = p0 + k * p1, q0 + k * q1
    err_a = abs(p1 / q1 - x)
    err_b = abs(pb / qb - x) if qb > 0 else float("inf")
    return (pb, qb) if err_b < err_a else (p1, q1)


def _shown_width(t: tuple) -> int:
    k = t[0]
    if k in ("lit", "fixed"):
        return len(t[1])
    if k in ("digit", "slash"):
        return 1
    return 0


def _blank_range(toks: List[tuple], lo: int, hi: int, spaces: bool, pieces: List[str], lits: List[Optional[str]]) -> None:
    for i in range(lo, min(hi, len(toks))):
        t = toks[i]
        s = " " * _shown_width(t) if spaces else ""
        if t[0] == "lit":
            lits[i] = s
        elif t[0] in ("digit", "slash", "fixed"):
            pieces[i] = s


def _render_fraction(sec: Section, x: float, pieces: List[str], lits: List[Optional[str]]) -> None:
    """A fraction section is ``[whole] [sep] num [lit] / [lit] den``, rendered the way Excel does."""
    toks = sec.toks

    def is_digit(i: int) -> bool:
        return 0 <= i < len(toks) and toks[i][0] == "digit"

    slash = next((i for i, t in enumerate(toks) if t[0] == "slash"), None)
    if slash is None:
        return
    num_last = next((i for i in range(slash - 1, -1, -1) if is_digit(i)), None)
    if num_last is None:
        return
    num_start = num_last
    while num_start > 0 and is_digit(num_start - 1):
        num_start -= 1
    whole_idx = _digit_indices(toks, 0, num_start)
    num_idx = _digit_indices(toks, num_start, num_last + 1)
    sep_start = whole_idx[-1] + 1 if whole_idx else num_start
    nxt = toks[slash + 1] if slash + 1 < len(toks) else None
    if nxt is not None and nxt[0] == "fixed":
        try:
            fixed_den: Optional[int] = min(max(int(nxt[1]), 1), 1_000_000_000)
        except ValueError:
            fixed_den = 1_000_000_000
        den_start = slash + 1
    else:
        fixed_den = None
        i = slash + 1
        while i < len(toks) and toks[i][0] == "lit":
            i += 1
        den_start = i
    den_end = den_start
    if fixed_den is not None:
        den_end += 1
    else:
        while is_digit(den_end):
            den_end += 1
    den_idx = _digit_indices(toks, den_start, den_end)

    def has_q(idx: List[int]) -> bool:
        return any(toks[i] == ("digit", "?") for i in idx)

    has_whole = bool(whole_idx)
    if has_whole:
        whole, f = math.floor(x), x - math.floor(x)
    else:
        whole, f = 0.0, x
    if fixed_den is not None:
        num, den = int(math.floor(f * fixed_den + 0.5)), fixed_den
    else:
        digits = min(max(len(den_idx), 1), 7)
        num, den = approximate(f, 10 ** digits - 1)
    if has_whole and num >= den and den > 0:
        whole += num // den
        num %= den
    whole_s = fixed(whole, 0)[0] if whole >= 1.0 else ""
    num_zero_ph = any(toks[i] == ("digit", "0") for i in num_idx)
    if has_whole and num == 0 and not num_zero_ph:
        _fill_int(toks, whole_idx, whole_s, sec.thousands, True, pieces)
        spaces = has_q(whole_idx) or has_q(num_idx) or has_q(den_idx)
        _blank_range(toks, sep_start, den_end, spaces, pieces, lits)
        return
    if has_whole:
        # A zero value shows `0` in a `?` or `0` units place, nothing in a `#` one.
        units_hash = bool(whole_idx) and toks[whole_idx[-1]] == ("digit", "#")
        force = whole_s == "" and num == 0 and not units_hash
        _fill_int(toks, whole_idx, whole_s, sec.thousands, force, pieces)
        whole_shown = any(any(ch.isdigit() for ch in pieces[i]) for i in whole_idx)
        if not whole_shown:
            spaces = has_q(whole_idx) or has_q(num_idx)
            _blank_range(toks, sep_start, num_start, spaces, pieces, lits)
    _fill_int(toks, num_idx, str(num), False, True, pieces)
    pieces[slash] = "/"
    den_s = str(den)
    if fixed_den is not None:
        pieces[den_start] = den_s
    else:
        # Denominator digits are zero padded to the `0` placeholders, then left aligned; `?`
        # pads with spaces after them.
        zeros = sum(1 for i in den_idx if toks[i] == ("digit", "0"))
        padded = den_s.rjust(zeros, "0")
        count = len(den_idx)
        for k, ti in enumerate(den_idx):
            if k + 1 == count:
                s = padded[k:]
            else:
                s = padded[k] if k < len(padded) else ""
            if not s:
                t = toks[ti]
                s = " " if t == ("digit", "?") else ("0" if t == ("digit", "0") else "")
            pieces[ti] = s
