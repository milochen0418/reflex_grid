"""Decimal-string rounding and the General format.

Excel works with 15 significant digits, so numbers are first written with 15 significant digits
and then rounded half away from zero on that decimal string. That makes 2.675 with ``0.00`` show
2.68 and 1.005 show 1.01, as in Excel, even though neither is exact in binary.
"""

from __future__ import annotations

import math
from typing import List, Optional, Tuple

MAX_DECIMALS = 400


def _significant_digits(a: float, n: int) -> Tuple[List[int], int]:
    """``n`` significant digits of ``a`` (rounded half to even on the exact binary value) and the
    power of ten of the first."""
    s = f"{a:.{max(n - 1, 0)}e}"
    mant, _, exp = s.partition("e")
    return [ord(ch) - 48 for ch in mant if ch.isdigit()], int(exp or 0)


def _is_exact_tie(a: float) -> bool:
    """True when ``a``'s exact decimal expansion has a 5 as its 16th significant digit and only
    zeros after it."""

    def tie(ds: List[int]) -> bool:
        return len(ds) > 15 and ds[15] == 5 and all(d == 0 for d in ds[16:])

    return tie(_significant_digits(a, 30)[0]) and tie(_significant_digits(a, 800)[0])


def decimal_digits(a: float) -> Tuple[List[int], int]:
    """The 15 significant digits of a positive finite number and the power of ten of the first
    digit: ``a ≈ d0.d1d2… × 10^exp``. An exact tie at the 16th digit goes toward zero, as Excel
    does."""
    if not math.isfinite(a) or a <= 0.0:
        return [0] * 15, 0
    d16, exp = _significant_digits(a, 16)
    digits = (d16 + [0] * 15)[:15]
    d = d16[15] if len(d16) > 15 else 0
    if d <= 4:
        return digits, exp
    if d == 5 and _is_exact_tie(a):
        return digits, exp
    if d == 5:
        d15, exp15 = _significant_digits(a, 15)
        return (d15 + [0] * 15)[:15], exp15
    carry = True
    for i in range(14, -1, -1):
        if digits[i] == 9:
            digits[i] = 0
        else:
            digits[i] += 1
            carry = False
            break
    if carry:
        digits = [1] + [0] * 14
        exp += 1
    return digits, exp


def fixed_from_digits(digits: List[int], exp: int, decimals: int) -> Tuple[str, str]:
    """Rounds ``digits × 10^exp`` to ``decimals`` places, half away from zero. Returns the integer
    digits (``""`` when the integer part is zero) and exactly ``decimals`` fraction digits."""
    decimals_calc = min(decimals, MAX_DECIMALS)
    keep = exp + decimals_calc + 1
    if keep <= 0:
        m = [1] if keep == 0 and digits and digits[0] >= 5 else []
    elif keep >= len(digits):
        m = list(digits) + [0] * (keep - len(digits))
    else:
        m = list(digits[:keep])
        if digits[keep] >= 5:
            carry = True
            for i in range(len(m) - 1, -1, -1):
                if m[i] == 9:
                    m[i] = 0
                else:
                    m[i] += 1
                    carry = False
                    break
            if carry:
                m.insert(0, 1)
    lead = 0
    while lead < len(m) and m[lead] == 0:
        lead += 1
    m = m[lead:]
    s = "".join(chr(48 + d) for d in m)
    if len(m) <= decimals_calc:
        int_part, frac = "", "0" * (decimals_calc - len(m)) + s
    else:
        split = len(m) - decimals_calc
        int_part, frac = s[:split], s[split:]
    if decimals > decimals_calc:
        frac += "0" * (decimals - decimals_calc)
    return int_part, frac


def fixed(a: float, decimals: int) -> Tuple[str, str]:
    """Rounds a non-negative number to ``decimals`` places (see :func:`fixed_from_digits`)."""
    if not math.isfinite(a) or a <= 0.0:
        return "", "0" * min(decimals, MAX_DECIMALS * 4)
    d, e = decimal_digits(a)
    return fixed_from_digits(d, e, decimals)


def format_general_fit(n: float, max_chars: int) -> Optional[str]:
    """General format squeezed into ``max_chars`` characters: fewer decimals, then scientific
    notation; ``None`` means the caller paints ``####``."""
    if not math.isfinite(n):
        return None
    if n == 0.0:
        return "0" if max_chars >= 1 else None
    sign = "-" if n < 0 else ""
    digits, x = decimal_digits(abs(n))
    int_len = len(sign) + max(x, 0) + 1
    if x >= -4 and int_len <= max_chars:
        avail = max_chars - int_len
        decimals = avail - 1 if avail >= 2 else 0
        while True:
            i, f = fixed_from_digits(digits, x, decimals)
            f = f.rstrip("0")
            if not i and not f:
                break  # rounded away to zero: use scientific
            s = sign + (i if i else "0")
            if f:
                s += "." + f
            if len(s) <= max_chars:
                return s
            if decimals == 0:
                break
            decimals -= 1
    for md in range(min(max_chars, 14), -1, -1):
        e = x
        mi, mf = fixed_from_digits(digits, 0, md)
        if len(mi) > 1:
            e += 1
            mi, mf = fixed_from_digits(digits, -1, md)
        mf = mf.rstrip("0")
        s = sign + (mi if mi else "0")
        if mf:
            s += "." + mf
        s += f"E{'-' if e < 0 else '+'}{abs(e):02d}"
        if len(s) <= max_chars:
            return s
    return None
