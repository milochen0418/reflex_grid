"""Engineering functions: base conversion, bit operations, complex numbers, error function,
Bessel functions and unit conversion."""

from __future__ import annotations

import cmath
import math
import re
import sys
from typing import Optional, Tuple

from ..core.value import Array, CellError, Value, XlError, scalar, to_number, to_text
from .registry import A, Arg, Ctx, S, VAR, fn
from .util import as_array, has, num, opt_num, text

E = "Engineering"

# ---------------------------------------------------------------- base conversion

_LIMITS = {2: (10, 2**9), 8: (10, 8**9 * 4), 16: (10, 16**9 * 8)}


def _from_base(s: str, base: int) -> int:
    s = s.strip().upper()
    if s == "":
        return 0
    if len(s) > 10:
        raise XlError(CellError.NUM)
    try:
        v = int(s, base)
    except ValueError:
        raise XlError(CellError.NUM) from None
    bits = 10 * {2: 1, 8: 3, 16: 4}[base]
    if len(s) == 10 and v >= 1 << (bits - 1):
        v -= 1 << bits
    return v


def _to_base(v: int, base: int, places: Optional[int]) -> str:
    bits = 10 * {2: 1, 8: 3, 16: 4}[base]
    lo, hi = -(1 << (bits - 1)), (1 << (bits - 1)) - 1
    if base == 2:
        lo, hi = -512, 511
    elif base == 8:
        lo, hi = -536870912, 536870911
    else:
        lo, hi = -549755813888, 549755813887
    if not lo <= v <= hi:
        raise XlError(CellError.NUM)
    if v < 0:
        v += 1 << bits
        digits = _fmt(v, base)
        return digits
    digits = _fmt(v, base)
    if places is not None:
        if places < len(digits) or places > 10:
            raise XlError(CellError.NUM)
        digits = digits.rjust(places, "0")
    return digits


def _fmt(v: int, base: int) -> str:
    if base == 2:
        return format(v, "b")
    if base == 8:
        return format(v, "o")
    return format(v, "X")


def _places(args, i: int) -> Optional[int]:
    if not has(args, i):
        return None
    p = num(args, i)
    if p < 0:
        raise XlError(CellError.NUM)
    return int(p)


def _src_text(args) -> str:
    v = scalar(args[0].value)
    if isinstance(v, CellError):
        raise XlError(v)
    if isinstance(v, bool):
        raise XlError(CellError.VALUE)
    if isinstance(v, float):
        if v != int(v):
            raise XlError(CellError.NUM)
        return str(int(v))
    return to_text(v)


def _conv(name: str, src: int, dst: int) -> None:
    if dst == 10:
        @fn(name, 1, 1, E, S, f"{name}(number)", f"Converts a base-{src} number to decimal.")
        def _impl(args, ctx, src=src):
            return float(_from_base(_src_text(args), src))
    else:
        @fn(name, 1, 2, E, S, f"{name}(number, [places])", f"Converts a base-{src} number to base {dst}.")
        def _impl(args, ctx, src=src, dst=dst):
            if src == 10:
                n = num(args, 0)
                v = math.trunc(n)
            else:
                v = _from_base(_src_text(args), src)
            return _to_base(v, dst, _places(args, 1))


for _name, _s, _d in [
    ("BIN2DEC", 2, 10), ("BIN2HEX", 2, 16), ("BIN2OCT", 2, 8), ("DEC2BIN", 10, 2), ("DEC2HEX", 10, 16),
    ("DEC2OCT", 10, 8), ("HEX2BIN", 16, 2), ("HEX2DEC", 16, 10), ("HEX2OCT", 16, 8), ("OCT2BIN", 8, 2),
    ("OCT2DEC", 8, 10), ("OCT2HEX", 8, 16),
]:
    _conv(_name, _s, _d)


def _bit_arg(args, i: int) -> int:
    n = num(args, i)
    if n < 0 or n != int(n) or n >= 2**48:
        raise XlError(CellError.NUM)
    return int(n)


@fn("BITAND", 2, 2, E, S, "BITAND(number1, number2)", "Returns a 'Bitwise And' of two numbers.")
def _bitand(args, ctx):
    return float(_bit_arg(args, 0) & _bit_arg(args, 1))


@fn("BITOR", 2, 2, E, S, "BITOR(number1, number2)", "Returns a bitwise OR of 2 numbers.")
def _bitor(args, ctx):
    return float(_bit_arg(args, 0) | _bit_arg(args, 1))


@fn("BITXOR", 2, 2, E, S, "BITXOR(number1, number2)", "Returns a bitwise 'Exclusive Or' of two numbers.")
def _bitxor(args, ctx):
    return float(_bit_arg(args, 0) ^ _bit_arg(args, 1))


@fn("BITLSHIFT", 2, 2, E, S, "BITLSHIFT(number, shift_amount)", "Returns a value number shifted left by shift_amount bits.")
def _bitlshift(args, ctx):
    n, k = _bit_arg(args, 0), int(num(args, 1))
    if abs(k) > 53:
        raise XlError(CellError.NUM)
    v = n << k if k >= 0 else n >> -k
    if v >= 2**48:
        raise XlError(CellError.NUM)
    return float(v)


@fn("BITRSHIFT", 2, 2, E, S, "BITRSHIFT(number, shift_amount)", "Returns a value number shifted right by shift_amount bits.")
def _bitrshift(args, ctx):
    n, k = _bit_arg(args, 0), int(num(args, 1))
    if abs(k) > 53:
        raise XlError(CellError.NUM)
    v = n >> k if k >= 0 else n << -k
    if v >= 2**48:
        raise XlError(CellError.NUM)
    return float(v)


@fn("DELTA", 1, 2, E, S, "DELTA(number1, [number2])", "Tests whether two values are equal.")
def _delta(args, ctx):
    return 1.0 if num(args, 0) == (num(args, 1) if has(args, 1) else 0.0) else 0.0


@fn("GESTEP", 1, 2, E, S, "GESTEP(number, [step])", "Tests whether a number is greater than a threshold value.")
def _gestep(args, ctx):
    return 1.0 if num(args, 0) >= (num(args, 1) if has(args, 1) else 0.0) else 0.0


@fn("ERF", 1, 2, E, S, "ERF(lower_limit, [upper_limit])", "Returns the error function.")
def _erf(args, ctx):
    lo = num(args, 0)
    if has(args, 1):
        return math.erf(num(args, 1)) - math.erf(lo)
    return math.erf(lo)


@fn("ERF.PRECISE", 1, 1, E, S, "ERF.PRECISE(x)", "Returns the error function.")
def _erf_precise(args, ctx):
    return math.erf(num(args, 0))


@fn("ERFC", 1, 1, E, S, "ERFC(x)", "Returns the complementary error function.")
def _erfc(args, ctx):
    return math.erfc(num(args, 0))


@fn("ERFC.PRECISE", 1, 1, E, S, "ERFC.PRECISE(x)", "Returns the complementary ERF function integrated between x and infinity.")
def _erfc_precise(args, ctx):
    return math.erfc(num(args, 0))


# ---------------------------------------------------------------- complex numbers

_COMPLEX_RE = re.compile(r"^([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)?(?:([+-])((?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)?([ij]))?$")


def parse_complex(v: Value) -> Tuple[complex, str]:
    if isinstance(v, CellError):
        raise XlError(v)
    if isinstance(v, bool):
        raise XlError(CellError.VALUE)
    if v is None:
        return complex(0, 0), "i"
    if isinstance(v, float):
        return complex(v, 0), "i"
    s = to_text(v).strip()
    if s == "":
        return complex(0, 0), "i"
    # Pure imaginary: "i", "-2j", "3.5i"
    m = re.fullmatch(r"([+-]?)((?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)?([ij])", s)
    if m:
        coef = float(m.group(2)) if m.group(2) else 1.0
        if m.group(1) == "-":
            coef = -coef
        return complex(0, coef), m.group(3)
    m = _COMPLEX_RE.match(s)
    if not m or m.group(1) is None:
        raise XlError(CellError.NUM)
    re_part = float(m.group(1))
    im = 0.0
    suffix = "i"
    if m.group(4):
        coef = float(m.group(3)) if m.group(3) else 1.0
        im = -coef if m.group(2) == "-" else coef
        suffix = m.group(4)
    return complex(re_part, im), suffix


def _num_text(x: float) -> str:
    from ..core.value import number_to_text

    return number_to_text(x)


def format_complex(z: complex, suffix: str = "i") -> str:
    re_part = 0.0 if abs(z.real) < 1e-300 else z.real
    im = 0.0 if abs(z.imag) < 1e-300 else z.imag
    if not math.isfinite(re_part) or not math.isfinite(im):
        raise XlError(CellError.NUM)
    if im == 0:
        return _num_text(re_part)
    if im == 1:
        ims = suffix
    elif im == -1:
        ims = "-" + suffix
    else:
        ims = _num_text(im) + suffix
    if re_part == 0:
        return ims
    if im > 0:
        return f"{_num_text(re_part)}+{ims}"
    return f"{_num_text(re_part)}{ims}"


@fn("COMPLEX", 2, 3, E, S, "COMPLEX(real_num, i_num, [suffix])", "Converts real and imaginary coefficients into a complex number.")
def _complex(args, ctx):
    suffix = text(args, 2) if has(args, 2) else "i"
    if suffix not in ("i", "j"):
        raise XlError(CellError.VALUE)
    return format_complex(complex(num(args, 0), num(args, 1)), suffix)


def _cx(args, i: int = 0) -> Tuple[complex, str]:
    return parse_complex(scalar(args[i].value))


@fn("IMREAL", 1, 1, E, S, "IMREAL(inumber)", "Returns the real coefficient of a complex number.")
def _imreal(args, ctx):
    return _cx(args)[0].real


@fn("IMAGINARY", 1, 1, E, S, "IMAGINARY(inumber)", "Returns the imaginary coefficient of a complex number.")
def _imaginary(args, ctx):
    return _cx(args)[0].imag


@fn("IMABS", 1, 1, E, S, "IMABS(inumber)", "Returns the absolute value (modulus) of a complex number.")
def _imabs(args, ctx):
    return abs(_cx(args)[0])


@fn("IMARGUMENT", 1, 1, E, S, "IMARGUMENT(inumber)", "Returns the argument theta, an angle expressed in radians.")
def _imargument(args, ctx):
    z = _cx(args)[0]
    if z == 0:
        raise XlError(CellError.DIV0)
    return cmath.phase(z)


@fn("IMCONJUGATE", 1, 1, E, S, "IMCONJUGATE(inumber)", "Returns the complex conjugate of a complex number.")
def _imconjugate(args, ctx):
    z, s = _cx(args)
    return format_complex(z.conjugate(), s)


def _cfun(name: str, f, desc: str, zero_err: bool = False) -> None:
    @fn(name, 1, 1, E, S, f"{name}(inumber)", desc)
    def _impl(args, ctx, f=f):
        z, s = _cx(args)
        if zero_err and z == 0:
            raise XlError(CellError.NUM)
        try:
            return format_complex(f(z), s)
        except (ValueError, ZeroDivisionError, OverflowError):
            raise XlError(CellError.NUM) from None


_cfun("IMCOS", cmath.cos, "Returns the cosine of a complex number.")
_cfun("IMSIN", cmath.sin, "Returns the sine of a complex number.")
_cfun("IMTAN", cmath.tan, "Returns the tangent of a complex number.")
_cfun("IMCOSH", cmath.cosh, "Returns the hyperbolic cosine of a complex number.")
_cfun("IMSINH", cmath.sinh, "Returns the hyperbolic sine of a complex number.")
_cfun("IMSEC", lambda z: 1 / cmath.cos(z), "Returns the secant of a complex number.")
_cfun("IMSECH", lambda z: 1 / cmath.cosh(z), "Returns the hyperbolic secant of a complex number.")
_cfun("IMCSC", lambda z: 1 / cmath.sin(z), "Returns the cosecant of a complex number.")
_cfun("IMCSCH", lambda z: 1 / cmath.sinh(z), "Returns the hyperbolic cosecant of a complex number.")
_cfun("IMCOT", lambda z: 1 / cmath.tan(z), "Returns the cotangent of a complex number.")
_cfun("IMEXP", cmath.exp, "Returns the exponential of a complex number.")
_cfun("IMLN", cmath.log, "Returns the natural logarithm of a complex number.", True)
_cfun("IMLOG10", cmath.log10, "Returns the base-10 logarithm of a complex number.", True)
_cfun("IMLOG2", lambda z: cmath.log(z) / math.log(2), "Returns the base-2 logarithm of a complex number.", True)
def _im_sqrt(z: complex) -> complex:
    """IMSQRT as Excel computes it, from the polar form: the square root of a negative real keeps
    the stray real part cos(π/2)·√r (IMSQRT("-4") is 1.22464679914735E-16+2i in Excel 16 for Mac);
    the squared magnitude counts as 0 below the smallest normal number (IMSQRT("1E-154") is 0) and
    overflows to #NUM! (IMSQRT("1E155"))."""
    s = z.real * z.real + z.imag * z.imag
    if not math.isfinite(s):
        raise XlError(CellError.NUM)
    if s < sys.float_info.min:
        return 0j
    r = math.sqrt(math.sqrt(s))
    th = math.atan2(z.imag, z.real) / 2.0
    return complex(r * math.cos(th), r * math.sin(th))


_cfun("IMSQRT", _im_sqrt, "Returns the square root of a complex number.")


@fn("IMPOWER", 2, 2, E, S, "IMPOWER(inumber, number)", "Returns a complex number raised to an integer power.")
def _impower(args, ctx):
    z, s = _cx(args)
    n = num(args, 1)
    if z == 0 and n <= 0:
        raise XlError(CellError.NUM)
    r, th = abs(z), cmath.phase(z)
    return format_complex(cmath.rect(r ** n, th * n), s)


def _cx_all(args):
    out = []
    suffix = None
    for a in args:
        items = a.value.data if isinstance(a.value, Array) else [a.value]
        for v in items:
            if v is None and a.from_ref:
                continue
            z, s = parse_complex(v)
            if isinstance(v, str) and ("i" in v or "j" in v):
                if suffix is not None and s != suffix:
                    raise XlError(CellError.VALUE)
                suffix = s
            out.append(z)
    return out, suffix or "i"


@fn("IMSUM", 1, VAR, E, A, "IMSUM(inumber1, [inumber2], …)", "Returns the sum of complex numbers.")
def _imsum(args, ctx):
    zs, s = _cx_all(args)
    return format_complex(sum(zs, complex(0, 0)), s)


@fn("IMPRODUCT", 1, VAR, E, A, "IMPRODUCT(inumber1, [inumber2], …)", "Returns the product of complex numbers.")
def _improduct(args, ctx):
    zs, s = _cx_all(args)
    p = complex(1, 0)
    for z in zs:
        p *= z
    return format_complex(p, s)


@fn("IMSUB", 2, 2, E, S, "IMSUB(inumber1, inumber2)", "Returns the difference between two complex numbers.")
def _imsub(args, ctx):
    a, s = _cx(args, 0)
    b, _ = _cx(args, 1)
    return format_complex(a - b, s)


@fn("IMDIV", 2, 2, E, S, "IMDIV(inumber1, inumber2)", "Returns the quotient of two complex numbers.")
def _imdiv(args, ctx):
    a, s = _cx(args, 0)
    b, _ = _cx(args, 1)
    if b == 0:
        raise XlError(CellError.NUM)
    return format_complex(a / b, s)


# ---------------------------------------------------------------- Bessel functions


def _bessel_j(n: int, x: float) -> float:
    # Series expansion (accurate for moderate x), fallback to integral form for large x.
    if abs(x) > 25:
        steps = 2000
        h = math.pi / steps
        s = 0.0
        for k in range(steps + 1):
            t = k * h
            w = 0.5 if k in (0, steps) else 1.0
            s += w * math.cos(n * t - x * math.sin(t))
        return s * h / math.pi
    total = 0.0
    term_x = (x / 2.0)
    for m in range(0, 200):
        t = (-1) ** m * term_x ** (2 * m + n) / (math.factorial(m) * math.gamma(m + n + 1))
        total += t
        if abs(t) < 1e-17 * max(1.0, abs(total)) and m > 5:
            break
    return total


def _bessel_i(n: int, x: float) -> float:
    total = 0.0
    for m in range(0, 300):
        t = (x / 2.0) ** (2 * m + n) / (math.factorial(m) * math.gamma(m + n + 1))
        total += t
        if abs(t) < 1e-17 * max(1.0, abs(total)) and m > 5:
            break
    return total


def _bessel_y(n: int, x: float) -> float:
    # Integral representation: Y_n(x) = 1/π ∫0^π sin(x sinθ − nθ) dθ − 1/π ∫0^∞ (e^{nt} + (−1)^n e^{−nt}) e^{−x sinh t} dt
    steps = 4000
    h = math.pi / steps
    a = 0.0
    for k in range(steps + 1):
        t = k * h
        w = 0.5 if k in (0, steps) else 1.0
        a += w * math.sin(x * math.sin(t) - n * t)
    a *= h / math.pi
    b = 0.0
    tmax = math.asinh(50.0 / x) + 1 if x > 0 else 10
    h2 = tmax / steps
    for k in range(steps + 1):
        t = k * h2
        w = 0.5 if k in (0, steps) else 1.0
        b += w * (math.exp(n * t) + (-1) ** n * math.exp(-n * t)) * math.exp(-x * math.sinh(t))
    b *= h2 / math.pi
    return a - b


def _bessel_k(n: int, x: float) -> float:
    # K_n(x) = ∫0^∞ e^{−x cosh t} cosh(nt) dt
    steps = 4000
    tmax = math.acosh(max(50.0 / x, 1.0)) + 2
    h = tmax / steps
    s = 0.0
    for k in range(steps + 1):
        t = k * h
        w = 0.5 if k in (0, steps) else 1.0
        s += w * math.exp(-x * math.cosh(t)) * math.cosh(n * t)
    return s * h


def _bessel(name: str, f, need_pos: bool) -> None:
    @fn(name, 2, 2, E, S, f"{name}(x, n)", f"Returns the modified/ordinary Bessel function {name[-1]}n(x).")
    def _impl(args, ctx, f=f):
        x, n = num(args, 0), math.trunc(num(args, 1))
        if n < 0 or (need_pos and x <= 0):
            raise XlError(CellError.NUM)
        return f(int(n), x)


_bessel("BESSELJ", _bessel_j, False)
_bessel("BESSELI", _bessel_i, False)
_bessel("BESSELY", _bessel_y, True)
_bessel("BESSELK", _bessel_k, True)


# ---------------------------------------------------------------- CONVERT

_PREFIXES = {
    "Y": 1e24, "Z": 1e21, "E": 1e18, "P": 1e15, "T": 1e12, "G": 1e9, "M": 1e6, "k": 1e3, "h": 1e2, "da": 1e1,
    "d": 1e-1, "c": 1e-2, "m": 1e-3, "u": 1e-6, "n": 1e-9, "p": 1e-12, "f": 1e-15, "a": 1e-18, "z": 1e-21, "y": 1e-24,
}
_BIN_PREFIXES = {"ki": 2**10, "Mi": 2**20, "Gi": 2**30, "Ti": 2**40, "Pi": 2**50, "Ei": 2**60, "Zi": 2**70, "Yi": 2**80}

# unit -> (category, factor to base unit, allows metric prefix)
_UNITS = {
    # mass (base kg)
    "g": ("mass", 1e-3, True), "sg": ("mass", 14.59390294, False), "lbm": ("mass", 0.45359237, False),
    "u": ("mass", 1.66053906660e-27, True), "ozm": ("mass", 0.028349523125, False), "grain": ("mass", 6.479891e-5, False),
    "cwt": ("mass", 45.359237, False), "shweight": ("mass", 45.359237, False), "uk_cwt": ("mass", 50.80234544, False),
    "lcwt": ("mass", 50.80234544, False), "hweight": ("mass", 50.80234544, False), "stone": ("mass", 6.35029318, False),
    "ton": ("mass", 907.18474, False), "uk_ton": ("mass", 1016.0469088, False), "LTON": ("mass", 1016.0469088, False),
    "brton": ("mass", 1016.0469088, False),
    # distance (base m)
    "m": ("distance", 1.0, True), "mi": ("distance", 1609.344, False), "Nmi": ("distance", 1852.0, False),
    "in": ("distance", 0.0254, False), "ft": ("distance", 0.3048, False), "yd": ("distance", 0.9144, False),
    "ang": ("distance", 1e-10, True), "ell": ("distance", 1.143, False), "ly": ("distance", 9.4607304725808e15, True),
    "parsec": ("distance", 3.0856775814671900e16, False), "pc": ("distance", 3.0856775814671900e16, False),
    "Pica": ("distance", 0.0254 / 6, False), "Picapt": ("distance", 0.0254 / 72, False), "pica": ("distance", 0.0254 / 6, False),
    "survey_mi": ("distance", 1609.3472186944, False),
    # time (base s)
    "yr": ("time", 31557600.0, False), "day": ("time", 86400.0, False), "d": ("time", 86400.0, False),
    "hr": ("time", 3600.0, False), "mn": ("time", 60.0, False), "min": ("time", 60.0, False),
    "sec": ("time", 1.0, True), "s": ("time", 1.0, True),
    # pressure (base Pa)
    "Pa": ("pressure", 1.0, True), "p": ("pressure", 1.0, True), "atm": ("pressure", 101325.0, True),
    "at": ("pressure", 101325.0, True), "mmHg": ("pressure", 133.322, True), "psi": ("pressure", 6894.757293168, False),
    "Torr": ("pressure", 133.322368421, False),
    # force (base N)
    "N": ("force", 1.0, True), "dyn": ("force", 1e-5, True), "dy": ("force", 1e-5, True), "lbf": ("force", 4.4482216152605, False),
    "pond": ("force", 0.00980665, True),
    # energy (base J)
    "J": ("energy", 1.0, True), "e": ("energy", 1e-7, True), "c": ("energy", 4.184, True), "cal": ("energy", 4.1868, True),
    "eV": ("energy", 1.602176634e-19, True), "ev": ("energy", 1.602176634e-19, True), "HPh": ("energy", 2684519.537696172792, False),
    "hh": ("energy", 2684519.537696172792, False), "Wh": ("energy", 3600.0, True), "wh": ("energy", 3600.0, True),
    "flb": ("energy", 0.0421401100938048, False), "BTU": ("energy", 1055.05585262, False), "btu": ("energy", 1055.05585262, False),
    # power (base W)
    "HP": ("power", 745.69987158227022, False), "h": ("power", 745.69987158227022, False), "PS": ("power", 735.49875, False),
    "W": ("power", 1.0, True), "w": ("power", 1.0, True),
    # magnetism (base T)
    "T": ("magnetism", 1.0, True), "ga": ("magnetism", 1e-4, True),
    # temperature handled specially
    "C": ("temperature", None, False), "cel": ("temperature", None, False), "F": ("temperature", None, False),
    "fah": ("temperature", None, False), "K": ("temperature", None, True), "kel": ("temperature", None, True),
    "Rank": ("temperature", None, False), "Reau": ("temperature", None, False),
    # volume (base m3)
    "tsp": ("volume", 4.92892159375e-6, False), "tspm": ("volume", 5e-6, False), "tbs": ("volume", 1.478676478125e-5, False),
    "oz": ("volume", 2.95735295625e-5, False), "cup": ("volume", 2.365882365e-4, False), "pt": ("volume", 4.73176473e-4, False),
    "us_pt": ("volume", 4.73176473e-4, False), "uk_pt": ("volume", 5.6826125e-4, False), "qt": ("volume", 9.46352946e-4, False),
    "uk_qt": ("volume", 1.1365225e-3, False), "gal": ("volume", 3.785411784e-3, False), "uk_gal": ("volume", 4.54609e-3, False),
    "l": ("volume", 1e-3, True), "L": ("volume", 1e-3, True), "lt": ("volume", 1e-3, True), "m3": ("volume", 1.0, True),
    "mi3": ("volume", 4.168181825440579584e9, False), "yd3": ("volume", 0.764554857984, False), "ft3": ("volume", 0.028316846592, False),
    "in3": ("volume", 1.6387064e-5, False), "ang3": ("volume", 1e-30, True), "barrel": ("volume", 0.158987294928, False),
    "bushel": ("volume", 0.03523907016688, False), "GRT": ("volume", 2.8316846592, False), "regton": ("volume", 2.8316846592, False),
    "MTON": ("volume", 1.13267386368, False),
    # area (base m2)
    "uk_acre": ("area", 4046.8564224, False), "us_acre": ("area", 4046.8726098742, False), "ang2": ("area", 1e-20, True),
    "ar": ("area", 100.0, True), "ft2": ("area", 0.09290304, False), "ha": ("area", 10000.0, False), "in2": ("area", 0.00064516, False),
    "ly2": ("area", 8.950542e31, False), "m2": ("area", 1.0, True), "Morgen": ("area", 2500.0, False), "mi2": ("area", 2589988.110336, False),
    "Nmi2": ("area", 3429904.0, False), "yd2": ("area", 0.83612736, False),
    # information (base bit)
    "bit": ("information", 1.0, True), "byte": ("information", 8.0, True),
    # speed (base m/s)
    "admkn": ("speed", 0.514773333, False), "kn": ("speed", 0.514444444, False), "m/h": ("speed", 1 / 3600.0, True),
    "m/hr": ("speed", 1 / 3600.0, True), "m/s": ("speed", 1.0, True), "m/sec": ("speed", 1.0, True), "mph": ("speed", 0.44704, False),
}


def _unit(u: str):
    if u in _UNITS:
        cat, f, _ = _UNITS[u]
        return cat, f, u, 1.0
    for p, mult in list(_BIN_PREFIXES.items()) + list(_PREFIXES.items()):
        if u.startswith(p) and len(u) > len(p):
            base = u[len(p):]
            if base in _UNITS and _UNITS[base][2]:
                cat, f, _ = _UNITS[base]
                if p in _BIN_PREFIXES and cat != "information":
                    continue
                return cat, f, base, mult
    raise XlError(CellError.NA)


def _to_kelvin(v: float, u: str) -> float:
    if u in ("C", "cel"):
        return v + 273.15
    if u in ("F", "fah"):
        return (v - 32) * 5 / 9 + 273.15
    if u in ("K", "kel"):
        return v
    if u == "Rank":
        return v * 5 / 9
    if u == "Reau":
        return v * 1.25 + 273.15
    raise XlError(CellError.NA)


def _from_kelvin(k: float, u: str) -> float:
    if u in ("C", "cel"):
        return k - 273.15
    if u in ("F", "fah"):
        return (k - 273.15) * 9 / 5 + 32
    if u in ("K", "kel"):
        return k
    if u == "Rank":
        return k * 9 / 5
    if u == "Reau":
        return (k - 273.15) * 0.8
    raise XlError(CellError.NA)


@fn("CONVERT", 3, 3, E, S, "CONVERT(number, from_unit, to_unit)", "Converts a number from one measurement system to another.")
def _convert(args, ctx):
    v = num(args, 0)
    fu, tu = text(args, 1), text(args, 2)
    c1, f1, b1, m1 = _unit(fu)
    c2, f2, b2, m2 = _unit(tu)
    if c1 != c2:
        raise XlError(CellError.NA)
    if c1 == "temperature":
        return _from_kelvin(_to_kelvin(v * m1, b1), b2) / m2
    return v * m1 * f1 / (f2 * m2)
