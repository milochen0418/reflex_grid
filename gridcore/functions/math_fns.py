"""Math & trigonometry functions."""

from __future__ import annotations

import math
from typing import List

from ..core.value import Array, CellError, Value, XlError, to_number
from .criteria import Criterion
from .registry import MAX_CELLS, A, Arg, Ctx, S, VAR, fn
from .util import (
    array_val,
    as_array,
    check_size,
    excel_mod,
    has,
    int_,
    num,
    numbers,
    opt_int,
    opt_num,
    round_half_away,
    text,
    to_int,
    to_sig_digits,
)

M = "Math & Trig"


def _f(n: float) -> float:
    if not math.isfinite(n):
        raise XlError(CellError.NUM)
    return n


# ---------------------------------------------------------------- sums and products


@fn("SUM", 1, VAR, M, A, "SUM(number1, [number2], …)", "Adds all the numbers in a range of cells.")
def _sum(args: List[Arg], ctx: Ctx) -> Value:
    return _f(math.fsum(numbers(args)))


@fn("SUMSQ", 1, VAR, M, A, "SUMSQ(number1, [number2], …)", "Returns the sum of the squares of the arguments.")
def _sumsq(args, ctx):
    return _f(math.fsum(x * x for x in numbers(args)))


def _pairs(args) -> List[tuple]:
    a, b = as_array(args[0].value), as_array(args[1].value)
    if len(a.data) != len(b.data):
        raise XlError(CellError.NA)
    out = []
    for x, y in zip(a.data, b.data):
        if isinstance(x, CellError):
            raise XlError(x)
        if isinstance(y, CellError):
            raise XlError(y)
        if type(x) is float and type(y) is float:
            out.append((x, y))
    return out


@fn("SUMX2MY2", 2, 2, M, A, "SUMX2MY2(array_x, array_y)", "Returns the sum of the difference of squares of corresponding values in two arrays.")
def _sumx2my2(args, ctx):
    return _f(math.fsum(x * x - y * y for x, y in _pairs(args)))


@fn("SUMX2PY2", 2, 2, M, A, "SUMX2PY2(array_x, array_y)", "Returns the sum of the sum of squares of corresponding values in two arrays.")
def _sumx2py2(args, ctx):
    return _f(math.fsum(x * x + y * y for x, y in _pairs(args)))


@fn("SUMXMY2", 2, 2, M, A, "SUMXMY2(array_x, array_y)", "Returns the sum of squares of differences of corresponding values in two arrays.")
def _sumxmy2(args, ctx):
    return _f(math.fsum((x - y) ** 2 for x, y in _pairs(args)))


@fn("PRODUCT", 1, VAR, M, A, "PRODUCT(number1, [number2], …)", "Multiplies its arguments.")
def _product(args, ctx):
    ns = numbers(args)
    if not ns:
        return 0.0
    p = 1.0
    for x in ns:
        p *= x
    return _f(p)


@fn("SUMPRODUCT", 1, VAR, M, A, "SUMPRODUCT(array1, [array2], …)", "Returns the sum of the products of corresponding array components.")
def _sumproduct(args, ctx):
    arrays = [as_array(a.value) for a in args]
    r, c = arrays[0].rows, arrays[0].cols
    for a in arrays:
        if a.rows != r or a.cols != c:
            raise XlError(CellError.VALUE)
    total = 0.0
    n = r * c
    for i in range(n):
        p = 1.0
        for a in arrays:
            v = a.data[i]
            if isinstance(v, CellError):
                raise XlError(v)
            p *= v if type(v) is float else 0.0
        total += p
    return _f(total)


@fn("SUMIF", 2, 3, M, (False, True, False), "SUMIF(range, criteria, [sum_range])", "Adds the cells specified by a given condition or criteria.")
def _sumif(args, ctx):
    rng = as_array(args[0].value)
    crit = Criterion.parse(args[1].value)
    src = as_array(args[2].value) if len(args) > 2 else rng
    total = 0.0
    for i, v in enumerate(rng.data):
        if crit.matches(v):
            r, c = divmod(i, rng.cols)
            x = src.get(r, c)
            if isinstance(x, CellError):
                raise XlError(x)
            if type(x) is float:
                total += x
    return _f(total)


def _ifs_mask(args, start: int) -> List[bool]:
    first = as_array(args[start].value)
    n = len(first.data)
    mask = [True] * n
    for i in range(start, len(args), 2):
        rng = as_array(args[i].value)
        if rng.rows != first.rows or rng.cols != first.cols:
            raise XlError(CellError.VALUE)
        crit = Criterion.parse(args[i + 1].value)
        for j, v in enumerate(rng.data):
            if mask[j] and not crit.matches(v):
                mask[j] = False
    return mask


def ifs_values(args, target_index: int) -> List[float]:
    """Numbers of the target range where every (range, criterion) pair matches."""
    if (len(args) - 1) % 2 != 0:
        raise XlError(CellError.VALUE)
    target = as_array(args[target_index].value)
    mask = _ifs_mask(args, 1)
    if len(mask) != len(target.data):
        raise XlError(CellError.VALUE)
    out = []
    for m, v in zip(mask, target.data):
        if m:
            if isinstance(v, CellError):
                raise XlError(v)
            if type(v) is float:
                out.append(v)
    return out


@fn("SUMIFS", 3, VAR, M, (False, False, True, False, True), "SUMIFS(sum_range, criteria_range1, criteria1, …)", "Adds the cells specified by a given set of conditions or criteria.")
def _sumifs(args, ctx):
    return _f(math.fsum(ifs_values(args, 0)))


# ---------------------------------------------------------------- rounding


@fn("ROUND", 2, 2, M, S, "ROUND(number, num_digits)", "Rounds a number to a specified number of digits.")
def _round(args, ctx):
    return round_half_away(num(args, 0), int_(args, 1))


def _round_dir(x: float, d: int, up: bool) -> float:
    """ROUNDUP/ROUNDDOWN/TRUNC: away from or toward zero at ``d`` decimals, on the
    15-significant-digit value (so binary noise such as 0.1+0.2 doesn't round up)."""
    if x == 0 or not math.isfinite(x):
        return 0.0 if x == 0 else x
    d = max(min(d, 400), -400)
    a = abs(x)
    if d > 300:
        return x
    y = a * (10.0 ** d) if d >= 0 else a / (10.0 ** -d)
    if not math.isfinite(y):
        return x
    y = to_sig_digits(y, 15)
    y = math.ceil(y) if up else math.floor(y)
    r = y / (10.0 ** d) if d >= 0 else y * (10.0 ** -d)
    return -r if x < 0 else r


@fn("ROUNDUP", 2, 2, M, S, "ROUNDUP(number, num_digits)", "Rounds a number up, away from zero.")
def _roundup(args, ctx):
    return _round_dir(num(args, 0), int_(args, 1), True)


@fn("ROUNDDOWN", 2, 2, M, S, "ROUNDDOWN(number, num_digits)", "Rounds a number down, toward zero.")
def _rounddown(args, ctx):
    return _round_dir(num(args, 0), int_(args, 1), False)


@fn("MROUND", 2, 2, M, S, "MROUND(number, multiple)", "Returns a number rounded to the desired multiple.")
def _mround(args, ctx):
    n, m = num(args, 0), num(args, 1)
    if m == 0:
        return 0.0
    if (n > 0 and m < 0) or (n < 0 and m > 0):
        raise XlError(CellError.NUM)
    return round_half_away(round_half_away(n / m, 0) * m, 12)


def _snap(x: float) -> float:
    return round_half_away(x, 12) if abs(x) < 1e15 else x


@fn("CEILING", 2, 2, M, S, "CEILING(number, significance)", "Rounds a number up to the nearest multiple of significance.")
def _ceiling(args, ctx):
    n, s = num(args, 0), num(args, 1)
    if s == 0 or n == 0:
        return 0.0
    if n > 0 and s < 0:
        raise XlError(CellError.NUM)
    return _snap(math.ceil(_snap(n / s)) * s)


@fn("CEILING.MATH", 1, 3, M, S, "CEILING.MATH(number, [significance], [mode])", "Rounds a number up to the nearest integer or nearest multiple of significance.")
def _ceiling_math(args, ctx):
    n = num(args, 0)
    s = abs(opt_num(args, 1, 1.0)) if has(args, 1) else 1.0
    mode = opt_num(args, 2, 0.0)
    if s == 0 or n == 0:
        return 0.0
    q = _snap(n / s)
    if n < 0 and mode != 0:
        return _snap(math.floor(q) * s)
    return _snap(math.ceil(q) * s)


@fn("CEILING.PRECISE", 1, 2, M, S, "CEILING.PRECISE(number, [significance])", "Rounds a number up to the nearest integer or multiple of significance, regardless of sign.")
def _ceiling_precise(args, ctx):
    n = num(args, 0)
    s = abs(opt_num(args, 1, 1.0)) if has(args, 1) else 1.0
    if s == 0 or n == 0:
        return 0.0
    return _snap(math.ceil(_snap(n / s)) * s)


@fn("ISO.CEILING", 1, 2, M, S, "ISO.CEILING(number, [significance])", "Rounds a number up to the nearest integer or multiple of significance.")
def _iso_ceiling(args, ctx):
    return _ceiling_precise(args, ctx)


@fn("FLOOR", 2, 2, M, S, "FLOOR(number, significance)", "Rounds a number down, toward zero.")
def _floor(args, ctx):
    n, s = num(args, 0), num(args, 1)
    if n == 0:
        return 0.0
    if s == 0:
        raise XlError(CellError.DIV0)
    if n > 0 and s < 0:
        raise XlError(CellError.NUM)
    q = _snap(n / s)
    return _snap(math.floor(q) * s)


@fn("FLOOR.MATH", 1, 3, M, S, "FLOOR.MATH(number, [significance], [mode])", "Rounds a number down to the nearest integer or nearest multiple of significance.")
def _floor_math(args, ctx):
    n = num(args, 0)
    s = abs(opt_num(args, 1, 1.0)) if has(args, 1) else 1.0
    mode = opt_num(args, 2, 0.0)
    if s == 0 or n == 0:
        return 0.0
    q = _snap(n / s)
    if n < 0 and mode != 0:
        return _snap(math.ceil(q) * s)
    return _snap(math.floor(q) * s)


@fn("FLOOR.PRECISE", 1, 2, M, S, "FLOOR.PRECISE(number, [significance])", "Rounds a number down to the nearest integer or multiple of significance, regardless of sign.")
def _floor_precise(args, ctx):
    n = num(args, 0)
    s = abs(opt_num(args, 1, 1.0)) if has(args, 1) else 1.0
    if s == 0 or n == 0:
        return 0.0
    return _snap(math.floor(_snap(n / s)) * s)


@fn("INT", 1, 1, M, S, "INT(number)", "Rounds a number down to the nearest integer.")
def _int(args, ctx):
    return float(math.floor(num(args, 0)))


@fn("TRUNC", 1, 2, M, S, "TRUNC(number, [num_digits])", "Truncates a number to an integer.")
def _trunc(args, ctx):
    return _round_dir(num(args, 0), opt_int(args, 1, 0), False)


@fn("MOD", 2, 2, M, S, "MOD(number, divisor)", "Returns the remainder from division.")
def _mod(args, ctx):
    n, d = num(args, 0), num(args, 1)
    r = excel_mod(n, d)
    if abs(r - d) < abs(d) * 1e-15:
        return 0.0
    return r


@fn("QUOTIENT", 2, 2, M, S, "QUOTIENT(numerator, denominator)", "Returns the integer portion of a division.")
def _quotient(args, ctx):
    n, d = num(args, 0), num(args, 1)
    if d == 0:
        raise XlError(CellError.DIV0)
    return float(math.trunc(n / d))


@fn("ABS", 1, 1, M, S, "ABS(number)", "Returns the absolute value of a number.")
def _abs(args, ctx):
    return abs(num(args, 0))


@fn("SIGN", 1, 1, M, S, "SIGN(number)", "Returns the sign of a number.")
def _sign(args, ctx):
    n = num(args, 0)
    return 1.0 if n > 0 else (-1.0 if n < 0 else 0.0)


@fn("SQRT", 1, 1, M, S, "SQRT(number)", "Returns a positive square root.")
def _sqrt(args, ctx):
    n = num(args, 0)
    if n < 0:
        raise XlError(CellError.NUM)
    return math.sqrt(n)


@fn("SQRTPI", 1, 1, M, S, "SQRTPI(number)", "Returns the square root of (number * pi).")
def _sqrtpi(args, ctx):
    n = num(args, 0)
    if n < 0:
        raise XlError(CellError.NUM)
    return math.sqrt(n * math.pi)


@fn("POWER", 2, 2, M, S, "POWER(number, power)", "Returns the result of a number raised to a power.")
def _power(args, ctx):
    x, y = num(args, 0), num(args, 1)
    if x == 0 and y == 0:
        raise XlError(CellError.NUM)
    if x == 0 and y < 0:
        raise XlError(CellError.DIV0)
    if x < 0 and y != int(y):
        raise XlError(CellError.NUM)
    return _f(x ** y)


@fn("EXP", 1, 1, M, S, "EXP(number)", "Returns e raised to the power of a given number.")
def _exp(args, ctx):
    return _f(math.exp(num(args, 0)))


@fn("LN", 1, 1, M, S, "LN(number)", "Returns the natural logarithm of a number.")
def _ln(args, ctx):
    n = num(args, 0)
    if n <= 0:
        raise XlError(CellError.NUM)
    return math.log(n)


@fn("LOG", 1, 2, M, S, "LOG(number, [base])", "Returns the logarithm of a number to a specified base.")
def _log(args, ctx):
    n = num(args, 0)
    b = opt_num(args, 1, 10.0)
    if n <= 0 or b <= 0:
        raise XlError(CellError.NUM)
    if b == 1:
        raise XlError(CellError.DIV0)
    return math.log(n) / math.log(b) if b != 10 else math.log10(n)


@fn("LOG10", 1, 1, M, S, "LOG10(number)", "Returns the base-10 logarithm of a number.")
def _log10(args, ctx):
    n = num(args, 0)
    if n <= 0:
        raise XlError(CellError.NUM)
    return math.log10(n)


@fn("PI", 0, 0, M, S, "PI()", "Returns the value of pi.")
def _pi(args, ctx):
    return math.pi


@fn("RAND", 0, 0, M, S, "RAND()", "Returns a random number between 0 and 1.", volatile=True)
def _rand(args, ctx):
    return ctx.random()


@fn("RANDBETWEEN", 2, 2, M, S, "RANDBETWEEN(bottom, top)", "Returns a random number between the numbers you specify.", volatile=True)
def _randbetween(args, ctx):
    lo, hi = math.ceil(num(args, 0)), math.floor(num(args, 1))
    if lo > hi:
        raise XlError(CellError.NUM)
    return float(lo + math.floor(ctx.random() * (hi - lo + 1)))


# ---------------------------------------------------------------- combinatorics


@fn("FACT", 1, 1, M, S, "FACT(number)", "Returns the factorial of a number.")
def _fact(args, ctx):
    n = num(args, 0)
    if n < 0:
        raise XlError(CellError.NUM)
    n = int(n)
    if n > 170:
        raise XlError(CellError.NUM)
    return float(math.factorial(n))


@fn("FACTDOUBLE", 1, 1, M, S, "FACTDOUBLE(number)", "Returns the double factorial of a number.")
def _factdouble(args, ctx):
    n = num(args, 0)
    if n < -1:
        raise XlError(CellError.NUM)
    n = int(n)
    if n > 300:
        raise XlError(CellError.NUM)
    p = 1
    while n > 1:
        p *= n
        n -= 2
    return _f(float(p))


def _comb(n: float, k: float) -> float:
    if n > 1e15 or k > 1e15:
        raise XlError(CellError.NUM)
    n, k = int(n), int(k)
    if n < 0 or k < 0 or k > n:
        raise XlError(CellError.NUM)
    if n > 2000:
        lg = math.lgamma(n + 1) - math.lgamma(k + 1) - math.lgamma(n - k + 1)
        if lg > 709.78:
            raise XlError(CellError.NUM)
        return _f(float(round(math.exp(lg))) if lg < 36 else math.exp(lg))
    return _f(float(math.comb(n, k)))


@fn("COMBIN", 2, 2, M, S, "COMBIN(number, number_chosen)", "Returns the number of combinations for a given number of objects.")
def _combin(args, ctx):
    return _comb(num(args, 0), num(args, 1))


@fn("COMBINA", 2, 2, M, S, "COMBINA(number, number_chosen)", "Returns the number of combinations with repetitions for a given number of items.")
def _combina(args, ctx):
    n0, k0 = num(args, 0), num(args, 1)
    if n0 > 1e15 or k0 > 1e15:
        raise XlError(CellError.NUM)
    n, k = int(n0), int(k0)
    if n < 0 or k < 0 or (n == 0 and k > 0):
        raise XlError(CellError.NUM)
    if n == 0 and k == 0:
        return 1.0
    return _comb(n + k - 1, k)


@fn("PERMUT", 2, 2, M, S, "PERMUT(number, number_chosen)", "Returns the number of permutations for a given number of objects.")
def _permut(args, ctx):
    n0, k0 = num(args, 0), num(args, 1)
    if n0 > 1e15 or k0 > 1e15:
        raise XlError(CellError.NUM)
    n, k = int(n0), int(k0)
    if n < 0 or k < 0 or k > n:
        raise XlError(CellError.NUM)
    if n > 2000:
        lg = math.lgamma(n + 1) - math.lgamma(n - k + 1)
        if lg > 709.78:
            raise XlError(CellError.NUM)
        return _f(math.exp(lg))
    return _f(float(math.perm(n, k)))


@fn("PERMUTATIONA", 2, 2, M, S, "PERMUTATIONA(number, number_chosen)", "Returns the number of permutations (with repetitions) for a given number of objects.")
def _permutationa(args, ctx):
    n, k = int(num(args, 0)), int(num(args, 1))
    if n < 0 or k < 0:
        raise XlError(CellError.NUM)
    return _f(float(n) ** k)


@fn("MULTINOMIAL", 1, VAR, M, A, "MULTINOMIAL(number1, [number2], …)", "Returns the multinomial of a set of numbers.")
def _multinomial(args, ctx):
    ns = [int(x) for x in numbers(args)]
    if any(x < 0 for x in ns):
        raise XlError(CellError.NUM)
    total = sum(ns)
    if total > 170:
        raise XlError(CellError.NUM)
    r = math.factorial(total)
    for x in ns:
        r //= math.factorial(x)
    return _f(float(r))


@fn("GCD", 1, VAR, M, A, "GCD(number1, [number2], …)", "Returns the greatest common divisor.")
def _gcd(args, ctx):
    ns = numbers(args)
    if any(x < 0 for x in ns):
        raise XlError(CellError.NUM)
    g = 0
    for x in ns:
        g = math.gcd(g, int(x))
    return float(g)


@fn("LCM", 1, VAR, M, A, "LCM(number1, [number2], …)", "Returns the least common multiple.")
def _lcm(args, ctx):
    ns = numbers(args)
    if any(x < 0 for x in ns):
        raise XlError(CellError.NUM)
    l = 1
    for x in ns:
        x = int(x)
        if x == 0:
            return 0.0
        l = l * x // math.gcd(l, x)
    return _f(float(l))


@fn("EVEN", 1, 1, M, S, "EVEN(number)", "Rounds a number up to the nearest even integer.")
def _even(args, ctx):
    n = num(args, 0)
    v = math.ceil(abs(n) / 2.0) * 2
    return float(-v if n < 0 else v)


@fn("ODD", 1, 1, M, S, "ODD(number)", "Rounds a number up to the nearest odd integer.")
def _odd(args, ctx):
    n = num(args, 0)
    a = math.ceil(abs(n))
    if a % 2 == 0:
        a += 1
    return float(-a if n < 0 else a)


# ---------------------------------------------------------------- trigonometry


def _unary(name: str, f, desc: str, domain=None):
    @fn(name, 1, 1, M, S, f"{name}(number)", desc)
    def _impl(args, ctx, f=f, domain=domain):
        x = num(args, 0)
        if domain is not None and not domain(x):
            raise XlError(CellError.NUM)
        try:
            return _f(f(x))
        except (ValueError, ZeroDivisionError):
            raise XlError(CellError.NUM) from None

    return _impl


def _div0(f):
    def g(x):
        d = f(x)
        if d == 0:
            raise XlError(CellError.DIV0)
        return 1.0 / d

    return g


_TRIG_LIMIT = 134217728.0
_unary("SIN", math.sin, "Returns the sine of the given angle.", lambda x: abs(x) < _TRIG_LIMIT)
_unary("COS", math.cos, "Returns the cosine of a number.", lambda x: abs(x) < _TRIG_LIMIT)
_unary("TAN", math.tan, "Returns the tangent of a number.", lambda x: abs(x) < _TRIG_LIMIT)
_unary("ASIN", math.asin, "Returns the arcsine of a number.", lambda x: -1 <= x <= 1)
_unary("ACOS", math.acos, "Returns the arccosine of a number.", lambda x: -1 <= x <= 1)
_unary("ATAN", math.atan, "Returns the arctangent of a number.")
_unary("SINH", math.sinh, "Returns the hyperbolic sine of a number.")
_unary("COSH", math.cosh, "Returns the hyperbolic cosine of a number.")
_unary("TANH", math.tanh, "Returns the hyperbolic tangent of a number.")
_unary("ASINH", math.asinh, "Returns the inverse hyperbolic sine of a number.")
_unary("ACOSH", math.acosh, "Returns the inverse hyperbolic cosine of a number.", lambda x: x >= 1)
_unary("ATANH", math.atanh, "Returns the inverse hyperbolic tangent of a number.", lambda x: -1 < x < 1)
_unary("COT", _div0(math.tan), "Returns the cotangent of an angle.", lambda x: abs(x) < _TRIG_LIMIT)
_unary("COTH", _div0(math.tanh), "Returns the hyperbolic cotangent of a number.")
_unary("CSC", _div0(math.sin), "Returns the cosecant of an angle.", lambda x: abs(x) < _TRIG_LIMIT)
_unary("CSCH", _div0(math.sinh), "Returns the hyperbolic cosecant of an angle.")
_unary("SEC", _div0(math.cos), "Returns the secant of an angle.", lambda x: abs(x) < _TRIG_LIMIT)
_unary("SECH", _div0(math.cosh), "Returns the hyperbolic secant of an angle.")
_unary("ACOT", lambda x: math.pi / 2 - math.atan(x), "Returns the arccotangent of a number.")
_unary("ACOTH", lambda x: 0.5 * math.log((x + 1) / (x - 1)), "Returns the hyperbolic arccotangent of a number.", lambda x: abs(x) > 1)
_unary("DEGREES", math.degrees, "Converts radians to degrees.")
_unary("RADIANS", math.radians, "Converts degrees to radians.")


@fn("ATAN2", 2, 2, M, S, "ATAN2(x_num, y_num)", "Returns the arctangent from x- and y-coordinates.")
def _atan2(args, ctx):
    x, y = num(args, 0), num(args, 1)
    if x == 0 and y == 0:
        raise XlError(CellError.DIV0)
    return math.atan2(y, x)


# ---------------------------------------------------------------- numerals and bases

_ROMAN = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]


@fn("ROMAN", 1, 2, M, S, "ROMAN(number, [form])", "Converts an Arabic numeral to roman, as text.")
def _roman(args, ctx):
    n = int(num(args, 0))
    if n < 0 or n > 3999:
        raise XlError(CellError.VALUE)
    out = []
    for v, s in _ROMAN:
        while n >= v:
            out.append(s)
            n -= v
    return "".join(out)


@fn("ARABIC", 1, 1, M, S, "ARABIC(text)", "Converts a Roman number to Arabic, as a number.")
def _arabic(args, ctx):
    t = text(args, 0).strip().upper()
    neg = t.startswith("-")
    if neg:
        t = t[1:]
    vals = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
    if any(ch not in vals for ch in t) or len(t) > 255:
        raise XlError(CellError.VALUE)
    total = 0
    for i, ch in enumerate(t):
        v = vals[ch]
        if i + 1 < len(t) and vals[t[i + 1]] > v:
            total -= v
        else:
            total += v
    return float(-total if neg else total)


_DIGITS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"


@fn("BASE", 2, 3, M, S, "BASE(number, radix, [min_length])", "Converts a number into a text representation with the given radix (base).")
def _base(args, ctx):
    n, radix = num(args, 0), int(num(args, 1))
    min_len = opt_int(args, 2, 0)
    if n < 0 or n >= 2**53 or not 2 <= radix <= 36 or min_len < 0 or min_len > 255:
        raise XlError(CellError.NUM)
    n = int(n)
    s = ""
    while n > 0:
        n, r = divmod(n, radix)
        s = _DIGITS[r] + s
    s = s or "0"
    return s.rjust(min_len, "0")


@fn("DECIMAL", 2, 2, M, S, "DECIMAL(text, radix)", "Converts a text representation of a number in a given base into a decimal number.")
def _decimal(args, ctx):
    t = text(args, 0).strip().upper()
    radix = int(num(args, 1))
    if not 2 <= radix <= 36 or len(t) > 255:
        raise XlError(CellError.NUM)
    if radix == 16 and t.startswith("0X"):
        t = t[2:]
    try:
        return float(int(t, radix)) if t else 0.0
    except ValueError:
        raise XlError(CellError.NUM) from None


# ---------------------------------------------------------------- matrices


def _matrix(v: Value) -> List[List[float]]:
    a = as_array(v)
    out = []
    for r in range(a.rows):
        row = []
        for c in range(a.cols):
            x = a.get(r, c)
            if isinstance(x, CellError):
                raise XlError(x)
            if type(x) is not float:
                raise XlError(CellError.VALUE)
            row.append(x)
        out.append(row)
    return out


@fn("MMULT", 2, 2, M, A, "MMULT(array1, array2)", "Returns the matrix product of two arrays.")
def _mmult(args, ctx):
    a, b = _matrix(args[0].value), _matrix(args[1].value)
    n, m, p = len(a), len(a[0]), len(b[0])
    if len(b) != m:
        raise XlError(CellError.VALUE)
    if n * p > MAX_CELLS:
        raise XlError(CellError.NUM)
    data = [math.fsum(a[i][k] * b[k][j] for k in range(m)) for i in range(n) for j in range(p)]
    return array_val(n, p, data)


def _lu_det(m: List[List[float]]) -> float:
    n = len(m)
    a = [row[:] for row in m]
    det = 1.0
    for i in range(n):
        piv = max(range(i, n), key=lambda r: abs(a[r][i]))
        if a[piv][i] == 0:
            return 0.0
        if piv != i:
            a[i], a[piv] = a[piv], a[i]
            det = -det
        det *= a[i][i]
        for r in range(i + 1, n):
            f = a[r][i] / a[i][i]
            for c in range(i, n):
                a[r][c] -= f * a[i][c]
    return det


@fn("MDETERM", 1, 1, M, A, "MDETERM(array)", "Returns the matrix determinant of an array.")
def _mdeterm(args, ctx):
    m = _matrix(args[0].value)
    if len(m) != len(m[0]):
        raise XlError(CellError.VALUE)
    return _f(_lu_det(m))


def invert(m: List[List[float]]) -> List[List[float]]:
    n = len(m)
    a = [row[:] + [1.0 if i == j else 0.0 for j in range(n)] for i, row in enumerate(m)]
    for i in range(n):
        piv = max(range(i, n), key=lambda r: abs(a[r][i]))
        if abs(a[piv][i]) < 1e-300:
            raise XlError(CellError.NUM)
        a[i], a[piv] = a[piv], a[i]
        p = a[i][i]
        a[i] = [x / p for x in a[i]]
        for r in range(n):
            if r != i and a[r][i] != 0:
                f = a[r][i]
                a[r] = [x - f * y for x, y in zip(a[r], a[i])]
    return [row[n:] for row in a]


@fn("MINVERSE", 1, 1, M, A, "MINVERSE(array)", "Returns the matrix inverse of an array.")
def _minverse(args, ctx):
    m = _matrix(args[0].value)
    n = len(m)
    if n != len(m[0]):
        raise XlError(CellError.VALUE)
    inv = invert(m)
    return array_val(n, n, [_f(x) for row in inv for x in row])


@fn("MUNIT", 1, 1, M, S, "MUNIT(dimension)", "Returns the unit matrix or the specified dimension.")
def _munit(args, ctx):
    n = int(num(args, 0))
    if n < 1:
        raise XlError(CellError.VALUE)
    check_size(n, n)
    return array_val(n, n, [1.0 if i == j else 0.0 for i in range(n) for j in range(n)])


@fn("SERIESSUM", 4, 4, M, (True, True, True, False), "SERIESSUM(x, n, m, coefficients)", "Returns the sum of a power series based on the formula.")
def _seriessum(args, ctx):
    x, n, m = num(args, 0), num(args, 1), num(args, 2)
    coefs = as_array(args[3].value).data
    total = 0.0
    for i, a in enumerate(coefs):
        if isinstance(a, CellError):
            raise XlError(a)
        if type(a) is not float:
            raise XlError(CellError.VALUE)
        total += a * x ** (n + i * m)
    return _f(total)
