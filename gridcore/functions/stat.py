"""Statistical functions: aggregates, order statistics, regression and probability
distributions (with their inverses), plus the SUBTOTAL/AGGREGATE helpers."""

from __future__ import annotations

import math
from typing import List, Optional, Sequence, Tuple

from ..core.value import Array, CellError, Value, XlError, scalar, to_number
from .criteria import Criterion
from .registry import A, Arg, Ctx, S, VAR, fn
from .special import (
    chisq_cdf,
    gamma_cdf,
    f_cdf,
    invert,
    ln_beta,
    norm_cdf,
    norm_inv,
    norm_pdf,
    reg_gamma_p,
    reg_inc_beta,
    t_cdf,
    t_pdf,
)
from .util import (
    array_numbers,
    array_val,
    as_array,
    boolean,
    has,
    num,
    numbers,
    numbers_a,
    opt_num,
    paired_numbers,
)

ST = "Statistical"
COMPAT = "Compatibility"


def _f(n: float) -> float:
    if not math.isfinite(n):
        raise XlError(CellError.NUM)
    return n


# ---------------------------------------------------------------- helpers over number lists


def mean(x: Sequence[float]) -> float:
    if not x:
        raise XlError(CellError.DIV0)
    return math.fsum(x) / len(x)


def devsq(x: Sequence[float]) -> float:
    if not x:
        return 0.0
    m = math.fsum(x) / len(x)
    return math.fsum((v - m) ** 2 for v in x)


def var_s(x: Sequence[float]) -> float:
    if len(x) < 2:
        raise XlError(CellError.DIV0)
    return devsq(x) / (len(x) - 1)


def var_p(x: Sequence[float]) -> float:
    if not x:
        raise XlError(CellError.DIV0)
    return devsq(x) / len(x)


def midpoint(a: float, b: float) -> float:
    """``(a + b) / 2`` without overflowing on its way: MEDIAN(1E+308, 1E+308) is 1E+308."""
    if abs(a) <= 8.98846567431158e307 and abs(b) <= 8.98846567431158e307:
        return (a + b) / 2.0
    return a / 2.0 + b / 2.0


def median(v: List[float]) -> float:
    if not v:
        raise XlError(CellError.NUM)
    s = sorted(v)
    n = len(s)
    if n % 2:
        return s[n // 2]
    return midpoint(s[n // 2 - 1], s[n // 2])


def modes(v: Sequence[float]) -> List[float]:
    counts: dict = {}
    order: List[float] = []
    for x in v:
        if x in counts:
            counts[x] += 1
        else:
            counts[x] = 1
            order.append(x)
    best = max(counts.values(), default=0)
    if best < 2:
        return []
    return [x for x in order if counts[x] == best]


def mode_numbers(args: List[Arg]) -> List[float]:
    out: List[float] = []
    for a in args:
        v = a.value
        if isinstance(v, Array):
            for x in v.data:
                if type(x) is float:
                    out.append(x)
                elif isinstance(x, CellError):
                    raise XlError(x)
            continue
        if type(v) is float:
            out.append(v)
        elif isinstance(v, CellError):
            raise XlError(v)
        elif v is None:
            raise XlError(CellError.VALUE)
        # Text (numeric text too) and logical values given as arguments don't count (Excel 365).
    return out


def kth(v: List[float], k: float, largest: bool) -> float:
    k = math.ceil(k)
    if k < 1 or k > len(v):
        raise XlError(CellError.NUM)
    s = sorted(v)
    return s[len(s) - k] if largest else s[k - 1]


def _interp(s: List[float], pos: float) -> float:
    i = math.floor(pos)
    frac = pos - i
    if i < 0 or i >= len(s):
        raise XlError(CellError.NUM)
    lo = s[i]
    if frac > 0 and i + 1 < len(s):
        return lo + frac * (s[i + 1] - lo)
    return lo


def percentile_inc(v: List[float], k: float) -> float:
    if not v or not 0.0 <= k <= 1.0:
        raise XlError(CellError.NUM)
    s = sorted(v)
    return _interp(s, k * (len(s) - 1))


def percentile_exc(v: List[float], k: float) -> float:
    n = len(v)
    if not v or not 0.0 < k < 1.0:
        raise XlError(CellError.NUM)
    pos = k * (n + 1)
    if pos < 1 or pos > n:
        raise XlError(CellError.NUM)
    return _interp(sorted(v), pos - 1)


def quartile_inc(v: List[float], q: float) -> float:
    q = math.trunc(q)
    if not 0 <= q <= 4:
        raise XlError(CellError.NUM)
    return percentile_inc(v, q / 4.0)


def quartile_exc(v: List[float], q: float) -> float:
    q = math.trunc(q)
    if not 1 <= q <= 3:
        raise XlError(CellError.NUM)
    return percentile_exc(v, q / 4.0)


def _arg_nums(args: List[Arg], i: int) -> List[float]:
    return array_numbers(args[i].value)


# ---------------------------------------------------------------- SUBTOTAL / AGGREGATE


def _aggregate_impl(code: int, values: List[Value], k: Optional[Value]) -> Value:
    if 101 <= code <= 111:
        code -= 100
    xs: List[float] = []
    nonempty = 0
    for v in values:
        items = v.data if isinstance(v, Array) else [v]
        for x in items:
            if type(x) is float:
                xs.append(x)
            elif isinstance(x, CellError):
                raise XlError(x)
            if x is not None:
                nonempty += 1
    kk = 0.0
    if 14 <= code <= 19:
        if k is None:
            raise XlError(CellError.VALUE)
        kk = to_number(scalar(k))
    if code == 1:
        return mean(xs)
    if code == 2:
        return float(len(xs))
    if code == 3:
        return float(nonempty)
    if code == 4:
        return max(xs) if xs else 0.0
    if code == 5:
        return min(xs) if xs else 0.0
    if code == 6:
        p = 1.0
        for x in xs:
            p *= x
        return p if xs else 0.0
    if code == 7:
        return math.sqrt(var_s(xs))
    if code == 8:
        return math.sqrt(var_p(xs))
    if code == 9:
        return math.fsum(xs)
    if code == 10:
        return var_s(xs)
    if code == 11:
        return var_p(xs)
    if code == 12:
        return median(xs)
    if code == 13:
        m = modes(xs)
        if not m:
            raise XlError(CellError.NA)
        return m[0]
    if code == 14:
        return kth(xs, kk, True)
    if code == 15:
        return kth(xs, kk, False)
    if code == 16:
        return percentile_inc(xs, kk)
    if code == 17:
        return quartile_inc(xs, kk)
    if code == 18:
        return percentile_exc(xs, kk)
    if code == 19:
        return quartile_exc(xs, kk)
    raise XlError(CellError.VALUE)


def aggregate_values(code: int, values: List[Value]) -> Value:
    try:
        return _f(_aggregate_impl(code, values, None))
    except XlError as e:
        return e.error


def aggregate_values_k(code: int, values: List[Value], k: Value) -> Value:
    try:
        return _f(_aggregate_impl(code, values, k))
    except XlError as e:
        return e.error


# ---------------------------------------------------------------- counting


@fn("COUNT", 1, VAR, ST, A, "COUNT(value1, [value2], …)", "Counts how many numbers are in the list of arguments.")
def _count(args, ctx):
    n = 0
    for a in args:
        v = a.value
        if isinstance(v, Array):
            n += sum(1 for x in v.data if type(x) is float)
        elif type(v) is float:
            n += 1
        elif a.from_ref or v is None or isinstance(v, CellError):
            continue
        else:
            try:
                to_number(v)
                n += 1
            except XlError:
                pass
    return float(n)


@fn("COUNTA", 1, VAR, ST, A, "COUNTA(value1, [value2], …)", "Counts how many values are in the list of arguments.")
def _counta(args, ctx):
    n = 0
    for a in args:
        v = a.value
        if isinstance(v, Array):
            n += sum(1 for x in v.data if x is not None)
        elif v is None and a.from_ref:
            continue
        else:
            n += 1
    return float(n)


@fn("COUNTBLANK", 1, 1, ST, A, "COUNTBLANK(range)", "Counts the number of blank cells within a range.")
def _countblank(args, ctx):
    v = args[0].value
    items = v.data if isinstance(v, Array) else [v]
    return float(sum(1 for x in items if x is None or x == ""))


@fn("COUNTIF", 2, 2, ST, (False, True), "COUNTIF(range, criteria)", "Counts the number of cells within a range that meet the given criteria.")
def _countif(args, ctx):
    rng = as_array(args[0].value)
    crit = Criterion.parse_arg(args[1].value)
    return float(sum(1 for v in rng.data if crit.matches(v)))


@fn("COUNTIFS", 2, VAR, ST, (False, True), "COUNTIFS(criteria_range1, criteria1, …)", "Counts the number of cells within a range that meet multiple criteria.", repeat=2)
def _countifs(args, ctx):
    if len(args) % 2:
        raise XlError(CellError.VALUE)
    first = as_array(args[0].value)
    mask = [True] * len(first.data)
    for i in range(0, len(args), 2):
        rng = as_array(args[i].value)
        if rng.rows != first.rows or rng.cols != first.cols:
            raise XlError(CellError.VALUE)
        crit = Criterion.parse_arg(args[i + 1].value)
        for j, v in enumerate(rng.data):
            if mask[j] and not crit.matches(v):
                mask[j] = False
    return float(sum(mask))


# ---------------------------------------------------------------- averages


@fn("AVERAGE", 1, VAR, ST, A, "AVERAGE(number1, [number2], …)", "Returns the average of its arguments.")
def _average(args, ctx):
    return mean(numbers(args))


@fn("AVERAGEA", 1, VAR, ST, A, "AVERAGEA(value1, [value2], …)", "Returns the average of its arguments, including numbers, text, and logical values.")
def _averagea(args, ctx):
    return mean(numbers_a(args))


@fn("AVERAGEIF", 2, 3, ST, (False, True, False), "AVERAGEIF(range, criteria, [average_range])", "Returns the average of all the cells in a range that meet a given criteria.")
def _averageif(args, ctx):
    rng = as_array(args[0].value)
    crit = Criterion.parse_arg(args[1].value)
    src = as_array(args[2].value) if len(args) > 2 else rng
    vals = []
    for i, v in enumerate(rng.data):
        if crit.matches(v):
            r, c = divmod(i, rng.cols)
            x = src.get(r, c)
            if isinstance(x, CellError):
                raise XlError(x)
            if type(x) is float:
                vals.append(x)
    return mean(vals)


@fn("AVERAGEIFS", 3, VAR, ST, (False, False, True), "AVERAGEIFS(average_range, criteria_range1, criteria1, …)", "Returns the average of all cells that meet multiple criteria.", repeat=2)
def _averageifs(args, ctx):
    from .math_fns import ifs_values

    return mean(ifs_values(args, 0))


@fn("MINIFS", 3, VAR, ST, (False, False, True), "MINIFS(min_range, criteria_range1, criteria1, …)", "Returns the minimum value among cells specified by a given set of conditions.", repeat=2)
def _minifs(args, ctx):
    from .math_fns import ifs_values

    v = ifs_values(args, 0)
    return min(v) if v else 0.0


@fn("MAXIFS", 3, VAR, ST, (False, False, True), "MAXIFS(max_range, criteria_range1, criteria1, …)", "Returns the maximum value among cells specified by a given set of conditions.", repeat=2)
def _maxifs(args, ctx):
    from .math_fns import ifs_values

    v = ifs_values(args, 0)
    return max(v) if v else 0.0


@fn("MEDIAN", 1, VAR, ST, A, "MEDIAN(number1, [number2], …)", "Returns the median of the given numbers.")
def _median(args, ctx):
    return median(numbers(args))


@fn("MODE.SNGL", 1, VAR, ST, A, "MODE.SNGL(number1, [number2], …)", "Returns the most common value in a data set.")
def _mode_sngl(args, ctx):
    m = modes(mode_numbers(args))
    if not m:
        raise XlError(CellError.NA)
    return m[0]


@fn("MODE.MULT", 1, VAR, ST, A, "MODE.MULT(number1, [number2], …)", "Returns a vertical array of the most frequently occurring values in a range of data.")
def _mode_mult(args, ctx):
    m = modes(mode_numbers(args))
    if not m:
        raise XlError(CellError.NA)
    return array_val(len(m), 1, list(m))


@fn("GEOMEAN", 1, VAR, ST, A, "GEOMEAN(number1, [number2], …)", "Returns the geometric mean.")
def _geomean(args, ctx):
    xs = numbers(args)
    if not xs or any(x <= 0 for x in xs):
        raise XlError(CellError.NUM)
    return math.exp(math.fsum(math.log(x) for x in xs) / len(xs))


@fn("HARMEAN", 1, VAR, ST, A, "HARMEAN(number1, [number2], …)", "Returns the harmonic mean.")
def _harmean(args, ctx):
    xs = numbers(args)
    if not xs or any(x <= 0 for x in xs):
        raise XlError(CellError.NUM)
    return len(xs) / math.fsum(1.0 / x for x in xs)


@fn("TRIMMEAN", 2, 2, ST, (False, True), "TRIMMEAN(array, percent)", "Returns the mean of the interior of a data set.")
def _trimmean(args, ctx):
    xs = sorted(_arg_nums(args, 0))
    p = num(args, 1)
    if not xs or not 0 <= p < 1:
        raise XlError(CellError.NUM)
    k = int(math.floor(len(xs) * p / 2.0))
    inner = xs[k:len(xs) - k]
    return mean(inner)


@fn("AVEDEV", 1, VAR, ST, A, "AVEDEV(number1, [number2], …)", "Returns the average of the absolute deviations of data points from their mean.")
def _avedev(args, ctx):
    xs = numbers(args)
    if not xs:
        raise XlError(CellError.NUM)
    m = mean(xs)
    return math.fsum(abs(x - m) for x in xs) / len(xs)


@fn("DEVSQ", 1, VAR, ST, A, "DEVSQ(number1, [number2], …)", "Returns the sum of squares of deviations.")
def _devsq(args, ctx):
    xs = numbers(args)
    if not xs:
        raise XlError(CellError.NUM)
    return devsq(xs)


# ---------------------------------------------------------------- spread


@fn("STDEV.S", 1, VAR, ST, A, "STDEV.S(number1, [number2], …)", "Estimates standard deviation based on a sample.")
def _stdev_s(args, ctx):
    return math.sqrt(var_s(numbers(args)))


@fn("STDEV.P", 1, VAR, ST, A, "STDEV.P(number1, [number2], …)", "Calculates standard deviation based on the entire population.")
def _stdev_p(args, ctx):
    return math.sqrt(var_p(numbers(args)))


@fn("STDEVA", 1, VAR, ST, A, "STDEVA(value1, [value2], …)", "Estimates standard deviation based on a sample, including numbers, text, and logical values.")
def _stdeva(args, ctx):
    return math.sqrt(var_s(numbers_a(args)))


@fn("STDEVPA", 1, VAR, ST, A, "STDEVPA(value1, [value2], …)", "Calculates standard deviation based on the entire population, including numbers, text, and logical values.")
def _stdevpa(args, ctx):
    return math.sqrt(var_p(numbers_a(args)))


@fn("VAR.S", 1, VAR, ST, A, "VAR.S(number1, [number2], …)", "Estimates variance based on a sample.")
def _var_s(args, ctx):
    return var_s(numbers(args))


@fn("VAR.P", 1, VAR, ST, A, "VAR.P(number1, [number2], …)", "Calculates variance based on the entire population.")
def _var_p(args, ctx):
    return var_p(numbers(args))


@fn("VARA", 1, VAR, ST, A, "VARA(value1, [value2], …)", "Estimates variance based on a sample, including numbers, text, and logical values.")
def _vara(args, ctx):
    return var_s(numbers_a(args))


@fn("VARPA", 1, VAR, ST, A, "VARPA(value1, [value2], …)", "Calculates variance based on the entire population, including numbers, text, and logical values.")
def _varpa(args, ctx):
    return var_p(numbers_a(args))


# ---------------------------------------------------------------- min / max / order


@fn("MIN", 1, VAR, ST, A, "MIN(number1, [number2], …)", "Returns the minimum value in a list of arguments.")
def _min(args, ctx):
    xs = numbers(args)
    return min(xs) if xs else 0.0


@fn("MAX", 1, VAR, ST, A, "MAX(number1, [number2], …)", "Returns the maximum value in a list of arguments.")
def _max(args, ctx):
    xs = numbers(args)
    return max(xs) if xs else 0.0


@fn("MINA", 1, VAR, ST, A, "MINA(value1, [value2], …)", "Returns the smallest value in a list of arguments, including numbers, text, and logical values.")
def _mina(args, ctx):
    xs = numbers_a(args)
    return min(xs) if xs else 0.0


@fn("MAXA", 1, VAR, ST, A, "MAXA(value1, [value2], …)", "Returns the maximum value in a list of arguments, including numbers, text, and logical values.")
def _maxa(args, ctx):
    xs = numbers_a(args)
    return max(xs) if xs else 0.0


@fn("LARGE", 2, 2, ST, (False, True), "LARGE(array, k)", "Returns the k-th largest value in a data set.")
def _large(args, ctx):
    return kth(_arg_nums(args, 0), num(args, 1), True)


@fn("SMALL", 2, 2, ST, (False, True), "SMALL(array, k)", "Returns the k-th smallest value in a data set.")
def _small(args, ctx):
    return kth(_arg_nums(args, 0), num(args, 1), False)


def _rank(args, avg: bool) -> float:
    x = num(args, 0)
    xs = _arg_nums(args, 1)
    order = 0.0
    if len(args) > 2:
        order = opt_num(args, 2, 0.0)
    if not any(v == x for v in xs):
        raise XlError(CellError.NA)
    if order == 0:
        better = sum(1 for v in xs if v > x)
    else:
        better = sum(1 for v in xs if v < x)
    ties = sum(1 for v in xs if v == x)
    if avg:
        return better + (ties + 1) / 2.0
    return float(better + 1)


@fn("RANK.EQ", 2, 3, ST, (True, False, True), "RANK.EQ(number, ref, [order])", "Returns the rank of a number in a list of numbers.")
def _rank_eq(args, ctx):
    return _rank(args, False)


@fn("RANK.AVG", 2, 3, ST, (True, False, True), "RANK.AVG(number, ref, [order])", "Returns the rank of a number in a list of numbers (average rank for ties).")
def _rank_avg(args, ctx):
    return _rank(args, True)


@fn("PERCENTILE.INC", 2, 2, ST, (False, True), "PERCENTILE.INC(array, k)", "Returns the k-th percentile of values in a range.")
def _percentile_inc(args, ctx):
    return percentile_inc(_arg_nums(args, 0), num(args, 1))


@fn("PERCENTILE.EXC", 2, 2, ST, (False, True), "PERCENTILE.EXC(array, k)", "Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive.")
def _percentile_exc(args, ctx):
    return percentile_exc(_arg_nums(args, 0), num(args, 1))


@fn("QUARTILE.INC", 2, 2, ST, (False, True), "QUARTILE.INC(array, quart)", "Returns the quartile of a data set.")
def _quartile_inc(args, ctx):
    return quartile_inc(_arg_nums(args, 0), num(args, 1))


@fn("QUARTILE.EXC", 2, 2, ST, (False, True), "QUARTILE.EXC(array, quart)", "Returns the quartile of the data set, based on percentile values from 0..1, exclusive.")
def _quartile_exc(args, ctx):
    return quartile_exc(_arg_nums(args, 0), num(args, 1))


def _percentrank(args, exclusive: bool) -> float:
    xs = sorted(_arg_nums(args, 0))
    x = num(args, 1)
    sig = int(opt_num(args, 2, 3.0)) if has(args, 2) else 3
    if not xs or sig < 1:
        raise XlError(CellError.NUM)
    n = len(xs)
    if x < xs[0] or x > xs[-1]:
        raise XlError(CellError.NA)
    below = sum(1 for v in xs if v < x)
    if x in xs:
        r = (below + 1) / (n + 1) if exclusive else (below / (n - 1) if n > 1 else 1.0)
    else:
        # Interpolate between neighbours.
        hi_i = below
        lo_v, hi_v = xs[hi_i - 1], xs[hi_i]
        frac = (x - lo_v) / (hi_v - lo_v)
        if exclusive:
            r = (hi_i + frac) / (n + 1)
        else:
            r = (hi_i - 1 + frac) / (n - 1)
    try:
        f = 10.0 ** sig
    except OverflowError:
        # A significance past 308 is #DIV/0! in Excel 16 for Mac (10^309 can't be stored).
        raise XlError(CellError.DIV0) from None
    return math.floor(r * f + 1e-9) / f


@fn("PERCENTRANK.INC", 2, 3, ST, (False, True, True), "PERCENTRANK.INC(array, x, [significance])", "Returns the rank of a value in a data set as a percentage (0..1, inclusive).")
def _percentrank_inc(args, ctx):
    return _percentrank(args, False)


@fn("PERCENTRANK.EXC", 2, 3, ST, (False, True, True), "PERCENTRANK.EXC(array, x, [significance])", "Returns the rank of a value in a data set as a percentage (0..1, exclusive).")
def _percentrank_exc(args, ctx):
    return _percentrank(args, True)


@fn("FREQUENCY", 2, 2, ST, A, "FREQUENCY(data_array, bins_array)", "Returns a frequency distribution as a vertical array.")
def _frequency(args, ctx):
    data = _arg_nums(args, 0)
    bins = _arg_nums(args, 1)
    order = sorted(range(len(bins)), key=lambda i: bins[i])
    counts = [0] * (len(bins) + 1)
    sb = [bins[i] for i in order]
    import bisect

    for x in data:
        j = bisect.bisect_left(sb, x)
        if j < len(sb):
            counts[order[j]] += 1
        else:
            counts[-1] += 1
    return array_val(len(counts), 1, [float(c) for c in counts])


# ---------------------------------------------------------------- shape


@fn("KURT", 1, VAR, ST, A, "KURT(number1, [number2], …)", "Returns the kurtosis of a data set.")
def _kurt(args, ctx):
    xs = numbers(args)
    n = len(xs)
    if n < 4:
        raise XlError(CellError.DIV0)
    m = mean(xs)
    s = math.sqrt(var_s(xs))
    if s == 0:
        raise XlError(CellError.DIV0)
    k = math.fsum(((x - m) / s) ** 4 for x in xs)
    return n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * k - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3))


@fn("SKEW", 1, VAR, ST, A, "SKEW(number1, [number2], …)", "Returns the skewness of a distribution.")
def _skew(args, ctx):
    xs = numbers(args)
    n = len(xs)
    if n < 3:
        raise XlError(CellError.DIV0)
    m = mean(xs)
    s = math.sqrt(var_s(xs))
    if s == 0:
        raise XlError(CellError.DIV0)
    return n / ((n - 1) * (n - 2)) * math.fsum(((x - m) / s) ** 3 for x in xs)


@fn("SKEW.P", 1, VAR, ST, A, "SKEW.P(number1, [number2], …)", "Returns the skewness of a distribution based on a population.")
def _skew_p(args, ctx):
    xs = numbers(args)
    n = len(xs)
    if n < 1:
        raise XlError(CellError.DIV0)
    m = mean(xs)
    s = math.sqrt(var_p(xs))
    if s == 0:
        raise XlError(CellError.DIV0)
    return math.fsum(((x - m) / s) ** 3 for x in xs) / n


@fn("STANDARDIZE", 3, 3, ST, S, "STANDARDIZE(x, mean, standard_dev)", "Returns a normalized value.")
def _standardize(args, ctx):
    x, m, s = num(args, 0), num(args, 1), num(args, 2)
    if s <= 0:
        raise XlError(CellError.NUM)
    return (x - m) / s


# ---------------------------------------------------------------- regression


def _xy(args, yi: int = 0, xi: int = 1) -> Tuple[List[float], List[float]]:
    ys, xs = paired_numbers(args[yi].value, args[xi].value)
    return xs, ys


def _slope_intercept(xs: List[float], ys: List[float]) -> Tuple[float, float]:
    if len(xs) < 1:
        raise XlError(CellError.DIV0)
    mx, my = mean(xs), mean(ys)
    sxx = math.fsum((x - mx) ** 2 for x in xs)
    if sxx == 0:
        raise XlError(CellError.DIV0)
    sxy = math.fsum((x - mx) * (y - my) for x, y in zip(xs, ys))
    b = sxy / sxx
    return b, my - b * mx


@fn("SLOPE", 2, 2, ST, A, "SLOPE(known_y's, known_x's)", "Returns the slope of the linear regression line.")
def _slope(args, ctx):
    xs, ys = _xy(args)
    return _slope_intercept(xs, ys)[0]


@fn("INTERCEPT", 2, 2, ST, A, "INTERCEPT(known_y's, known_x's)", "Returns the intercept of the linear regression line.")
def _intercept(args, ctx):
    xs, ys = _xy(args)
    return _slope_intercept(xs, ys)[1]


def _correl(xs: List[float], ys: List[float]) -> float:
    if len(xs) < 2:
        raise XlError(CellError.DIV0)
    mx, my = mean(xs), mean(ys)
    sxy = math.fsum((x - mx) * (y - my) for x, y in zip(xs, ys))
    sxx = math.fsum((x - mx) ** 2 for x in xs)
    syy = math.fsum((y - my) ** 2 for y in ys)
    if sxx == 0 or syy == 0:
        raise XlError(CellError.DIV0)
    return sxy / math.sqrt(sxx * syy)


@fn("CORREL", 2, 2, ST, A, "CORREL(array1, array2)", "Returns the correlation coefficient between two data sets.")
def _correl_fn(args, ctx):
    xs, ys = paired_numbers(args[0].value, args[1].value)
    return _correl(xs, ys)


@fn("PEARSON", 2, 2, ST, A, "PEARSON(array1, array2)", "Returns the Pearson product moment correlation coefficient.")
def _pearson(args, ctx):
    return _correl_fn(args, ctx)


@fn("RSQ", 2, 2, ST, A, "RSQ(known_y's, known_x's)", "Returns the square of the Pearson product moment correlation coefficient.")
def _rsq(args, ctx):
    ys, xs = paired_numbers(args[0].value, args[1].value)
    return _correl(xs, ys) ** 2


@fn("STEYX", 2, 2, ST, A, "STEYX(known_y's, known_x's)", "Returns the standard error of the predicted y-value for each x in the regression.")
def _steyx(args, ctx):
    xs, ys = _xy(args)
    n = len(xs)
    if n < 3:
        raise XlError(CellError.DIV0)
    b, a = _slope_intercept(xs, ys)
    sse = math.fsum((y - (a + b * x)) ** 2 for x, y in zip(xs, ys))
    return math.sqrt(sse / (n - 2))


@fn("COVARIANCE.P", 2, 2, ST, A, "COVARIANCE.P(array1, array2)", "Returns covariance, the average of the products of paired deviations.")
def _cov_p(args, ctx):
    xs, ys = paired_numbers(args[0].value, args[1].value)
    if not xs:
        raise XlError(CellError.DIV0)
    mx, my = mean(xs), mean(ys)
    return math.fsum((x - mx) * (y - my) for x, y in zip(xs, ys)) / len(xs)


@fn("COVARIANCE.S", 2, 2, ST, A, "COVARIANCE.S(array1, array2)", "Returns the sample covariance.")
def _cov_s(args, ctx):
    xs, ys = paired_numbers(args[0].value, args[1].value)
    if len(xs) < 2:
        raise XlError(CellError.DIV0)
    mx, my = mean(xs), mean(ys)
    return math.fsum((x - mx) * (y - my) for x, y in zip(xs, ys)) / (len(xs) - 1)


@fn("FORECAST.LINEAR", 3, 3, ST, (True, False, False), "FORECAST.LINEAR(x, known_y's, known_x's)", "Returns a value along a linear trend.")
def _forecast_linear(args, ctx):
    x = num(args, 0)
    ys, xs = paired_numbers(args[1].value, args[2].value)
    b, a = _slope_intercept(xs, ys)
    return a + b * x


def _lstsq(X: List[List[float]], y: List[float]) -> List[float]:
    """Least squares by the normal equations (small problems)."""
    from .math_fns import invert

    k = len(X[0])
    xtx = [[math.fsum(row[i] * row[j] for row in X) for j in range(k)] for i in range(k)]
    xty = [math.fsum(row[i] * yy for row, yy in zip(X, y)) for i in range(k)]
    inv = invert(xtx)
    return [math.fsum(inv[i][j] * xty[j] for j in range(k)) for i in range(k)]


MAX_REGRESSORS = 255
MAX_REGRESSION_WORK = 1_000_000_000


def _design(args, yi: int, xi: Optional[int], const: bool):
    ya = as_array(args[yi].value)
    ys = []
    for v in ya.data:
        if isinstance(v, CellError):
            raise XlError(v)
        if type(v) is not float:
            raise XlError(CellError.VALUE)
        ys.append(v)
    n = len(ys)
    if xi is not None and has(args, xi):
        xa = as_array(args[xi].value)
        if len(xa.data) == n:
            cols = [[x for x in xa.data]]
        elif xa.rows == n and (ya.cols == 1):
            cols = [xa.col_list(c) for c in range(xa.cols)]
        elif xa.cols == n and (ya.rows == 1):
            cols = [xa.row_list(r) for r in range(xa.rows)]
        else:
            raise XlError(CellError.REF)
    else:
        cols = [[float(i + 1) for i in range(n)]]
    for c in cols:
        for v in c:
            if isinstance(v, CellError):
                raise XlError(v)
            if type(v) is not float:
                raise XlError(CellError.VALUE)
    # The normal equations are a k×k matrix solved in O(k³): one `SEQUENCE(1,16384)` of predictors
    # would take gigabytes and minutes, so more than MAX_REGRESSORS predictors or more than
    # MAX_REGRESSION_WORK multiply-adds (rows × k²) is #NUM! before anything is built.
    k = len(cols)
    if k > MAX_REGRESSORS or n * k * k > MAX_REGRESSION_WORK:
        raise XlError(CellError.NUM)
    X = []
    for i in range(n):
        row = [c[i] for c in cols]
        if const:
            row.append(1.0)
        X.append(row)
    return X, ys, cols


@fn("TREND", 1, 4, ST, A, "TREND(known_y's, [known_x's], [new_x's], [const])", "Returns values along a linear trend.")
def _trend(args, ctx):
    const = boolean(args, 3) if has(args, 3) else True
    X, ys, cols = _design(args, 0, 1, const)
    coef = _lstsq(X, ys)
    if has(args, 2):
        nx = as_array(args[2].value)
        k = len(cols)
        out = []
        if k == 1:
            for v in nx.data:
                x = to_number(v)
                out.append(coef[0] * x + (coef[1] if const else 0.0))
            return array_val(nx.rows, nx.cols, out)
        rows = nx.rows if nx.cols == k else nx.cols
        for r in range(rows):
            xs = nx.row_list(r) if nx.cols == k else nx.col_list(r)
            out.append(math.fsum(coef[i] * to_number(xs[i]) for i in range(k)) + (coef[k] if const else 0.0))
        return array_val(len(out), 1, out)
    out = [math.fsum(c * x for c, x in zip(coef, row)) for row in X]
    ya = as_array(args[0].value)
    return array_val(ya.rows, ya.cols, out)


@fn("GROWTH", 1, 4, ST, A, "GROWTH(known_y's, [known_x's], [new_x's], [const])", "Returns values along an exponential trend.")
def _growth(args, ctx):
    const = boolean(args, 3) if has(args, 3) else True
    X, ys, cols = _design(args, 0, 1, const)
    if any(y <= 0 for y in ys):
        raise XlError(CellError.NUM)
    coef = _lstsq(X, [math.log(y) for y in ys])
    k = len(cols)
    if has(args, 2):
        nx = as_array(args[2].value)
        out = []
        if k == 1:
            for v in nx.data:
                out.append(math.exp(coef[0] * to_number(v) + (coef[1] if const else 0.0)))
            return array_val(nx.rows, nx.cols, out)
        rows = nx.rows if nx.cols == k else nx.cols
        for r in range(rows):
            xs = nx.row_list(r) if nx.cols == k else nx.col_list(r)
            out.append(math.exp(math.fsum(coef[i] * to_number(xs[i]) for i in range(k)) + (coef[k] if const else 0.0)))
        return array_val(len(out), 1, out)
    out = [math.exp(math.fsum(c * x for c, x in zip(coef, row))) for row in X]
    ya = as_array(args[0].value)
    return array_val(ya.rows, ya.cols, out)


def _linest(args, log: bool):
    const = boolean(args, 2) if has(args, 2) else True
    stats = boolean(args, 3) if has(args, 3) else False
    X, ys, cols = _design(args, 0, 1, const)
    if log:
        if any(y <= 0 for y in ys):
            raise XlError(CellError.NUM)
        ys = [math.log(y) for y in ys]
    coef = _lstsq(X, ys)
    k = len(cols)
    # Excel lists the slopes in reverse column order, then the intercept.
    slopes = list(reversed(coef[:k]))
    b = coef[k] if const else 0.0
    if log:
        first = [math.exp(s) for s in slopes] + [math.exp(b) if const else 1.0]
    else:
        first = slopes + [b]
    if not stats:
        return array_val(1, k + 1, first)
    n = len(ys)
    fitted = [math.fsum(c * x for c, x in zip(coef, row)) for row in X]
    sse = math.fsum((y - f) ** 2 for y, f in zip(ys, fitted))
    my = mean(ys)
    sst = math.fsum((y - my) ** 2 for y in ys) if const else math.fsum(y * y for y in ys)
    ssr = sst - sse
    df = n - k - (1 if const else 0)
    r2 = ssr / sst if sst else 1.0
    sey = math.sqrt(sse / df) if df > 0 else 0.0
    from .math_fns import invert

    kk = len(X[0])
    xtx = [[math.fsum(row[i] * row[j] for row in X) for j in range(kk)] for i in range(kk)]
    try:
        inv = invert(xtx)
        ses = [math.sqrt(max(inv[i][i], 0.0)) * sey for i in range(kk)]
    except XlError:
        ses = [CellError.NUM] * kk
    se_slopes = list(reversed(ses[:k]))
    se_b = ses[k] if const else CellError.NA
    fstat = (ssr / k) / (sse / df) if df > 0 and sse > 0 else CellError.NUM
    na = CellError.NA
    rows = [
        first,
        se_slopes + [se_b],
        [r2, sey] + [na] * (k - 1),
        [fstat, float(df)] + [na] * (k - 1),
        [ssr, sse] + [na] * (k - 1),
    ]
    return array_val(5, k + 1, [v for r in rows for v in r])


@fn("LINEST", 1, 4, ST, A, "LINEST(known_y's, [known_x's], [const], [stats])", "Returns the parameters of a linear trend.")
def _linest_fn(args, ctx):
    return _linest(args, False)


@fn("LOGEST", 1, 4, ST, A, "LOGEST(known_y's, [known_x's], [const], [stats])", "Returns the parameters of an exponential trend.")
def _logest_fn(args, ctx):
    return _linest(args, True)


# ---------------------------------------------------------------- distributions


@fn("NORM.DIST", 4, 4, ST, S, "NORM.DIST(x, mean, standard_dev, cumulative)", "Returns the normal cumulative distribution.")
def _norm_dist(args, ctx):
    x, m, s = num(args, 0), num(args, 1), num(args, 2)
    if s <= 0:
        raise XlError(CellError.NUM)
    z = (x - m) / s
    return norm_cdf(z) if boolean(args, 3) else norm_pdf(z) / s


@fn("NORM.INV", 3, 3, ST, S, "NORM.INV(probability, mean, standard_dev)", "Returns the inverse of the normal cumulative distribution.")
def _norm_inv(args, ctx):
    p, m, s = num(args, 0), num(args, 1), num(args, 2)
    if s <= 0:
        raise XlError(CellError.NUM)
    return m + s * norm_inv(p)


@fn("NORM.S.DIST", 2, 2, ST, S, "NORM.S.DIST(z, cumulative)", "Returns the standard normal cumulative distribution.")
def _norm_s_dist(args, ctx):
    z = num(args, 0)
    return norm_cdf(z) if boolean(args, 1) else norm_pdf(z)


@fn("NORM.S.INV", 1, 1, ST, S, "NORM.S.INV(probability)", "Returns the inverse of the standard normal cumulative distribution.")
def _norm_s_inv(args, ctx):
    return norm_inv(num(args, 0))


@fn("NORMSDIST", 1, 1, COMPAT, S, "NORMSDIST(z)", "Returns the standard normal cumulative distribution.")
def _normsdist(args, ctx):
    return norm_cdf(num(args, 0))


@fn("PHI", 1, 1, ST, S, "PHI(x)", "Returns the value of the density function for a standard normal distribution.")
def _phi(args, ctx):
    return norm_pdf(num(args, 0))


@fn("GAUSS", 1, 1, ST, S, "GAUSS(x)", "Returns 0.5 less than the standard normal cumulative distribution.")
def _gauss(args, ctx):
    return norm_cdf(num(args, 0)) - 0.5


@fn("LOGNORM.DIST", 4, 4, ST, S, "LOGNORM.DIST(x, mean, standard_dev, cumulative)", "Returns the cumulative lognormal distribution.")
def _lognorm_dist(args, ctx):
    x, m, s = num(args, 0), num(args, 1), num(args, 2)
    if x <= 0 or s <= 0:
        raise XlError(CellError.NUM)
    z = (math.log(x) - m) / s
    return norm_cdf(z) if boolean(args, 3) else norm_pdf(z) / (x * s)


@fn("LOGNORM.INV", 3, 3, ST, S, "LOGNORM.INV(probability, mean, standard_dev)", "Returns the inverse of the lognormal cumulative distribution.")
def _lognorm_inv(args, ctx):
    p, m, s = num(args, 0), num(args, 1), num(args, 2)
    if s <= 0:
        raise XlError(CellError.NUM)
    return math.exp(m + s * norm_inv(p))


@fn("LOGNORMDIST", 3, 3, COMPAT, S, "LOGNORMDIST(x, mean, standard_dev)", "Returns the cumulative lognormal distribution.")
def _lognormdist(args, ctx):
    x, m, s = num(args, 0), num(args, 1), num(args, 2)
    if x <= 0 or s <= 0:
        raise XlError(CellError.NUM)
    return norm_cdf((math.log(x) - m) / s)


def _df(v: float, lo: float = 1.0) -> float:
    v = math.trunc(v)
    if v < lo or v >= 1e10:
        raise XlError(CellError.NUM)
    return v


@fn("T.DIST", 3, 3, ST, S, "T.DIST(x, deg_freedom, cumulative)", "Returns the Student's left-tailed t-distribution.")
def _t_dist(args, ctx):
    x, v = num(args, 0), _df(num(args, 1))
    return t_cdf(x, v) if boolean(args, 2) else t_pdf(x, v)


@fn("T.DIST.2T", 2, 2, ST, S, "T.DIST.2T(x, deg_freedom)", "Returns the two-tailed Student's t-distribution.")
def _t_dist_2t(args, ctx):
    x, v = num(args, 0), _df(num(args, 1))
    if x < 0:
        raise XlError(CellError.NUM)
    return 2 * (1 - t_cdf(x, v))


@fn("T.DIST.RT", 2, 2, ST, S, "T.DIST.RT(x, deg_freedom)", "Returns the right-tailed Student's t-distribution.")
def _t_dist_rt(args, ctx):
    x, v = num(args, 0), _df(num(args, 1))
    return 1 - t_cdf(x, v)


@fn("TDIST", 3, 3, COMPAT, S, "TDIST(x, deg_freedom, tails)", "Returns the Student's t-distribution.")
def _tdist(args, ctx):
    x, v, tails = num(args, 0), _df(num(args, 1)), math.trunc(num(args, 2))
    if x < 0 or tails not in (1, 2):
        raise XlError(CellError.NUM)
    return (1 - t_cdf(x, v)) * tails


def _t_inv(p: float, v: float) -> float:
    if not 0 < p < 1:
        raise XlError(CellError.NUM)
    return invert(lambda t: t_cdf(t, v), p, -10.0, 10.0)


@fn("T.INV", 2, 2, ST, S, "T.INV(probability, deg_freedom)", "Returns the left-tailed inverse of the Student's t-distribution.")
def _t_inv_fn(args, ctx):
    return _t_inv(num(args, 0), _df(num(args, 1)))


@fn("T.INV.2T", 2, 2, ST, S, "T.INV.2T(probability, deg_freedom)", "Returns the two-tailed inverse of the Student's t-distribution.")
def _t_inv_2t(args, ctx):
    p, v = num(args, 0), _df(num(args, 1))
    if not 0 < p <= 1:
        raise XlError(CellError.NUM)
    return abs(_t_inv(1 - p / 2, v))


@fn("T.TEST", 4, 4, ST, (False, False, True, True), "T.TEST(array1, array2, tails, type)", "Returns the probability associated with a Student's t-test.")
def _t_test(args, ctx):
    tails, typ = math.trunc(num(args, 2)), math.trunc(num(args, 3))
    if tails not in (1, 2) or typ not in (1, 2, 3):
        raise XlError(CellError.NUM)
    if typ == 1:
        xs, ys = paired_numbers(args[0].value, args[1].value)
        d = [a - b for a, b in zip(xs, ys)]
        n = len(d)
        if n < 2:
            raise XlError(CellError.DIV0)
        sd = math.sqrt(var_s(d))
        if sd == 0:
            raise XlError(CellError.DIV0)
        t = abs(mean(d)) / (sd / math.sqrt(n))
        v = n - 1.0
    else:
        a, b = _arg_nums(args, 0), _arg_nums(args, 1)
        n1, n2 = len(a), len(b)
        if n1 < 2 or n2 < 2:
            raise XlError(CellError.DIV0)
        v1, v2 = var_s(a), var_s(b)
        if typ == 2:
            v = n1 + n2 - 2.0
            sp = ((n1 - 1) * v1 + (n2 - 1) * v2) / v
            se = math.sqrt(sp * (1 / n1 + 1 / n2))
        else:
            se = math.sqrt(v1 / n1 + v2 / n2)
            v = (v1 / n1 + v2 / n2) ** 2 / ((v1 / n1) ** 2 / (n1 - 1) + (v2 / n2) ** 2 / (n2 - 1))
        if se == 0:
            raise XlError(CellError.DIV0)
        t = abs(mean(a) - mean(b)) / se
    return (1 - t_cdf(t, v)) * tails


@fn("CHISQ.DIST", 3, 3, ST, S, "CHISQ.DIST(x, deg_freedom, cumulative)", "Returns the cumulative beta probability density function.")
def _chisq_dist(args, ctx):
    x, k = num(args, 0), _df(num(args, 1))
    if x < 0:
        raise XlError(CellError.NUM)
    if boolean(args, 2):
        return chisq_cdf(x, k)
    if x == 0:
        return 0.5 if k == 2 else (0.0 if k > 2 else math.inf)
    return math.exp((k / 2 - 1) * math.log(x) - x / 2 - (k / 2) * math.log(2) - math.lgamma(k / 2))


@fn("CHISQ.DIST.RT", 2, 2, ST, S, "CHISQ.DIST.RT(x, deg_freedom)", "Returns the one-tailed probability of the chi-squared distribution.")
def _chisq_dist_rt(args, ctx):
    x, k = num(args, 0), _df(num(args, 1))
    if x < 0:
        raise XlError(CellError.NUM)
    return 1 - chisq_cdf(x, k)


@fn("CHISQ.INV", 2, 2, ST, S, "CHISQ.INV(probability, deg_freedom)", "Returns the cumulative beta probability density function.")
def _chisq_inv(args, ctx):
    p, k = num(args, 0), _df(num(args, 1))
    if not 0 <= p < 1:
        raise XlError(CellError.NUM)
    return invert(lambda x: chisq_cdf(x, k), p, 0.0, max(k * 2, 10.0), 0.0)


@fn("CHISQ.INV.RT", 2, 2, ST, S, "CHISQ.INV.RT(probability, deg_freedom)", "Returns the inverse of the one-tailed probability of the chi-squared distribution.")
def _chisq_inv_rt(args, ctx):
    p, k = num(args, 0), _df(num(args, 1))
    if not 0 < p <= 1:
        raise XlError(CellError.NUM)
    return invert(lambda x: chisq_cdf(x, k), 1 - p, 0.0, max(k * 2, 10.0), 0.0)


@fn("CHISQ.TEST", 2, 2, ST, A, "CHISQ.TEST(actual_range, expected_range)", "Returns the test for independence.")
def _chisq_test(args, ctx):
    a, e = as_array(args[0].value), as_array(args[1].value)
    if a.rows != e.rows or a.cols != e.cols:
        raise XlError(CellError.NA)
    stat = 0.0
    for x, y in zip(a.data, e.data):
        if type(x) is float and type(y) is float:
            if y == 0:
                raise XlError(CellError.DIV0)
            stat += (x - y) ** 2 / y
    if a.rows > 1 and a.cols > 1:
        df = (a.rows - 1) * (a.cols - 1)
    else:
        df = max(a.rows, a.cols) - 1
    if df < 1:
        raise XlError(CellError.NA)
    return 1 - chisq_cdf(stat, df)


@fn("F.DIST", 4, 4, ST, S, "F.DIST(x, deg_freedom1, deg_freedom2, cumulative)", "Returns the F probability distribution.")
def _f_dist(args, ctx):
    x, d1, d2 = num(args, 0), _df(num(args, 1)), _df(num(args, 2))
    if x < 0:
        raise XlError(CellError.NUM)
    if boolean(args, 3):
        return f_cdf(x, d1, d2)
    if x == 0:
        return 0.0 if d1 > 2 else (1.0 if d1 == 2 else math.inf)
    return math.exp(0.5 * (d1 * math.log(d1 * x) + d2 * math.log(d2) - (d1 + d2) * math.log(d1 * x + d2)) - math.log(x) - ln_beta(d1 / 2, d2 / 2))


@fn("F.DIST.RT", 3, 3, ST, S, "F.DIST.RT(x, deg_freedom1, deg_freedom2)", "Returns the (right-tailed) F probability distribution.")
def _f_dist_rt(args, ctx):
    x, d1, d2 = num(args, 0), _df(num(args, 1)), _df(num(args, 2))
    if x < 0:
        raise XlError(CellError.NUM)
    return 1 - f_cdf(x, d1, d2)


@fn("F.INV", 3, 3, ST, S, "F.INV(probability, deg_freedom1, deg_freedom2)", "Returns the inverse of the F probability distribution.")
def _f_inv(args, ctx):
    p, d1, d2 = num(args, 0), _df(num(args, 1)), _df(num(args, 2))
    if not 0 <= p < 1:
        raise XlError(CellError.NUM)
    return invert(lambda x: f_cdf(x, d1, d2), p, 0.0, 10.0, 0.0)


@fn("F.INV.RT", 3, 3, ST, S, "F.INV.RT(probability, deg_freedom1, deg_freedom2)", "Returns the inverse of the (right-tailed) F probability distribution.")
def _f_inv_rt(args, ctx):
    p, d1, d2 = num(args, 0), _df(num(args, 1)), _df(num(args, 2))
    if not 0 < p <= 1:
        raise XlError(CellError.NUM)
    return invert(lambda x: f_cdf(x, d1, d2), 1 - p, 0.0, 10.0, 0.0)


@fn("F.TEST", 2, 2, ST, A, "F.TEST(array1, array2)", "Returns the result of an F-test.")
def _f_test(args, ctx):
    a, b = _arg_nums(args, 0), _arg_nums(args, 1)
    if len(a) < 2 or len(b) < 2:
        raise XlError(CellError.DIV0)
    va, vb = var_s(a), var_s(b)
    if va == 0 or vb == 0:
        raise XlError(CellError.DIV0)
    f = va / vb
    p = f_cdf(f, len(a) - 1, len(b) - 1)
    return 2 * min(p, 1 - p)


def _binom_pmf(k: int, n: int, p: float) -> float:
    if p == 0:
        return 1.0 if k == 0 else 0.0
    if p == 1:
        return 1.0 if k == n else 0.0
    return math.exp(math.lgamma(n + 1) - math.lgamma(k + 1) - math.lgamma(n - k + 1) + k * math.log(p) + (n - k) * math.log1p(-p))


@fn("BINOM.DIST", 4, 4, ST, S, "BINOM.DIST(number_s, trials, probability_s, cumulative)", "Returns the individual term binomial distribution probability.")
def _binom_dist(args, ctx):
    k, n, p = math.trunc(num(args, 0)), math.trunc(num(args, 1)), num(args, 2)
    if k < 0 or k > n or not 0 <= p <= 1 or n > 1e15:
        raise XlError(CellError.NUM)
    if boolean(args, 3):
        return _binom_cdf(k, n, p)
    return _binom_pmf(k, n, p)


def _binom_cdf(k: int, n: int, p: float) -> float:
    if k < 0:
        return 0.0
    if k >= n or p == 0:
        return 1.0
    if p == 1:
        return 0.0
    if n <= 1000:
        return min(math.fsum(_binom_pmf(i, n, p) for i in range(0, k + 1)), 1.0)
    return reg_inc_beta(1.0 - p, n - k, k + 1)


@fn("BINOM.DIST.RANGE", 3, 4, ST, S, "BINOM.DIST.RANGE(trials, probability_s, number_s, [number_s2])", "Returns the probability of a trial result using a binomial distribution.")
def _binom_range(args, ctx):
    n, p, s = math.trunc(num(args, 0)), num(args, 1), math.trunc(num(args, 2))
    s2 = math.trunc(num(args, 3)) if has(args, 3) else s
    if n < 0 or not 0 <= p <= 1 or s < 0 or s > n or s2 < s or s2 > n or n > 1e15:
        raise XlError(CellError.NUM)
    if s2 - s <= 1000:
        return math.fsum(_binom_pmf(i, n, p) for i in range(s, s2 + 1))
    return _binom_cdf(s2, n, p) - _binom_cdf(s - 1, n, p)


@fn("BINOM.INV", 3, 3, ST, S, "BINOM.INV(trials, probability_s, alpha)", "Returns the smallest value for which the cumulative binomial distribution is less than or equal to a criterion value.")
def _binom_inv(args, ctx):
    n, p, alpha = math.trunc(num(args, 0)), num(args, 1), num(args, 2)
    if n < 0 or not 0 <= p <= 1 or not 0 <= alpha <= 1 or n > 1e15:
        raise XlError(CellError.NUM)
    lo, hi = 0, n
    while lo < hi:
        mid = (lo + hi) // 2
        if _binom_cdf(mid, n, p) >= alpha - 1e-15:
            hi = mid
        else:
            lo = mid + 1
    return float(lo)


@fn("NEGBINOM.DIST", 4, 4, ST, S, "NEGBINOM.DIST(number_f, number_s, probability_s, cumulative)", "Returns the negative binomial distribution.")
def _negbinom_dist(args, ctx):
    f, s, p = math.trunc(num(args, 0)), math.trunc(num(args, 1)), num(args, 2)
    if f < 0 or s < 1 or not 0 < p < 1:
        raise XlError(CellError.NUM)

    def pmf(k: int) -> float:
        return math.exp(math.lgamma(k + s) - math.lgamma(k + 1) - math.lgamma(s) + s * math.log(p) + k * math.log1p(-p))

    if boolean(args, 3):
        return reg_inc_beta(p, s, f + 1)
    return pmf(f)


@fn("HYPGEOM.DIST", 5, 5, ST, S, "HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative)", "Returns the hypergeometric distribution.")
def _hypgeom_dist(args, ctx):
    k, n, K, N = (math.trunc(num(args, i)) for i in range(4))
    if k < 0 or n < 0 or K < 0 or N < 0 or k > n or k > K or n > N or K > N or n - k > N - K:
        raise XlError(CellError.NUM)

    def lc(a, b):
        return math.lgamma(a + 1) - math.lgamma(b + 1) - math.lgamma(a - b + 1)

    def pmf(i):
        if i < max(0, n - (N - K)) or i > min(n, K):
            return 0.0
        return math.exp(lc(K, i) + lc(N - K, n - i) - lc(N, n))

    if boolean(args, 4):
        lo = max(0, n - (N - K))
        if k - lo > 1_000_000:
            raise XlError(CellError.NUM)
        return min(math.fsum(pmf(i) for i in range(lo, k + 1)), 1.0)
    return pmf(k)


@fn("POISSON.DIST", 3, 3, ST, S, "POISSON.DIST(x, mean, cumulative)", "Returns the Poisson distribution.")
def _poisson_dist(args, ctx):
    x, m = math.trunc(num(args, 0)), num(args, 1)
    if x < 0 or m < 0:
        raise XlError(CellError.NUM)

    def pmf(k):
        if m == 0:
            return 1.0 if k == 0 else 0.0
        return math.exp(k * math.log(m) - m - math.lgamma(k + 1))

    if boolean(args, 2):
        if m == 0:
            return 1.0
        return 1.0 - reg_gamma_p(x + 1, m) if x < 1e6 else 1.0
    return pmf(x)


@fn("EXPON.DIST", 3, 3, ST, S, "EXPON.DIST(x, lambda, cumulative)", "Returns the exponential distribution.")
def _expon_dist(args, ctx):
    x, l = num(args, 0), num(args, 1)
    if x < 0 or l <= 0:
        raise XlError(CellError.NUM)
    return -math.expm1(-l * x) if boolean(args, 2) else l * math.exp(-l * x)


@fn("GAMMA", 1, 1, ST, S, "GAMMA(x)", "Returns the Gamma function value.")
def _gamma(args, ctx):
    x = num(args, 0)
    if x <= 0 and x == math.floor(x):
        raise XlError(CellError.NUM)
    return _f(math.gamma(x))


@fn("GAMMALN", 1, 1, ST, S, "GAMMALN(x)", "Returns the natural logarithm of the gamma function, Γ(x).")
def _gammaln(args, ctx):
    x = num(args, 0)
    if x <= 0:
        raise XlError(CellError.NUM)
    return math.lgamma(x)


@fn("GAMMA.DIST", 4, 4, ST, S, "GAMMA.DIST(x, alpha, beta, cumulative)", "Returns the gamma distribution.")
def _gamma_dist(args, ctx):
    x, a, b = num(args, 0), num(args, 1), num(args, 2)
    if x < 0 or a <= 0 or b <= 0:
        raise XlError(CellError.NUM)
    if boolean(args, 3):
        return gamma_cdf(x / b, a)
    if x == 0:
        return 0.0 if a > 1 else (1 / b if a == 1 else math.inf)
    return math.exp((a - 1) * math.log(x) - x / b - math.lgamma(a) - a * math.log(b))


@fn("GAMMA.INV", 3, 3, ST, S, "GAMMA.INV(probability, alpha, beta)", "Returns the inverse of the gamma cumulative distribution.")
def _gamma_inv(args, ctx):
    p, a, b = num(args, 0), num(args, 1), num(args, 2)
    if not 0 <= p < 1 or a <= 0 or b <= 0:
        raise XlError(CellError.NUM)
    return invert(lambda x: gamma_cdf(x / b, a), p, 0.0, max(a * b * 2, 1.0), 0.0)


@fn("BETA.DIST", 4, 6, ST, S, "BETA.DIST(x, alpha, beta, cumulative, [A], [B])", "Returns the beta cumulative distribution function.")
def _beta_dist(args, ctx):
    x, a, b = num(args, 0), num(args, 1), num(args, 2)
    lo = opt_num(args, 4, 0.0) if has(args, 4) else 0.0
    hi = opt_num(args, 5, 1.0) if has(args, 5) else 1.0
    if a <= 0 or b <= 0 or not lo <= x <= hi or lo == hi:
        raise XlError(CellError.NUM)
    z = (x - lo) / (hi - lo)
    if boolean(args, 3):
        return reg_inc_beta(z, a, b)
    if z in (0.0, 1.0):
        return 0.0
    return math.exp((a - 1) * math.log(z) + (b - 1) * math.log1p(-z) - ln_beta(a, b)) / (hi - lo)


@fn("BETA.INV", 3, 5, ST, S, "BETA.INV(probability, alpha, beta, [A], [B])", "Returns the inverse of the cumulative distribution function for a specified beta distribution.")
def _beta_inv(args, ctx):
    p, a, b = num(args, 0), num(args, 1), num(args, 2)
    lo = opt_num(args, 3, 0.0) if has(args, 3) else 0.0
    hi = opt_num(args, 4, 1.0) if has(args, 4) else 1.0
    if not 0 <= p <= 1 or a <= 0 or b <= 0 or lo >= hi:
        raise XlError(CellError.NUM)
    z = invert(lambda t: reg_inc_beta(t, a, b), p, 0.0, 1.0, 0.0, 1.0)
    return lo + z * (hi - lo)


@fn("WEIBULL.DIST", 4, 4, ST, S, "WEIBULL.DIST(x, alpha, beta, cumulative)", "Returns the Weibull distribution.")
def _weibull_dist(args, ctx):
    x, a, b = num(args, 0), num(args, 1), num(args, 2)
    if x < 0 or a <= 0 or b <= 0:
        raise XlError(CellError.NUM)
    if boolean(args, 3):
        return -math.expm1(-((x / b) ** a))
    return a / b ** a * x ** (a - 1) * math.exp(-((x / b) ** a))


@fn("Z.TEST", 2, 3, ST, (False, True, True), "Z.TEST(array, x, [sigma])", "Returns the one-tailed P-value of a z-test.")
def _z_test(args, ctx):
    xs = _arg_nums(args, 0)
    x = num(args, 1)
    n = len(xs)
    if n < 1:
        raise XlError(CellError.NA)
    s = num(args, 2) if has(args, 2) else math.sqrt(var_s(xs))
    if s == 0:
        raise XlError(CellError.DIV0)
    return 1 - norm_cdf((mean(xs) - x) / (s / math.sqrt(n)))


@fn("CONFIDENCE.NORM", 3, 3, ST, S, "CONFIDENCE.NORM(alpha, standard_dev, size)", "Returns the confidence interval for a population mean.")
def _confidence_norm(args, ctx):
    a, s, n = num(args, 0), num(args, 1), math.trunc(num(args, 2))
    if not 0 < a < 1 or s <= 0 or n < 1:
        raise XlError(CellError.NUM)
    return norm_inv(1 - a / 2) * s / math.sqrt(n)


@fn("CONFIDENCE.T", 3, 3, ST, S, "CONFIDENCE.T(alpha, standard_dev, size)", "Returns the confidence interval for a population mean, using a Student's t distribution.")
def _confidence_t(args, ctx):
    a, s, n = num(args, 0), num(args, 1), math.trunc(num(args, 2))
    if not 0 < a < 1 or s <= 0 or n < 1:
        raise XlError(CellError.NUM)
    if n == 1:
        raise XlError(CellError.DIV0)
    return abs(_t_inv(a / 2, n - 1)) * s / math.sqrt(n)


@fn("FISHER", 1, 1, ST, S, "FISHER(x)", "Returns the Fisher transformation.")
def _fisher(args, ctx):
    x = num(args, 0)
    if not -1 < x < 1:
        raise XlError(CellError.NUM)
    return 0.5 * math.log((1 + x) / (1 - x))


@fn("FISHERINV", 1, 1, ST, S, "FISHERINV(y)", "Returns the inverse of the Fisher transformation.")
def _fisherinv(args, ctx):
    return math.tanh(num(args, 0))


@fn("PROB", 3, 4, ST, (False, False, True, True), "PROB(x_range, prob_range, lower_limit, [upper_limit])", "Returns the probability that values in a range are between two limits.")
def _prob(args, ctx):
    xs, ps = paired_numbers(args[0].value, args[1].value)
    # Excel checks only that the probabilities add up to 1 (to 15 digits): 0 and even -0.1 with
    # 1.1 pass, 1.00000000000001 doesn't (Excel 16 for Mac).
    if abs(math.fsum(ps) - 1) > 5e-15:
        raise XlError(CellError.NUM)
    lo = num(args, 2)
    hi = num(args, 3) if has(args, 3) else lo
    return math.fsum(p for x, p in zip(xs, ps) if lo <= x <= hi)


# ---------------------------------------------------------------- legacy (Compatibility) names


def _compat(name: str, target: str, desc: Optional[str] = None) -> None:
    from .registry import alias

    alias(name, target, desc, COMPAT)


for _n, _t in [
    ("STDEV", "STDEV.S"), ("STDEVP", "STDEV.P"), ("VAR", "VAR.S"), ("VARP", "VAR.P"), ("MODE", "MODE.SNGL"),
    ("RANK", "RANK.EQ"), ("PERCENTILE", "PERCENTILE.INC"), ("QUARTILE", "QUARTILE.INC"),
    ("PERCENTRANK", "PERCENTRANK.INC"), ("COVAR", "COVARIANCE.P"), ("FORECAST", "FORECAST.LINEAR"),
    ("NORMDIST", "NORM.DIST"), ("NORMINV", "NORM.INV"), ("NORMSINV", "NORM.S.INV"), ("LOGINV", "LOGNORM.INV"),
    ("TINV", "T.INV.2T"), ("TTEST", "T.TEST"), ("CHIDIST", "CHISQ.DIST.RT"), ("CHIINV", "CHISQ.INV.RT"),
    ("CHITEST", "CHISQ.TEST"), ("FDIST", "F.DIST.RT"), ("FINV", "F.INV.RT"), ("FTEST", "F.TEST"),
    ("BINOMDIST", "BINOM.DIST"), ("CRITBINOM", "BINOM.INV"), ("NEGBINOMDIST", "NEGBINOM.DIST"),
    ("POISSON", "POISSON.DIST"), ("EXPONDIST", "EXPON.DIST"), ("GAMMADIST", "GAMMA.DIST"), ("GAMMAINV", "GAMMA.INV"),
    ("WEIBULL", "WEIBULL.DIST"), ("ZTEST", "Z.TEST"), ("CONFIDENCE", "CONFIDENCE.NORM"),
    ("GAMMALN.PRECISE", "GAMMALN"),
]:
    _compat(_n, _t)


@fn("HYPGEOMDIST", 4, 4, COMPAT, S, "HYPGEOMDIST(sample_s, number_sample, population_s, number_pop)", "Returns the hypergeometric distribution.")
def _hypgeomdist(args, ctx):
    return _hypgeom_dist(list(args) + [Arg(False)], ctx)


@fn("BETADIST", 3, 5, COMPAT, S, "BETADIST(x, alpha, beta, [A], [B])", "Returns the beta cumulative distribution function.")
def _betadist(args, ctx):
    a = list(args)
    rest = a[3:]
    return _beta_dist(a[:3] + [Arg(True)] + rest, ctx)


@fn("BETAINV", 3, 5, COMPAT, S, "BETAINV(probability, alpha, beta, [A], [B])", "Returns the inverse of the cumulative distribution function for a specified beta distribution.")
def _betainv(args, ctx):
    return _beta_inv(args, ctx)
