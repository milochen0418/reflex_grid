"""Financial functions: time value of money, cash flows, depreciation, securities."""

from __future__ import annotations

import math
from typing import Callable, List, Optional

from ..core.date import DateSystem, datetime_from_serial, serial_from_ymd
from ..core.value import Array, CellError, Value, XlError, scalar, to_number
from .registry import A, Arg, Ctx, S, VAR, fn
from .util import array_numbers, as_array, has, num, opt_num

F = "Financial"


def _f(n: float) -> float:
    if not math.isfinite(n):
        raise XlError(CellError.NUM)
    return n


def _type(args, i: int) -> int:
    return 1 if has(args, i) and num(args, i) != 0 else 0


def fv(rate: float, nper: float, pmt: float, pv: float, typ: int) -> float:
    if rate == 0:
        return -(pv + pmt * nper)
    f = (1 + rate) ** nper
    return -(pv * f + pmt * (1 + rate * typ) * (f - 1) / rate)


def pmt(rate: float, nper: float, pv: float, fv_: float, typ: int) -> float:
    if nper == 0:
        raise XlError(CellError.NUM)
    if rate == 0:
        return -(pv + fv_) / nper
    f = (1 + rate) ** nper
    return -(rate * (pv * f + fv_)) / ((1 + rate * typ) * (f - 1))


def pv(rate: float, nper: float, pmt_: float, fv_: float, typ: int) -> float:
    if rate == 0:
        return -(fv_ + pmt_ * nper)
    f = (1 + rate) ** nper
    return -(fv_ + pmt_ * (1 + rate * typ) * (f - 1) / rate) / f


def ipmt(rate: float, per: float, nper: float, pv_: float, fv_: float, typ: int) -> float:
    p = pmt(rate, nper, pv_, fv_, typ)
    if per == 1 and typ == 1:
        return 0.0
    if typ == 1:
        balance = fv(rate, per - 2, p, pv_, 1) - p
    else:
        balance = fv(rate, per - 1, p, pv_, 0)
    return balance * rate


@fn("PMT", 3, 5, F, S, "PMT(rate, nper, pv, [fv], [type])", "Returns the periodic payment for an annuity.")
def _pmt(args, ctx):
    return _f(pmt(num(args, 0), num(args, 1), num(args, 2), opt_num(args, 3, 0.0), _type(args, 4)))


@fn("FV", 3, 5, F, S, "FV(rate, nper, pmt, [pv], [type])", "Returns the future value of an investment.")
def _fv(args, ctx):
    return _f(fv(num(args, 0), num(args, 1), num(args, 2), opt_num(args, 3, 0.0), _type(args, 4)))


@fn("PV", 3, 5, F, S, "PV(rate, nper, pmt, [fv], [type])", "Returns the present value of an investment.")
def _pv(args, ctx):
    return _f(pv(num(args, 0), num(args, 1), num(args, 2), opt_num(args, 3, 0.0), _type(args, 4)))


def _per_check(per: float, nper: float) -> None:
    if per < 1 or per > nper:
        raise XlError(CellError.NUM)


@fn("IPMT", 4, 6, F, S, "IPMT(rate, per, nper, pv, [fv], [type])", "Returns the interest payment for an investment for a given period.")
def _ipmt(args, ctx):
    rate, per, nper, p = num(args, 0), num(args, 1), num(args, 2), num(args, 3)
    _per_check(per, nper)
    return _f(ipmt(rate, per, nper, p, opt_num(args, 4, 0.0), _type(args, 5)))


@fn("PPMT", 4, 6, F, S, "PPMT(rate, per, nper, pv, [fv], [type])", "Returns the payment on the principal for an investment for a given period.")
def _ppmt(args, ctx):
    rate, per, nper, p = num(args, 0), num(args, 1), num(args, 2), num(args, 3)
    _per_check(per, nper)
    f, t = opt_num(args, 4, 0.0), _type(args, 5)
    return _f(pmt(rate, nper, p, f, t) - ipmt(rate, per, nper, p, f, t))


@fn("CUMIPMT", 6, 6, F, S, "CUMIPMT(rate, nper, pv, start_period, end_period, type)", "Returns the cumulative interest paid between two periods.")
def _cumipmt(args, ctx):
    rate, nper, p, s, e = num(args, 0), num(args, 1), num(args, 2), math.trunc(num(args, 3)), math.trunc(num(args, 4))
    t = int(num(args, 5))
    if rate <= 0 or nper <= 0 or p <= 0 or s < 1 or e < s or e > nper or t not in (0, 1) or e - s > 1_000_000:
        raise XlError(CellError.NUM)
    return _f(math.fsum(ipmt(rate, k, nper, p, 0.0, t) for k in range(s, e + 1)))


@fn("CUMPRINC", 6, 6, F, S, "CUMPRINC(rate, nper, pv, start_period, end_period, type)", "Returns the cumulative principal paid on a loan between two periods.")
def _cumprinc(args, ctx):
    rate, nper, p, s, e = num(args, 0), num(args, 1), num(args, 2), math.trunc(num(args, 3)), math.trunc(num(args, 4))
    t = int(num(args, 5))
    if rate <= 0 or nper <= 0 or p <= 0 or s < 1 or e < s or e > nper or t not in (0, 1) or e - s > 1_000_000:
        raise XlError(CellError.NUM)
    pay = pmt(rate, nper, p, 0.0, t)
    return _f(math.fsum(pay - ipmt(rate, k, nper, p, 0.0, t) for k in range(s, e + 1)))


@fn("NPER", 3, 5, F, S, "NPER(rate, pmt, pv, [fv], [type])", "Returns the number of periods for an investment.")
def _nper(args, ctx):
    rate, p, v = num(args, 0), num(args, 1), num(args, 2)
    f, t = opt_num(args, 3, 0.0), _type(args, 4)
    if rate == 0:
        if p == 0:
            raise XlError(CellError.NUM)
        return -(v + f) / p
    num_ = p * (1 + rate * t) - f * rate
    den = v * rate + p * (1 + rate * t)
    if den == 0 or num_ / den <= 0:
        raise XlError(CellError.NUM)
    return _f(math.log(num_ / den) / math.log(1 + rate))


def _solve(f: Callable[[float], float], guess: float) -> float:
    """Newton's method with numeric derivative, then a bracketed fallback."""
    x = guess
    for _ in range(100):
        y = f(x)
        if not math.isfinite(y):
            break
        if abs(y) < 1e-12:
            return x
        h = max(abs(x) * 1e-7, 1e-9)
        d = (f(x + h) - f(x - h)) / (2 * h)
        if d == 0 or not math.isfinite(d):
            break
        nx = x - y / d
        if nx <= -1:
            nx = (x - 1) / 2
        if abs(nx - x) < 1e-12 * max(1.0, abs(x)):
            return nx
        x = nx
    else:
        if abs(f(x)) < 1e-7:
            return x
    # Bracket search on (-0.9999, 10].
    grid = [-0.99, -0.9, -0.5, -0.2, -0.1, 0.0, 0.01, 0.05, 0.1, 0.2, 0.5, 1.0, 2.0, 5.0, 10.0]
    vals = []
    for g in grid:
        try:
            vals.append(f(g))
        except (ZeroDivisionError, OverflowError):
            vals.append(math.nan)
    for i in range(len(grid) - 1):
        a, b = vals[i], vals[i + 1]
        if math.isfinite(a) and math.isfinite(b) and a * b <= 0:
            lo, hi = grid[i], grid[i + 1]
            for _ in range(200):
                mid = (lo + hi) / 2
                fm = f(mid)
                if (fm < 0) == (f(lo) < 0):
                    lo = mid
                else:
                    hi = mid
            return (lo + hi) / 2
    raise XlError(CellError.NUM)


@fn("RATE", 3, 6, F, S, "RATE(nper, pmt, pv, [fv], [type], [guess])", "Returns the interest rate per period of an annuity.")
def _rate(args, ctx):
    nper, p, v = num(args, 0), num(args, 1), num(args, 2)
    f, t = opt_num(args, 3, 0.0), _type(args, 4)
    guess = opt_num(args, 5, 0.1) if has(args, 5) else 0.1
    if nper <= 0:
        raise XlError(CellError.NUM)

    def g(r: float) -> float:
        if abs(r) < 1e-12:
            return v + p * nper + f
        q = (1 + r) ** nper
        return v * q + p * (1 + r * t) * (q - 1) / r + f

    return _f(_solve(g, guess))


def _cash_flows(v: Value) -> List[float]:
    return array_numbers(v)


@fn("NPV", 2, VAR, F, (True, False), "NPV(rate, value1, [value2], …)", "Returns the net present value of an investment based on a series of periodic cash flows and a discount rate.")
def _npv(args, ctx):
    rate = num(args, 0)
    flows: List[float] = []
    for a in args[1:]:
        v = a.value
        if isinstance(v, Array):
            flows.extend(x for x in v.data if type(x) is float)
            for x in v.data:
                if isinstance(x, CellError):
                    raise XlError(x)
        elif isinstance(v, CellError):
            raise XlError(v)
        elif type(v) is float:
            flows.append(v)
        elif v is None or a.from_ref:
            continue
        else:
            flows.append(to_number(v))
    if rate == -1:
        raise XlError(CellError.DIV0)
    return _f(math.fsum(cf / (1 + rate) ** (i + 1) for i, cf in enumerate(flows)))


@fn("IRR", 1, 2, F, (False, True), "IRR(values, [guess])", "Returns the internal rate of return for a series of cash flows.")
def _irr(args, ctx):
    flows = _cash_flows(args[0].value)
    if not any(x > 0 for x in flows) or not any(x < 0 for x in flows):
        raise XlError(CellError.NUM)
    guess = opt_num(args, 1, 0.1) if has(args, 1) else 0.1

    def g(r: float) -> float:
        return math.fsum(cf / (1 + r) ** i for i, cf in enumerate(flows))

    return _f(_solve(g, guess))


@fn("MIRR", 3, 3, F, (False, True, True), "MIRR(values, finance_rate, reinvest_rate)", "Returns the internal rate of return where positive and negative cash flows are financed at different rates.")
def _mirr(args, ctx):
    flows = _cash_flows(args[0].value)
    fr, rr = num(args, 1), num(args, 2)
    n = len(flows)
    pos = math.fsum(cf * (1 + rr) ** (n - 1 - i) for i, cf in enumerate(flows) if cf > 0)
    neg = math.fsum(cf / (1 + fr) ** i for i, cf in enumerate(flows) if cf < 0)
    if pos == 0 or neg == 0 or n < 2:
        raise XlError(CellError.DIV0)
    return _f((pos / -neg) ** (1.0 / (n - 1)) - 1)


def _dated_flows(args, vi: int, di: int):
    vals = as_array(args[vi].value).data
    dates = as_array(args[di].value).data
    if len(vals) != len(dates):
        raise XlError(CellError.NUM)
    vs, ds = [], []
    for v, d in zip(vals, dates):
        if isinstance(v, CellError):
            raise XlError(v)
        if isinstance(d, CellError):
            raise XlError(d)
        vs.append(to_number(v))
        dn = to_number(d)
        if dn < 0:
            raise XlError(CellError.NUM)
        ds.append(math.floor(dn))
    if any(d < ds[0] for d in ds):
        raise XlError(CellError.NUM)
    return vs, ds


@fn("XNPV", 3, 3, F, (True, False, False), "XNPV(rate, values, dates)", "Returns the net present value for a schedule of cash flows that is not necessarily periodic.")
def _xnpv(args, ctx):
    rate = num(args, 0)
    vs, ds = _dated_flows(args, 1, 2)
    if rate <= -1:
        raise XlError(CellError.NUM)
    return _f(math.fsum(v / (1 + rate) ** ((d - ds[0]) / 365.0) for v, d in zip(vs, ds)))


@fn("XIRR", 2, 3, F, (False, False, True), "XIRR(values, dates, [guess])", "Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic.")
def _xirr(args, ctx):
    vs, ds = _dated_flows(args, 0, 1)
    if not any(v > 0 for v in vs) or not any(v < 0 for v in vs):
        raise XlError(CellError.NUM)
    guess = opt_num(args, 2, 0.1) if has(args, 2) else 0.1

    def g(r: float) -> float:
        return math.fsum(v / (1 + r) ** ((d - ds[0]) / 365.0) for v, d in zip(vs, ds))

    return _f(_solve(g, guess))


@fn("EFFECT", 2, 2, F, S, "EFFECT(nominal_rate, npery)", "Returns the effective annual interest rate.")
def _effect(args, ctx):
    r, n = num(args, 0), math.trunc(num(args, 1))
    if r <= 0 or n < 1:
        raise XlError(CellError.NUM)
    return _f((1 + r / n) ** n - 1)


@fn("NOMINAL", 2, 2, F, S, "NOMINAL(effect_rate, npery)", "Returns the annual nominal interest rate.")
def _nominal(args, ctx):
    r, n = num(args, 0), math.trunc(num(args, 1))
    if r <= 0 or n < 1:
        raise XlError(CellError.NUM)
    return _f(n * ((1 + r) ** (1.0 / n) - 1))


@fn("ISPMT", 4, 4, F, S, "ISPMT(rate, per, nper, pv)", "Calculates the interest paid during a specific period of an investment.")
def _ispmt(args, ctx):
    rate, per, nper, v = num(args, 0), num(args, 1), num(args, 2), num(args, 3)
    if nper == 0:
        raise XlError(CellError.DIV0)
    return v * rate * (per / nper - 1)


@fn("PDURATION", 3, 3, F, S, "PDURATION(rate, pv, fv)", "Returns the number of periods required by an investment to reach a specified value.")
def _pduration(args, ctx):
    r, p, f = num(args, 0), num(args, 1), num(args, 2)
    if r <= 0 or p <= 0 or f <= 0:
        raise XlError(CellError.NUM)
    return _f((math.log(f) - math.log(p)) / math.log1p(r))


@fn("RRI", 3, 3, F, S, "RRI(nper, pv, fv)", "Returns an equivalent interest rate for the growth of an investment.")
def _rri(args, ctx):
    n, p, f = num(args, 0), num(args, 1), num(args, 2)
    if n <= 0 or p == 0:
        raise XlError(CellError.NUM)
    return _f((f / p) ** (1.0 / n) - 1)


@fn("FVSCHEDULE", 2, 2, F, (True, False), "FVSCHEDULE(principal, schedule)", "Returns the future value of an initial principal after applying a series of compound interest rates.")
def _fvschedule(args, ctx):
    p = num(args, 0)
    for v in as_array(args[1].value).data:
        if isinstance(v, CellError):
            raise XlError(v)
        p *= 1 + (to_number(v) if v is not None else 0.0)
    return _f(p)


# ---------------------------------------------------------------- depreciation


@fn("SLN", 3, 3, F, S, "SLN(cost, salvage, life)", "Returns the straight-line depreciation of an asset for one period.")
def _sln(args, ctx):
    c, s, l = num(args, 0), num(args, 1), num(args, 2)
    if l == 0:
        raise XlError(CellError.DIV0)
    return (c - s) / l


@fn("SYD", 4, 4, F, S, "SYD(cost, salvage, life, per)", "Returns the sum-of-years' digits depreciation of an asset for a specified period.")
def _syd(args, ctx):
    c, s, l, p = num(args, 0), num(args, 1), num(args, 2), num(args, 3)
    if l <= 0 or p <= 0 or p > l:
        raise XlError(CellError.NUM)
    return (c - s) * (l - p + 1) * 2 / (l * (l + 1))


@fn("DB", 4, 5, F, S, "DB(cost, salvage, life, period, [month])", "Returns the depreciation of an asset for a specified period by using the fixed-declining balance method.")
def _db(args, ctx):
    c, s, l, per = num(args, 0), num(args, 1), num(args, 2), math.trunc(num(args, 3))
    month = math.trunc(opt_num(args, 4, 12.0)) if has(args, 4) else 12
    if c < 0 or s < 0 or l <= 0 or per < 1 or month < 1 or month > 12 or per > l + 1:
        raise XlError(CellError.NUM)
    if c == 0:
        return 0.0
    from .util import round_half_away

    rate = round_half_away(1 - (s / c) ** (1 / l), 3)
    first = c * rate * month / 12
    if per == 1:
        return first
    life = int(l)
    # Book value after period p (2 <= p <= life): (c - first) * (1 - rate)^(p - 1).
    if per <= life:
        return (c - first) * (1 - rate) ** (per - 2) * rate
    book = (c - first) * (1 - rate) ** (life - 1)
    return book * rate * (12 - month) / 12


@fn("DDB", 4, 5, F, S, "DDB(cost, salvage, life, period, [factor])", "Returns the depreciation of an asset for a specified period by using the double-declining balance method.")
def _ddb(args, ctx):
    c, s, l, per = num(args, 0), num(args, 1), num(args, 2), num(args, 3)
    factor = opt_num(args, 4, 2.0) if has(args, 4) else 2.0
    if c < 0 or s < 0 or l <= 0 or per <= 0 or per > l or factor <= 0:
        raise XlError(CellError.NUM)
    r = factor / l
    p = math.ceil(per)

    def book(k: int) -> float:
        if r >= 1:
            return c if k == 0 else min(c, s)
        return max(c * (1 - r) ** k, min(s, c))

    return max(book(p - 1) - book(p), 0.0)


def _vdb_core(c: float, s: float, l: float, start: float, end: float, factor: float, no_switch: bool) -> float:
    """Declining-balance schedule per whole period (switching to straight line over the remaining
    life when that gives more, unless ``no_switch``), summed over the overlap of each period with
    ``(start, end]``."""
    n = math.ceil(end)
    book = c
    switched = False
    sl_dep = 0.0
    total = 0.0
    for k in range(1, n + 1):
        if switched:
            dep = sl_dep
        else:
            ddb = min(book * factor / l, max(book - s, 0.0))
            remaining = l - (k - 1)
            sl = (book - s) / remaining if remaining > 0 else 0.0
            if not no_switch and sl > ddb:
                switched = True
                sl_dep = sl
                dep = sl
            else:
                dep = ddb
        dep = min(dep, max(book - s, 0.0))
        lo, hi = max(start, k - 1), min(end, k)
        if hi > lo:
            total += dep * (hi - lo)
        book -= dep
    return total


@fn("VDB", 5, 7, F, S, "VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])", "Returns the depreciation of an asset for a specified or partial period by using a declining balance method.")
def _vdb(args, ctx):
    c, s, l, a, b = num(args, 0), num(args, 1), num(args, 2), num(args, 3), num(args, 4)
    factor = opt_num(args, 5, 2.0) if has(args, 5) else 2.0
    no_switch = bool(opt_num(args, 6, 0.0)) if has(args, 6) else False
    if c < 0 or s < 0 or l <= 0 or a < 0 or b < a or b > l or factor <= 0 or b > 1_000_000:
        raise XlError(CellError.NUM)
    return _f(_vdb_core(c, s, l, a, b, factor, no_switch))


# ---------------------------------------------------------------- dollar fractions, T-bills, discounts


@fn("DOLLARDE", 2, 2, F, S, "DOLLARDE(fractional_dollar, fraction)", "Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number.")
def _dollarde(args, ctx):
    x, f = num(args, 0), math.trunc(num(args, 1))
    if f < 0:
        raise XlError(CellError.NUM)
    if f == 0:
        raise XlError(CellError.DIV0)
    ip = math.trunc(x)
    digits = 10 ** math.ceil(math.log10(f))
    return ip + (x - ip) * digits / f


@fn("DOLLARFR", 2, 2, F, S, "DOLLARFR(decimal_dollar, fraction)", "Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction.")
def _dollarfr(args, ctx):
    x, f = num(args, 0), math.trunc(num(args, 1))
    if f < 0:
        raise XlError(CellError.NUM)
    if f == 0:
        raise XlError(CellError.DIV0)
    ip = math.trunc(x)
    digits = 10 ** math.ceil(math.log10(f))
    return ip + (x - ip) * f / digits


def _days(args, a: int, b: int):
    s, m = math.floor(num(args, a)), math.floor(num(args, b))
    if s >= m:
        raise XlError(CellError.NUM)
    return s, m


@fn("TBILLPRICE", 3, 3, F, S, "TBILLPRICE(settlement, maturity, discount)", "Returns the price per $100 face value for a Treasury bill.")
def _tbillprice(args, ctx):
    s, m = _days(args, 0, 1)
    d = num(args, 2)
    if d <= 0 or m - s > 365:
        raise XlError(CellError.NUM)
    return 100 * (1 - d * (m - s) / 360)


@fn("TBILLYIELD", 3, 3, F, S, "TBILLYIELD(settlement, maturity, pr)", "Returns the yield for a Treasury bill.")
def _tbillyield(args, ctx):
    s, m = _days(args, 0, 1)
    p = num(args, 2)
    if p <= 0 or m - s > 365:
        raise XlError(CellError.NUM)
    return (100 - p) / p * 360 / (m - s)


@fn("TBILLEQ", 3, 3, F, S, "TBILLEQ(settlement, maturity, discount)", "Returns the bond-equivalent yield for a Treasury bill.")
def _tbilleq(args, ctx):
    s, m = _days(args, 0, 1)
    d = num(args, 2)
    days = m - s
    if d <= 0 or days > 365:
        raise XlError(CellError.NUM)
    if days <= 182:
        return 365 * d / (360 - d * days)
    price = 100 * (1 - d * days / 360)
    t = days / 365
    return (-2 * t + 2 * math.sqrt(t * t - (2 * t - 1) * (1 - 100 / price))) / (2 * t - 1)


def _basis_days(sys, s: int, m: int, basis: int):
    from .date_fns import days360_core, yearfrac_core

    if basis not in (0, 1, 2, 3, 4):
        raise XlError(CellError.NUM)
    return yearfrac_core(sys, s, m, basis)


@fn("DISC", 4, 5, F, S, "DISC(settlement, maturity, pr, redemption, [basis])", "Returns the discount rate for a security.")
def _disc(args, ctx):
    s, m = _days(args, 0, 1)
    pr, red = num(args, 2), num(args, 3)
    basis = int(opt_num(args, 4, 0.0)) if has(args, 4) else 0
    if pr <= 0 or red <= 0:
        raise XlError(CellError.NUM)
    return (red - pr) / red / _basis_days(ctx.date_system, s, m, basis)


@fn("INTRATE", 4, 5, F, S, "INTRATE(settlement, maturity, investment, redemption, [basis])", "Returns the interest rate for a fully invested security.")
def _intrate(args, ctx):
    s, m = _days(args, 0, 1)
    inv, red = num(args, 2), num(args, 3)
    basis = int(opt_num(args, 4, 0.0)) if has(args, 4) else 0
    if inv <= 0 or red <= 0:
        raise XlError(CellError.NUM)
    return (red - inv) / inv / _basis_days(ctx.date_system, s, m, basis)


@fn("RECEIVED", 4, 5, F, S, "RECEIVED(settlement, maturity, investment, discount, [basis])", "Returns the amount received at maturity for a fully invested security.")
def _received(args, ctx):
    s, m = _days(args, 0, 1)
    inv, d = num(args, 2), num(args, 3)
    basis = int(opt_num(args, 4, 0.0)) if has(args, 4) else 0
    if inv <= 0 or d <= 0:
        raise XlError(CellError.NUM)
    den = 1 - d * _basis_days(ctx.date_system, s, m, basis)
    if den <= 0:
        raise XlError(CellError.NUM)
    return inv / den


@fn("PRICEDISC", 4, 5, F, S, "PRICEDISC(settlement, maturity, discount, redemption, [basis])", "Returns the price per $100 face value of a discounted security.")
def _pricedisc(args, ctx):
    s, m = _days(args, 0, 1)
    d, red = num(args, 2), num(args, 3)
    basis = int(opt_num(args, 4, 0.0)) if has(args, 4) else 0
    if d <= 0 or red <= 0:
        raise XlError(CellError.NUM)
    return red * (1 - d * _basis_days(ctx.date_system, s, m, basis))


@fn("YIELDDISC", 4, 5, F, S, "YIELDDISC(settlement, maturity, pr, redemption, [basis])", "Returns the annual yield for a discounted security.")
def _yielddisc(args, ctx):
    s, m = _days(args, 0, 1)
    pr, red = num(args, 2), num(args, 3)
    basis = int(opt_num(args, 4, 0.0)) if has(args, 4) else 0
    if pr <= 0 or red <= 0:
        raise XlError(CellError.NUM)
    return (red / pr - 1) / _basis_days(ctx.date_system, s, m, basis)
