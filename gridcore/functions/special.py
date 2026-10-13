"""Special functions for the statistical distributions: regularized incomplete gamma and beta
functions, the normal distribution and a robust inverse for monotone CDFs."""

from __future__ import annotations

import math
from statistics import NormalDist
from typing import Callable

from ..core.value import CellError, XlError

_STD = NormalDist()
EPS = 1e-15
FPMIN = 1e-300


def norm_cdf(x: float) -> float:
    return 0.5 * math.erfc(-x / math.sqrt(2.0))


def norm_pdf(x: float) -> float:
    return math.exp(-0.5 * x * x) / math.sqrt(2.0 * math.pi)


def norm_inv(p: float) -> float:
    if not 0.0 < p < 1.0:
        raise XlError(CellError.NUM)
    return _STD.inv_cdf(p)


def ln_gamma(x: float) -> float:
    return math.lgamma(x)


def ln_beta(a: float, b: float) -> float:
    return math.lgamma(a) + math.lgamma(b) - math.lgamma(a + b)


def reg_gamma_p(a: float, x: float) -> float:
    """Regularized lower incomplete gamma P(a, x)."""
    if x <= 0:
        return 0.0
    if a <= 0:
        raise XlError(CellError.NUM)
    if x < a + 1.0:
        # Series.
        ap = a
        s = d = 1.0 / a
        for _ in range(100000):
            ap += 1.0
            d *= x / ap
            s += d
            if abs(d) < abs(s) * EPS:
                break
        return min(max(s * math.exp(-x + a * math.log(x) - math.lgamma(a)), 0.0), 1.0)
    return 1.0 - reg_gamma_q(a, x)


def reg_gamma_q(a: float, x: float) -> float:
    """Regularized upper incomplete gamma Q(a, x) = 1 - P(a, x)."""
    if x <= 0:
        return 1.0
    if x < a + 1.0:
        return 1.0 - reg_gamma_p(a, x)
    # Continued fraction (Lentz).
    b = x + 1.0 - a
    c = 1.0 / FPMIN
    d = 1.0 / b
    h = d
    for i in range(1, 100000):
        an = -i * (i - a)
        b += 2.0
        d = an * d + b
        if abs(d) < FPMIN:
            d = FPMIN
        c = b + an / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d
        de = d * c
        h *= de
        if abs(de - 1.0) < EPS:
            break
    return min(max(math.exp(-x + a * math.log(x) - math.lgamma(a)) * h, 0.0), 1.0)


def _betacf(a: float, b: float, x: float) -> float:
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c = 1.0
    d = 1.0 - qab * x / qap
    if abs(d) < FPMIN:
        d = FPMIN
    d = 1.0 / d
    h = d
    for m in range(1, 100000):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < FPMIN:
            d = FPMIN
        c = 1.0 + aa / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        if abs(d) < FPMIN:
            d = FPMIN
        c = 1.0 + aa / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d
        de = d * c
        h *= de
        if abs(de - 1.0) < EPS:
            break
    return h


def reg_inc_beta(x: float, a: float, b: float) -> float:
    """Regularized incomplete beta I_x(a, b)."""
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0
    if a + b > 1e7:
        # Normal approximation of the beta distribution.
        m = a / (a + b)
        sd = math.sqrt(a * b / ((a + b) ** 2 * (a + b + 1)))
        return norm_cdf((x - m) / sd) if sd > 0 else (1.0 if x >= m else 0.0)
    bt = math.exp(math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log1p(-x))
    if x < (a + 1.0) / (a + b + 2.0):
        return min(max(bt * _betacf(a, b, x) / a, 0.0), 1.0)
    return min(max(1.0 - bt * _betacf(b, a, 1.0 - x) / b, 0.0), 1.0)


def invert(cdf: Callable[[float], float], p: float, lo: float, hi: float, lo_bound: float = -math.inf, hi_bound: float = math.inf) -> float:
    """Solves ``cdf(x) = p`` for a non-decreasing ``cdf`` by bracketing and bisection."""
    if not 0.0 <= p <= 1.0:
        raise XlError(CellError.NUM)
    width = max(hi - lo, 1.0)
    n = 0
    while lo > lo_bound and cdf(lo) > p and n < 2000:
        lo = max(lo - width, lo_bound)
        width *= 2.0
        n += 1
    width = max(hi - lo, 1.0)
    n = 0
    while hi < hi_bound and cdf(hi) < p and n < 2000:
        hi = min(hi + width, hi_bound)
        width *= 2.0
        n += 1
    for _ in range(400):
        mid = (lo + hi) / 2.0
        if mid == lo or mid == hi:
            break
        if cdf(mid) < p:
            lo = mid
        else:
            hi = mid
        if abs(hi - lo) <= 1e-15 * max(1.0, abs(mid)):
            break
    return (lo + hi) / 2.0


def _wilson_hilferty(x: float, k: float) -> float:
    """Chi-square CDF by the Wilson-Hilferty normal approximation (huge degrees of freedom)."""
    if x <= 0:
        return 0.0
    z = ((x / k) ** (1.0 / 3.0) - (1.0 - 2.0 / (9.0 * k))) / math.sqrt(2.0 / (9.0 * k))
    return norm_cdf(z)


def t_cdf(t: float, v: float) -> float:
    if v > 1e7:
        return norm_cdf(t)
    x = v / (v + t * t)
    tail = 0.5 * reg_inc_beta(x, v / 2.0, 0.5)
    return 1.0 - tail if t > 0 else tail


def t_pdf(t: float, v: float) -> float:
    return math.exp(math.lgamma((v + 1) / 2) - math.lgamma(v / 2) - 0.5 * math.log(v * math.pi) - (v + 1) / 2 * math.log1p(t * t / v))


def chisq_cdf(x: float, k: float) -> float:
    if k > 2e6:
        return _wilson_hilferty(x, k)
    return reg_gamma_p(k / 2.0, x / 2.0) if x > 0 else 0.0


def gamma_cdf(x: float, a: float) -> float:
    """Regularized lower incomplete gamma P(a, x), approximated for huge ``a``."""
    if a > 1e6:
        return _wilson_hilferty(2.0 * x, 2.0 * a)
    return reg_gamma_p(a, x)


def f_cdf(x: float, d1: float, d2: float) -> float:
    if x <= 0:
        return 0.0
    if d1 > 1e6 or d2 > 1e6:
        # Paulson's normal approximation.
        a, b = 2.0 / (9.0 * d1), 2.0 / (9.0 * d2)
        c = x ** (1.0 / 3.0)
        z = ((1.0 - b) * c - (1.0 - a)) / math.sqrt(a + c * c * b)
        return norm_cdf(z)
    return reg_inc_beta(d1 * x / (d1 * x + d2), d1 / 2.0, d2 / 2.0)


def erf(x: float) -> float:
    return math.erf(x)
