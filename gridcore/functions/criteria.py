"""Excel criteria (COUNTIF, SUMIFS, DCOUNT…) and wildcard matching."""

from __future__ import annotations

from functools import lru_cache
from typing import List, Optional, Tuple

from ..core.parse import parse_number_text
from ..core.value import Array, CellError, Value, XlError, compare_numbers, compare_text, fold

_OPS = (">=", "<=", "<>", "=", ">", "<")

# Excel's limit on a text criterion of COUNTIF, SUMIF, AVERAGEIF and the *IFS functions, in UTF-16
# code units (as Excel counts text).
MAX_CRITERION_UNITS = 255


def _utf16_len(s: str) -> int:
    return len(s) + sum(1 for c in s if ord(c) > 0xFFFF)


def _cmp_op(op: str, o: int) -> bool:
    if op == "=":
        return o == 0
    if op == "<>":
        return o != 0
    if op == "<":
        return o < 0
    if op == "<=":
        return o <= 0
    if op == ">":
        return o > 0
    return o >= 0


class Criterion:
    """A parsed criterion: operator, operand kind and operand."""

    __slots__ = ("op", "kind", "operand", "bare_empty", "pattern")

    def __init__(self, op: str, kind: str, operand, bare_empty: bool = False):
        self.op = op
        self.kind = kind  # number bool error text blank
        self.operand = operand
        self.bare_empty = bare_empty
        self.pattern = _tokenize(operand) if kind == "text" else None

    @staticmethod
    def parse(v: Value) -> "Criterion":
        if isinstance(v, Array):
            v = v.data[0] if v.data else None
        if v is None:
            return Criterion("=", "number", 0.0)
        if isinstance(v, bool):
            return Criterion("=", "bool", v)
        if isinstance(v, (int, float)):
            return Criterion("=", "number", float(v))
        if isinstance(v, CellError):
            return Criterion("=", "error", v)
        return _parse_text(str(v))

    @staticmethod
    def parse_arg(v: Value) -> "Criterion":
        """The criteria argument of COUNTIF, SUMIF, AVERAGEIF and the *IFS functions: a text
        criterion longer than 255 UTF-16 units is #VALUE!, however many criteria pairs there are
        (Excel 16 for Mac)."""
        if isinstance(v, Array):
            v = v.data[0] if v.data else None
        if isinstance(v, str) and len(v) > MAX_CRITERION_UNITS // 2 and _utf16_len(v) > MAX_CRITERION_UNITS:
            raise XlError(CellError.VALUE)
        return Criterion.parse(v)

    def matches(self, v: Value) -> bool:
        if isinstance(v, Array):
            v = v.data[0] if v.data else None
        if self.bare_empty:
            return v is None or v == ""
        k = self.kind
        op = self.op
        if k == "blank":
            blank = v is None
            if op == "=":
                return blank
            if op == "<>":
                return not blank
            if isinstance(v, str):
                # A bare "<=" or ">=" matches no empty text (Excel 365).
                return v != "" and _cmp_op(op, compare_text(v, ""))
            return False
        if k == "number":
            if isinstance(v, bool):
                cell = None
            elif isinstance(v, (int, float)):
                cell = float(v)
            elif isinstance(v, str) and op in ("=", "<>"):
                cell = parse_number_text(v)
            else:
                cell = None
            if cell is None:
                return op == "<>"
            return _cmp_op(op, compare_numbers(cell, self.operand))
        if k == "bool":
            if isinstance(v, bool):
                return _cmp_op(op, (v > self.operand) - (v < self.operand))
            return op == "<>"
        if k == "error":
            if isinstance(v, CellError):
                return _cmp_op(op, (v.code > self.operand.code) - (v.code < self.operand.code))
            return op == "<>"
        # text
        if op in ("=", "<>"):
            m = isinstance(v, str) and _match(self.pattern, fold(v), True) is not None
            return m if op == "=" else not m
        if isinstance(v, str):
            return _cmp_op(op, compare_text(v, self.operand))
        return False


@lru_cache(maxsize=4096)
def _parse_text_cached(t: str) -> Tuple[str, str, object, bool]:
    for o in _OPS:
        if t.startswith(o):
            op, rest = o, t[len(o):]
            break
    else:
        if t == "":
            return "=", "blank", None, True
        op, rest = "=", t
    if rest == "":
        return op, "blank", None, False
    n = parse_number_text(rest)
    if n is not None:
        return op, "number", n, False
    u = rest.upper()
    if u == "TRUE":
        return op, "bool", True, False
    if u == "FALSE":
        return op, "bool", False, False
    e = CellError.parse(rest)
    if e is not None:
        return op, "error", e, False
    return op, "text", rest, False


def _parse_text(t: str) -> Criterion:
    op, kind, operand, bare = _parse_text_cached(t)
    return Criterion(op, kind, operand, bare)


# ---------------------------------------------------------------- wildcards

_STAR = object()
_ANY = object()


@lru_cache(maxsize=4096)
def _tokenize(pattern: str) -> tuple:
    out: list = []
    it = iter(pattern)
    for c in it:
        if c == "~":
            n = next(it, None)
            if n is None:
                out.append("~")
            else:
                out.append(fold(n))
        elif c == "*":
            out.append(_STAR)
        elif c == "?":
            out.append(_ANY)
        else:
            out.append(fold(c))
    return tuple(out)


def has_wildcards(pattern: str) -> bool:
    return any(t is _STAR or t is _ANY for t in _tokenize(pattern)) or "~" in pattern


# Steps one match may take before it gives up and reports no match. An honest match against
# Excel's 32,767-character cell ceiling costs a few hundred thousand; a hostile pattern (star
# churn: `*a*a*a*a*…b` over a text of a's) is quadratic, so it is capped instead of running for
# minutes — quiet degradation, as with the other work caps on far-past-realistic input.
_MAX_STEPS = 1 << 22


def _match(p: tuple, t: str, whole: bool, budget: Optional[List[int]] = None) -> Optional[int]:
    pi = ti = 0
    star: Optional[Tuple[int, int]] = None
    np, nt = len(p), len(t)
    steps = budget if budget is not None else [_MAX_STEPS]
    while True:
        steps[0] -= 1
        if steps[0] < 0:
            return None
        if pi == np and (not whole or ti == nt):
            return ti
        tok = p[pi] if pi < np else None
        if tok is _STAR:
            star = (pi, ti)
            pi += 1
            continue
        if tok is not None and ti < nt and (tok is _ANY or tok == t[ti]):
            pi += 1
            ti += 1
            continue
        if star is not None and star[1] < nt:
            sp, st = star
            star = (sp, st + 1)
            pi = sp + 1
            ti = st + 1
            continue
        return None


def wildcard_match(pattern: str, text: str) -> bool:
    """Case-insensitive whole-text match with ``*``, ``?`` and ``~`` escapes."""
    return _match(_tokenize(pattern), fold(text), True) is not None


def _find_from(p: tuple, t: str, start: int) -> Optional[Tuple[int, int]]:
    """The first position at or after ``start`` where ``p`` matches a prefix of ``t[i:]``, with the
    matched length; ``None`` when ``start`` is past the text. A leading run of stars absorbs the
    text before the rest matches, so the match starts at ``start`` (SEARCH("*b","aab") is 1 in
    Excel); it is dropped and the first token that consumes a character anchors the scan: a literal
    skips straight to its occurrences, so the attempts stay linear, and every attempt shares one
    step budget, so star churn stays flat instead of quadratic."""
    if start > len(t):
        return None
    skip = 0
    while skip < len(p) and p[skip] is _STAR:
        skip += 1
    p = p[skip:]
    if not p:
        return start, len(t) - start if skip else 0
    anys = 0
    while anys < len(p) and p[anys] is _ANY:
        anys += 1
    anchor = p[anys] if anys < len(p) else None
    budget = [_MAX_STEPS]
    hit = None
    if isinstance(anchor, str):
        j = start + anys
        while hit is None:
            j = t.find(anchor, j)
            if j < 0:
                return None
            n = _match(p, t[j - anys:], False, budget)
            if n is not None:
                hit = (j - anys, n)
            elif budget[0] < 0:
                return None
            j += 1
    else:
        for i in range(start, len(t) + 1):
            n = _match(p, t[i:], False, budget)
            if n is not None:
                hit = (i, n)
                break
            if budget[0] < 0:
                return None
    if hit is None:
        return None
    return (start, hit[0] - start + hit[1]) if skip else hit


def wildcard_find(pattern: str, text: str, start: int) -> Optional[int]:
    """SEARCH-style find: index of the first position at or after ``start`` where the pattern
    matches; ``None`` from the end of the text on (SEARCH("","abc",4) is #VALUE! in Excel 16)."""
    t = fold(text)
    if start >= len(t):
        return None
    hit = _find_from(_tokenize(pattern), t, start)
    return None if hit is None else hit[0]


def wildcard_find_span(pattern: str, text: str, start: int) -> Optional[Tuple[int, int]]:
    return _find_from(_tokenize(pattern), fold(text), start)


def lookup_equal(needle: Value, hay: Value, wild: bool) -> bool:
    """Exact-or-wildcard equality used by MATCH/VLOOKUP/XLOOKUP exact modes."""
    tn, th = type(needle), type(hay)
    if tn is str and th is str:
        if wild:
            return wildcard_match(needle, hay)  # type: ignore[arg-type]
        return fold(needle) == fold(hay)  # type: ignore[arg-type]
    if tn is float and th is float:
        return compare_numbers(needle, hay) == 0  # type: ignore[arg-type]
    if tn is bool and th is bool:
        return needle == hay
    if tn is CellError and th is CellError:
        return needle is hay
    if needle is None and hay is None:
        return True
    return False


def criteria_pairs(args, start: int) -> List[Tuple[Array, Criterion]]:
    """``(range, criterion)`` pairs of a *IFS function from argument ``start``."""
    from .util import as_array

    out = []
    for i in range(start, len(args), 2):
        rng = as_array(args[i].value)
        crit = Criterion.parse(args[i + 1].value if i + 1 < len(args) else None)
        out.append((rng, crit))
    return out
