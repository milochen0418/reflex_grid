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


def _match(p: tuple, t: str, whole: bool) -> Optional[int]:
    pi = ti = 0
    star: Optional[Tuple[int, int]] = None
    np, nt = len(p), len(t)
    while True:
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


def wildcard_find(pattern: str, text: str, start: int) -> Optional[int]:
    """SEARCH-style find: index of the first position at or after ``start`` where the pattern
    matches."""
    p = _tokenize(pattern)
    t = fold(text)
    if not p:
        return start if start <= len(t) else None
    for i in range(start, len(t)):
        if _match(p, t[i:], False) is not None:
            return i
    return None


def wildcard_find_span(pattern: str, text: str, start: int) -> Optional[Tuple[int, int]]:
    p = _tokenize(pattern)
    t = fold(text)
    for i in range(start, len(t) + 1):
        n = _match(p, t[i:], False)
        if n is not None:
            return i, n
    return None


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
