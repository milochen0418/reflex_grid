"""Shared argument coercion and collection helpers."""

from __future__ import annotations

import math
from typing import Iterable, List, Optional, Tuple

from ..core.value import Array, CellError, Value, XlError, scalar, to_bool, to_number, to_text
from .registry import MAX_CELLS, Arg


def err(e: CellError) -> XlError:
    return XlError(e)


def arg(args: List[Arg], i: int) -> Arg:
    if i >= len(args):
        raise XlError(CellError.VALUE)
    return args[i]


def has(args: List[Arg], i: int) -> bool:
    """Whether argument ``i`` was supplied (present and not an omitted literal like ``F(1,)``)."""
    if i >= len(args):
        return False
    a = args[i]
    return a.value is not None or a.from_ref


def val(args: List[Arg], i: int) -> Value:
    return scalar(arg(args, i).value)


def num(args: List[Arg], i: int) -> float:
    """Number coercion for a scalar parameter: empty = 0, TRUE = 1, numeric text parses."""
    return to_number(scalar(arg(args, i).value))


def opt_num(args: List[Arg], i: int, default: float) -> float:
    """Optional number: ``default`` when missing, otherwise coerced (an explicit empty is 0)."""
    if i >= len(args):
        return default
    return to_number(scalar(args[i].value))


def to_int(n: float) -> int:
    if not math.isfinite(n) or abs(n) > 9.007_199_254_740_992e15:
        raise XlError(CellError.NUM)
    return int(n)  # truncates toward zero


def int_(args: List[Arg], i: int) -> int:
    return to_int(num(args, i))


def opt_int(args: List[Arg], i: int, default: int) -> int:
    if not has(args, i):
        return default
    return int_(args, i)


def text(args: List[Arg], i: int) -> str:
    return to_text(scalar(arg(args, i).value))


def opt_text(args: List[Arg], i: int, default: str) -> str:
    if i >= len(args):
        return default
    return to_text(scalar(args[i].value))


def boolean(args: List[Arg], i: int) -> bool:
    return to_bool(scalar(arg(args, i).value))


def opt_bool(args: List[Arg], i: int, default: bool) -> bool:
    if not has(args, i):
        return default
    return to_bool(scalar(args[i].value))


def as_array(v: Value) -> Array:
    """The value viewed as an array (scalars become 1×1)."""
    if isinstance(v, Array):
        return v
    return Array(1, 1, [v])


def array_val(rows: int, cols: int, data: List[Value]) -> Value:
    """Wraps an array result. Rejects empty and oversized arrays."""
    if rows * cols > MAX_CELLS:
        raise XlError(CellError.NUM)
    if rows <= 0 or cols <= 0 or rows * cols != len(data):
        raise XlError(CellError.CALC)
    if rows == 1 and cols == 1:
        return Array(1, 1, data)
    return Array(rows, cols, data)


def check_size(rows: float, cols: float) -> Tuple[int, int]:
    if any(map(math.isnan, (rows, cols))) or rows < 1 or cols < 1 or not math.isfinite(rows) or not math.isfinite(cols):
        raise XlError(CellError.VALUE)
    r, c = int(rows), int(cols)
    if r > 1_048_576 or c > 16_384 or r * c > MAX_CELLS:
        raise XlError(CellError.NUM)
    return r, c


def flatten(args: Iterable[Arg]) -> List[Value]:
    out: List[Value] = []
    for a in args:
        v = a.value
        if isinstance(v, Array):
            out.extend(v.data)
        else:
            out.append(v)
    return out


def numbers(args: Iterable[Arg]) -> List[float]:
    """Numbers for aggregate functions (SUM, AVERAGE, MIN, STDEV…): inside references and arrays
    only numbers count (text, booleans and blanks are skipped); direct scalar arguments are
    coerced ("3" → 3, TRUE → 1, other text → #VALUE!). Errors propagate."""
    out: List[float] = []
    for a in args:
        v = a.value
        if isinstance(v, Array):
            for x in v.data:
                t = type(x)
                if t is float:
                    out.append(x)  # type: ignore[arg-type]
                elif t is CellError:
                    raise XlError(x)  # type: ignore[arg-type]
                elif t is int:
                    out.append(float(x))  # type: ignore[arg-type]
            continue
        t = type(v)
        if t is CellError:
            raise XlError(v)  # type: ignore[arg-type]
        if t is float:
            out.append(v)  # type: ignore[arg-type]
        elif a.from_ref or v is None:
            continue
        else:
            out.append(to_number(v))
    return out


def numbers_a(args: Iterable[Arg]) -> List[float]:
    """Numbers for the "A" aggregates (AVERAGEA, MAXA…): inside references and arrays text
    counts as 0 and booleans as 1/0, blanks are skipped."""
    out: List[float] = []
    for a in args:
        v = a.value
        if isinstance(v, Array):
            for x in v.data:
                t = type(x)
                if t is float:
                    out.append(x)  # type: ignore[arg-type]
                elif t is bool:
                    out.append(1.0 if x else 0.0)
                elif t is str:
                    out.append(0.0)
                elif t is CellError:
                    raise XlError(x)  # type: ignore[arg-type]
            continue
        if isinstance(v, CellError):
            raise XlError(v)
        if v is None:
            continue
        if isinstance(v, str) and a.from_ref:
            out.append(0.0)
            continue
        out.append(to_number(v))
    return out


def array_numbers(v: Value) -> List[float]:
    """Numbers of a single array-like argument where only numbers count, errors propagate."""
    out: List[float] = []
    if isinstance(v, Array):
        for x in v.data:
            t = type(x)
            if t is float:
                out.append(x)  # type: ignore[arg-type]
            elif t is CellError:
                raise XlError(x)  # type: ignore[arg-type]
        return out
    if v is None:
        return out
    if isinstance(v, CellError):
        raise XlError(v)
    if isinstance(v, float):
        return [v]
    return [to_number(v)]


def paired_numbers(a: Value, b: Value) -> Tuple[List[float], List[float]]:
    """Two equally sized arrays, keeping positions where both hold numbers (CORREL, SLOPE…)."""
    x, y = as_array(a), as_array(b)
    if len(x.data) != len(y.data):
        raise XlError(CellError.NA)
    xs: List[float] = []
    ys: List[float] = []
    for p, q in zip(x.data, y.data):
        if isinstance(p, CellError):
            raise XlError(p)
        if isinstance(q, CellError):
            raise XlError(q)
        if type(p) is float and type(q) is float:
            xs.append(p)
            ys.append(q)
    return xs, ys


def round_half_away(x: float, digits: int) -> float:
    """Rounds half away from zero at ``digits`` decimals using the 15-significant-digit decimal
    representation of ``x``, so ROUND(2.675, 2) = 2.68 like Excel."""
    if not math.isfinite(x) or x == 0.0:
        return x
    digits = min(max(digits, -400), 400)
    s = f"{abs(x):.14e}"
    mant, _, exp = s.partition("e")
    e = int(exp)
    d = [ord(c) - 48 for c in mant if c.isdigit()]
    keep = e + 1 + digits
    if keep >= len(d):
        return to_sig_digits(x, len(d))
    if keep < 0:
        return 0.0
    round_up = d[keep] >= 5
    d = d[:keep]
    exp10 = e + 1
    if round_up:
        i = keep
        while True:
            if i == 0:
                d.insert(0, 1)
                exp10 += 1
                break
            i -= 1
            if d[i] == 9:
                d[i] = 0
            else:
                d[i] += 1
                break
    if not d:
        return 0.0
    r = float("0." + "".join(chr(48 + v) for v in d) + f"e{exp10}")
    return -r if x < 0 else r


def to_sig_digits(x: float, sig: int) -> float:
    if not math.isfinite(x) or x == 0.0:
        return x
    return float(f"{x:.{max(min(sig - 1, 30), 0)}e}")


def excel_mod(n: float, d: float) -> float:
    if d == 0:
        raise XlError(CellError.DIV0)
    r = n - d * math.floor(n / d)
    return r


def values_of(v: Value) -> List[Value]:
    return v.data if isinstance(v, Array) else [v]
