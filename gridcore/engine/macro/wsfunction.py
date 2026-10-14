"""``WorksheetFunction`` (and ``Application.Sum``-style late-bound calls): VBA calling the
workbook's own worksheet functions.

Arguments are bound by name into the formula evaluator (Range arguments stay references, so
SUMIF, INDEX, MATCH, COUNTIF… see cells, not copies), the call is evaluated as a formula would
be, and the result comes back as VBA values. ``WorksheetFunction.X`` raises run-time error 1004
when the result is an error; ``Application.X`` returns the error value instead, as in Excel.
"""

from __future__ import annotations

from typing import Any, List, Optional

from ...core.value import Array, CellError
from ...vba import values as V
from ...vba.errors import VbaError
from ...vba.runtime import Execution, HostObject
from ...vba.values import EMPTY, MISSING, NOTHING, NULL, Currency, Date, ErrorValue, VArray, is_object
from .host import CELL_ERRORS, ERROR_CODES, cell_to_vba, ole_to_excel

# Worksheet functions VBA can't reach through WorksheetFunction (VBA has its own, or they only
# make sense inside a cell).
_EXCLUDED = {"IF", "LET", "LAMBDA", "OFFSET", "INDIRECT", "ROW", "COLUMN", "ROWS", "COLUMNS", "CELL", "INFO", "NOW", "TODAY", "RAND", "RANDARRAY", "RANDBETWEEN"}


def function_spec(name: str) -> Optional[str]:
    """The worksheet function a member name stands for (upper case), or None."""
    from ...calc.eval import is_known_function

    u = name.upper()
    alias = {"STDEV_S": "STDEV.S", "STDEV_P": "STDEV.P", "VAR_S": "VAR.S", "VAR_P": "VAR.P", "PERCENTILE_INC": "PERCENTILE.INC", "PERCENTILE_EXC": "PERCENTILE.EXC", "QUARTILE_INC": "QUARTILE.INC", "QUARTILE_EXC": "QUARTILE.EXC", "MODE_SNGL": "MODE.SNGL", "RANK_EQ": "RANK.EQ", "RANK_AVG": "RANK.AVG", "NORM_DIST": "NORM.DIST", "NORM_INV": "NORM.INV", "NORM_S_DIST": "NORM.S.DIST", "NORM_S_INV": "NORM.S.INV", "T_TEST": "T.TEST", "CEILING_MATH": "CEILING.MATH", "FLOOR_MATH": "FLOOR.MATH"}.get(u, u)
    if alias in _EXCLUDED:
        return None
    return alias if is_known_function(alias) else None


def to_cell_value(v: Any, d1904: bool) -> Any:
    """A VBA value as a worksheet function argument."""
    if is_object(v) and v is not NOTHING:
        v = V.value_of(v)
    t = type(v)
    if v is EMPTY or v is MISSING:
        return None
    if v is NULL:
        raise VbaError(94)
    if t is bool or t is str:
        return v
    if t is Date:
        return ole_to_excel(float(v), d1904)
    if t is ErrorValue:
        return CELL_ERRORS.get(v.code, CellError.VALUE)
    if t is Currency:
        return float(v)
    if t in (int, V.Integer, V.Byte, V.LongLong, float, V.Single):
        return float(v)
    import decimal

    if t is decimal.Decimal:
        return float(v)
    raise VbaError(13)


def to_array(a: VArray, d1904: bool) -> Array:
    rows = a.rows() if a.dims in (1, 2) else None
    if rows is None:
        raise VbaError(13)
    data = [to_cell_value(x, d1904) for row in rows for x in row]
    return Array(len(rows), len(rows[0]) if rows else 0, data)


def from_result(r: Any, d1904: bool) -> Any:
    if isinstance(r, Array):
        if r.rows == 1 and r.cols == 1:
            return from_result(r.get(0, 0), d1904)
        if r.rows == 1:
            # A single row comes back as a 1-D array, as Excel's Transpose of a column does.
            return VArray([(1, r.cols)], [cell_to_vba(x, "", d1904, True) for x in r.data])
        return VArray.from_rows([[cell_to_vba(r.get(i, j), "", d1904, True) for j in range(r.cols)] for i in range(r.rows)], 1)
    if r is None:
        return EMPTY
    return cell_to_vba(r, "", d1904, True)


class WorksheetFunction(HostObject):
    TYPE_NAME = "WorksheetFunction"

    def __init__(self, h, raise_errors: bool = True):
        self.h = h
        self.raise_errors = raise_errors

    def vba_get(self, ex: Execution, name: str, args: list, named: Optional[dict] = None) -> Any:
        fn = function_spec(name)
        if fn is None:
            raise VbaError(438, f"Object doesn't support this property or method: WorksheetFunction.{name}")
        if named:
            raise VbaError(446, "WorksheetFunction members take positional arguments")
        r = call_function(self.h, fn, list(args))
        if isinstance(r, ErrorValue) and self.raise_errors:
            raise VbaError(1004, f"Unable to get the {name} property of the WorksheetFunction class")
        return r


def call_function(h, fn: str, args: List[Any]) -> Any:
    """Evaluates worksheet function ``fn`` on VBA arguments."""
    from ...calc.eval import Area, Evaluator, RefVal
    from ...calc.recalc import Calc, _PassHost
    from ...core.date import now_serial
    from ...formula import Call, Missing, Name
    from .range import Range

    while args and args[-1] is MISSING:
        args.pop()
    flush = getattr(h, "flush", None)
    if flush is not None and getattr(h, "pending", None) is not None:
        flush()
    d1904 = h.d1904()
    env = []
    exprs = []
    for i, a in enumerate(args):
        if a is MISSING:
            exprs.append(Missing())
            continue
        key = f"__vba_arg{i}"
        if isinstance(a, Range):
            si = h.index_of(a.ws.name)
            env.append((key, RefVal([Area(si, r) for r in a.areas])))
        elif type(a) is VArray:
            env.append((key, to_array(a, d1904)))
        else:
            env.append((key, to_cell_value(a, d1904)))
        exprs.append(Name(key))
    ev_host = getattr(h, "ev", None)
    wb = h.wb()
    if ev_host is not None:
        host = ev_host.host
        at = h.caller
        ev = Evaluator(host, at[0], at[1])
    else:
        host = _PassHost(wb, {}, Calc(), now_serial())
        ev = Evaluator(host, wb.active_sheet, (0, 0))
    ev.env = env
    r = ev.eval(Call(fn, tuple(exprs)))
    if isinstance(r, RefVal):
        if len(r.areas) == 1:
            a0 = r.areas[0]
            return Range(h, h.sheet_object(wb.sheets[a0.sheet].name), [a0.rng])
        r = ev.deref(r)
    return from_result(r, d1904)
