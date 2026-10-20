"""The formula evaluator.

Expressions evaluate to a value, a reference (:class:`RefVal`, one or more areas kept unresolved
so ``ROW``, ``OFFSET``, ``INDEX``, ``:`` and friends can work on references) or a
:class:`Lambda`. Cell reads go through a host (see :class:`Host`), which lets the recalculation
engine evaluate dirty precedents on demand.
"""

from __future__ import annotations

import math
from typing import Any, Callable, Dict, List, NamedTuple, Optional, Sequence, Tuple

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef, col_to_letters
from ..core.value import Array, CellError, Value, XlError, compare, scalar, to_bool, to_number, to_number_in, to_text
from ..formula import (
    ArrayLit,
    Binary,
    Bool,
    Call,
    Err,
    Expr,
    Invoke,
    Missing,
    Name,
    Num,
    ParseError,
    Paren,
    Ref,
    Str,
    StructRef,
    Unary,
    parse,
)
from ..formula.adjust import has_relative_refs, shift_relative_wrapping
from ..functions import registry as freg
from ..functions.registry import Arg

# Largest array a reference may expand to.
MAX_CELLS = 16_000_000
# References up to this many cells are read at their full size; larger ones (A:A, ranges reserved
# for data to come) stop at the sheet's used range.
FULL_SIZE_CELLS = 16_384
MAX_DEPTH = 400
# The argument of GROUPBY and PIVOTBY that takes a function (a worksheet function by its name).
_FUNCTION_ARG = {"GROUPBY": 2, "PIVOTBY": 3}
# Operators whose left-nested chains are evaluated in a loop (reference operators keep recursing).
_CHAINED = frozenset(("+", "-", "*", "/", "^", "&", "=", "<>", "<", ">", "<=", ">="))


# Worksheet functions defined outside the function library (VBA Functions of the workbook's
# macros): set by gridcore.engine.macro. Called as hook(evaluator, NAME, arg exprs); returns the
# result, or None when it doesn't know the name.
UDF_HOOK: Optional[Callable[["Evaluator", str, Sequence[Expr]], Any]] = None


class Area(NamedTuple):
    sheet: int
    rng: RangeRef


class RefVal:
    __slots__ = ("areas",)

    def __init__(self, areas: List[Area]):
        self.areas = areas

    def __repr__(self) -> str:
        return f"RefVal({self.areas!r})"


class Lambda:
    __slots__ = ("params", "body", "env")

    def __init__(self, params: List[str], body: Expr, env: List[Tuple[str, Any]]):
        self.params = params
        self.body = body
        self.env = env


class Host:
    """What the evaluator needs from its surroundings (a base class with default behaviour)."""

    wb: Any

    def cell_value(self, sheet: int, rc: Tuple[int, int]) -> Value:
        s = self.wb.sheet(sheet)
        return s.value(rc) if s is not None else None

    def now_serial(self) -> float:
        from ..core.date import now_serial

        return now_serial()

    def random(self) -> float:
        import random

        return random.random()

    def note_volatile(self) -> None:
        pass

    def spill_range(self, sheet: int, anchor: Tuple[int, int]) -> Optional[RangeRef]:
        s = self.wb.sheet(sheet)
        return s.spill_ranges.get(anchor) if s is not None else None

    def range_values(self, sheet: int, rng: RangeRef) -> List[Value]:
        return [self.cell_value(sheet, rc) for rc in rng.cells()]

    def used_range(self, sheet: int) -> Optional[RangeRef]:
        s = self.wb.sheet(sheet)
        return s.used_range() if s is not None else None

    def visit_range(self, sheet: int, rng: RangeRef, f: Callable[[Value], bool]) -> bool:
        return False


class _FnCtx(freg.Ctx):
    __slots__ = ("ev",)

    def __init__(self, ev: "Evaluator"):
        self.ev = ev
        self.date_system = ev.host.wb.date_system

    def now_serial(self) -> float:
        return self.ev.host.now_serial()

    def random(self) -> float:
        return self.ev.host.random()


def _err(e: CellError) -> CellError:
    return e


class Evaluator:
    __slots__ = ("host", "sheet", "at", "env", "depth")

    def __init__(self, host: Host, sheet: int, at: Tuple[int, int]):
        self.host = host
        self.sheet = sheet
        self.at = at
        self.env: List[Tuple[str, Any]] = []
        self.depth = 0

    # ------------------------------------------------------------ values

    def value(self, e: Expr) -> Value:
        """Evaluates to a value (references are dereferenced; multi-cell → array)."""
        return self.deref(self.eval(e))

    def deref(self, ev: Any) -> Value:
        if isinstance(ev, RefVal):
            if len(ev.areas) == 1:
                return self.area_value(ev.areas[0])
            return CellError.VALUE
        if isinstance(ev, Lambda):
            return CellError.CALC
        return ev

    def area_value(self, a: Area) -> Value:
        rng = a.rng
        if rng.r0 == rng.r1 and rng.c0 == rng.c1:
            return self.host.cell_value(a.sheet, (rng.r0, rng.c0))
        cells = rng.height * rng.width
        if cells > FULL_SIZE_CELLS:
            rng = self.trim(a)
        h, w = rng.height, rng.width
        if h * w > MAX_CELLS:
            return CellError.NUM
        data = self.host.range_values(a.sheet, rng)
        if len(data) != h * w:
            return CellError.VALUE
        return Array(h, w, data)

    def trim(self, a: Area) -> RangeRef:
        """Trims the bottom/right of an area to the sheet's used range (keeps the top-left)."""
        if self.host.wb.sheet(a.sheet) is None:
            return a.rng
        used = self.host.used_range(a.sheet)
        r1, c1 = a.rng.r1, a.rng.c1
        if used is not None:
            r1 = min(r1, max(used.r1, a.rng.r0))
            c1 = min(c1, max(used.c1, a.rng.c0))
        else:
            r1, c1 = a.rng.r0, a.rng.c0
        return RangeRef(a.rng.r0, a.rng.c0, r1, c1)

    def resolve_sheet(self, s) -> List[int]:
        wb = self.host.wb
        if s is None:
            return [self.sheet]
        if isinstance(s, str):
            i = wb.sheet_index(s)
            if i is None:
                raise XlError(CellError.REF)
            return [i]
        a, b = s
        i, j = wb.sheet_index(a), wb.sheet_index(b)
        if i is None or j is None:
            raise XlError(CellError.REF)
        return list(range(min(i, j), max(i, j) + 1))

    def reference_areas(self, r: Ref) -> List[Area]:
        rng = r.range()
        return [Area(s, rng) for s in self.resolve_sheet(r.sheet)]

    # ------------------------------------------------------------ eval

    def eval(self, e: Expr) -> Any:
        self.depth += 1
        if self.depth > MAX_DEPTH:
            self.depth -= 1
            return CellError.CALC
        try:
            return _DISPATCH[type(e)](self, e)
        except XlError as x:
            return x.error
        except RecursionError:
            return CellError.CALC
        except Exception:  # never crash: an engine bug is this formula's #VALUE!
            return CellError.VALUE
        finally:
            self.depth -= 1

    def _num(self, e: Num):
        v = e.value
        return v if math.isfinite(v) else CellError.NUM

    def _str(self, e: Str):
        return e.value

    def _bool(self, e: Bool):
        return e.value

    def _errv(self, e: Err):
        return e.error

    def _missing(self, e: Missing):
        return None

    def _paren(self, e: Paren):
        return self.eval(e.expr)

    def _array(self, e: ArrayLit):
        h = len(e.rows)
        w = len(e.rows[0]) if h else 0
        data = [self.value(el) for row in e.rows for el in row]
        if h * w != len(data) or h == 0 or w == 0:
            return CellError.VALUE
        return Array(h, w, data)

    def _ref(self, e: Ref):
        return RefVal(self.reference_areas(e))

    def _invoke(self, e: Invoke):
        f = self.eval(e.callee)
        if isinstance(f, Lambda):
            return self.apply_lambda(f, e.args)
        if isinstance(f, CellError):
            return f
        return CellError.VALUE

    def _name(self, e: Name):
        return self.name(e.name)

    def name(self, n: str):
        nl = n.lower()
        for k, v in reversed(self.env):
            if k.lower() == nl:
                return v
        wb = self.host.wb
        if "!" in n:
            s, bare = n.rsplit("!", 1)
            si = wb.sheet_index(s.strip("'").replace("''", "'"))
            if si is None:
                return CellError.NAME
            scope_sheet = si
        else:
            scope_sheet, bare = self.sheet, n
        d = wb.name(bare, scope_sheet)
        if d is not None:
            try:
                expr = parse(d.formula)
            except ParseError:
                return CellError.NAME
            # Relative references in a name are relative to A1: used here, they move by this
            # cell's offset from it.
            if has_relative_refs(expr):
                expr = shift_relative_wrapping(expr, self.at[0], self.at[1])
            saved = self.sheet
            self.sheet = d.scope if d.scope is not None else self.sheet
            try:
                return self.eval(expr)
            finally:
                self.sheet = saved
        # A table name alone means its data body.
        t = wb.table(bare)
        if t is not None:
            si, ti = t
            tab = wb.sheets[si].tables[ti]
            dr = tab.data_range()
            if dr is not None:
                return RefVal([Area(si, dr)])
        return CellError.NAME

    def _struct(self, s: StructRef):
        wb = self.host.wb
        if not s.table:
            sh = wb.sheet(self.sheet)
            found = None
            if sh is not None:
                for ti, t in enumerate(sh.tables):
                    if t.range.contains(self.at):
                        found = (self.sheet, ti)
                        break
        else:
            found = wb.table(s.table)
        if found is None:
            return CellError.REF
        si, ti = found
        t = wb.sheets[si].tables[ti]
        r0, r1 = t.range.r0, t.range.r1
        data = t.data_range()
        items = s.specifiers
        if not items or items == ("#Data",):
            if data is None:
                return CellError.REF
            r0, r1 = data.r0, data.r1
        elif "#This Row" in items:
            if not t.range.r0 <= self.at[0] <= t.range.r1:
                return CellError.VALUE
            r0 = r1 = self.at[0]
        elif "#All" not in items:
            lo, hi = 1 << 30, -1
            if "#Headers" in items:
                if not t.header_row:
                    return CellError.REF
                lo, hi = min(lo, t.range.r0), max(hi, t.range.r0)
            if "#Data" in items and data is not None:
                lo, hi = min(lo, data.r0), max(hi, data.r1)
            if "#Totals" in items:
                if not t.totals_row:
                    return CellError.REF
                lo, hi = min(lo, t.range.r1), max(hi, t.range.r1)
            if lo > hi:
                return CellError.REF
            r0, r1 = lo, hi
        c0, c1 = t.range.c0, t.range.c1
        if s.col_start is not None:
            i = t.column_index(s.col_start)
            if i is None:
                return CellError.REF
            c0 = c1 = t.range.c0 + i
            if s.col_end is not None:
                j = t.column_index(s.col_end)
                if j is None:
                    return CellError.REF
                cj = t.range.c0 + j
                c0, c1 = min(c0, cj), max(c1, cj)
        return RefVal([Area(si, RangeRef(r0, c0, r1, c1))])

    def _unary(self, e: Unary):
        op = e.op
        if op in (".:", ":.", ".:."):
            return self.trim_ref(op, e.expr)
        if op == "#":
            x = self.eval(e.expr)
            if isinstance(x, RefVal):
                if len(x.areas) == 1 and x.areas[0].rng.is_single():
                    a = x.areas[0]
                    r = self.host.spill_range(a.sheet, (a.rng.r0, a.rng.c0))
                    if r is not None:
                        return RefVal([Area(a.sheet, r)])
                return CellError.REF
            if isinstance(x, CellError):
                return x
            return CellError.REF
        if op == "@":
            x = self.eval(e.expr)
            if isinstance(x, RefVal):
                if len(x.areas) == 1:
                    c = self.implicit_intersect(x.areas[0])
                    if c is not None:
                        return self.host.cell_value(x.areas[0].sheet, c)
                return CellError.VALUE
            if isinstance(x, Lambda):
                return x
            return scalar(x)
        v = self.value(e.expr)
        sys = self.host.wb.date_system
        if op == "-":
            return _map_unary(v, lambda n: -n, False, sys)
        if op == "+":
            return _map_unary(v, lambda n: n, True, sys)
        return _map_unary(v, lambda n: n / 100.0, False, sys)

    def implicit_intersect(self, a: Area) -> Optional[Tuple[int, int]]:
        r = a.rng
        if r.is_single():
            return r.start
        if r.width == 1 and r.r0 <= self.at[0] <= r.r1:
            return (self.at[0], r.c0)
        if r.height == 1 and r.c0 <= self.at[1] <= r.c1:
            return (r.r0, self.at[1])
        return None

    def _binary(self, e: Binary):
        op = e.op
        if op == ":":
            x, y = self.eval(e.left), self.eval(e.right)
            if not isinstance(x, RefVal) or not isinstance(y, RefVal):
                if isinstance(x, CellError):
                    return x
                if isinstance(y, CellError):
                    return y
                return CellError.VALUE
            allv = x.areas + y.areas
            if not allv:
                return CellError.REF
            first = allv[0]
            if any(a.sheet != first.sheet for a in allv):
                return CellError.VALUE
            r = first.rng
            for a in allv[1:]:
                r = r.union(a.rng)
            return RefVal([Area(first.sheet, r)])
        if op == ",":
            x, y = self.eval(e.left), self.eval(e.right)
            if isinstance(x, RefVal) and isinstance(y, RefVal):
                return RefVal(x.areas + y.areas)
            if isinstance(x, CellError):
                return x
            if isinstance(y, CellError):
                return y
            return CellError.VALUE
        if op == " ":
            x, y = self.eval(e.left), self.eval(e.right)
            if isinstance(x, RefVal) and isinstance(y, RefVal):
                out = []
                for p in x.areas:
                    for q in y.areas:
                        if p.sheet == q.sheet:
                            i = p.rng.intersection(q.rng)
                            if i is not None:
                                out.append(Area(p.sheet, i))
                return RefVal(out) if out else CellError.NULL
            if isinstance(x, CellError):
                return x
            if isinstance(y, CellError):
                return y
            return CellError.VALUE
        sys = self.host.wb.date_system
        if op in _CHAINED and isinstance(e.left, Binary) and e.left.op in _CHAINED:
            # A long chain such as A1+A2+…+A2000 nests one level per operator: its left spine is
            # walked in a loop, so it stays within the depth limit as in Excel (8,192 characters).
            chain = []
            x: Expr = e
            while isinstance(x, Binary) and x.op in _CHAINED:
                chain.append(x)
                x = x.left
            acc = self.value(x)
            for b in reversed(chain):
                acc = binary_values(b.op, acc, self.value(b.right), sys)
            return acc
        return binary_values(op, self.value(e.left), self.value(e.right), sys)

    # ------------------------------------------------------------ calls

    def arg_num(self, args: Sequence[Expr], i: int, default: float) -> float:
        if i >= len(args) or isinstance(args[i], Missing):
            return default
        return to_number(scalar(self.value(args[i])))

    def _call(self, e: Call):
        name = e.name
        h = _SPECIAL.get(name)
        if h is not None:
            return h(self, name, e.args)
        return self.builtin(name, e.args)

    def builtin(self, name: str, args: Sequence[Expr]):
        spec = freg.lookup(name)
        if spec is None:
            # A defined name holding a LAMBDA, or a LET-bound lambda.
            f = self.name(name)
            if isinstance(f, Lambda):
                return self.apply_lambda(f, args)
            hook = UDF_HOOK
            if hook is not None:
                self.host.note_volatile()
                r = hook(self, name, args)
                if r is not None:
                    return r
            return CellError.NAME
        if len(args) < spec.min_args or len(args) > spec.max_args:
            return CellError.VALUE
        if spec.volatile:
            self.host.note_volatile()
        if name in _AGG:
            return self.aggregate(name, args)
        evaluated: List[Arg] = []
        criteria_shape = None
        resizes = len(args) == 3 and name in ("SUMIF", "AVERAGEIF")
        fpos = _FUNCTION_ARG.get(name)
        for i, a in enumerate(args):
            if i == fpos and isinstance(a, Name) and self.is_function_name(a.name):
                # GROUPBY(…, SUM): the function by its name.
                evaluated.append(Arg(a.name.upper(), False))
                continue
            ev = self.eval(a)
            if resizes:
                if i == 0:
                    criteria_shape = _ev_shape(ev)
                elif i == 2 and criteria_shape is not None and isinstance(ev, RefVal) and len(ev.areas) == 1:
                    ar = ev.areas[0]
                    ev = RefVal([Area(ar.sheet, resized_sum_range(ar.rng, *criteria_shape))])
            from_ref = isinstance(ev, RefVal)
            if from_ref and len(ev.areas) > 1:
                vals: List[Value] = []
                for ar in ev.areas:
                    v = self.area_value(ar)
                    if isinstance(v, Array):
                        vals.extend(v.data)
                    else:
                        vals.append(v)
                value: Value = Array.row(vals)
            else:
                value = self.deref(ev)
            evaluated.append(Arg(value, from_ref))
        return freg.call(spec, evaluated, _FnCtx(self))

    def aggregate(self, name: str, args: Sequence[Expr]):
        """SUM, COUNT and AVERAGE fold references over the stored cells without building arrays."""
        total = 0.0
        count = 0
        first_error = None
        for a in args:
            ev = self.eval(a)
            if isinstance(ev, RefVal):
                for area in ev.areas:
                    r = self._agg_area(name, area)
                    if isinstance(r, CellError):
                        if first_error is None:
                            first_error = r
                    else:
                        total += r[0]
                        count += r[1]
            else:
                v = self.deref(ev)
                r = _agg_value(name, v, False)
                if isinstance(r, CellError):
                    if first_error is None:
                        first_error = r
                else:
                    total += r[0]
                    count += r[1]
        if first_error is not None:
            return first_error
        if name == "SUM":
            return total if math.isfinite(total) else CellError.NUM
        if name == "COUNT":
            return float(count)
        if count == 0:
            return CellError.DIV0
        return total / count

    def _agg_area(self, name: str, a: Area):
        if not a.rng.is_single():
            acc = [0.0, 0]
            err: List[CellError] = []
            count_only = name == "COUNT"

            def visit(v: Value) -> bool:
                t = type(v)
                if t is float:
                    acc[0] += v  # type: ignore[operator]
                    acc[1] += 1
                elif t is CellError and not count_only:
                    err.append(v)  # type: ignore[arg-type]
                    return False
                return True

            if self.host.visit_range(a.sheet, a.rng, visit):
                if err:
                    return err[0]
                return acc[0], acc[1]
        return _agg_value(name, self.area_value(a), True)

    def apply_lambda(self, l: Lambda, args: Sequence[Expr]):
        # One argument position per parameter: ``(1,)`` leaves the second omitted, ``(1)`` is
        # #VALUE! (Excel); ``f()`` of a one-parameter LAMBDA counts as one empty position.
        n = len(args)
        if n > len(l.params) or (n < len(l.params) and not (n == 0 and len(l.params) == 1)):
            return CellError.VALUE
        bound = []
        for i, p in enumerate(l.params):
            if i >= len(args) or isinstance(args[i], Missing):
                v: Any = None
            else:
                v = self.eval(args[i])
            bound.append((p, v))
        return self.call_lambda_values(l, bound)

    def call_lambda_values(self, l: Lambda, bound: List[Tuple[str, Any]]):
        saved = self.env
        self.env = list(l.env) + bound
        try:
            return self.eval(l.body)
        finally:
            self.env = saved

    def lambda_arg(self, e: Optional[Expr], arity: int = 1) -> Optional[Lambda]:
        """The function argument of MAP, REDUCE, BYROW…: a LAMBDA, or a worksheet function by its
        name (``BYROW(A1:B9, SUM)``) taking ``arity`` arguments."""
        if e is None:
            return None
        v = self.eval(e)
        if isinstance(v, Lambda):
            return v
        if isinstance(e, Name) and self.is_function_name(e.name):
            params = [f"\x00{i}" for i in range(arity)]
            return Lambda(params, Call(e.name.upper(), tuple(Name(p) for p in params)), [])
        return None

    def is_function_name(self, n: str) -> bool:
        """``n`` names a worksheet function, and no LET variable or defined name."""
        nl = n.lower()
        if "!" in n or any(k.lower() == nl for k, _ in self.env) or self.host.wb.name(n, self.sheet) is not None:
            return False
        return is_worksheet_function(n)

    def trim_lines(self, a: Area) -> RangeRef:
        if a.rng.is_full_cols() or a.rng.is_full_rows():
            return self.trim(a)
        return a.rng

    def trim_ref(self, mode: str, x: Expr):
        """``A1:.E10``-style trim references (TRIMRANGE with trim codes)."""
        ev = self.eval(x)
        if not isinstance(ev, RefVal):
            return ev if isinstance(ev, CellError) else CellError.VALUE
        code = {".:": 1, ":.": 2, ".:.": 3}[mode]
        out = []
        for a in ev.areas:
            r = self._trim_area(a, code, code)
            if r is None:
                return CellError.REF
            out.append(Area(a.sheet, r))
        return RefVal(out)

    def _trim_area(self, a: Area, trim_rows: int, trim_cols: int) -> Optional[RangeRef]:
        sh = self.host.wb.sheet(a.sheet)
        rng = a.rng
        if sh is None:
            return rng
        used = self.host.used_range(a.sheet)
        if used is None:
            return RangeRef(rng.r0, rng.c0, rng.r0, rng.c0)
        inter = rng.intersection(used)
        if inter is None:
            return RangeRef(rng.r0, rng.c0, rng.r0, rng.c0)
        filled_rows = set()
        filled_cols = set()
        for (r, c), cell in sh.cells.iter_range(inter):
            if cell.value is not None and cell.value != "" or cell.formula is not None:
                filled_rows.add(r)
                filled_cols.add(c)
        for (r, c) in sh.spill:
            if inter.contains((r, c)):
                filled_rows.add(r)
                filled_cols.add(c)
        if not filled_rows:
            return RangeRef(rng.r0, rng.c0, rng.r0, rng.c0)
        r0, r1, c0, c1 = rng.r0, rng.r1, rng.c0, rng.c1
        if trim_rows in (1, 3):
            r0 = max(r0, min(filled_rows))
        if trim_rows in (2, 3):
            r1 = min(r1, max(filled_rows))
        if trim_cols in (1, 3):
            c0 = max(c0, min(filled_cols))
        if trim_cols in (2, 3):
            c1 = min(c1, max(filled_cols))
        return RangeRef(r0, c0, max(r1, r0), max(c1, c0))


# ------------------------------------------------------------------ special forms


def _sf_if(ev: Evaluator, name: str, args):
    if not 1 <= len(args) <= 3:
        return CellError.VALUE
    c = ev.value(args[0])
    if isinstance(c, Array):
        t = ev.value(args[1]) if len(args) > 1 else True
        f = ev.value(args[2]) if len(args) > 2 else False
        if isinstance(t, Lambda) or isinstance(f, Lambda):
            return CellError.CALC

        def pick(cv, tv, fv):
            try:
                return tv if to_bool(cv) else fv
            except XlError as x:
                return x.error

        return zip3(c, t, f, pick)
    try:
        b = to_bool(c)
    except XlError as x:
        return x.error
    if b:
        if len(args) > 1:
            return 0.0 if isinstance(args[1], Missing) else ev.eval(args[1])
        return True
    if len(args) > 2:
        return 0.0 if isinstance(args[2], Missing) else ev.eval(args[2])
    return False


def _sf_ifs(ev: Evaluator, name: str, args):
    if len(args) < 2 or len(args) % 2:
        return CellError.VALUE
    for i in range(0, len(args), 2):
        cv = scalar(ev.value(args[i]))
        try:
            if to_bool(cv):
                return ev.eval(args[i + 1])
        except XlError as x:
            return x.error
    return CellError.NA


def _sf_iferror(ev: Evaluator, name: str, args):
    if len(args) != 2:
        return CellError.VALUE
    na_only = name == "IFNA"

    def catch(e) -> bool:
        return isinstance(e, CellError) and (not na_only or e is CellError.NA)

    v = ev.eval(args[0])
    if isinstance(v, RefVal):
        v = ev.deref(v)
    if catch(v):
        return ev.eval(args[1])
    if isinstance(v, Array) and any(catch(x) for x in v.data):
        alt = scalar(ev.value(args[1]))
        return Array(v.rows, v.cols, [alt if catch(x) else x for x in v.data])
    return v


def _sf_choose(ev: Evaluator, name: str, args):
    if len(args) < 2:
        return CellError.VALUE
    index = ev.value(args[0])
    if isinstance(index, Array):
        spec = freg.lookup("CHOOSE")
        vals = [Arg(index)] + [Arg(ev.value(a)) for a in args[1:]]
        return freg.call(spec, vals, _FnCtx(ev))
    try:
        i = int(to_number(index))
    except XlError as x:
        return x.error
    if i < 1 or i >= len(args):
        return CellError.VALUE
    return ev.eval(args[i])


def _sf_switch(ev: Evaluator, name: str, args):
    if len(args) < 3:
        return CellError.VALUE
    v = scalar(ev.value(args[0]))
    if isinstance(v, CellError):
        return v
    rest = args[1:]
    i = 0
    while i + 1 < len(rest):
        cand = scalar(ev.value(rest[i]))
        if not isinstance(cand, CellError) and compare(v, cand) == 0 and _same_kind(v, cand):
            return ev.eval(rest[i + 1])
        i += 2
    if len(rest) % 2 == 1:
        return ev.eval(rest[-1])
    return CellError.NA


def _same_kind(a: Value, b: Value) -> bool:
    if a is None or b is None:
        return True
    return type(a) is type(b) or (isinstance(a, (int, float)) and not isinstance(a, bool) and isinstance(b, (int, float)) and not isinstance(b, bool))


def _sf_row_col(ev: Evaluator, name: str, args):
    is_row = name == "ROW"
    if not args or isinstance(args[0], Missing):
        return float(ev.at[0] + 1 if is_row else ev.at[1] + 1)
    x = ev.eval(args[0])
    if isinstance(x, RefVal):
        if not x.areas:
            return CellError.REF
        r = ev.trim_lines(x.areas[0])
        if is_row:
            vals = [float(i + 1) for i in range(r.r0, r.r1 + 1)]
            return vals[0] if len(vals) == 1 else Array.column(vals)
        vals = [float(i + 1) for i in range(r.c0, r.c1 + 1)]
        return vals[0] if len(vals) == 1 else Array.row(vals)
    if isinstance(x, CellError):
        return x
    return CellError.VALUE


def _sf_rows_cols(ev: Evaluator, name: str, args):
    if len(args) != 1:
        return CellError.VALUE
    is_rows = name == "ROWS"
    x = ev.eval(args[0])
    if isinstance(x, RefVal):
        if len(x.areas) != 1:
            return CellError.REF
        r = x.areas[0].rng
        return float(r.height if is_rows else r.width)
    if isinstance(x, Array):
        return float(x.rows if is_rows else x.cols)
    if isinstance(x, CellError):
        return x
    if isinstance(x, Lambda):
        return CellError.VALUE
    return 1.0


def _sf_areas(ev: Evaluator, name: str, args):
    if not args:
        return CellError.VALUE
    x = ev.eval(args[0])
    if isinstance(x, RefVal):
        return float(len(x.areas))
    if isinstance(x, CellError):
        return x
    return CellError.VALUE


def _sf_isref(ev: Evaluator, name: str, args):
    if not args:
        return CellError.VALUE
    return isinstance(ev.eval(args[0]), RefVal)


def _sf_formula_info(ev: Evaluator, name: str, args):
    x = ev.eval(args[0]) if args else None
    if not isinstance(x, RefVal) or not x.areas:
        return CellError.NA
    a = x.areas[0]
    sh = ev.host.wb.sheet(a.sheet)
    cell = sh.cell(a.rng.start) if sh is not None else None
    f = cell.formula if cell is not None else None
    if name == "ISFORMULA":
        return f is not None
    if f is None:
        return CellError.NA
    return "=" + f.text


def _sf_sheet(ev: Evaluator, name: str, args):
    wb = ev.host.wb
    if not args:
        return float(ev.sheet + 1 if name == "SHEET" else len(wb.sheets))
    x = ev.eval(args[0])
    if isinstance(x, RefVal):
        if name == "SHEET":
            return float(x.areas[0].sheet + 1) if x.areas else CellError.REF
        return float(len({a.sheet for a in x.areas}))
    if isinstance(x, str) and name == "SHEET":
        i = wb.sheet_index(x)
        return CellError.NA if i is None else float(i + 1)
    return CellError.NA


def _sf_offset(ev: Evaluator, name: str, args):
    ev.host.note_volatile()
    if not 3 <= len(args) <= 5:
        return CellError.VALUE
    base = ev.eval(args[0])
    if isinstance(base, CellError):
        return base
    if not isinstance(base, RefVal) or len(base.areas) != 1:
        return CellError.VALUE
    b = base.areas[0]
    try:
        rows = int(ev.arg_num(args, 1, 0.0))
        cols = int(ev.arg_num(args, 2, 0.0))
        h = int(ev.arg_num(args, 3, float(b.rng.height)))
        w = int(ev.arg_num(args, 4, float(b.rng.width)))
    except XlError as x:
        return x.error
    if h == 0 or w == 0:
        return CellError.REF
    r0 = b.rng.r0 + rows
    c0 = b.rng.c0 + cols
    r0, r1 = (r0, r0 + h - 1) if h > 0 else (r0 + h + 1, r0)
    c0, c1 = (c0, c0 + w - 1) if w > 0 else (c0 + w + 1, c0)
    if r0 < 0 or c0 < 0 or r1 >= MAX_ROWS or c1 >= MAX_COLS:
        return CellError.REF
    return RefVal([Area(b.sheet, RangeRef(r0, c0, r1, c1))])


def _sf_indirect(ev: Evaluator, name: str, args):
    ev.host.note_volatile()
    if not 1 <= len(args) <= 2:
        return CellError.VALUE
    v = scalar(ev.value(args[0]))
    try:
        t = to_text(v)
    except XlError as x:
        return x.error
    a1 = True
    if len(args) > 1 and not isinstance(args[1], Missing):
        try:
            a1 = to_bool(scalar(ev.value(args[1])))
        except XlError:
            a1 = True
    conv = None
    if not a1:
        conv = r1c1_to_a1(t, ev.at)
        if conv is not None:
            t = conv
    try:
        e = parse(t)
    except ParseError:
        return CellError.REF
    if isinstance(e, Ref) and not a1 and conv is None:
        # A1-style text when R1C1 is asked for; a defined name still resolves (Excel 365).
        return CellError.REF
    if isinstance(e, (Ref, Name, StructRef)):
        r = ev.eval(e)
        return r if isinstance(r, RefVal) else CellError.REF
    return CellError.REF


def _sf_index(ev: Evaluator, name: str, args):
    if not 1 <= len(args) <= 4:
        return CellError.VALUE
    src = ev.eval(args[0])

    def opt(i):
        if i >= len(args) or isinstance(args[i], Missing):
            return None
        return ev.value(args[i])

    rows, cols = opt(1), opt(2)
    try:
        area_n = int(ev.arg_num(args, 3, 1.0))
    except XlError as x:
        return x.error
    two = len(args) == 2

    def num(v):
        return 0 if v is None else int(to_number(v))

    if isinstance(rows, Array) or isinstance(cols, Array):
        ra = rows if isinstance(rows, Array) else Array.scalar(rows if rows is not None else 0.0)
        ca = cols if isinstance(cols, Array) else Array.scalar(cols if cols is not None else 0.0)
        h, w = max(ra.rows, ca.rows), max(ra.cols, ca.cols)
        if h * w > MAX_CELLS:
            return CellError.NUM
        out = []
        for r in range(h):
            for c in range(w):
                try:
                    one = _index_one(ev, src, num(ra.get_broadcast(r, c)), num(ca.get_broadcast(r, c)), area_n, two)
                    v = ev.deref(one)
                    if isinstance(v, Array) and v.rows * v.cols > 1:
                        v = CellError.CALC
                    else:
                        v = scalar(v)
                except XlError as x:
                    v = x.error
                out.append(v)
        return Array(h, w, out)
    try:
        return _index_one(ev, src, num(rows), num(cols), area_n, two)
    except XlError as x:
        return x.error


def _index_one(ev: Evaluator, src, row: int, col: int, area_n: int, two_args: bool):
    if row < 0 or col < 0:
        return CellError.VALUE
    if isinstance(src, RefVal):
        if area_n < 1:
            # An area number below 1 (0.5 counts as 0) is #VALUE!; past the last area, #REF! (Excel).
            return CellError.VALUE
        idx = area_n - 1
        if idx >= len(src.areas):
            return CellError.REF
        a = src.areas[idx]
        r = a.rng
        if two_args and r.height == 1:
            row, col = 1, row
        if row > r.height or col > r.width:
            return CellError.REF
        r0, r1 = (r.r0, r.r1) if row == 0 else (r.r0 + row - 1, r.r0 + row - 1)
        c0, c1 = (r.c0, r.c1) if col == 0 else (r.c0 + col - 1, r.c0 + col - 1)
        return RefVal([Area(a.sheet, RangeRef(r0, c0, r1, c1))])
    if isinstance(src, Array):
        if two_args and src.rows == 1:
            row, col = 1, row
        if row > src.rows or col > src.cols:
            return CellError.REF
        if row == 0 and col == 0:
            return src
        if row == 0:
            return Array.column([src.get(r, col - 1) for r in range(src.rows)])
        if col == 0:
            return Array.row([src.get(row - 1, c) for c in range(src.cols)])
        return src.get(row - 1, col - 1, CellError.REF)
    if isinstance(src, CellError):
        return src
    if isinstance(src, Lambda):
        return CellError.VALUE
    if row <= 1 and col <= 1:
        return src
    return CellError.REF


def _sf_xlookup(ev: Evaluator, name: str, args):
    """XLOOKUP whose result range is a reference returns a reference, as in Excel."""
    if 3 <= len(args) <= 6:
        look, ret = ev.eval(args[1]), ev.eval(args[2])
        if isinstance(look, RefVal) and isinstance(ret, RefVal) and len(look.areas) == 1 and len(ret.areas) == 1:
            key = ev.value(args[0])
            l, r = look.areas[0], ret.areas[0]
            vertical = l.rng.width == 1 and l.rng.height > 1
            horizontal = l.rng.height == 1 and l.rng.width > 1
            same_len = r.rng.height == l.rng.height if vertical else r.rng.width == l.rng.width
            if not isinstance(key, Array) and (vertical or horizontal) and same_len:
                xargs = [Arg(key), Arg(ev.area_value(l), True)]

                def given(i):
                    return i < len(args) and not isinstance(args[i], Missing)

                last = 5 if given(5) else (4 if given(4) else 3)
                for i in range(4, last + 1):
                    xargs.append(Arg(0.0) if not given(i) else Arg(ev.value(args[i])))
                spec = freg.lookup("XMATCH")
                p = freg.call(spec, xargs, _FnCtx(ev))
                if isinstance(p, float) and p >= 1:
                    k = int(p) - 1
                    if vertical:
                        rng = RangeRef(r.rng.r0 + k, r.rng.c0, r.rng.r0 + k, r.rng.c1)
                    else:
                        rng = RangeRef(r.rng.r0, r.rng.c0 + k, r.rng.r1, r.rng.c0 + k)
                    return RefVal([Area(r.sheet, rng)])
                if p is CellError.NA and given(3):
                    return ev.eval(args[3])
                if isinstance(p, CellError):
                    return p
                return CellError.VALUE
    return ev.builtin("XLOOKUP", args)


def _sf_cell(ev: Evaluator, name: str, args):
    try:
        kind = to_text(scalar(ev.value(args[0]))).lower() if args else None
    except XlError as x:
        return x.error
    if kind is None:
        return CellError.VALUE
    if len(args) > 1:
        x = ev.eval(args[1])
        if not isinstance(x, RefVal) or not x.areas:
            return CellError.VALUE
        a = x.areas[0]
    else:
        a = Area(ev.sheet, RangeRef(ev.at[0], ev.at[1], ev.at[0], ev.at[1]))
    c = a.rng.start
    wb = ev.host.wb
    sh = wb.sheet(a.sheet)
    if kind == "address":
        return f"${col_to_letters(c[1])}${c[0] + 1}"
    if kind == "row":
        return float(c[0] + 1)
    if kind == "col":
        return float(c[1] + 1)
    if kind == "contents":
        return ev.host.cell_value(a.sheet, c)
    if kind == "type":
        v = ev.host.cell_value(a.sheet, c)
        return "b" if v is None else ("l" if isinstance(v, str) else "v")
    if kind == "width":
        if sh is None:
            return 8.0
        return float(round(max(sh.col_width(c[1]) - 5.0, 0.0) / 7.0))
    if kind == "filename":
        return ""
    if kind == "sheetname":
        return sh.name if sh is not None else ""
    if kind == "protect":
        if sh is None:
            return 0.0
        return 1.0 if wb.styles.get(sh.style_id(c)).protection.locked else 0.0
    if kind == "format":
        code = wb.styles.get(sh.style_id(c)).num_fmt if sh is not None else "General"
        return cell_format_code(code)
    if kind in ("color", "parentheses"):
        return 0.0
    if kind == "prefix":
        return ""
    return CellError.VALUE


def _sf_trimrange(ev: Evaluator, name: str, args):
    if not 1 <= len(args) <= 3:
        return CellError.VALUE
    x = ev.eval(args[0])
    if not isinstance(x, RefVal):
        return x if isinstance(x, CellError) else CellError.VALUE
    try:
        tr = int(ev.arg_num(args, 1, 3.0))
        tc = int(ev.arg_num(args, 2, 3.0))
    except XlError as e:
        return e.error
    if tr not in (0, 1, 2, 3) or tc not in (0, 1, 2, 3):
        return CellError.VALUE
    out = []
    for a in x.areas:
        r = ev._trim_area(a, tr, tc)
        if r is None:
            return CellError.REF
        out.append(Area(a.sheet, r))
    return RefVal(out)


def _sf_subtotal(ev: Evaluator, name: str, args):
    from ..functions.stat import aggregate_values, aggregate_values_k

    if len(args) < (3 if name == "AGGREGATE" else 2):
        return CellError.VALUE
    try:
        code = int(ev.arg_num(args, 0, 0.0))
    except XlError as x:
        return x.error
    if name == "AGGREGATE":
        try:
            options = int(ev.arg_num(args, 1, 0.0))
        except XlError as x:
            return x.error
        func, rest = code, list(args[2:])
    else:
        func, ignore_hidden = (code - 100, True) if code > 100 else (code, False)
        options = 5 if ignore_hidden else 0
        rest = list(args[1:])
    if not 1 <= func <= 19 or (name == "SUBTOTAL" and func > 11):
        return CellError.VALUE
    skip_hidden = options in (1, 3, 5, 7)
    skip_filtered = name == "SUBTOTAL"
    skip_errors = options in (2, 3, 6, 7)
    # SUBTOTAL leaves out nested SUBTOTAL and AGGREGATE cells whatever its function number.
    skip_nested = name == "SUBTOTAL" or options <= 3
    values: List[Value] = []
    k = None
    refs = rest
    if name == "AGGREGATE" and func >= 14:
        refs = rest[:1]
        if len(rest) > 1:
            k = ev.value(rest[1])
    wb = ev.host.wb
    for e in refs:
        x = ev.eval(e)
        if isinstance(x, RefVal):
            for a in x.areas:
                rng = ev.trim(a)
                s = wb.sheet(a.sheet)
                for (r, c) in rng.cells():
                    if s is not None:
                        hidden = s.is_row_hidden(r) if skip_hidden else (skip_filtered and s.is_row_filtered(r))
                        if hidden:
                            continue
                        if skip_nested:
                            cell = s.cell((r, c))
                            if cell is not None and cell.formula is not None:
                                u = cell.formula.text.upper()
                                if "SUBTOTAL(" in u or "AGGREGATE(" in u:
                                    continue
                    v = ev.host.cell_value(a.sheet, (r, c))
                    if skip_errors and isinstance(v, CellError):
                        continue
                    values.append(v)
        elif name == "SUBTOTAL" or func <= 13:
            # These take references only: AGGREGATE(9,0,{1,2}) is #VALUE! (Excel refuses such a
            # SUBTOTAL as it is typed); only the array form (14-19) takes an array.
            return x if isinstance(x, CellError) else CellError.VALUE
        elif isinstance(x, Array):
            values.extend(v for v in x.data if not (skip_errors and isinstance(v, CellError)))
        elif isinstance(x, Lambda):
            return CellError.VALUE
        else:
            values.append(x)
    if k is not None:
        return aggregate_values_k(func, values, k)
    return aggregate_values(func, values)


def _sf_text(ev: Evaluator, name: str, args):
    from ..numfmt import text_function

    if len(args) != 2:
        return CellError.VALUE
    v = ev.value(args[0])
    f = ev.value(args[1])
    sys = ev.host.wb.date_system

    def one(x, fc):
        if isinstance(x, CellError):
            return x
        try:
            return text_function(x, to_text(fc), sys)
        except XlError as e:
            return e.error

    return zip2(v, f, one)


def _sf_let(ev: Evaluator, name: str, args):
    if len(args) < 3 or len(args) % 2 == 0:
        return CellError.VALUE
    base = len(ev.env)
    try:
        for i in range(0, len(args) - 1, 2):
            n = args[i]
            if not isinstance(n, Name):
                return CellError.VALUE
            ev.env.append((n.name, ev.eval(args[i + 1])))
        return ev.eval(args[-1])
    finally:
        del ev.env[base:]


def _sf_lambda(ev: Evaluator, name: str, args):
    if not args:
        return CellError.VALUE
    params = []
    for p in args[:-1]:
        if not isinstance(p, Name):
            return CellError.VALUE
        params.append(p.name.strip("[]"))
    return Lambda(params, args[-1], list(ev.env))


def _sf_isomitted(ev: Evaluator, name: str, args):
    if args and isinstance(args[0], Name):
        n = args[0].name.lower()
        for k, v in reversed(ev.env):
            if k.lower() == n:
                return v is None
    return False


def _as_arr(v: Value) -> Array:
    return v if isinstance(v, Array) else Array.scalar(v)


def _sf_higher_order(ev: Evaluator, name: str, args):
    if name == "MAP":
        if len(args) < 2:
            return CellError.VALUE
        l = ev.lambda_arg(args[-1], len(args) - 1)
        if l is None:
            return CellError.VALUE
        arrays = [_as_arr(ev.value(a)) for a in args[:-1]]
        rows = max(a.rows for a in arrays)
        cols = max(a.cols for a in arrays)
        if rows * cols > MAX_CELLS:
            return CellError.NUM
        out = []
        for r in range(rows):
            for c in range(cols):
                bound = [(p, a.get_broadcast(r, c)) for p, a in zip(l.params, arrays)]
                out.append(scalar(ev.deref(ev.call_lambda_values(l, bound))))
        return Array(rows, cols, out)
    if name in ("REDUCE", "SCAN"):
        if len(args) != 3:
            return CellError.VALUE
        acc = ev.value(args[0])
        arr = _as_arr(ev.value(args[1]))
        l = ev.lambda_arg(args[2], 2)
        if l is None or len(l.params) != 2:
            return CellError.VALUE
        scan = []
        for v in arr.data:
            acc = ev.deref(ev.call_lambda_values(l, [(l.params[0], acc), (l.params[1], v)]))
            if name == "SCAN":
                scan.append(scalar(acc))
        return Array(arr.rows, arr.cols, scan) if name == "SCAN" else acc
    if name in ("BYROW", "BYCOL"):
        if len(args) != 2:
            return CellError.VALUE
        arr = _as_arr(ev.value(args[0]))
        l = ev.lambda_arg(args[1])
        if l is None or not l.params:
            return CellError.VALUE
        by_row = name == "BYROW"
        n = arr.rows if by_row else arr.cols
        out = []
        for i in range(n):
            sl = Array.row(arr.row_list(i)) if by_row else Array.column(arr.col_list(i))
            out.append(scalar(ev.deref(ev.call_lambda_values(l, [(l.params[0], sl)]))))
        return Array.column(out) if by_row else Array.row(out)
    if name == "MAKEARRAY":
        if len(args) != 3:
            return CellError.VALUE
        try:
            rows, cols = int(ev.arg_num(args, 0, 0.0)), int(ev.arg_num(args, 1, 0.0))
        except XlError as x:
            return x.error
        if rows < 1 or cols < 1 or rows * cols > MAX_CELLS:
            return CellError.VALUE
        l = ev.lambda_arg(args[2], 2)
        if l is None or len(l.params) != 2:
            return CellError.VALUE
        out = []
        for r in range(rows):
            for c in range(cols):
                out.append(scalar(ev.deref(ev.call_lambda_values(l, [(l.params[0], float(r + 1)), (l.params[1], float(c + 1))]))))
        return Array(rows, cols, out)
    return CellError.NAME


def _sf_volatile(ev: Evaluator, name: str, args):
    ev.host.note_volatile()
    return ev.builtin(name, args)


_SPECIAL: Dict[str, Callable] = {
    "IF": _sf_if,
    "IFS": _sf_ifs,
    "IFERROR": _sf_iferror,
    "IFNA": _sf_iferror,
    "CHOOSE": _sf_choose,
    "SWITCH": _sf_switch,
    "ROW": _sf_row_col,
    "COLUMN": _sf_row_col,
    "ROWS": _sf_rows_cols,
    "COLUMNS": _sf_rows_cols,
    "AREAS": _sf_areas,
    "ISREF": _sf_isref,
    "ISFORMULA": _sf_formula_info,
    "FORMULATEXT": _sf_formula_info,
    "SHEET": _sf_sheet,
    "SHEETS": _sf_sheet,
    "OFFSET": _sf_offset,
    "INDIRECT": _sf_indirect,
    "INDEX": _sf_index,
    "XLOOKUP": _sf_xlookup,
    "CELL": _sf_cell,
    "TRIMRANGE": _sf_trimrange,
    "SUBTOTAL": _sf_subtotal,
    "AGGREGATE": _sf_subtotal,
    "TEXT": _sf_text,
    "LET": _sf_let,
    "LAMBDA": _sf_lambda,
    "ISOMITTED": _sf_isomitted,
    "MAP": _sf_higher_order,
    "REDUCE": _sf_higher_order,
    "SCAN": _sf_higher_order,
    "BYROW": _sf_higher_order,
    "BYCOL": _sf_higher_order,
    "MAKEARRAY": _sf_higher_order,
    "NOW": _sf_volatile,
    "TODAY": _sf_volatile,
    "RAND": _sf_volatile,
    "RANDBETWEEN": _sf_volatile,
    "RANDARRAY": _sf_volatile,
}

SPECIAL_FUNCTIONS = sorted(k for k in _SPECIAL if k not in ("NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY"))

_DISPATCH = {
    Num: Evaluator._num,
    Str: Evaluator._str,
    Bool: Evaluator._bool,
    Err: Evaluator._errv,
    Missing: Evaluator._missing,
    Paren: Evaluator._paren,
    ArrayLit: Evaluator._array,
    Ref: Evaluator._ref,
    Name: Evaluator._name,
    StructRef: Evaluator._struct,
    Unary: Evaluator._unary,
    Binary: Evaluator._binary,
    Call: Evaluator._call,
    Invoke: Evaluator._invoke,
}

_AGG = {"SUM", "COUNT", "AVERAGE"}


def _agg_value(name: str, v: Value, from_ref: bool):
    """Adds one argument's value by calling SUM and COUNT on it (so coercion rules stay in one
    place); returns ``(sum, count)`` or an error."""
    s = 0.0
    c = 0
    if name != "COUNT":
        r = freg.call(freg.lookup("SUM"), [Arg(v, from_ref)], freg.Ctx())
        if isinstance(r, float):
            s = r
        else:
            return r if isinstance(r, CellError) else CellError.VALUE
    if name != "SUM":
        r = freg.call(freg.lookup("COUNT"), [Arg(v, from_ref)], freg.Ctx())
        if isinstance(r, float):
            c = int(r)
        else:
            return r if isinstance(r, CellError) else CellError.VALUE
    return s, c


def cell_format_code(code: str) -> str:
    """CELL("format") code for a number format."""
    c = code.lower()
    if c == "general":
        return "G"

    def decimals(x: str) -> int:
        if "." not in x:
            return 0
        d = x.split(".", 1)[1]
        n = 0
        for ch in d:
            if ch == "0":
                n += 1
            else:
                break
        return n

    if "%" in c:
        return f"P{decimals(c)}"
    if "e" in c and "0" in c:
        return f"S{decimals(c)}"
    if "$" in c:
        return f"C{decimals(c)}"
    if "y" in c or "d" in c:
        return "D1"
    if "h" in c:
        return "D9"
    if "," in c:
        return f",{decimals(c)}"
    return f"F{decimals(c)}"


def is_worksheet_function(name: str) -> bool:
    """A built-in worksheet function's name (functions are passed by name to GROUPBY, BYROW…)."""
    u = name.upper()
    return freg.lookup(u) is not None or u in _SPECIAL


def r1c1_to_a1(s: str, at: Tuple[int, int]) -> Optional[str]:
    """Converts an R1C1 reference text to A1 (for INDIRECT(…, FALSE))."""
    if "!" in s:
        sh, body = s.rsplit("!", 1)
        sheet = sh + "!"
    else:
        sheet, body = "", s

    def conv(part: str) -> Optional[str]:
        u = part.upper()
        if not u.startswith("R") or "C" not in u:
            return None
        rpart, cpart = u[1:].split("C", 1)

        def num(p: str, base: int) -> Optional[Tuple[int, bool]]:
            if p == "":
                return base, False
            if p.startswith("[") and p.endswith("]"):
                try:
                    d = int(p[1:-1])
                except ValueError:
                    return None
                v = base + d
                return (v, False) if v >= 0 else None
            try:
                v = int(p)
            except ValueError:
                return None
            return (v - 1, True) if v >= 1 else None

        r = num(rpart, at[0])
        c = num(cpart, at[1])
        if r is None or c is None:
            return None
        return f"{'$' if c[1] else ''}{col_to_letters(c[0])}{'$' if r[1] else ''}{r[0] + 1}"

    if ":" in body:
        a, b = body.split(":", 1)
        ca, cb = conv(a), conv(b)
        if ca is None or cb is None:
            return None
        out = f"{ca}:{cb}"
    else:
        out = conv(body)
        if out is None:
            return None
    return sheet + out


# ------------------------------------------------------------------ operators


def _map_unary(v: Value, f: Callable[[float], float], plus: bool, sys=None) -> Value:
    if isinstance(v, Array):
        return Array(v.rows, v.cols, [_map_unary(x, f, plus, sys) for x in v.data])
    if isinstance(v, CellError):
        return v
    if plus and isinstance(v, str):
        return v
    try:
        n = f(to_number_in(v, sys))
    except XlError as x:
        return x.error
    if not math.isfinite(n):
        return CellError.NUM
    return 0.0 if n == 0 else n


def _broadcast_dim(a: int, b: int) -> int:
    if a == 1:
        return b
    if b == 1:
        return a
    return max(a, b)


def zip2(a: Value, b: Value, f: Callable[[Value, Value], Value]) -> Value:
    """Element-wise combination with Excel's broadcasting."""
    if isinstance(a, Array):
        if isinstance(b, Array):
            rows = _broadcast_dim(a.rows, b.rows)
            cols = _broadcast_dim(a.cols, b.cols)
            data = [f(a.get_broadcast(r, c), b.get_broadcast(r, c)) for r in range(rows) for c in range(cols)]
            return Array(rows, cols, data)
        return Array(a.rows, a.cols, [f(x, b) for x in a.data])
    if isinstance(b, Array):
        return Array(b.rows, b.cols, [f(a, y) for y in b.data])
    return f(a, b)


def zip3(c: Array, t: Value, f: Value, g) -> Value:
    def dims(v):
        return (v.rows, v.cols) if isinstance(v, Array) else (1, 1)

    tr, tc = dims(t)
    fr, fc = dims(f)
    rows = _broadcast_dim(_broadcast_dim(c.rows, tr), fr)
    cols = _broadcast_dim(_broadcast_dim(c.cols, tc), fc)

    def get(v, r, k):
        return v.get_broadcast(r, k) if isinstance(v, Array) else v

    data = [g(c.get_broadcast(r, k), get(t, r, k), get(f, r, k)) for r in range(rows) for k in range(cols)]
    return Array(rows, cols, data)


def _num_or_err(v: Value, sys=None):
    t = type(v)
    if t is float:
        return v
    try:
        return to_number_in(v, sys)
    except XlError as x:
        return x.error


def binary_scalar(op: str, a: Value, b: Value, sys=None) -> Value:
    """``a op b`` for scalars; ``sys`` is the date system date text is read in."""
    if isinstance(a, CellError):
        return a
    if isinstance(b, CellError):
        return b
    if op in ("+", "-", "*", "/", "^"):
        x = a if type(a) is float else _num_or_err(a, sys)
        if isinstance(x, CellError):
            return x
        y = b if type(b) is float else _num_or_err(b, sys)
        if isinstance(y, CellError):
            return y
        if op == "+":
            r = x + y
        elif op == "-":
            r = x - y
        elif op == "*":
            r = x * y
        elif op == "/":
            if y == 0:
                return CellError.DIV0
            r = x / y
        else:
            if x == 0 and y == 0:
                return CellError.NUM
            if x == 0 and y < 0:
                return CellError.DIV0
            try:
                r = x ** y
            except (OverflowError, ZeroDivisionError):
                return CellError.NUM
            if isinstance(r, complex):
                return CellError.NUM
        if r != r or r in (math.inf, -math.inf):
            return CellError.NUM
        return 0.0 if r == 0 else r
    if op == "&":
        try:
            s = to_text(a) + to_text(b)
        except XlError as x:
            return x.error
        return CellError.VALUE if len(s) > 32767 * 4 else s
    o = compare(a, b)
    if op == "=":
        return o == 0
    if op == "<>":
        return o != 0
    if op == "<":
        return o < 0
    if op == ">":
        return o > 0
    if op == "<=":
        return o <= 0
    if op == ">=":
        return o >= 0
    return CellError.VALUE


def binary_values(op: str, x: Value, y: Value, sys=None) -> Value:
    if not isinstance(x, Array) and not isinstance(y, Array):
        return binary_scalar(op, x, y, sys)
    return zip2(x, y, lambda a, b: binary_scalar(op, a, b, sys))


def _ev_shape(ev) -> Optional[Tuple[int, int]]:
    if isinstance(ev, RefVal):
        if len(ev.areas) == 1:
            r = ev.areas[0].rng
            return r.height, r.width
        return None
    if isinstance(ev, Array):
        return ev.rows, ev.cols
    if isinstance(ev, Lambda):
        return None
    return 1, 1


def resized_sum_range(sum_rng: RangeRef, height: int, width: int) -> RangeRef:
    """The range SUMIF/AVERAGEIF add up: from the sum range's top-left cell, as many rows and
    columns as the criteria range."""
    row = min(sum_rng.r0 + max(height - 1, 0), MAX_ROWS - 1)
    col = min(sum_rng.c0 + max(width - 1, 0), MAX_COLS - 1)
    return RangeRef(sum_rng.r0, sum_rng.c0, row, col)


_DYNAMIC_FNS = {"INDIRECT", "OFFSET", "NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "CELL", "INFO"}


def precedents(wb, sheet: int, e: Expr) -> Tuple[List[Area], bool]:
    """The static references of an expression with their sheets resolved; ``dynamic`` is set when
    the formula uses INDIRECT/OFFSET/names/tables whose targets can change."""
    out: List[Area] = []
    dynamic = False

    def push(r: Ref, rng: RangeRef) -> None:
        s = r.sheet
        if s is None:
            out.append(Area(sheet, rng))
        elif isinstance(s, str):
            i = wb.sheet_index(s)
            if i is not None:
                out.append(Area(i, rng))
        else:
            i, j = wb.sheet_index(s[0]), wb.sheet_index(s[1])
            if i is not None and j is not None:
                for k in range(min(i, j), max(i, j) + 1):
                    out.append(Area(k, rng))

    for x in e.walk():
        t = type(x)
        if t is Ref:
            push(x, x.range())
        elif t is Call:
            n = x.name
            if len(x.args) == 3 and n in ("SUMIF", "AVERAGEIF"):
                c, sr = x.args[0], x.args[2]
                if isinstance(c, Ref) and isinstance(sr, Ref):
                    crit, srng = c.range(), sr.range()
                    push(sr, resized_sum_range(srng, crit.height, crit.width))
                elif not isinstance(sr, (Num, Str, Bool, Err, ArrayLit, Missing)):
                    dynamic = True
            if n in _DYNAMIC_FNS:
                dynamic = True
            elif n not in _SPECIAL and freg.lookup(n) is None:
                dynamic = True
            else:
                spec = freg.lookup(n)
                if spec is not None and spec.volatile:
                    dynamic = True
        elif t is Name or t is StructRef:
            dynamic = True
        elif t is Binary and x.op == ":" and not (isinstance(x.left, Ref) and isinstance(x.right, Ref)):
            inner = precedents(wb, sheet, x.left)[0] + precedents(wb, sheet, x.right)[0]
            boxes: Dict[int, RangeRef] = {}
            for a in inner:
                boxes[a.sheet] = a.rng if a.sheet not in boxes else boxes[a.sheet].union(a.rng)
            out.extend(Area(s, r) for s, r in boxes.items())
    return out, dynamic


def is_known_function(n: str) -> bool:
    u = n.upper()
    return u in _SPECIAL or freg.lookup(u) is not None
