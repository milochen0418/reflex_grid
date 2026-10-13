"""Export-time implicit-intersection analysis, without evaluating the workbook.

A range is harmless in SUM, but not in SQRT or arithmetic. Formulas whose results may be arrays,
or that apply scalar operations to ranges, are written as dynamic-array formulas (``cm="1"``)
so Excel 365 evaluates them with array semantics instead of adding ``@``."""

from __future__ import annotations

from typing import List, Optional, Tuple

from ..formula import ArrayLit, Binary, Call, Expr, Invoke, Name, Num, ParseError, Paren, Ref, StructRef, Unary, parse

_ARRAY_RESULT = {
    "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY", "VSTACK", "HSTACK", "TOROW", "TOCOL",
    "WRAPROWS", "WRAPCOLS", "TAKE", "DROP", "CHOOSEROWS", "CHOOSECOLS", "EXPAND", "GROUPBY", "PIVOTBY", "TRANSPOSE",
    "MMULT", "MINVERSE", "FREQUENCY", "LINEST", "LOGEST", "TREND", "GROWTH", "MAP", "SCAN", "BYROW", "BYCOL",
    "MAKEARRAY", "ANCHORARRAY", "TEXTSPLIT", "MODE.MULT", "MUNIT", "REGEXEXTRACT",
}
_LEGACY_ARRAY = {"SUMPRODUCT", "MMULT", "MINVERSE", "MDETERM", "TRANSPOSE", "FREQUENCY", "LINEST", "LOGEST", "TREND", "GROWTH"}

_RANGE_ARGS_ALL = {
    "SUM", "AVERAGE", "AVERAGEA", "MIN", "MINA", "MAX", "MAXA", "COUNT", "COUNTA", "COUNTBLANK", "PRODUCT", "SUMSQ",
    "MEDIAN", "MODE", "MODE.SNGL", "STDEV", "STDEV.S", "STDEV.P", "STDEVP", "VAR", "VAR.S", "VAR.P", "VARP", "AVEDEV",
    "DEVSQ", "GEOMEAN", "HARMEAN", "KURT", "SKEW", "SKEW.P", "AND", "OR", "XOR", "SUMPRODUCT", "SUMX2MY2", "SUMX2PY2",
    "SUMXMY2", "CORREL", "COVAR", "COVARIANCE.P", "COVARIANCE.S", "PEARSON", "RSQ", "SLOPE", "INTERCEPT", "STEYX",
    "ROWS", "COLUMNS", "AREAS", "ROW", "COLUMN", "TRANSPOSE", "ISREF", "MMULT", "MINVERSE", "MDETERM", "FREQUENCY",
    "CONCAT", "VSTACK", "HSTACK", "ANCHORARRAY", "STDEVA", "STDEVPA", "VARA", "VARPA",
}


def _scalar_arg(name: str, i: int) -> bool:
    """Legacy reference/array-taking parameters; other parameters consume a single value."""
    if name in _RANGE_ARGS_ALL:
        return False
    if name in ("SUMIF", "AVERAGEIF"):
        return i == 1
    if name == "COUNTIF":
        return i != 0
    if name in ("SUMIFS", "AVERAGEIFS", "MINIFS", "MAXIFS"):
        return i > 0 and i % 2 == 0
    if name == "COUNTIFS":
        return i % 2 == 1
    if name in ("INDEX", "LARGE", "SMALL", "PERCENTILE", "PERCENTILE.INC", "PERCENTILE.EXC", "QUARTILE",
                "QUARTILE.INC", "QUARTILE.EXC", "IRR", "MIRR"):
        return i != 0
    if name.startswith("RANK"):
        return i != 1
    if name in ("VLOOKUP", "HLOOKUP", "MATCH", "XMATCH"):
        return i != 1
    if name in ("LOOKUP", "NPV", "XNPV", "SUBTOTAL"):
        return i == 0
    if name == "XLOOKUP":
        return i not in (1, 2)
    if name == "XIRR":
        return i > 1
    if name == "OFFSET":
        return i != 0
    if name in ("AGGREGATE", "TEXTJOIN"):
        return i < 2
    if name in ("FILTER", "SORTBY"):
        return i > 1
    if name in ("SORT", "UNIQUE", "TAKE", "DROP", "CHOOSECOLS", "CHOOSEROWS", "TOROW", "TOCOL", "WRAPROWS", "WRAPCOLS", "EXPAND"):
        return i != 0
    if name in ("LINEST", "LOGEST", "TREND", "GROWTH"):
        return i > 1
    if name in ("FORECAST", "FORECAST.LINEAR"):
        return i == 0
    return True


class _Scan:
    def __init__(self, wb, at: Tuple[int, int]):
        self.wb = wb
        self.at = at
        self.remaining = 4096
        self.locals: List[Tuple[str, Tuple[bool, bool]]] = []

    def expr(self, e: Expr, sheet: int, ctx: bool, depth: int) -> Tuple[bool, bool]:
        """(array, affected): whether the result may be an array, and whether a scalar operation
        is applied to an array outside an array context."""
        if depth >= 64 or self.remaining == 0:
            return True, True
        self.remaining -= 1
        if isinstance(e, Ref):
            return (not e.range().is_single()) or isinstance(e.sheet, tuple), False
        if isinstance(e, ArrayLit):
            return len(e.rows) > 1 or (bool(e.rows) and len(e.rows[0]) > 1), False
        if isinstance(e, Name):
            nl = e.name.lower()
            for k, shape in reversed(self.locals):
                if k.lower() == nl:
                    return shape
            if "!" in e.name:
                s, bare = e.name.rsplit("!", 1)
                si = self.wb.sheet_index(s.strip("'").replace("''", "'"))
                scope = si if si is not None else sheet
            else:
                scope, bare = sheet, e.name
            d = self.wb.name(bare, scope)
            if d is not None:
                try:
                    x = parse(d.formula.lstrip("="))
                except ParseError:
                    return True, False
                return self.expr(x, d.scope if d.scope is not None else scope, ctx, depth + 1)
            return self.wb.table(bare) is not None, False
        if isinstance(e, StructRef):
            one_row = any(s in ("#This Row", "#Headers", "#Totals") for s in e.specifiers)
            if e.col_start is not None and e.col_end is not None:
                one_col = e.col_start.lower() == e.col_end.lower()
            elif e.col_start is not None:
                one_col = True
            else:
                t = None
                if e.table:
                    ti = self.wb.table(e.table)
                    if ti is not None:
                        t = self.wb.sheets[ti[0]].tables[ti[1]]
                one_col = t is not None and len(t.columns) == 1
            return not (one_row and one_col), False
        if isinstance(e, Paren):
            return self.expr(e.expr, sheet, ctx, depth + 1)
        if isinstance(e, Unary):
            a, aff = self.expr(e.expr, sheet, ctx, depth + 1)
            if e.op == "@":
                return False, aff
            if e.op in ("#", ".:", ":.", ".:."):
                return True, aff
            return a, aff or (a and not ctx)
        if isinstance(e, Binary):
            a1, f1 = self.expr(e.left, sheet, ctx, depth + 1)
            a2, f2 = self.expr(e.right, sheet, ctx, depth + 1)
            reference = e.op in (":", ",", " ")
            arr = reference or a1 or a2
            return arr, f1 or f2 or (not reference and arr and not ctx)
        if isinstance(e, Call):
            return self.call(e.name, e.args, sheet, ctx, depth + 1)
        if isinstance(e, Invoke):
            return True, False
        return False, False

    def call(self, name: str, args, sheet: int, ctx: bool, depth: int) -> Tuple[bool, bool]:
        if name == "LET" and len(args) >= 3:
            saved = len(self.locals)
            affected = False
            for i in range(0, len(args) - 1, 2):
                n = args[i]
                if isinstance(n, Name):
                    shape = self.expr(args[i + 1], sheet, ctx, depth)
                    affected = affected or shape[1]
                    self.locals.append((n.name, shape))
            arr, aff = self.expr(args[-1], sheet, ctx, depth)
            del self.locals[saved:]
            return arr, aff or affected
        legacy = name in _LEGACY_ARRAY
        affected = False
        lifted = False
        shapes = []
        if len(args) > 255:
            return True, True
        for i, a in enumerate(args):
            c = ctx or legacy or (name == "INDEX" and i == 0)
            arr, aff = self.expr(a, sheet, c, depth)
            affected = affected or aff
            if arr and _scalar_arg(name, i):
                lifted = True
                affected = affected or not c
            shapes.append(arr)
        first = bool(shapes) and shapes[0]
        if name == "INDIRECT":
            a0 = args[0] if args else None
            if a0 is not None and type(a0).__name__ == "Str" and not (len(args) > 1 and type(args[1]).__name__ == "Bool" and not args[1].value):
                try:
                    arr = self.expr(parse(a0.value), sheet, ctx, depth)[0]
                except ParseError:
                    arr = True
            else:
                arr = True
        elif name == "OFFSET":
            if len(args) <= 3:
                arr = first
            else:
                h = args[3] if len(args) > 3 else None
                w = args[4] if len(args) > 4 else None
                one = lambda x: isinstance(x, Num) and x.value == 1.0  # noqa: E731
                arr = (not one(h)) or (not one(w) if w is not None else first)
        elif name == "INDEX":
            row = args[1] if len(args) > 1 else None
            col = args[2] if len(args) > 2 else None
            pos = lambda x: isinstance(x, Num) and x.value >= 1.0  # noqa: E731
            single_col = bool(args) and isinstance(args[0], Ref) and args[0].range().width == 1
            arr = (not pos(row)) or (col is not None and not pos(col)) or (col is None and not single_col)
        elif name in ("ROW", "COLUMN"):
            affected = affected or (first and not ctx)
            arr = first
        elif name in ("IF", "IFERROR", "IFNA", "IFS", "SWITCH", "CHOOSE"):
            arr = any(shapes)
        elif name == "XLOOKUP":
            if lifted:
                arr = True
            elif len(args) > 2 and isinstance(args[1], Ref) and isinstance(args[2], Ref):
                lk, rs = args[1].range(), args[2].range()
                arr = rs.width > 1 if lk.width == 1 else rs.height > 1
            else:
                arr = True
        elif name in _ARRAY_RESULT:
            arr = True
        else:
            arr = lifted or self.wb.name(name, sheet) is not None
        return arr, affected


def needs_array(wb, sheet: int, at: Tuple[int, int], formula) -> bool:
    e = formula.get_expr()
    if e is None:
        return False
    arr, affected = _Scan(wb, at).expr(e, sheet, False, 0)
    has_at = any(isinstance(x, Unary) and x.op == "@" for x in e.walk())
    return arr or affected or has_at
