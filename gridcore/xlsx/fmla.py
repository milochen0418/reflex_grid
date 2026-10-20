"""Formula text conversion between the model (canonical text, no ``=``) and files (``_xlfn.``
prefixes for newer functions, ``_xlpm.`` LET/LAMBDA parameters, ``ANCHORARRAY``/``SINGLE``)."""

from __future__ import annotations

from typing import List, Optional

from ..formula import Call, Err, Expr, Invoke, Name, Num, ParseError, Paren, Unary, Binary, parse, print_formula
from ..formula.ast import StructRef
from ..model.cell import Formula
from .tables import file_function_name

PARAM_PREFIX = "_xlpm."


def to_file(f: Formula, table: Optional[str] = None) -> str:
    """``table``: the table holding the cell, which unqualified structured references name."""
    e = f.get_expr()
    return expr_to_file(e, table) if e is not None else f.text


def text_to_file(text: str, table: Optional[str] = None) -> str:
    """Arbitrary formula text (names, CF, validation) for a file. Unparseable text is kept."""
    body = text[1:] if text.startswith("=") else text
    try:
        return expr_to_file(parse(body), table)
    except ParseError:
        return body


def struct_ref_file(s: StructRef, table: Optional[str]) -> str:
    """A structured reference as files spell it: always the table name, never ``@``
    (``Table1[[#This Row],[Units]]``)."""

    def col(c: str) -> str:
        return "[" + c.replace("'", "''").replace("[", "'[").replace("]", "']").replace("#", "'#") + "]"

    name = s.table or table or ""
    parts = [f"[{i}]" for i in s.specifiers]
    if s.col_start is not None and s.col_end is not None:
        parts.append(col(s.col_start) + ":" + col(s.col_end))
    elif s.col_start is not None:
        parts.append(col(s.col_start))
    if not parts:
        return name + "[]"
    if len(parts) == 1 and not (s.col_start is not None and s.col_start != s.col_start.strip()):
        return name + parts[0]
    return name + "[" + ",".join(parts) + "]"


def unqualify_own_table(e: Expr, table: str) -> Expr:
    """``Table1[@Units]`` inside Table1 reads as ``[@Units]``, as Excel shows it."""
    tl = table.lower()

    def f(x: Expr) -> Expr:
        if isinstance(x, StructRef) and x.table.lower() == tl and x.specifiers == ("#This Row",):
            return StructRef("", x.specifiers, x.col_start, x.col_end)
        return x

    return e.map(f)


def _prefix_params(e: Expr, scope: List[str]) -> Expr:
    sl = [s.lower() for s in scope]
    if isinstance(e, Name) and e.name.lower() in sl:
        return Name(PARAM_PREFIX + e.name)
    if isinstance(e, Unary):
        return Unary(e.op, _prefix_params(e.expr, scope))
    if isinstance(e, Paren):
        return Paren(_prefix_params(e.expr, scope))
    if isinstance(e, Binary):
        return Binary(e.op, _prefix_params(e.left, scope), _prefix_params(e.right, scope))
    if isinstance(e, Invoke):
        return Invoke(_prefix_params(e.callee, scope), tuple(_prefix_params(a, scope) for a in e.args))
    if isinstance(e, Call) and e.name in ("LET", "LAMBDA"):
        last = len(e.args) - 1
        is_let = e.name == "LET"
        sc = list(scope)
        pending = None
        out: List[Expr] = []
        for i, a in enumerate(e.args):
            declares = i < last and (not is_let or i % 2 == 0)
            if declares and isinstance(a, Name):
                out.append(Name(PARAM_PREFIX + a.name))
                if is_let:
                    pending = a.name
                else:
                    sc.append(a.name)
            else:
                out.append(_prefix_params(a, sc))
                if pending is not None:
                    sc.append(pending)
                    pending = None
        return Call(e.name, tuple(out))
    if isinstance(e, Call) and e.name.lower() in sl:
        return Call(PARAM_PREFIX + e.name, tuple(_prefix_params(a, scope) for a in e.args))
    if isinstance(e, Call):
        return Call(e.name, tuple(_prefix_params(a, scope) for a in e.args))
    return e


# Where a function takes a function: files write a worksheet function passed by name with
# ``_xleta.`` (``GROUPBY(A2:A9,B2:B9,_xleta.SUM)``); -1 is the last argument.
_FUNCTION_ARGS = {"BYROW": 1, "BYCOL": 1, "MAP": -1, "REDUCE": 2, "SCAN": 2, "MAKEARRAY": 2, "GROUPBY": 2, "PIVOTBY": 3}
ETA_PREFIX = "_xleta."


def _eta_args(x: Call) -> tuple:
    from ..calc.eval import is_worksheet_function

    pos = _FUNCTION_ARGS.get(x.name)
    if pos is None or not x.args:
        return x.args
    i = pos if pos >= 0 else len(x.args) + pos
    a = x.args[i] if 0 <= i < len(x.args) else None
    if not isinstance(a, Name) or a.name.lower().startswith(PARAM_PREFIX) or not is_worksheet_function(a.name):
        return x.args
    return x.args[:i] + (Name(ETA_PREFIX + a.name.upper()),) + x.args[i + 1:]


def expr_to_file(e: Expr, table: Optional[str] = None) -> str:
    e = _prefix_params(e, [])

    def f(x: Expr) -> Expr:
        if isinstance(x, StructRef):
            return Name(struct_ref_file(x, table))
        if isinstance(x, Call):
            n = file_function_name(x.name)
            args = _eta_args(x)
            return Call(n if n is not None else x.name, args) if n is not None or args is not x.args else x
        if isinstance(x, Unary):
            if x.op == "#":
                if isinstance(x.expr, Err):
                    return x.expr
                return Call("_xlfn.ANCHORARRAY", (x.expr,))
            if x.op == "@":
                return Call("_xlfn.SINGLE", (x.expr,))
            if x.op in (".:", ":.", ".:."):
                code = Num(float({".:": 1, ":.": 2, ".:.": 3}[x.op]))
                return Call("_xlfn.TRIMRANGE", (x.expr, code, code))
        return x

    return print_formula(e.map(f))


def _strip_param_prefix(n: str) -> str:
    if n[: len(PARAM_PREFIX)].lower() == PARAM_PREFIX:
        return n[len(PARAM_PREFIX):]
    return n


def from_file(text: str) -> str:
    """Formula text read from a file → model text (prefixes stripped when it parses)."""
    body = text[1:] if text.startswith("=") else text
    try:
        e = parse(body)
    except ParseError:
        return body

    def f(x: Expr) -> Expr:
        if isinstance(x, Name):
            if x.name[: len(ETA_PREFIX)].lower() == ETA_PREFIX:
                return Name(x.name[len(ETA_PREFIX):])
            return Name(_strip_param_prefix(x.name))
        if isinstance(x, Call):
            return Call(_strip_param_prefix(x.name), x.args)
        return x

    return print_formula(e.map(f))
