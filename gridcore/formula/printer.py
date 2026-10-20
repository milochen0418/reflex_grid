"""Expr → formula text (without ``=``), A1 or R1C1."""

from __future__ import annotations

from typing import List, Optional, Tuple

from ..core.addr import col_to_letters, parse_cell
from ..core.value import number_to_text
from .ast import (
    TRIM_MODES,
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
    Paren,
    Ref,
    Str,
    StructRef,
    Unary,
)


def needs_quotes(name: str) -> bool:
    """Whether a sheet name needs quotes in a reference."""
    if not name:
        return True
    if name[0].isdigit() or any(not (c.isalnum() or c in "_.") for c in name):
        return True
    if parse_cell(name) is not None or name.upper() in ("TRUE", "FALSE"):
        return True
    # R1C1-looking names (R, C, R1C1) are ambiguous too.
    u = name.upper()
    if u in ("R", "C") or (u.startswith("R") and "C" in u and u[1:].replace("C", "", 1).isdigit()):
        return True
    return False


def quote_sheet(name: str) -> str:
    if needs_quotes(name):
        return "'" + name.replace("'", "''") + "'"
    return name


def sheet_prefix(s) -> str:
    if s is None:
        return ""
    if isinstance(s, str):
        return quote_sheet(s) + "!"
    a, b = s
    joined = f"{a}:{b}"
    if needs_quotes(a) or needs_quotes(b):
        return "'" + joined.replace("'", "''") + "'!"
    return joined + "!"


def _anchor_a1(row: int, col: int, row_abs: bool, col_abs: bool) -> str:
    return f"{'$' if col_abs else ''}{col_to_letters(col)}{'$' if row_abs else ''}{row + 1}"


def range_a1(r: Ref, sep: str = ":") -> str:
    k = r.kind
    if k == "cell":
        body = _anchor_a1(r.r0, r.c0, r.r0_abs, r.c0_abs)
    elif k == "range":
        body = _anchor_a1(r.r0, r.c0, r.r0_abs, r.c0_abs) + sep + _anchor_a1(r.r1, r.c1, r.r1_abs, r.c1_abs)
    elif k == "rows":
        body = f"{'$' if r.r0_abs else ''}{r.r0 + 1}{sep}{'$' if r.r1_abs else ''}{r.r1 + 1}"
    else:
        body = f"{'$' if r.c0_abs else ''}{col_to_letters(r.c0)}{sep}{'$' if r.c1_abs else ''}{col_to_letters(r.c1)}"
    return sheet_prefix(r.sheet) + body


def reference_a1(r: Ref) -> str:
    return range_a1(r, ":")


def _r1c1_part(letter: str, v: int, absolute: bool, base: int) -> str:
    if absolute:
        return f"{letter}{v + 1}"
    d = v - base
    return letter if d == 0 else f"{letter}[{d}]"


def range_r1c1(r: Ref, at: Tuple[int, int], sep: str = ":") -> str:
    k = r.kind
    if k == "cell":
        body = _r1c1_part("R", r.r0, r.r0_abs, at[0]) + _r1c1_part("C", r.c0, r.c0_abs, at[1])
    elif k == "range":
        a = _r1c1_part("R", r.r0, r.r0_abs, at[0]) + _r1c1_part("C", r.c0, r.c0_abs, at[1])
        b = _r1c1_part("R", r.r1, r.r1_abs, at[0]) + _r1c1_part("C", r.c1, r.c1_abs, at[1])
        body = a + sep + b
    elif k == "rows":
        x, y = _r1c1_part("R", r.r0, r.r0_abs, at[0]), _r1c1_part("R", r.r1, r.r1_abs, at[0])
        body = x if (x == y and sep == ":") else x + sep + y
    else:
        x, y = _r1c1_part("C", r.c0, r.c0_abs, at[1]), _r1c1_part("C", r.c1, r.c1_abs, at[1])
        body = x if (x == y and sep == ":") else x + sep + y
    return sheet_prefix(r.sheet) + body


def struct_ref(s: StructRef) -> str:
    def esc(c: str) -> str:
        return c.replace("'", "''").replace("[", "'[").replace("]", "']").replace("#", "'#")

    def simple_col(c: str) -> bool:
        return not any(ch in " []#',:" for ch in c)

    this_row = s.specifiers == ("#This Row",)
    out = s.table
    if s.col_start is None and s.col_end is None and not s.specifiers:
        return out + "[]"
    if s.col_start is None and s.col_end is None and this_row:
        return out + "[@]"
    if s.col_start is None and s.col_end is None and len(s.specifiers) == 1:
        return out + f"[{s.specifiers[0]}]"
    if s.col_start is not None and s.col_end is None and not s.specifiers:
        if s.col_start != s.col_start.strip():
            # Spaces at the edges of a column name: the double-bracket form, as Excel writes it.
            return out + f"[[{esc(s.col_start)}]]"
        return out + f"[{esc(s.col_start)}]"
    if s.col_start is not None and s.col_end is None and this_row:
        if simple_col(s.col_start):
            return out + f"[@{esc(s.col_start)}]"
        return out + f"[@[{esc(s.col_start)}]]"
    parts = [f"[{i}]" for i in s.specifiers]
    if s.col_start is not None and s.col_end is not None:
        parts.append(f"[{esc(s.col_start)}]:[{esc(s.col_end)}]")
    elif s.col_start is not None:
        parts.append(f"[{esc(s.col_start)}]")
    return out + "[" + ",".join(parts) + "]"


def to_text(e: Expr, r1c1: Optional[Tuple[int, int]] = None) -> str:
    """Prints an expression in A1 notation (or R1C1 relative to ``r1c1``)."""
    out: List[str] = []
    _write(out, e, r1c1)
    return "".join(out)


def print_a1(e: Expr) -> str:
    return to_text(e)


def _write(out: List[str], e: Expr, r1c1) -> None:
    t = type(e)
    if t is Num:
        out.append(number_to_text(e.value))
    elif t is Str:
        out.append('"' + e.value.replace('"', '""') + '"')
    elif t is Bool:
        out.append("TRUE" if e.value else "FALSE")
    elif t is Err:
        out.append(e.error.value)
    elif t is ArrayLit:
        out.append("{")
        for ri, row in enumerate(e.rows):
            if ri:
                out.append(";")
            for ci, el in enumerate(row):
                if ci:
                    out.append(",")
                _write(out, el, r1c1)
        out.append("}")
    elif t is Ref:
        out.append(range_r1c1(e, r1c1) if r1c1 is not None else range_a1(e))
    elif t is Name:
        n = e.name
        if "!" in n:
            # A sheet-scoped name keeps its sheet bare (`My Sheet!Total`): quote it as references do.
            sh, bare = n.rsplit("!", 1)
            if len(sh) > 1 and sh.startswith("'") and sh.endswith("'"):
                sh = sh[1:-1].replace("''", "'")
            out.append(quote_sheet(sh) + "!" + bare)
        else:
            out.append(n)
    elif t is StructRef:
        out.append(struct_ref(e))
    elif t is Unary:
        op = e.op
        x = e.expr
        if op in ("-", "+", "@"):
            out.append(op)
            _write(out, x, r1c1)
        elif op == "%":
            _write(out, x, r1c1)
            out.append("%")
        elif op == "#":
            _write(out, x, r1c1)
            if not isinstance(x, Err):
                out.append("#")
        else:  # trim reference
            if isinstance(x, Ref) and x.kind != "cell":
                out.append(range_r1c1(x, r1c1, op) if r1c1 is not None else range_a1(x, op))
            elif isinstance(x, Binary) and x.op == ":":
                _write(out, x.left, r1c1)
                out.append(op)
                _write(out, x.right, r1c1)
            elif isinstance(x, Err):
                _write(out, x, r1c1)
            else:
                code = TRIM_MODES.get(op, 3)
                out.append("TRIMRANGE(")
                _write(out, x, r1c1)
                out.append(f",{code},{code})")
    elif t is Binary:
        _write(out, e.left, r1c1)
        out.append(e.op)
        _write(out, e.right, r1c1)
    elif t is Call:
        out.append("_xludf." + e.name[7:] if e.name.startswith("_XLUDF.") else e.name)
        out.append("(")
        for i, a in enumerate(e.args):
            if i:
                out.append(",")
            _write(out, a, r1c1)
        out.append(")")
    elif t is Invoke:
        _write(out, e.callee, r1c1)
        out.append("(")
        for i, a in enumerate(e.args):
            if i:
                out.append(",")
            _write(out, a, r1c1)
        out.append(")")
    elif t is Missing:
        pass
    elif t is Paren:
        out.append("(")
        _write(out, e.expr, r1c1)
        out.append(")")
    else:  # pragma: no cover
        raise TypeError(f"unknown node {e!r}")
