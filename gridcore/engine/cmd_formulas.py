"""Formulas commands: AutoSum, defined names, the function list, calculation, auditing and
evaluating formulas."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

from ..calc import evaluate
from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef, a1, letters_to_col, parse_cell
from ..core.value import Array, CellError, display
from ..formula import ParseError, Ref, parse, print_formula
from ..formula.printer import quote_sheet
from ..functions import all_functions
from ..model.features import DefinedName
from .cmd_edit import cell_set, _json_value
from .selection import Selection
from .session import EngineError, Session, command, edit, need, p_bool, p_str, param_cell, param_sheet

AUTOSUM_FUNCTIONS = ("SUM", "AVERAGE", "COUNT", "MAX", "MIN")


def _is_num(sh, rc) -> bool:
    return isinstance(sh.value(rc), float)


def guess_sum_range(sh, at) -> Optional[RangeRef]:
    """The numbers AutoSum proposes for a single cell: the run above it, else the run left of it."""
    r, c = at
    if r > 0 and _is_num(sh, (r - 1, c)):
        top = r - 1
        while top > 0 and _is_num(sh, (top - 1, c)):
            top -= 1
        return RangeRef(top, c, r - 1, c)
    if c > 0 and _is_num(sh, (r, c - 1)):
        left = c - 1
        while left > 0 and _is_num(sh, (r, left - 1)):
            left -= 1
        return RangeRef(r, left, r, c - 1)
    # A header right above, then numbers: still sum the numbers above the header? Excel stops.
    return None


@command("formulas.autoSum", "AutoSum", menu=("Home", "Editing"), shortcut="Alt+=", params="{function?: SUM|AVERAGE|COUNT|MAX|MIN, preview?: bool}")
def autosum(s: Session, p: Dict[str, Any]):
    fn = (p_str(p, "function", "SUM") or "SUM").upper()
    if fn not in AUTOSUM_FUNCTIONS:
        raise EngineError(f"formulas.autoSum: unknown function `{fn}`")
    d = s.doc()
    sh = d.sheet
    sel = d.selection.current()
    merged = sh.merge_at(sel.start)
    if sel.is_single() or (merged is not None and merged == sel):
        at = sel.start
        rng = guess_sum_range(sh, at)
        formula = f"={fn}({rng.a1() if rng is not None else ''})"
        if p_bool(p, "preview"):
            return {"cell": a1(*at), "formula": formula, "range": rng.a1() if rng is not None else None}
        if rng is None:
            return {"cell": a1(*at), "formula": formula, "edit": True}
        cell_set(s, {"cell": a1(*at), "input": formula})
        return {"cell": a1(*at), "formula": formula}

    def empty(rc) -> bool:
        return sh.value(rc) is None

    writes = []
    last_row_empty = sel.height > 1 and all(empty((sel.r1, c)) for c in range(sel.c0, sel.c1 + 1))
    last_col_empty = sel.width > 1 and all(empty((r, sel.c1)) for r in range(sel.r0, sel.r1 + 1))
    if last_row_empty or last_col_empty:
        r_end = sel.r1 - 1 if last_row_empty else sel.r1
        c_end = sel.c1 - 1 if last_col_empty else sel.c1
        if last_row_empty:
            for c in range(sel.c0, c_end + 1):
                writes.append(((sel.r1, c), RangeRef(sel.r0, c, r_end, c)))
        if last_col_empty:
            for r in range(sel.r0, r_end + 1):
                writes.append(((r, sel.c1), RangeRef(r, sel.c0, r, c_end)))
        if last_row_empty and last_col_empty:
            writes.append(((sel.r1, sel.c1), RangeRef(sel.r0, sel.c0, r_end, c_end)))
    elif sel.height > 1:
        if sel.r1 + 1 >= MAX_ROWS:
            raise EngineError("There's no room below the selection.")
        for c in range(sel.c0, sel.c1 + 1):
            writes.append(((sel.r1 + 1, c), RangeRef(sel.r0, c, sel.r1, c)))
    else:
        if sel.c1 + 1 >= MAX_COLS:
            raise EngineError("There's no room right of the selection.")
        writes.append(((sel.r0, sel.c1 + 1), RangeRef(sel.r0, sel.c0, sel.r0, sel.c1)))
    if len(writes) > 16384:
        raise EngineError("The selection is too large.")
    from .entry import input_to_cell

    def f(ctx):
        wb = ctx.wb
        sm = ctx.sheet_mut()
        for rc, rng in writes:
            cell = input_to_cell(wb, sm, rc, f"={fn}({rng.a1()})")
            sm.set_cell(rc, cell)
            ctx.touch(wb.active_sheet, rc)

    edit(s, f)
    return {"cells": [a1(*rc) for rc, _ in writes]}


# ---------------------------------------------------------------- names

_NAME_RE = re.compile(r"^[A-Za-z_\\À-￿][A-Za-z0-9_.\\?À-￿]*$")


def check_name(name: str) -> Optional[str]:
    if not name or len(name) > 255:
        return "The name that you entered is not valid."
    if not _NAME_RE.match(name):
        return "The name that you entered is not valid."
    u = name.upper()
    if u in ("TRUE", "FALSE", "R", "C"):
        return "The name that you entered is not valid."
    if parse_cell(name) is not None:
        return "The name that you entered is not valid."
    if re.match(r"^[Rr]\d*[Cc]\d*$", name) or re.match(r"^[Rr]\d+$", name) or re.match(r"^[Cc]\d+$", name):
        return "The name that you entered is not valid."
    m = re.match(r"^([A-Za-z]{1,3})(\d+)$", name)
    if m and letters_to_col(m.group(1)) is not None and 1 <= int(m.group(2)) <= MAX_ROWS:
        return "The name that you entered is not valid."
    return None


def _scope(s: Session, p: Dict[str, Any]) -> Optional[int]:
    sc = p.get("scope")
    if sc is None or sc == "" or str(sc).lower() == "workbook":
        return None
    d = s.doc()
    if isinstance(sc, int):
        return sc
    i = d.wb.sheet_index(str(sc))
    if i is None:
        raise EngineError(f"no sheet named `{sc}`")
    return i


def selection_reference(s: Session) -> str:
    d = s.doc()
    r = d.selection.current()
    return f"{quote_sheet(d.sheet.name)}!{r.abs_a1()}"


@command("formulas.defineName", "Define Name…", menu=("Formulas", "Defined Names"), params='{name, refersTo?: "=Sheet1!$A$1:$A$10" (default: the selection), scope?: sheet name, comment?, replace?: bool}')
def define_name(s: Session, p: Dict[str, Any]):
    d = s.doc()
    name = str(need(p, "name", "formulas.defineName")).strip()
    err = check_name(name)
    if err:
        raise EngineError(err)
    scope = _scope(s, p)
    refers = p_str(p, "refersTo") or ("=" + selection_reference(s))
    body = refers[1:] if refers.startswith("=") else refers
    try:
        e = parse(body)
    except ParseError as ex:
        raise EngineError(f"There's a problem with this formula: {ex}")
    text = print_formula(e)
    existing = [dn for dn in d.wb.names if dn.name.lower() == name.lower() and dn.scope == scope]
    if existing and not p_bool(p, "replace"):
        raise EngineError("The name that you entered already exists. Enter a unique name.")
    if d.wb.table(name) is not None:
        raise EngineError("The name that you entered already exists. Enter a unique name.")

    def f(ctx):
        wb = ctx.wb
        wb.names = [dn for dn in wb.names if not (dn.name.lower() == name.lower() and dn.scope == scope)]
        wb.names.append(DefinedName(name, text, scope, p_str(p, "comment", "") or ""))
        ctx.structural = True

    edit(s, f)
    return {"name": name, "refersTo": "=" + text}


@command("formulas.deleteName", "Delete Name", menu=("Formulas", "Defined Names"), params="{name, scope?}")
def delete_name(s: Session, p: Dict[str, Any]):
    d = s.doc()
    name = str(need(p, "name", "formulas.deleteName"))
    scope = _scope(s, p)
    if not any(dn.name.lower() == name.lower() and dn.scope == scope for dn in d.wb.names):
        raise EngineError(f"no name `{name}`")

    def f(ctx):
        ctx.wb.names = [dn for dn in ctx.wb.names if not (dn.name.lower() == name.lower() and dn.scope == scope)]
        ctx.structural = True

    edit(s, f)
    return None


@command("formulas.nameManager", "Name Manager", menu=("Formulas", "Defined Names"), shortcut="Ctrl+F3", undoable=False, journal=False)
def name_manager(s: Session, p: Dict[str, Any]):
    d = s.doc()
    out = []
    for dn in d.wb.names:
        if dn.hidden:
            continue
        si = dn.scope if dn.scope is not None else d.wb.active_sheet
        try:
            v = evaluate(d.wb, si, (0, 0), dn.formula)
        except Exception:
            v = CellError.VALUE
        if isinstance(v, Array):
            preview = "{" + ";".join(",".join(display(v.get(r, c)) for c in range(min(v.cols, 5))) for r in range(min(v.rows, 5))) + ("…}" if v.rows > 5 or v.cols > 5 else "}")
        else:
            preview = display(v)
        out.append({
            "name": dn.name,
            "refersTo": "=" + dn.formula,
            "scope": d.wb.sheets[dn.scope].name if dn.scope is not None and dn.scope < len(d.wb.sheets) else "Workbook",
            "value": preview,
            "comment": dn.comment,
        })
    for si, sh in enumerate(d.wb.sheets):
        for t in sh.tables:
            out.append({"name": t.name, "refersTo": f"={quote_sheet(sh.name)}!{t.range.abs_a1()}", "scope": "Workbook", "value": "", "comment": "", "table": True})
    out.sort(key=lambda x: x["name"].lower())
    return {"names": out}


@command("formulas.createFromSelection", "Create from Selection", menu=("Formulas", "Defined Names"), params="{top?, left?, bottom?, right?}")
def create_from_selection(s: Session, p: Dict[str, Any]):
    d = s.doc()
    sh = d.sheet
    rng = d.selection.current()
    top, left = p_bool(p, "top", True), p_bool(p, "left", False)
    bottom, right = p_bool(p, "bottom"), p_bool(p, "right")
    sheet = quote_sheet(sh.name)
    made = []

    def clean(t: str) -> Optional[str]:
        n = re.sub(r"[^A-Za-z0-9_.\\À-￿]", "_", t.strip())
        if n and not (n[0].isalpha() or n[0] in "_\\"):
            n = "_" + n
        return n if n and check_name(n) is None else None

    r0, r1, c0, c1 = rng.r0 + (1 if top else 0), rng.r1 - (1 if bottom else 0), rng.c0 + (1 if left else 0), rng.c1 - (1 if right else 0)
    if r0 > r1 or c0 > c1:
        raise EngineError("The selection is too small.")
    if top or bottom:
        hr = rng.r0 if top else rng.r1
        for c in range(c0, c1 + 1):
            v = sh.value((hr, c))
            n = clean(display(v)) if v is not None else None
            if n:
                made.append((n, f"{sheet}!{RangeRef(r0, c, r1, c).abs_a1()}"))
    if left or right:
        hc = rng.c0 if left else rng.c1
        for r in range(r0, r1 + 1):
            v = sh.value((r, hc))
            n = clean(display(v)) if v is not None else None
            if n:
                made.append((n, f"{sheet}!{RangeRef(r, c0, r, c1).abs_a1()}"))
    if not made:
        raise EngineError("No names were created: the labels aren't valid names.")

    def f(ctx):
        wb = ctx.wb
        for n, ref in made:
            wb.names = [dn for dn in wb.names if not (dn.name.lower() == n.lower() and dn.scope is None)]
            wb.names.append(DefinedName(n, ref, None))
        ctx.structural = True

    edit(s, f)
    return {"names": [n for n, _ in made]}


# ---------------------------------------------------------------- functions & calculation


@command("formulas.functions", "List Functions", menu=("Formulas", "Function Library"), params="{category?, search?}", undoable=False, journal=False)
def functions(s: Session, p: Dict[str, Any]):
    cat = p_str(p, "category")
    q = (p_str(p, "search") or "").lower()
    out = []
    for spec in all_functions():
        if cat and spec.category.lower() != cat.lower():
            continue
        if q and q not in spec.name.lower() and q not in spec.description.lower():
            continue
        out.append({"name": spec.name, "category": spec.category, "signature": spec.signature, "description": spec.description})
    return {"functions": out}


@command("formulas.calculateNow", "Calculate Now", menu=("Formulas", "Calculation"), shortcut="F9", undoable=False)
def calculate_now(s: Session, p: Dict[str, Any]):
    d = s.doc()

    def f(ctx):
        ctx.structural = True

    edit(s, f)
    return {"circular": len(d.calc.circular)}


@command("formulas.calculationOptions", "Calculation Options", menu=("Formulas", "Calculation"), params="{mode: auto|autoNoTable|manual, iterative?: bool, maxIterations?, maxChange?}")
def calculation_options(s: Session, p: Dict[str, Any]):
    mode = p_str(p, "mode")
    if mode is not None and mode not in ("auto", "autoNoTable", "manual"):
        raise EngineError(f"unknown calculation mode `{mode}`")

    def f(ctx):
        c = ctx.wb.calc
        if mode is not None:
            c.mode = mode
        if "iterative" in p:
            c.iterative = p_bool(p, "iterative")
        if p.get("maxIterations") is not None:
            c.max_iterations = max(1, min(int(p["maxIterations"]), 32767))
        if p.get("maxChange") is not None:
            c.max_change = max(0.0, float(p["maxChange"]))
        ctx.structural = c.mode != "manual"

    edit(s, f)
    c = s.doc().wb.calc
    return {"mode": c.mode, "iterative": c.iterative}


@command("formulas.evaluate", "Evaluate", params="{formula, cell?, sheet?}", undoable=False, journal=False)
def evaluate_cmd(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    at = param_cell(s, p)
    text = str(need(p, "formula", "formulas.evaluate"))
    v = evaluate(d.wb, si, at, text)
    return {"value": _json_value(v), "text": display(v) if not isinstance(v, Array) else None}


@command("formulas.tracePrecedents", "Trace Precedents", menu=("Formulas", "Formula Auditing"), params="{cell?}", undoable=False, journal=False)
def trace_precedents(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    at = param_cell(s, p)
    cell = d.wb.sheet(si).cell(at)
    if cell is None or cell.formula is None:
        return {"precedents": []}
    e = cell.formula.get_expr()
    out = []
    if e is not None:
        for r in e.references():
            sheet = r.sheet if isinstance(r.sheet, str) else (None if r.sheet is None else f"{r.sheet[0]}:{r.sheet[1]}")
            out.append({"sheet": sheet or d.wb.sheets[si].name, "range": r.range().a1()})
    return {"precedents": out}


@command("formulas.traceDependents", "Trace Dependents", menu=("Formulas", "Formula Auditing"), params="{cell?}", undoable=False, journal=False)
def trace_dependents(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    at = param_cell(s, p)
    keys = d.calc.graph.dependents(si, at[0], at[1])
    out = [{"sheet": d.wb.sheets[k[0]].name, "cell": a1(k[1], k[2])} for k in sorted(set(keys))[:1000]]
    return {"dependents": out}


@command("formulas.errorChecking", "Error Checking", menu=("Formulas", "Formula Auditing"), undoable=False, journal=False)
def error_checking(s: Session, p: Dict[str, Any]):
    d = s.doc()
    out = []
    for si, sh in enumerate(d.wb.sheets):
        for rc, cell in sh.cells.iter_sorted():
            if cell.formula is not None and isinstance(cell.value, CellError):
                out.append({"sheet": sh.name, "cell": a1(*rc), "error": cell.value.value, "formula": "=" + cell.formula.text})
                if len(out) >= 1000:
                    return {"errors": out, "circular": [a1(k[1], k[2]) for k in d.calc.circular[:50]]}
    return {"errors": out, "circular": [f"{d.wb.sheets[k[0]].name}!{a1(k[1], k[2])}" for k in d.calc.circular[:50]]}
