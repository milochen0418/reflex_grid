"""Editing commands: entering cells, selection, undo/redo, the clipboard, clearing, fill, find
and replace, Go To and Go To Special."""

from __future__ import annotations

import math
import re
from dataclasses import replace
from typing import Any, Dict, List, Optional, Tuple

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef, a1, parse_cell
from ..core.value import Array, CellError, display
from ..formula import Name, ParseError, Ref, parse
from ..formula.adjust import shift_relative
from ..formula.printer import quote_sheet
from ..model.cell import Cell, Formula
from ..model.style import Style
from .display import cell_text, edit_text, range_text, value_text
from .entry import check_entry, input_to_cell, parse_formula_input
from .fill import fill, series
from .ranges import cells_to_ranges, copy_rules, remove_rules_in
from .selection import Selection, current_region, jump
from .session import (
    PROTECTED,
    Clipboard,
    EngineError,
    HistoryEntry,
    Session,
    command,
    edit,
    need,
    p_bool,
    p_float,
    p_int,
    p_str,
    param_cell,
    param_range,
    param_sheet,
)
from .structure import move_block
from .tables import auto_expand, calculated_column

MAX_EDIT_CELLS = 2_000_000


def json_to_input(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v == int(v) and abs(v) < 1e15 else repr(v)
    return str(v)


def cell_ref_text(wb, si: int, rc: Tuple[int, int]) -> str:
    return f"{quote_sheet(wb.sheets[si].name)}!{a1(*rc)}"


def array_at(sheet, rc: Tuple[int, int]) -> Optional[RangeRef]:
    """The legacy (Ctrl+Shift+Enter) array formula range covering ``rc``."""
    window = RangeRef(max(rc[0] - 5000, 0), 0, rc[0], rc[1])
    for (r, c), cell in sheet.cells.iter_range(window):
        f = cell.formula
        if f is not None and f.array is not None and f.array.contains(rc):
            return f.array
    return None


def check_locked(wb, si: int, rng: RangeRef) -> None:
    sh = wb.sheet(si)
    if sh is None or sh.protection is None:
        return
    if rng.count() <= 100_000:
        for rc in rng.cells():
            if wb.styles.get(sh.style_id(rc)).protection.locked:
                raise EngineError(PROTECTED)
    else:
        raise EngineError(PROTECTED)


def check_arrays(sheet, rng: RangeRef) -> None:
    """Refuses changing part of a legacy array formula."""
    window = RangeRef(max(rng.r0 - 5000, 0), 0, rng.r1, rng.c1)
    for (r, c), cell in sheet.cells.iter_range(window):
        f = cell.formula
        if f is not None and f.array is not None and not f.array.is_single() and f.array.intersects(rng) and not rng.contains_range(f.array):
            raise EngineError("You can't change part of an array.")


def auto_widen(ctx, si: int, rc: Tuple[int, int]) -> None:
    """Like Excel, a number typed into a column of standard width widens the column when it
    would otherwise show as ####."""
    from ..model.features import LineInfo
    from ..numfmt import GENERAL_WIDTH, format_value, parse as parse_fmt
    from .layout import approx_text_width

    wb = ctx.wb
    sh = wb.sheet(si)
    info = sh.cols.get(rc[1])
    if (info is not None and (info.custom or info.hidden)) or sh.merge_at(rc) is not None:
        return
    cell = sh.cell(rc)
    st = wb.styles.get(cell.style)
    if st.align.wrap or st.align.shrink or st.align.rotation:
        return
    text = format_value(cell.value, parse_fmt(st.num_fmt), wb.date_system, GENERAL_WIDTH).text
    need = math.ceil(approx_text_width(text, st.font.size, st.font.bold) + 6.0)
    if need > sh.col_width(rc[1]) and need <= 400:
        ctx.sheet_mut(si).set_col(rc[1], replace(info or LineInfo(), size=float(need), custom=False))


def _set(ctx, si: int, rc: Tuple[int, int], cell: Optional[Cell]) -> None:
    sm = ctx.sheet_mut(si)
    if cell is None or cell.is_blank():
        sm.remove_cell(rc)
    else:
        sm.set_cell(rc, cell)
    ctx.touch(si, rc)


# ---------------------------------------------------------------- entering


@command("cell.set", "Enter Cell", params='{cell?: "B2", input: "text, number or =formula", sheet?, array?: bool}')
def cell_set(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    at = param_cell(s, p)
    raw = p.get("input", p.get("value"))
    if raw is None and "input" not in p and "value" not in p:
        raise EngineError("cell.set: missing `input`")
    text = json_to_input(raw)
    array = p_bool(p, "array")
    sh = d.wb.sheet(si)
    if sh.protection is not None and d.wb.styles.get(sh.style_id(at)).protection.locked:
        raise EngineError(PROTECTED)
    arr = array_at(sh, at)
    if arr is not None and not arr.is_single() and not array:
        raise EngineError("You can't change part of an array.")

    def f(ctx):
        wb = ctx.wb
        sheet = wb.sheet(si)
        if array:
            sel = ctx.sel.current()
            rng = sel if si == wb.active_sheet and sel.contains(at) else RangeRef.of(at)
            if not rng.is_single() and any(t.range.intersects(rng) for t in sheet.tables):
                raise EngineError("Multi-cell array formulas aren't allowed in tables.")
            if not rng.is_single():
                check_arrays(sheet, rng)
            if not (text.startswith("=") or text.startswith("+") or text.startswith("-")):
                cell = input_to_cell(wb, sheet, rng.start, text)
            else:
                fm = parse_formula_input(text)
                cell = Cell(None, Formula(fm.text, fm.expr, rng), sheet.style_id(rng.start))
            sm = ctx.sheet_mut(si)
            if rng.count() > MAX_EDIT_CELLS:
                raise EngineError("The selection is too large.")
            for rc in rng.cells():
                if rc == rng.start:
                    continue
                old = sm.cell(rc)
                if old is not None and (old.value is not None or old.formula is not None):
                    _set(ctx, si, rc, Cell(None, None, old.style))
            _set(ctx, si, rng.start, cell)
            return
        cell = input_to_cell(wb, sheet, at, text)
        _set(ctx, si, at, cell)
        if cell is None or cell.formula is None:
            # Validation sees the new value in place (custom rules may refer to the cell itself).
            bad = check_entry(ctx.wb, si, at, cell.value if cell is not None else None)
            if bad is not None:
                if bad[0].error_style == "stop":
                    raise EngineError(bad[1])
                s.messages.append(bad[1])
        if cell is not None and cell.formula is None and isinstance(cell.value, float):
            auto_widen(ctx, si, at)
        if sheet.tables:
            auto_expand(ctx, si, at)
            calculated_column(ctx, si, at)

    edit(s, f)
    sh = s.doc().wb.sheet(si)
    return {"cell": a1(*at), "value": display(sh.value(at))}


@command("cell.get", "Read Cell", params='{cell?, sheet?}', undoable=False, journal=False)
def cell_get(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    at = param_cell(s, p)
    sh = d.wb.sheet(si)
    cell = sh.cell(at)
    v = sh.value(at)
    return {
        "cell": a1(*at),
        "sheet": sh.name,
        "value": _json_value(v),
        "text": value_text(d.wb, sh, at),
        "input": edit_text(d.wb, sh, at),
        "formula": ("=" + cell.formula.text) if cell is not None and cell.formula is not None else None,
        "format": d.wb.styles.get(sh.style_id(at)).num_fmt,
    }


def _json_value(v):
    if isinstance(v, CellError):
        return v.value
    if isinstance(v, Array):
        return [[_json_value(v.get(r, c)) for c in range(v.cols)] for r in range(v.rows)]
    return v


@command("range.setValues", "Set Values", params='{range|cell: "A1", values: [[1, "text", "=A1*2"], …]}')
def range_set_values(s: Session, p: Dict[str, Any]):
    si = param_sheet(s, p)
    start = param_range(s, p).start
    rows = need(p, "values", "range.setValues")
    if not isinstance(rows, list):
        raise EngineError("range.setValues: `values` must be a list of rows")
    if len(rows) > MAX_ROWS:
        raise EngineError("too many rows")
    d = s.doc()
    h = len(rows)
    w = max((len(r) if isinstance(r, list) else 1 for r in rows), default=0)
    if h and w:
        check_locked(d.wb, si, RangeRef(start[0], start[1], min(start[0] + h - 1, MAX_ROWS - 1), min(start[1] + w - 1, MAX_COLS - 1)))

    def f(ctx):
        wb = ctx.wb
        sheet = wb.sheet(si)
        for i, row in enumerate(rows):
            if not isinstance(row, list):
                row = [row]
            for j, v in enumerate(row):
                rc = (start[0] + i, start[1] + j)
                if rc[0] >= MAX_ROWS or rc[1] >= MAX_COLS:
                    continue
                if isinstance(v, bool) or isinstance(v, (int, float)) and not isinstance(v, bool):
                    cell = Cell(float(v) if not isinstance(v, bool) else v, None, sheet.style_id(rc))
                else:
                    cell = input_to_cell(wb, sheet, rc, json_to_input(v))
                _set(ctx, si, rc, cell)

    edit(s, f)
    return {"range": RangeRef(start[0], start[1], start[0] + max(h - 1, 0), start[1] + max(w - 1, 0)).a1()}


@command("range.get", "Read Range", params='{range?, sheet?, formulas?: bool}', undoable=False, journal=False)
def range_get(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    rng = param_range(s, p)
    sh = d.wb.sheet(si)
    used = sh.used_range()
    if used is None:
        return {"range": rng.a1(), "values": []}
    inter = rng.intersection(RangeRef(rng.r0, rng.c0, max(used.r1, rng.r0), max(used.c1, rng.c0)))
    if inter is None or inter.count() > 1_000_000:
        raise EngineError("range too large")
    formulas = p_bool(p, "formulas")
    out = []
    for r in range(inter.r0, inter.r1 + 1):
        row = []
        for c in range(inter.c0, inter.c1 + 1):
            cell = sh.cell((r, c))
            if formulas and cell is not None and cell.formula is not None:
                row.append("=" + cell.formula.text)
            else:
                row.append(_json_value(sh.value((r, c))))
        out.append(row)
    return {"range": inter.a1(), "values": out}


@command("range.fill", "Fill Selection with Entry", params='{range?, input} (Ctrl+Enter)')
def range_fill(s: Session, p: Dict[str, Any]):
    si = param_sheet(s, p)
    rng = param_range(s, p)
    text = json_to_input(p.get("input", p.get("value")))
    d = s.doc()
    if rng.count() > MAX_EDIT_CELLS:
        raise EngineError("range too large")
    check_locked(d.wb, si, rng)
    check_arrays(d.wb.sheet(si), rng)

    def f(ctx):
        wb = ctx.wb
        sheet = wb.sheet(si)
        first = input_to_cell(wb, sheet, rng.start, text)
        for rc in rng.cells():
            if first is not None and first.formula is not None:
                e = first.formula.get_expr()
                fm = Formula.from_expr(shift_relative(e, rc[0] - rng.r0, rc[1] - rng.c0)) if e is not None else first.formula
                cell = Cell(None, fm, wb.sheet(si).style_id(rc) if rc != rng.start else first.style)
            else:
                cell = input_to_cell(wb, wb.sheet(si), rc, text)
            _set(ctx, si, rc, cell)

    edit(s, f)
    return {"range": rng.a1()}


# ---------------------------------------------------------------- selection


def _visible_step(sh, rc: Tuple[int, int], dr: int, dc: int) -> Tuple[int, int]:
    r, c = rc
    m = sh.merge_at(rc)
    if m is not None:
        if dr > 0:
            r = m.r1
        elif dr < 0:
            r = m.r0
        if dc > 0:
            c = m.c1
        elif dc < 0:
            c = m.c0
    for _ in range(abs(dr)):
        n = r + (1 if dr > 0 else -1)
        guard = 0
        while 0 <= n < MAX_ROWS and sh.is_row_hidden(n) and guard < 100000:
            n += 1 if dr > 0 else -1
            guard += 1
        if not 0 <= n < MAX_ROWS:
            break
        r = n
    for _ in range(abs(dc)):
        n = c + (1 if dc > 0 else -1)
        guard = 0
        while 0 <= n < MAX_COLS and sh.is_col_hidden(n) and guard < 20000:
            n += 1 if dc > 0 else -1
            guard += 1
        if not 0 <= n < MAX_COLS:
            break
        c = n
    return r, c


def _select_cell(d, rc: Tuple[int, int]) -> None:
    sh = d.sheet
    m = sh.merge_at(rc)
    if m is not None and not m.is_single():
        d.selection = Selection(m.start, m.start, [m])
    else:
        d.selection = Selection.at(rc)


@command("selection.set", "Select", params='{range: "A1:B2,D4" | cell, active?: "A1", sheet?}', undoable=False, journal=False)
def selection_set(s: Session, p: Dict[str, Any]):
    d = s.doc()
    if p.get("sheet") is not None or (isinstance(p.get("range"), str) and "!" in p["range"]):
        si = param_sheet(s, p)
        if si != d.wb.active_sheet:
            _activate(s, si)
    t = p.get("range") or p.get("cell")
    if not t:
        raise EngineError("selection.set: missing `range`")
    body = str(t).rsplit("!", 1)[-1]
    ranges = []
    for part in body.split(","):
        r = RangeRef.parse(part)
        if r is None:
            raise EngineError(f"not a range: `{part}`")
        ranges.append(r)
    active = None
    if p.get("active"):
        active = parse_cell(str(p["active"]).rsplit("!", 1)[-1])
    if active is None or not any(r.contains(active) for r in ranges):
        active = ranges[-1].start
    anchor = active
    if p.get("anchor"):
        a = parse_cell(str(p["anchor"]))
        if a is not None:
            anchor = a
    sel = Selection(active, anchor, ranges)
    sel.expand_merges(d.sheet)
    d.selection = sel
    return sel.to_json()


@command("selection.move", "Move Selection", params="{dr, dc, extend?: bool, jump?: bool (Ctrl+arrow)}", undoable=False, journal=False)
def selection_move(s: Session, p: Dict[str, Any]):
    d = s.doc()
    sh = d.sheet
    sel = d.selection
    dr, dc = p_int(p, "dr", 0), p_int(p, "dc", 0)
    jmp = p_bool(p, "jump")
    if p_bool(p, "extend"):
        r = sel.current()
        anchor = sel.active
        cr = r.r1 if anchor[0] <= r.r0 else r.r0
        cc = r.c1 if anchor[1] <= r.c0 else r.c0
        if anchor[0] == r.r1 and anchor[0] != r.r0:
            cr = r.r0
        if anchor[1] == r.c1 and anchor[1] != r.c0:
            cc = r.c0
        corner = jump(sh, (cr, cc), dr, dc) if jmp else _visible_step(sh, (cr, cc), dr, dc)
        new = RangeRef.of(anchor, corner)
        ns = Selection(anchor, anchor, sel.ranges[:-1] + [new])
        ns.expand_merges(sh)
        d.selection = ns
    else:
        n = jump(sh, sel.active, dr, dc) if jmp else _visible_step(sh, sel.active, dr, dc)
        _select_cell(d, n)
    return d.selection.to_json()


@command("selection.next", "Next Cell in Selection", params="{forward?: true, byRow?: false} (Enter/Tab)", undoable=False, journal=False)
def selection_next(s: Session, p: Dict[str, Any]):
    d = s.doc()
    forward = p_bool(p, "forward", True)
    by_row = p_bool(p, "byRow", False)
    sel = d.selection
    if sel.is_single_cell() or (len(sel.ranges) == 1 and sel.current() == d.sheet.merge_at(sel.active)):
        dr, dc = (0, 1 if forward else -1) if by_row else (1 if forward else -1, 0)
        _select_cell(d, _visible_step(d.sheet, sel.active, dr, dc))
    else:
        ns = sel.copy()
        ns.cycle(forward, by_row)
        d.selection = ns
    return d.selection.to_json()


@command("edit.selectAll", "Select All", menu=("Edit",), shortcut="Ctrl+A", params="{} (current region first, then the whole sheet)", undoable=False, journal=False)
def select_all(s: Session, p: Dict[str, Any]):
    d = s.doc()
    sh = d.sheet
    region = current_region(sh, d.selection.active)
    if not region.is_single() and d.selection.current() != region and not p_bool(p, "sheet"):
        d.selection = Selection(d.selection.active, d.selection.active, [region])
    else:
        d.selection = Selection(d.selection.active, d.selection.active, [RangeRef.all()])
    return d.selection.to_json()


@command("selection.currentRegion", "Select Current Region", shortcut="Ctrl+Shift+8", undoable=False, journal=False)
def select_region(s: Session, p: Dict[str, Any]):
    d = s.doc()
    region = current_region(d.sheet, d.selection.active)
    d.selection = Selection(d.selection.active, d.selection.active, [region])
    return d.selection.to_json()


@command("selection.row", "Select Entire Row", shortcut="Shift+Space", undoable=False, journal=False)
def select_row(s: Session, p: Dict[str, Any]):
    d = s.doc()
    rs = [RangeRef.rows(r.r0, r.r1) for r in d.selection.ranges]
    d.selection = Selection(d.selection.active, d.selection.active, rs)
    return d.selection.to_json()


@command("selection.column", "Select Entire Column", shortcut="Ctrl+Space", undoable=False, journal=False)
def select_col(s: Session, p: Dict[str, Any]):
    d = s.doc()
    rs = [RangeRef.cols(r.c0, r.c1) for r in d.selection.ranges]
    d.selection = Selection(d.selection.active, d.selection.active, rs)
    return d.selection.to_json()


def resolve_reference(wb, si: int, text: str) -> Tuple[int, RangeRef]:
    """``B5``, ``Sheet2!A1:C3``, ``A:A``, a defined name or a table name → (sheet, range)."""
    t = text.strip().lstrip("=")
    if not t:
        raise EngineError("Reference isn't valid.")
    r = RangeRef.parse(t)
    if r is not None:
        return si, r
    loc = wb.table(t)
    if loc is not None:
        return loc[0], wb.sheets[loc[0]].tables[loc[1]].range
    try:
        e = parse(t)
    except ParseError:
        raise EngineError("Reference isn't valid.")
    seen = 0
    while isinstance(e, Name) and seen < 10:
        dn = wb.name(e.name, si)
        if dn is None:
            raise EngineError("Reference isn't valid.")
        try:
            e = parse(dn.formula)
        except ParseError:
            raise EngineError("Reference isn't valid.")
        seen += 1
    if isinstance(e, Ref):
        target = si
        if isinstance(e.sheet, str):
            i = wb.sheet_index(e.sheet)
            if i is None:
                raise EngineError("Reference isn't valid.")
            target = i
        return target, e.range()
    raise EngineError("Reference isn't valid.")


def _activate(s: Session, si: int) -> None:
    d = s.doc()
    if si == d.wb.active_sheet:
        return
    d.sheet_selections[d.wb.active_sheet] = d.selection
    d.wb.active_sheet = si
    d.selection = d.sheet_selections.get(si) or Selection.at(d.wb.sheets[si].view_active)


@command("edit.goTo", "Go To…", menu=("Home", "Editing", "Find & Select"), shortcut="Ctrl+G", params='{reference: "B5" | "Sheet2!A1:C3" | name}', undoable=False)
def go_to(s: Session, p: Dict[str, Any]):
    d = s.doc()
    ref = p.get("reference") or p.get("range")
    if not ref:
        raise EngineError("edit.goTo: missing `reference`")
    si, rng = resolve_reference(d.wb, d.wb.active_sheet, str(ref))
    sh = d.wb.sheets[si]
    if sh.visibility != "visible":
        raise EngineError("Reference isn't valid.")
    _activate(s, si)
    d.selection = Selection.of_range(rng)
    d.selection.expand_merges(sh)
    return {"sheet": sh.name, "range": rng.a1()}


@command("edit.goToSpecial", "Go To Special…", menu=("Home", "Editing", "Find & Select"), params="{kind: blanks|constants|formulas|comments|lastCell|currentRegion|visible|conditionalFormats|dataValidation|errors|numbers|text}", undoable=False)
def go_to_special(s: Session, p: Dict[str, Any]):
    d = s.doc()
    sh = d.sheet
    kind = p_str(p, "kind", "blanks")
    used = sh.used_range()
    sel = d.selection.current()
    area = sel if not sel.is_single() else used
    if kind == "lastCell":
        if used is None:
            d.selection = Selection.at((0, 0))
        else:
            d.selection = Selection.at(used.end)
        return d.selection.to_json()
    if kind == "currentRegion":
        return select_region(s, p)
    if kind in ("conditionalFormats", "dataValidation"):
        rules = sh.cond_formats if kind == "conditionalFormats" else sh.validations
        ranges = [r for rule in rules for r in rule.ranges]
        if not ranges:
            raise EngineError("No cells were found.")
        d.selection = Selection(ranges[0].start, ranges[0].start, ranges)
        return d.selection.to_json()
    if area is None or used is None:
        raise EngineError("No cells were found.")
    area = area.intersection(used) if kind != "visible" else area
    if area is None:
        raise EngineError("No cells were found.")
    if kind == "visible":
        ranges = []
        r = area.r0
        limit = min(area.r1, used.r1 if used else area.r1)
        while r <= limit:
            if sh.is_row_hidden(r):
                r += 1
                continue
            start = r
            while r + 1 <= limit and not sh.is_row_hidden(r + 1):
                r += 1
            ranges.append(RangeRef(start, area.c0, r, area.c1))
            r += 1
        if not ranges:
            raise EngineError("No cells were found.")
        d.selection = Selection(ranges[0].start, ranges[0].start, ranges)
        return d.selection.to_json()
    cells = []
    if kind == "blanks":
        if area.count() > MAX_EDIT_CELLS:
            raise EngineError("range too large")
        for rc in area.cells():
            if sh.value(rc) is None and sh.cell(rc) is None or (sh.cell(rc) is not None and sh.cell(rc).value is None and sh.cell(rc).formula is None and rc not in sh.spill):
                cells.append(rc)
    elif kind == "comments":
        cells = [rc for rc in sh.comments if area.contains(rc)]
    else:
        for rc, cell in sh.cells.iter_range(area):
            v = cell.value
            if kind == "formulas" and cell.formula is not None:
                cells.append(rc)
            elif kind == "constants" and cell.formula is None and v is not None:
                cells.append(rc)
            elif kind == "errors" and isinstance(v, CellError):
                cells.append(rc)
            elif kind == "numbers" and isinstance(v, float):
                cells.append(rc)
            elif kind == "text" and isinstance(v, str):
                cells.append(rc)
    if not cells:
        raise EngineError("No cells were found.")
    ranges = cells_to_ranges(cells)
    d.selection = Selection(ranges[0].start, ranges[0].start, ranges)
    return d.selection.to_json()


# ---------------------------------------------------------------- undo / redo


@command("edit.undo", "Undo", menu=("Edit",), shortcut="Ctrl+Z", params="{steps?: 1}", undoable=False)
def undo(s: Session, p: Dict[str, Any]):
    d = s.doc()
    if not d.undo:
        raise EngineError("Nothing to undo.")
    label = ""
    for _ in range(max(p_int(p, "steps", 1), 1)):
        if not d.undo:
            break
        h = d.undo.pop()
        d.redo.append(HistoryEntry(h.label, d.wb, d.selection.copy(), d.wb.active_sheet))
        d.wb = h.wb
        d.selection = h.selection
        label = h.label
    d.calc.rebuild(d.wb)
    d.revision += 1
    return {"undone": label}


@command("edit.redo", "Redo", menu=("Edit",), shortcut="Ctrl+Y", params="{steps?: 1}", undoable=False)
def redo(s: Session, p: Dict[str, Any]):
    d = s.doc()
    if not d.redo:
        raise EngineError("Nothing to redo.")
    label = ""
    for _ in range(max(p_int(p, "steps", 1), 1)):
        if not d.redo:
            break
        h = d.redo.pop()
        d.undo.append(HistoryEntry(h.label, d.wb, d.selection.copy(), d.wb.active_sheet))
        d.wb = h.wb
        d.selection = h.selection
        label = h.label
    d.calc.rebuild(d.wb)
    d.revision += 1
    return {"redone": label}


# ---------------------------------------------------------------- clipboard


def _copy(s: Session, p: Dict[str, Any], cut: bool):
    d = s.doc()
    si = param_sheet(s, p)
    if p.get("range"):
        rng = param_range(s, p)
    else:
        if len(d.selection.ranges) > 1:
            raise EngineError("This action won't work on multiple selections.")
        rng = d.selection.current()
    sh = d.wb.sheet(si)
    used = sh.used_range()
    if rng.count() > MAX_EDIT_CELLS and used is not None:
        clip = rng.intersection(RangeRef(0, 0, used.r1, used.c1))
        rng = clip if clip is not None else RangeRef.of(rng.start)
    # The text the system clipboard got (from the UI) identifies our copy when it's pasted back.
    text = p["text"] if isinstance(p.get("text"), str) else range_text(d.wb, sh, rng)
    s.clipboard = Clipboard(d.wb, si, rng, cut, d.uid, text)
    return {"range": rng.a1(), "text": text, "cut": cut}


@command("edit.copy", "Copy", menu=("Home", "Clipboard"), shortcut="Ctrl+C", params="{range?, text?: what the system clipboard got} → {range, text}", undoable=False)
def copy_cmd(s: Session, p: Dict[str, Any]):
    return _copy(s, p, False)


@command("edit.cut", "Cut", menu=("Home", "Clipboard"), shortcut="Ctrl+X", params="{range?} → {range, text}", undoable=False)
def cut_cmd(s: Session, p: Dict[str, Any]):
    return _copy(s, p, True)


@command("edit.clearClipboard", "Cancel Copy", shortcut="Escape", undoable=False, journal=False)
def clear_clipboard(s: Session, p: Dict[str, Any]):
    s.clipboard = None
    return None


def parse_tsv(text: str) -> List[List[str]]:
    """Tab-separated clipboard text; quoted fields may hold tabs, newlines and doubled quotes."""
    rows: List[List[str]] = []
    row: List[str] = []
    i, n = 0, len(text)
    field = []
    at_start = True
    while i < n:
        ch = text[i]
        if at_start and ch == '"':
            # A quoted field runs to the closing quote that is followed by a separator.
            j = i + 1
            buf = []
            closed = False
            while j < n:
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    if j + 1 >= n or text[j + 1] in "\t\r\n":
                        closed = True
                        j += 1
                        break
                buf.append(text[j])
                j += 1
            if closed:
                field = buf
                i = j
                at_start = False
                continue
        if ch == "\t":
            row.append("".join(field))
            field = []
            at_start = True
        elif ch in "\r\n":
            row.append("".join(field))
            rows.append(row)
            row, field = [], []
            at_start = True
            if ch == "\r" and i + 1 < n and text[i + 1] == "\n":
                i += 1
        else:
            field.append(ch)
            at_start = False
        i += 1
    if field or row:
        row.append("".join(field))
        rows.append(row)
    return rows


@command("edit.paste", "Paste", menu=("Home", "Clipboard"), shortcut="Ctrl+V", params='{text?: "system clipboard text", at?: "B2"}')
def paste(s: Session, p: Dict[str, Any]):
    text = p.get("text")
    clip = s.clipboard
    if isinstance(text, str) and (clip is None or clip.text.rstrip("\r\n") != text.rstrip("\r\n")):
        return paste_text(s, p, text)
    if clip is None:
        raise EngineError("Nothing to paste.")
    return paste_special(s, {"what": "all", **({"at": p["at"]} if p.get("at") else {})})


def paste_text(s: Session, p: Dict[str, Any], text: str):
    d = s.doc()
    si = param_sheet(s, p)
    at = param_cell(s, p, "at") if p.get("at") else d.selection.current().start
    rows = parse_tsv(text)
    if not rows:
        return {"range": a1(*at)}
    h = len(rows)
    w = max(len(r) for r in rows)
    if at[0] + h > MAX_ROWS or at[1] + w > MAX_COLS:
        raise EngineError("The data you're pasting doesn't fit on the sheet.")
    dest = RangeRef(at[0], at[1], at[0] + h - 1, at[1] + w - 1)
    check_locked(d.wb, si, dest)
    check_arrays(d.wb.sheet(si), dest)

    def f(ctx):
        wb = ctx.wb
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                rc = (at[0] + i, at[1] + j)
                try:
                    cell = input_to_cell(wb, wb.sheet(si), rc, v)
                except EngineError:
                    cell = Cell(v, None, wb.sheet(si).style_id(rc))
                _set(ctx, si, rc, cell)
        ctx.sel = Selection(at, at, [dest])

    edit(s, f)
    return {"range": dest.a1()}


_OPS = {"add": "+", "subtract": "-", "multiply": "*", "divide": "/"}


def _operate(op: str, dest: Optional[Cell], src_v, style: int) -> Optional[Cell]:
    """Paste Special › Operation on one cell."""
    sym = _OPS[op]
    if not isinstance(src_v, (float, bool)) and src_v is not None:
        return None  # text and errors in the copy are pasted as they are
    sv = float(src_v) if src_v is not None else 0.0
    if dest is not None and dest.formula is not None:
        text = f"({dest.formula.text}){sym}{display(sv)}"
        return Cell(None, Formula.new(text), style)
    dv = dest.value if dest is not None else None
    if dv is None:
        dv = 0.0
    if isinstance(dv, bool):
        dv = float(dv)
    if not isinstance(dv, float):
        return dest  # text and errors stay
    if op == "add":
        r = dv + sv
    elif op == "subtract":
        r = dv - sv
    elif op == "multiply":
        r = dv * sv
    else:
        r = CellError.DIV0 if sv == 0 else dv / sv
    return Cell(r, None, style)


@command("edit.pasteSpecial", "Paste Special…", menu=("Home", "Clipboard", "Paste"), shortcut="Ctrl+Alt+V",
         params="{what: all|formulas|values|formats|comments|validation|allExceptBorders|columnWidths|formulasAndNumberFormats|valuesAndNumberFormats, operation?: none|add|subtract|multiply|divide, skipBlanks?, transpose?, at?}")
def paste_special(s: Session, p: Dict[str, Any]):
    clip = s.clipboard
    if clip is None:
        raise EngineError("Nothing to paste.")
    d = s.doc()
    what = p_str(p, "what", "all")
    if what == "notes":
        what = "comments"
    if what not in ("all", "formulas", "values", "formats", "comments", "validation", "allExceptBorders", "columnWidths", "formulasAndNumberFormats", "valuesAndNumberFormats"):
        raise EngineError(f"edit.pasteSpecial: unknown `what`: {what}")
    op = p_str(p, "operation", "none")
    if op not in ("none", "add", "subtract", "multiply", "divide"):
        raise EngineError(f"edit.pasteSpecial: unknown `operation`: {op}")
    skip_blanks = p_bool(p, "skipBlanks")
    transpose = p_bool(p, "transpose")
    si = param_sheet(s, p)
    src = clip.range
    csh = clip.wb.sheet(clip.sheet)
    if csh is None:
        raise EngineError("The copied cells are gone.")
    sel = d.selection.current()
    at = param_cell(s, p, "at") if p.get("at") else sel.start
    h, w = (src.width, src.height) if transpose else (src.height, src.width)
    tile = not p.get("at") and not sel.is_single() and sel.height % h == 0 and sel.width % w == 0 and sel.start == at
    dest = RangeRef(at[0], at[1], at[0] + (sel.height if tile else h) - 1, at[1] + (sel.width if tile else w) - 1)
    if dest.r1 >= MAX_ROWS or dest.c1 >= MAX_COLS:
        raise EngineError("The data you're pasting doesn't fit on the sheet.")
    if dest.count() > MAX_EDIT_CELLS:
        raise EngineError("The paste area is too large.")
    same_doc = clip.doc_uid == d.uid
    if what != "columnWidths":
        check_locked(d.wb, si, dest)
        check_arrays(d.wb.sheet(si), dest)

    if clip.cut and same_doc and what == "all" and not transpose and op == "none":
        if src.count() > MAX_EDIT_CELLS:
            raise EngineError("The paste area is too large.")
        csi = clip.sheet

        def mv(ctx):
            move_block(ctx, csi, src, at[0], at[1], si)
            ctx.sel = Selection(at, at, [RangeRef(at[0], at[1], at[0] + src.height - 1, at[1] + src.width - 1)])

        edit(s, mv)
        s.clipboard = None
        return {"range": RangeRef(at[0], at[1], at[0] + src.height - 1, at[1] + src.width - 1).a1(), "moved": True}

    def f(ctx):
        wb = ctx.wb
        sm = ctx.sheet_mut(si)
        cstyles = clip.wb.styles

        def src_of(rc):
            i = (rc[0] - dest.r0) % h
            j = (rc[1] - dest.c0) % w
            return (src.r0 + j, src.c0 + i) if transpose else (src.r0 + i, src.c0 + j)

        if what == "columnWidths":
            for c in range(dest.c0, dest.c1 + 1):
                sc = src_of((dest.r0, c))[1]
                sm.set_col(c, replace(sm.col_info(c), size=csh.col_info(sc).size))
            return
        for rc in dest.cells():
            sr = src_of(rc)
            scell = csh.cell(sr)
            sv = csh.value(sr)
            spilled = (scell is None or (scell.value is None and scell.formula is None)) and sr in csh.spill
            if spilled:
                anchor = next((a for a, r in csh.spill_ranges.items() if r.contains(sr)), None)
                if anchor is not None and src.contains(anchor):
                    sv = None  # re-spills from the pasted anchor
                    spilled = False
            blank = (scell is None or (scell.value is None and scell.formula is None)) and not spilled
            if skip_blanks and blank:
                continue
            old = sm.cell(rc)
            old_style = sm.style_id(rc)
            src_style = cstyles.get(csh.style_id(sr))
            sid = wb.styles.intern(src_style)
            # Content.
            if scell is not None and scell.formula is not None and not spilled:
                e = scell.formula.get_expr()
                fm = Formula.from_expr(shift_relative(e, rc[0] - sr[0], rc[1] - sr[1])) if e is not None else Formula(scell.formula.text)
                if scell.formula.array is not None:
                    if not scell.formula.array.is_single() and not src.contains_range(scell.formula.array):
                        fm = None
                    else:
                        a = scell.formula.array
                        fm = Formula(fm.text, fm.expr, RangeRef(a.r0 + rc[0] - sr[0], a.c0 + rc[1] - sr[1], a.r1 + rc[0] - sr[0], a.c1 + rc[1] - sr[1]))
                content = (None, fm) if fm is not None else (scell.value, None)
            else:
                content = (sv if (scell is not None or spilled) else None, None)
            if what == "values" or what == "valuesAndNumberFormats":
                content = (sv if not isinstance(sv, Array) else sv.get(0, 0), None)
            # Style.
            if what in ("all",):
                style = sid
            elif what == "allExceptBorders":
                style = wb.styles.intern(replace(src_style, borders=wb.styles.get(old_style).borders))
            elif what == "formats":
                style = sid
            elif what in ("formulasAndNumberFormats", "valuesAndNumberFormats"):
                style = wb.styles.intern(replace(wb.styles.get(old_style), num_fmt=src_style.num_fmt))
            else:
                style = old_style
            if what == "formats":
                new = old.with_style(style) if old is not None else Cell(None, None, style)
            elif what in ("comments", "validation"):
                new = old
            elif op != "none":
                srcv = content[0] if content[1] is None else sv
                new = _operate(op, old, srcv, style)
                if new is None:
                    new = Cell(content[0], content[1], style)
            else:
                new = Cell(content[0], content[1], style)
            if new is not old:
                _set(ctx, si, rc, new)
        dr, dc = dest.r0 - src.r0, dest.c0 - src.c0
        if what in ("all", "comments"):
            for rc0, cm in list(csh.comments.items()):
                if src.contains(rc0):
                    rr = (src.r0 + (rc0[1] - src.c0), src.c0 + (rc0[0] - src.r0)) if transpose else rc0
                    sm.comments[(rr[0] + dr, rr[1] + dc)] = cm
        if what == "all":
            for rc0, hl in list(csh.hyperlinks.items()):
                if src.contains(rc0):
                    rr = (src.r0 + (rc0[1] - src.c0), src.c0 + (rc0[0] - src.r0)) if transpose else rc0
                    sm.hyperlinks[(rr[0] + dr, rr[1] + dc)] = hl
        if what in ("all", "formats"):
            sm.merges = [m for m in sm.merges if not m.intersects(dest)]
            for m in csh.merges:
                if src.contains_range(m) and not transpose:
                    sm.merges.append(RangeRef(m.r0 + dr, m.c0 + dc, m.r1 + dr, m.c1 + dc))
            if csh.cond_formats:
                sm.cond_formats = remove_rules_in(sm.cond_formats, dest) + copy_rules(csh.cond_formats, src, dr, dc, transpose)
        if what in ("all", "validation") and csh.validations:
            sm.validations = remove_rules_in(sm.validations, dest) + copy_rules(csh.validations, src, dr, dc, transpose)
        ctx.sel = Selection(at, at, [dest])

    edit(s, f)
    if clip.cut and what == "all":
        s.clipboard = None
    return {"range": dest.a1()}


@command("edit.moveCells", "Move Cells", params='{from: "A1:B2", to: "D1" (top-left), sheet?, copy?: bool}')
def move_cells(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    src = RangeRef.parse(str(need(p, "from", "edit.moveCells")))
    to = parse_cell(str(need(p, "to", "edit.moveCells")))
    if src is None or to is None:
        raise EngineError("edit.moveCells: invalid range")
    if to[0] + src.height > MAX_ROWS or to[1] + src.width > MAX_COLS:
        raise EngineError("The data doesn't fit on the sheet.")
    dest = RangeRef(to[0], to[1], to[0] + src.height - 1, to[1] + src.width - 1)
    check_locked(d.wb, si, dest)
    if p_bool(p, "copy"):
        saved = s.clipboard
        s.clipboard = Clipboard(d.wb, si, src, False, d.uid, "")
        d.selection = Selection(to, to, [dest])
        try:
            return paste_special(s, {"what": "all"})
        finally:
            s.clipboard = saved
    check_locked(d.wb, si, src)

    def f(ctx):
        move_block(ctx, si, src, to[0], to[1])
        ctx.sel = Selection(to, to, [dest])

    edit(s, f)
    return {"range": dest.a1()}


# ---------------------------------------------------------------- clear


def _clear(s: Session, p: Dict[str, Any], what: str):
    d = s.doc()
    si = param_sheet(s, p)
    ranges = [param_range(s, p)] if p.get("range") else list(d.selection.ranges)
    sh = d.wb.sheet(si)
    for rng in ranges:
        if what in ("contents", "all"):
            if sh.protection is not None:
                used = sh.used_range()
                inter = rng.intersection(used) if used is not None else None
                if inter is not None:
                    check_locked(d.wb, si, inter)
            check_arrays(sh, rng)

    def f(ctx):
        sm = ctx.sheet_mut(si)
        for rng in ranges:
            for rc, cell in list(sm.cells.iter_range(rng)):
                if what == "contents":
                    if cell.value is not None or cell.formula is not None:
                        _set(ctx, si, rc, Cell(None, None, cell.style))
                elif what == "formats":
                    if cell.style != 0:
                        _set(ctx, si, rc, Cell(cell.value, cell.formula, 0))
                        ctx.fit_rows.append((si, rc[0]))
                elif what == "all":
                    _set(ctx, si, rc, None)
            if what == "contents":
                # Spilled values belong to their anchor; clearing them is a no-op like Excel.
                pass
            if what in ("formats", "all"):
                sm.merges = [m for m in sm.merges if not m.intersects(rng)]
                sm.cond_formats = remove_rules_in(sm.cond_formats, rng)
                if rng.is_full_rows() or rng.is_full_cols():
                    if rng.is_full_rows():
                        for r in [r for r in sm.rows if rng.r0 <= r <= rng.r1]:
                            info = sm.rows[r]
                            if info.style is not None:
                                sm.set_row(r, replace(info, style=None))
                    if rng.is_full_cols():
                        for c in [c for c in sm.cols if rng.c0 <= c <= rng.c1]:
                            info = sm.cols[c]
                            if info.style is not None:
                                sm.set_col(c, replace(info, style=None))
            if what in ("comments", "all"):
                for rc in [rc for rc in sm.comments if rng.contains(rc)]:
                    del sm.comments[rc]
            if what in ("hyperlinks", "all"):
                for rc in [rc for rc in sm.hyperlinks if rng.contains(rc)]:
                    del sm.hyperlinks[rc]

    edit(s, f)
    return None


@command("edit.clearAll", "Clear All", menu=("Home", "Editing", "Clear"), params="{range?}")
def clear_all(s, p):
    return _clear(s, p, "all")


@command("edit.clearFormats", "Clear Formats", menu=("Home", "Editing", "Clear"), params="{range?}")
def clear_formats(s, p):
    return _clear(s, p, "formats")


@command("edit.clearContents", "Clear Contents", menu=("Home", "Editing", "Clear"), shortcut="Delete", params="{range?}")
def clear_contents(s, p):
    return _clear(s, p, "contents")


@command("edit.clearComments", "Clear Comments and Notes", menu=("Home", "Editing", "Clear"), params="{range?}")
def clear_comments(s, p):
    return _clear(s, p, "comments")


@command("edit.clearHyperlinks", "Clear Hyperlinks", menu=("Home", "Editing", "Clear"), params="{range?}")
def clear_hyperlinks(s, p):
    return _clear(s, p, "hyperlinks")


# ---------------------------------------------------------------- fill


def _fill_dir(s: Session, p: Dict[str, Any], dr: int, dc: int):
    d = s.doc()
    si = param_sheet(s, p)
    rng = param_range(s, p)
    sh = d.wb.sheet(si)
    used = sh.used_range()
    if rng.is_full_cols() and used is not None:
        rng = RangeRef(rng.r0, rng.c0, max(used.r1, rng.r0), rng.c1)
    if rng.is_full_rows() and used is not None:
        rng = RangeRef(rng.r0, rng.c0, rng.r1, max(used.c1, rng.c0))
    if rng.count() > MAX_EDIT_CELLS:
        raise EngineError("range too large")
    if dr != 0:
        single = rng.height == 1
        if dr > 0:
            src = RangeRef(rng.r0 - 1, rng.c0, rng.r0 - 1, rng.c1) if single else RangeRef(rng.r0, rng.c0, rng.r0, rng.c1)
        else:
            src = RangeRef(rng.r1 + 1, rng.c0, rng.r1 + 1, rng.c1) if single else RangeRef(rng.r1, rng.c0, rng.r1, rng.c1)
    else:
        single = rng.width == 1
        if dc > 0:
            src = RangeRef(rng.r0, rng.c0 - 1, rng.r1, rng.c0 - 1) if single else RangeRef(rng.r0, rng.c0, rng.r1, rng.c0)
        else:
            src = RangeRef(rng.r0, rng.c1 + 1, rng.r1, rng.c1 + 1) if single else RangeRef(rng.r0, rng.c1, rng.r1, rng.c1)
    if src.r0 < 0 or src.c0 < 0 or src.r1 >= MAX_ROWS or src.c1 >= MAX_COLS:
        return None
    target = src.union(rng)
    check_locked(d.wb, si, rng)
    check_arrays(sh, rng)
    edit(s, lambda ctx: fill(ctx, si, src, target, "copy"))
    return {"range": rng.a1()}


@command("edit.fillDown", "Fill Down", menu=("Home", "Editing", "Fill"), shortcut="Ctrl+D", params="{range?}")
def fill_down(s, p):
    return _fill_dir(s, p, 1, 0)


@command("edit.fillRight", "Fill Right", menu=("Home", "Editing", "Fill"), shortcut="Ctrl+R", params="{range?}")
def fill_right(s, p):
    return _fill_dir(s, p, 0, 1)


@command("edit.fillUp", "Fill Up", menu=("Home", "Editing", "Fill"), params="{range?}")
def fill_up(s, p):
    return _fill_dir(s, p, -1, 0)


@command("edit.fillLeft", "Fill Left", menu=("Home", "Editing", "Fill"), params="{range?}")
def fill_left(s, p):
    return _fill_dir(s, p, 0, -1)


@command("edit.autoFill", "AutoFill", params='{source?: "A1:A2", target: "A1:A10", mode?: series|alternate|copy|formats|values}')
def auto_fill(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    src = RangeRef.parse(str(p["source"])) if p.get("source") else d.selection.current()
    target = RangeRef.parse(str(need(p, "target", "edit.autoFill")))
    if src is None or target is None:
        raise EngineError("edit.autoFill: invalid range")
    mode = p_str(p, "mode", "series")
    if target.count() > MAX_EDIT_CELLS:
        raise EngineError("range too large")
    sh = d.wb.sheet(si)
    if src.contains_range(target):
        # Dragging the handle back into the selection clears what it leaves.
        if target == src:
            return {"range": src.a1()}
        cut = RangeRef(target.r1 + 1, src.c0, src.r1, src.c1) if target.width == src.width else RangeRef(src.r0, target.c1 + 1, src.r1, src.c1)
        check_locked(d.wb, si, cut)

        def clr(ctx):
            sm = ctx.sheet_mut(si)
            for rc, cell in list(sm.cells.iter_range(cut)):
                if cell.value is not None or cell.formula is not None:
                    _set(ctx, si, rc, Cell(None, None, cell.style))
            ctx.sel = Selection(ctx.sel.active if target.contains(ctx.sel.active) else target.start, target.start, [target])

        edit(s, clr)
        return {"range": target.a1()}
    if not target.contains_range(src):
        raise EngineError("edit.autoFill: the target must extend the source")
    check_locked(d.wb, si, target)
    check_arrays(sh, target)

    def f(ctx):
        fill(ctx, si, src, target, mode)
        ctx.sel = Selection(ctx.sel.active, ctx.sel.active, [target]) if target.contains(ctx.sel.active) else Selection.of_range(target)

    edit(s, f)
    return {"range": target.a1()}


@command("edit.fillSeries", "Series…", menu=("Home", "Editing", "Fill"), params="{range?, rows?: bool, type?: linear|growth|date, step?: 1, stop?, unit?: day|weekday|month|year}")
def fill_series(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    rng = param_range(s, p)
    rows = p_bool(p, "rows", rng.width > rng.height)
    kind = p_str(p, "type", "linear")
    step = p_float(p, "step", 1.0)
    stop = p_float(p, "stop")
    unit = p_str(p, "unit", "day")
    if rng.count() > MAX_EDIT_CELLS:
        raise EngineError("range too large")
    check_locked(d.wb, si, rng)
    edit(s, lambda ctx: series(ctx, si, rng, rows, kind, step, stop, unit))
    return {"range": rng.a1()}


# ---------------------------------------------------------------- find & replace


def _pattern(what: str, match_case: bool, whole: bool) -> "re.Pattern":
    out = []
    i = 0
    while i < len(what):
        ch = what[i]
        if ch == "~" and i + 1 < len(what) and what[i + 1] in "*?~":
            out.append(re.escape(what[i + 1]))
            i += 2
            continue
        if ch == "*":
            out.append(".*")
        elif ch == "?":
            out.append(".")
        else:
            out.append(re.escape(ch))
        i += 1
    body = "".join(out)
    if whole:
        body = "^(?:" + body + ")$"
    return re.compile(body, (0 if match_case else re.IGNORECASE) | re.DOTALL)


def _search_text(wb, sh, rc, cell, look_in: str) -> Optional[str]:
    if look_in == "comments":
        c = sh.comments.get(rc)
        return c.text if c is not None else None
    if look_in == "formulas":
        if cell is None:
            return None
        return cell.input_text() if cell.value is not None or cell.formula is not None else None
    if cell is None and rc not in sh.spill:
        return None
    if sh.value(rc) is None:
        return None
    return value_text(wb, sh, rc)


def _find_all(wb, sheets: List[int], pat, look_in: str, by_cols: bool, limit: int = 10000):
    out = []
    for si in sheets:
        sh = wb.sheets[si]
        if look_in == "comments":
            keys = sorted(sh.comments)
        else:
            keys = [rc for rc, _ in sh.cells.iter_sorted()]
            if look_in == "values" and sh.spill:
                keys = sorted(set(keys) | set(sh.spill))
        if by_cols:
            keys.sort(key=lambda rc: (rc[1], rc[0]))
        for rc in keys:
            t = _search_text(wb, sh, rc, sh.cell(rc), look_in)
            if t is not None and pat.search(t):
                out.append((si, rc))
                if len(out) >= limit:
                    return out
    return out


def _find_args(s: Session, p: Dict[str, Any], cid: str):
    what = p.get("what")
    if what is None or what == "":
        raise EngineError(f"{cid}: missing `what`")
    pat = _pattern(str(what), p_bool(p, "matchCase"), p_bool(p, "wholeCell"))
    look_in = p_str(p, "lookIn", "formulas")
    by_cols = p_bool(p, "byColumns")
    d = s.doc()
    if p_str(p, "within", "sheet") == "workbook":
        n = len(d.wb.sheets)
        a = d.wb.active_sheet
        sheets = [(a + i) % n for i in range(n) if d.wb.sheets[(a + i) % n].visibility == "visible"]
    else:
        sheets = [d.wb.active_sheet]
    return pat, look_in, by_cols, sheets


@command("edit.find", "Find…", menu=("Home", "Editing", "Find & Select"), shortcut="Ctrl+F",
         params="{what, matchCase?, wholeCell?, lookIn?: formulas|values|comments, byColumns?, all?: bool, within?: sheet|workbook, previous?: bool}", undoable=False)
def find(s: Session, p: Dict[str, Any]):
    d = s.doc()
    pat, look_in, by_cols, sheets = _find_args(s, p, "edit.find")
    s.last_find = dict(p)
    hits = _find_all(d.wb, sheets, pat, look_in, by_cols)
    if p_bool(p, "all"):
        return {
            "matches": [
                {"sheet": d.wb.sheets[si].name, "cell": a1(*rc), "value": value_text(d.wb, d.wb.sheets[si], rc), "formula": d.wb.sheets[si].input_text(rc)}
                for si, rc in hits
            ]
        }
    if not hits:
        raise EngineError("We couldn't find what you were looking for.")
    cur = (sheets.index(d.wb.active_sheet) if d.wb.active_sheet in sheets else 0, d.selection.active)

    def key(h):
        si, rc = h
        k = (rc[1], rc[0]) if by_cols else rc
        return (sheets.index(si), k)

    ck = (cur[0], (cur[1][1], cur[1][0]) if by_cols else cur[1])
    if p_bool(p, "previous"):
        before = [h for h in hits if key(h) < ck]
        hit = max(before, key=key) if before else max(hits, key=key)
    else:
        after = [h for h in hits if key(h) > ck]
        hit = min(after, key=key) if after else min(hits, key=key)
    si, rc = hit
    _activate(s, si)
    sel = d.selection
    if not sel.is_single_cell() and sel.contains(rc) and si == d.wb.active_sheet and len(hits) > 1:
        d.selection = Selection(rc, sel.anchor, sel.ranges)
    else:
        d.selection = Selection.at(rc)
    return {"sheet": d.wb.sheets[si].name, "cell": a1(*rc), "count": len(hits)}


@command("edit.replace", "Replace…", menu=("Home", "Editing", "Find & Select"), shortcut="Ctrl+H",
         params="{what, with, all?: bool, matchCase?, wholeCell?, within?: sheet|workbook, byColumns?}")
def replace_cmd(s: Session, p: Dict[str, Any]):
    d = s.doc()
    p = dict(p)
    p["lookIn"] = "formulas"
    pat, _look, by_cols, sheets = _find_args(s, p, "edit.replace")
    repl = str(p.get("with", ""))
    hits = _find_all(d.wb, sheets, pat, "formulas", by_cols, limit=1_000_000)
    if not p_bool(p, "all"):
        a = d.selection.active
        here = [(si, rc) for si, rc in hits if si == d.wb.active_sheet and rc == a]
        if not here:
            return find(s, {**p, "lookIn": "formulas"})
        hits = here
    if not hits:
        raise EngineError("We couldn't find anything to replace.")
    for si, rc in hits:
        sh = d.wb.sheets[si]
        if sh.protection is not None and d.wb.styles.get(sh.style_id(rc)).protection.locked:
            raise EngineError(PROTECTED)
    count = 0

    def f(ctx):
        nonlocal count
        wb = ctx.wb
        for si, rc in hits:
            sh = wb.sheets[si]
            cell = sh.cell(rc)
            if cell is None:
                continue
            old = cell.input_text()
            new = pat.sub(lambda m: repl, old)
            if new == old:
                continue
            try:
                nc = input_to_cell(wb, sh, rc, new)
            except EngineError:
                continue
            _set(ctx, si, rc, nc)
            count += 1

    edit(s, f)
    if not p_bool(p, "all"):
        try:
            find(s, {**p, "lookIn": "formulas"})
        except EngineError:
            pass
    return {"replaced": count}


# ---------------------------------------------------------------- format painter


@command("edit.formatPainter", "Format Painter", menu=("Home", "Clipboard"), params="{range?, keep?: bool (double-click keeps it on)}", undoable=False)
def format_painter(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    rng = param_range(s, p)
    s.format_painter = (d.wb, si, rng, p_bool(p, "keep"))
    return {"range": rng.a1()}


@command("edit.applyFormatPainter", "Paint Formats", params="{range?}")
def apply_format_painter(s: Session, p: Dict[str, Any]):
    fp = s.format_painter
    if fp is None:
        raise EngineError("Pick up a format with the Format Painter first.")
    wb0, si0, src, keep = fp
    d = s.doc()
    si = param_sheet(s, p)
    at = param_range(s, p)
    saved = s.clipboard
    s.clipboard = Clipboard(wb0, si0, src, False, d.uid if wb0 is d.wb else -1, "")
    try:
        d.selection = Selection(at.start, at.start, [at if not at.is_single() else RangeRef(at.r0, at.c0, at.r0 + src.height - 1, at.c0 + src.width - 1)])
        r = paste_special(s, {"what": "formats", "sheet": si})
    finally:
        s.clipboard = saved
    if not keep:
        s.format_painter = None
    return r


@command("edit.cancelFormatPainter", "Cancel Format Painter", undoable=False, journal=False)
def cancel_format_painter(s: Session, p: Dict[str, Any]):
    s.format_painter = None
    return None
