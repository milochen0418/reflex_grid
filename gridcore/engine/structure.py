"""Structural edits: inserting and deleting rows, columns and cells, moving blocks. Everything
that points at cells follows: formulas on every sheet, defined names, merges, conditional
formats, validation, tables, AutoFilter, notes, hyperlinks, charts and frozen panes."""

from __future__ import annotations

from dataclasses import replace
from typing import Dict, List, Optional, Tuple

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef
from ..formula import ParseError, parse, print_formula
from ..formula.adjust import Edit, adjust
from ..model.cell import Cell, Formula
from ..model.features import LineInfo


def _adjust_formula_text(text: str, host: str, target: str, edit: Edit) -> str:
    body = text[1:] if text.startswith("=") else text
    try:
        e = parse(body)
    except ParseError:
        return text
    out = print_formula(adjust(e, host, target, edit))
    return ("=" + out) if text.startswith("=") else out


def _map_index(v: int, at: int, count: int, insert: bool, mx: int) -> Optional[int]:
    if insert:
        if v >= at:
            n = v + count
            return None if n >= mx else n
        return v
    if v < at:
        return v
    if v >= at + count:
        return v - count
    return None


def _map_range(r: RangeRef, axis: str, at: int, count: int, insert: bool) -> Optional[RangeRef]:
    """A range after inserting/deleting lines; ``None`` when it is deleted entirely."""
    if axis == "rows":
        lo, hi, mx = r.r0, r.r1, MAX_ROWS
    else:
        lo, hi, mx = r.c0, r.c1, MAX_COLS
    if insert:
        lo2 = lo + count if lo >= at else lo
        hi2 = min(hi + count if hi >= at else hi, mx - 1)
        if lo2 > hi2:
            return None
    else:
        end = at + count
        if lo >= at and hi < end:
            return None
        lo2 = lo if lo < at else (lo - count if lo >= end else at)
        hi2 = hi if hi < at else (hi - count if hi >= end else max(at - 1, 0))
        hi2 = max(hi2, lo2)
    if axis == "rows":
        return RangeRef(lo2, r.c0, hi2, r.c1)
    return RangeRef(r.r0, lo2, r.r1, hi2)


def _adjust_everything(ctx, si: int, edit: Edit) -> None:
    """Formulas on every sheet and defined names follow ``edit`` on sheet ``si``."""
    wb = ctx.wb
    target = wb.sheets[si].name
    for i, sh in enumerate(wb.sheets):
        host = sh.name
        changes = []
        for rc, cell in sh.cells.items():
            if cell.formula is None:
                continue
            e = cell.formula.get_expr()
            if e is None:
                continue
            ne = adjust(e, host, target, edit)
            if ne != e:
                changes.append((rc, Cell(cell.value, Formula.from_expr(ne, cell.formula.array), cell.style)))
        if changes:
            sm = wb.sheet_mut(i)
            for rc, c in changes:
                sm.cells.set(rc, c)
        # Conditional-format and validation formulas are relative to their ranges; references
        # to other cells in them follow the edit.
        if any(cf.rule.formulas for cf in sh.cond_formats) or any(v.f1 or v.f2 for v in sh.validations):
            sm = wb.sheet_mut(i)
            for cf in sm.cond_formats:
                cf.rule.formulas = [_adjust_formula_text(f, host, target, edit) for f in cf.rule.formulas]
            for v in sm.validations:
                if v.f1 and not v.f1.startswith('"'):
                    v.f1 = _adjust_formula_text(v.f1, host, target, edit)
                if v.f2:
                    v.f2 = _adjust_formula_text(v.f2, host, target, edit)
        if sh.charts:
            sm = wb.sheet_mut(i)
            for ch in sm.charts:
                for ser in ch.series:
                    if ser.values:
                        ser.values = _adjust_formula_text(ser.values, host, target, edit)
                    if ser.categories:
                        ser.categories = _adjust_formula_text(ser.categories, host, target, edit)
                    if ser.name and not ser.name.startswith('"'):
                        ser.name = _adjust_formula_text(ser.name, host, target, edit)
    for d in wb.names:
        host = wb.sheets[d.scope].name if d.scope is not None and d.scope < len(wb.sheets) else target
        # Names hold absolute references qualified by sheet; unqualified ones follow the scope.
        try:
            e = parse(d.formula)
        except ParseError:
            continue
        ne = adjust(e, host if d.scope is not None else "\u0000", target, edit)
        if ne != e:
            d.formula = print_formula(ne)


def shift_lines(ctx, si: int, axis: str, at: int, count: int, insert: bool) -> None:
    """Inserts (``insert``) or deletes ``count`` rows/columns at index ``at`` of sheet ``si``."""
    edit = Edit("insert" if insert else "delete", axis=axis, at=at, count=count)
    _adjust_everything(ctx, si, edit)
    sm = ctx.sheet_mut(si)
    rows = axis == "rows"
    # Cells.
    moved: List[Tuple[Tuple[int, int], Cell]] = []
    for rc, cell in list(sm.cells.items()):
        v = rc[0] if rows else rc[1]
        if v < at:
            continue
        moved.append((rc, cell))
    for rc, _ in moved:
        sm.cells.remove(rc)
    for rc, cell in moved:
        v = rc[0] if rows else rc[1]
        nv = _map_index(v, at, count, insert, MAX_ROWS if rows else MAX_COLS)
        if nv is None:
            continue
        nrc = (nv, rc[1]) if rows else (rc[0], nv)
        sm.cells.set(nrc, cell)
    # Spills are laid out again by the full recalculation.
    sm.spill = {}
    sm.spill_ranges = {}
    # Line info.
    lines = sm.rows if rows else sm.cols
    new_lines: Dict[int, LineInfo] = {}
    for k, info in lines.items():
        nk = _map_index(k, at, count, insert, MAX_ROWS if rows else MAX_COLS)
        if nk is not None:
            new_lines[nk] = info
    if insert and at > 0 and (at - 1) in lines:
        # New lines take the format of the line before them (Excel's Insert Options default).
        prev = lines[at - 1]
        if prev.style is not None or prev.size is not None:
            for k in range(at, min(at + count, MAX_ROWS if rows else MAX_COLS)):
                new_lines[k] = LineInfo(size=prev.size, style=prev.style, custom=prev.custom)
    if rows:
        sm.rows = new_lines
    else:
        sm.cols = new_lines
    sm._invalidate_geometry()

    def mr(r: RangeRef) -> Optional[RangeRef]:
        return _map_range(r, axis, at, count, insert)

    sm.merges = [m for m in (mr(m) for m in sm.merges) if m is not None and not m.is_single()]
    for cf in sm.cond_formats:
        cf.ranges = [r for r in (mr(r) for r in cf.ranges) if r is not None]
    sm.cond_formats = [cf for cf in sm.cond_formats if cf.ranges]
    for v in sm.validations:
        v.ranges = [r for r in (mr(r) for r in v.ranges) if r is not None]
    sm.validations = [v for v in sm.validations if v.ranges]
    tables = []
    for t in sm.tables:
        nr = mr(t.range)
        if nr is None:
            continue
        if not rows and nr.width != t.range.width:
            # Columns inserted or deleted inside a table add or drop table columns.
            if insert and t.range.c0 < at <= t.range.c1:
                pos = at - t.range.c0
                from ..model.features import TableColumn

                used = {c.name.lower() for c in t.columns}
                new_cols = []
                k = 1
                for _ in range(count):
                    while f"column{k}" in used:
                        k += 1
                    used.add(f"column{k}")
                    new_cols.append(TableColumn(f"Column{k}"))
                t.columns = t.columns[:pos] + new_cols + t.columns[pos:]
            elif not insert:
                lo = max(at, t.range.c0) - t.range.c0
                hi = min(at + count - 1, t.range.c1) - t.range.c0
                t.columns = t.columns[:lo] + t.columns[hi + 1:]
        t.range = nr
        tables.append(t)
    sm.tables = tables
    if sm.autofilter is not None:
        nr = mr(sm.autofilter.range)
        sm.autofilter = None if nr is None else replace(sm.autofilter, range=nr)
    for name in ("comments", "hyperlinks"):
        d = getattr(sm, name)
        nd = {}
        for rc, x in d.items():
            v = rc[0] if rows else rc[1]
            nv = _map_index(v, at, count, insert, MAX_ROWS if rows else MAX_COLS)
            if nv is not None:
                nd[(nv, rc[1]) if rows else (rc[0], nv)] = x
        setattr(sm, name, nd)
    for obj in list(sm.charts) + list(sm.images):
        a = obj.anchor
        if a.mode == "absolute":
            continue
        v = a.row if rows else a.col
        nv = _map_index(v, at, count, insert, MAX_ROWS if rows else MAX_COLS)
        if nv is None:
            nv = at
        if rows:
            a.row = nv
        else:
            a.col = nv
    if sm.freeze is not None:
        fr, fc = sm.freeze
        if rows and at < fr:
            fr = max(fr + count, 0) if insert else max(fr - min(count, fr - at), 0)
        if not rows and at < fc:
            fc = max(fc + count, 0) if insert else max(fc - min(count, fc - at), 0)
        sm.freeze = (fr, fc) if fr or fc else None
    if sm.print.print_area is not None:
        sm.print.print_area = mr(sm.print.print_area)
    ctx.structural = True
    ctx.moved = True


def shift_cells(ctx, si: int, rng: RangeRef, axis: str, insert: bool) -> None:
    """Insert Cells shifting down/right (``axis`` rows/cols) or Delete Cells shifting up/left."""
    edit = Edit("insert_cells" if insert else "delete_cells", axis=axis, range=rng)
    _adjust_everything(ctx, si, edit)
    sm = ctx.sheet_mut(si)
    rows = axis == "rows"
    span = rng.height if rows else rng.width
    moved = []
    for rc, cell in list(sm.cells.items()):
        if rows:
            if rng.c0 <= rc[1] <= rng.c1 and rc[0] >= rng.r0:
                moved.append((rc, cell))
        else:
            if rng.r0 <= rc[0] <= rng.r1 and rc[1] >= rng.c0:
                moved.append((rc, cell))
    for rc, _ in moved:
        sm.cells.remove(rc)
    for rc, cell in moved:
        if rows:
            if insert:
                nr = rc[0] + span
            else:
                if rc[0] <= rng.r1:
                    continue
                nr = rc[0] - span
            if 0 <= nr < MAX_ROWS:
                sm.cells.set((nr, rc[1]), cell)
        else:
            if insert:
                nc = rc[1] + span
            else:
                if rc[1] <= rng.c1:
                    continue
                nc = rc[1] - span
            if 0 <= nc < MAX_COLS:
                sm.cells.set((rc[0], nc), cell)
    sm.spill = {}
    sm.spill_ranges = {}
    merges = []
    for m in sm.merges:
        inside = (rng.c0 <= m.c0 and m.c1 <= rng.c1 and m.r0 >= rng.r0) if rows else (rng.r0 <= m.r0 and m.r1 <= rng.r1 and m.c0 >= rng.c0)
        if not inside:
            merges.append(m)
            continue
        if rows:
            if not insert and m.r0 <= rng.r1:
                continue
            d = span if insert else -span
            if m.r1 + d < MAX_ROWS:
                merges.append(RangeRef(m.r0 + d, m.c0, m.r1 + d, m.c1))
        else:
            if not insert and m.c0 <= rng.c1:
                continue
            d = span if insert else -span
            if m.c1 + d < MAX_COLS:
                merges.append(RangeRef(m.r0, m.c0 + d, m.r1, m.c1 + d))
    sm.merges = merges
    for name in ("comments", "hyperlinks"):
        d = getattr(sm, name)
        nd = {}
        for rc, x in d.items():
            inside = (rng.c0 <= rc[1] <= rng.c1 and rc[0] >= rng.r0) if rows else (rng.r0 <= rc[0] <= rng.r1 and rc[1] >= rng.c0)
            if not inside:
                nd[rc] = x
                continue
            if rows:
                if not insert and rc[0] <= rng.r1:
                    continue
                n = (rc[0] + span, rc[1]) if insert else (rc[0] - span, rc[1])
            else:
                if not insert and rc[1] <= rng.c1:
                    continue
                n = (rc[0], rc[1] + span) if insert else (rc[0], rc[1] - span)
            if 0 <= n[0] < MAX_ROWS and 0 <= n[1] < MAX_COLS:
                nd[n] = x
        setattr(sm, name, nd)
    ctx.structural = True
    ctx.moved = True


def move_block(ctx, si: int, src: RangeRef, dst_row: int, dst_col: int, dst_sheet: Optional[int] = None) -> None:
    """Cut-paste of a block: the cells move, and every reference into the block follows."""
    dsi = si if dst_sheet is None else dst_sheet
    wb = ctx.wb
    if dsi == si:
        edit = Edit("move", range=src, to_row=dst_row, to_col=dst_col)
    else:
        edit = Edit("move_to_sheet", range=src, to_row=dst_row, to_col=dst_col, to_sheet=wb.sheets[dsi].name)
    _adjust_everything(ctx, si, edit)
    sm = ctx.sheet_mut(si)
    taken = sm.cells.take_range(src)
    dm = ctx.sheet_mut(dsi)
    dr, dc = dst_row - src.r0, dst_col - src.c0
    dst = RangeRef(src.r0 + dr, src.c0 + dc, src.r1 + dr, src.c1 + dc)
    for rc, _cell in list(dm.cells.iter_range(dst)):
        dm.cells.remove(rc)
    for rc, cell in taken:
        n = (rc[0] + dr, rc[1] + dc)
        if 0 <= n[0] < MAX_ROWS and 0 <= n[1] < MAX_COLS:
            # A formula inside the moved block keeps pointing at what it pointed at: references
            # into the block were adjusted above; others stay.
            dm.cells.set(n, cell)
            ctx.touch(dsi, n)
        ctx.touch(si, rc)
    merges = [m for m in sm.merges if src.contains_range(m)]
    if merges:
        sm.merges = [m for m in sm.merges if m not in merges]
        dm.merges = [m for m in dm.merges if not m.intersects(dst)] + [RangeRef(m.r0 + dr, m.c0 + dc, m.r1 + dr, m.c1 + dc) for m in merges]
    for name in ("comments", "hyperlinks"):
        sd = getattr(sm, name)
        moved = {rc: x for rc, x in sd.items() if src.contains(rc)}
        if moved:
            for rc in moved:
                del sd[rc]
            dd = getattr(dm, name)
            for rc, x in moved.items():
                dd[(rc[0] + dr, rc[1] + dc)] = x
    ctx.structural = True
    ctx.moved = True
