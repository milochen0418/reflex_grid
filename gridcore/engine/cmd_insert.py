"""Insert and Review commands: tables, charts, pictures, hyperlinks, notes and comments,
sparklines and floating objects."""

from __future__ import annotations

import base64
import struct
from dataclasses import replace
from typing import Any, Dict, List, Optional, Tuple

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef, a1, parse_cell
from ..formula.printer import quote_sheet
from ..model.cell import Cell, Formula
from ..model.features import CHART_KINDS, Anchor, Chart, Comment, Hyperlink, Image, Series, Sparkline, Table, TableColumn
from ..model.style import Color
from .cmd_data import guess_header
from .cmd_edit import _activate, _set, resolve_reference
from .cmd_format import _ranges
from .display import value_text
from .selection import Selection, current_region
from .session import EngineError, Session, command, edit, need, p_bool, p_float, p_int, p_str, param_cell, param_sheet
from .structure import shift_cells
from .tables import TABLE_STYLE_NAMES, totals_formula, unique_column_name


def _region(s: Session, p: Dict[str, Any]) -> RangeRef:
    d = s.doc()
    if p.get("range"):
        return _ranges(s, p)[0]
    sel = d.selection.current()
    if not sel.is_single():
        return sel
    return current_region(d.sheet, sel.start)


# ---------------------------------------------------------------- tables


def next_table_name(wb) -> str:
    used = {t.name.lower() for sh in wb.sheets for t in sh.tables} | {dn.name.lower() for dn in wb.names}
    n = 1
    while f"table{n}" in used:
        n += 1
    return f"Table{n}"


@command("insert.table", "Table", menu=("Insert", "Tables"), shortcut="Ctrl+T", params="{range?, header?: bool, style?: TableStyleMedium2, name?}")
def insert_table(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)
    rng = _region(s, p)
    if rng.is_full_cols() or rng.is_full_rows():
        used = sh.used_range()
        if used is None:
            raise EngineError("Select a range with data first.")
        rng = rng.intersection(used) or rng
    if rng.count() > 50_000_000:
        raise EngineError("The range is too large for a table.")
    if any(t.range.intersects(rng) for t in sh.tables):
        raise EngineError("A table can't overlap another table.")
    if any(m.intersects(rng) for m in sh.merges):
        raise EngineError("A table can't overlap a range that contains merged cells.")
    header = p_bool(p, "header") if "header" in p else (guess_header(sh, rng, d.wb) or all(isinstance(sh.value((rng.r0, c)), str) for c in range(rng.c0, rng.c1 + 1)))
    style = p_str(p, "style", "TableStyleMedium2")
    if style not in TABLE_STYLE_NAMES and style != "None":
        raise EngineError(f"unknown table style `{style}`")
    name = p_str(p, "name") or next_table_name(d.wb)
    from .cmd_formulas import check_name

    if check_name(name) is not None or d.wb.table(name) is not None or any(dn.name.lower() == name.lower() for dn in d.wb.names):
        raise EngineError("The table name is not valid or is already used.")

    def f(ctx):
        wb = ctx.wb
        r = rng
        if not header:
            # A header row is added above the data.
            shift_cells(ctx, si, RangeRef(r.r0, r.c0, r.r0, r.c1), "rows", True)
            r = RangeRef(r.r0, r.c0, r.r1 + 1, r.c1)
        sm = ctx.sheet_mut(si)
        t = Table(wb.next_object_id(), name, r, header_row=True, style=style if style != "None" else "TableStyleLight1")
        cols: List[TableColumn] = []
        for i in range(r.width):
            rc = (r.r0, r.c0 + i)
            text = value_text(wb, sm, rc).replace("\n", " ").strip() if (header and sm.value(rc) is not None) else ""
            base = text or f"Column{i + 1}"
            name_i, k = base, 2
            while name_i.lower() in {c.name.lower() for c in cols}:
                name_i = f"{base}{k}"
                k += 1
            cols.append(TableColumn(name_i))
            cell = sm.cell(rc)
            sm.set_cell(rc, Cell(name_i, None, cell.style if cell is not None else sm.style_id(rc)))
            ctx.touch(si, rc)
        t.columns = cols
        sm.tables.append(t)
        if sm.autofilter is not None and sm.autofilter.range.intersects(r):
            sm.autofilter = None
        ctx.sel = Selection(ctx.sel.active, ctx.sel.active, [r]) if r.contains(ctx.sel.active) else Selection.of_range(r)
        ctx.structural = True

    edit(s, f)
    return {"name": name, "range": rng.a1()}


def _table(s: Session, p: Dict[str, Any]) -> Tuple[int, Table]:
    d = s.doc()
    if p.get("table"):
        loc = d.wb.table(str(p["table"]))
        if loc is None:
            raise EngineError(f"no table `{p['table']}`")
        return loc[0], d.wb.sheets[loc[0]].tables[loc[1]]
    si = d.wb.active_sheet
    t = d.sheet.table_at(d.selection.active)
    if t is None:
        raise EngineError("Select a cell in a table first.")
    return si, t


def _table_edit(s: Session, p: Dict[str, Any], fn):
    si, t = _table(s, p)
    tid = t.id

    def f(ctx):
        sm = ctx.sheet_mut(si)
        tt = next(x for x in sm.tables if x.id == tid)
        fn(ctx, sm, tt)
        ctx.structural = True

    edit(s, f)
    return {"table": t.name}


def _flag(attr: str):
    def run(s, p):
        si, t = _table(s, p)
        on = p_bool(p, "on", not getattr(t, attr))

        def fn(ctx, sm, tt):
            setattr(tt, attr, on)

        return _table_edit(s, p, fn)

    return run


command("table.bandedRows", "Banded Rows", menu=("Table Design", "Table Style Options"), params="{on?, table?}")(_flag("banded_rows"))
command("table.bandedColumns", "Banded Columns", menu=("Table Design", "Table Style Options"), params="{on?, table?}")(_flag("banded_cols"))
command("table.firstColumn", "First Column", menu=("Table Design", "Table Style Options"), params="{on?, table?}")(_flag("first_col"))
command("table.lastColumn", "Last Column", menu=("Table Design", "Table Style Options"), params="{on?, table?}")(_flag("last_col"))
command("table.filterButton", "Filter Button", menu=("Table Design", "Table Style Options"), params="{on?, table?}")(_flag("filter_button"))


@command("table.totalRow", "Total Row", menu=("Table Design", "Table Style Options"), shortcut="Ctrl+Shift+T", params="{on?, table?}")
def total_row(s, p):
    si, t = _table(s, p)
    on = p_bool(p, "on", not t.totals_row)
    if on == t.totals_row:
        return {"table": t.name}

    def fn(ctx, sm, tt):
        r = tt.range
        if on:
            below = RangeRef(r.r1 + 1, r.c0, r.r1 + 1, r.c1)
            if below.r1 >= MAX_ROWS:
                raise EngineError("There's no room for a total row.")
            if any(sm.cell(rc) is not None and (sm.cell(rc).value is not None or sm.cell(rc).formula is not None) for rc in below.cells()):
                shift_cells(ctx, si, below, "rows", True)
                sm = ctx.sheet_mut(si)
                tt = next(x for x in sm.tables if x.id == t.id)
            tt.range = RangeRef(r.r0, r.c0, r.r1 + 1, r.c1)
            tt.totals_row = True
            row = tt.range.r1
            # Label in the first column, a sum (or count) in the last.
            if tt.columns:
                tt.columns[0].totals = "none"
                tt.columns[0].totals_label = "Total"
                last = len(tt.columns) - 1
                data = tt.data_range()
                numeric = data is not None and any(isinstance(sm.value((rr, r.c0 + last)), float) for rr in range(data.r0, min(data.r1, data.r0 + 1000) + 1))
                if last > 0:
                    tt.columns[last].totals = "sum" if numeric else "count"
            for i, col in enumerate(tt.columns):
                rc = (row, r.c0 + i)
                fm = totals_formula(tt, i)
                if i == 0 and col.totals == "none" and col.totals_label:
                    sm.set_cell(rc, Cell(col.totals_label, None, sm.style_id(rc)))
                elif fm:
                    sm.set_cell(rc, Cell(None, Formula.new(fm), sm.style_id(rc)))
                ctx.touch(si, rc)
        else:
            row = r.r1
            for c in range(r.c0, r.c1 + 1):
                cell = sm.cell((row, c))
                if cell is not None:
                    _set(ctx, si, (row, c), Cell(None, None, cell.style))
            tt.range = RangeRef(r.r0, r.c0, r.r1 - 1, r.c1)
            tt.totals_row = False

    return _table_edit(s, p, fn)


@command("table.totalFunction", "Total Row Function", params="{col: offset | name, function: none|sum|average|count|countNums|max|min|stdDev|var, table?}")
def total_function(s, p):
    si, t = _table(s, p)
    if not t.totals_row:
        raise EngineError("Turn on the total row first.")
    colp = need(p, "col", "table.totalFunction")
    i = t.column_index(str(colp)) if isinstance(colp, str) and not str(colp).isdigit() else int(colp)
    if i is None or not 0 <= i < len(t.columns):
        raise EngineError("no such table column")
    fn_name = p_str(p, "function", "sum")

    def fn(ctx, sm, tt):
        tt.columns[i].totals = fn_name
        rc = (tt.range.r1, tt.range.c0 + i)
        fm = totals_formula(tt, i)
        _set(ctx, si, rc, Cell(None, Formula.new(fm), sm.style_id(rc)) if fm else Cell(None, None, sm.style_id(rc)))

    return _table_edit(s, p, fn)


@command("table.headerRow", "Header Row", menu=("Table Design", "Table Style Options"), params="{on?, table?}")
def header_row(s, p):
    si, t = _table(s, p)
    on = p_bool(p, "on", not t.header_row)
    if on == t.header_row:
        return {"table": t.name}

    def fn(ctx, sm, tt):
        r = tt.range
        if not on:
            for c in range(r.c0, r.c1 + 1):
                cell = sm.cell((r.r0, c))
                if cell is not None:
                    _set(ctx, si, (r.r0, c), Cell(None, None, cell.style))
            tt.range = RangeRef(r.r0 + 1, r.c0, r.r1, r.c1)
            tt.header_row = False
        else:
            if r.r0 == 0:
                raise EngineError("There's no room for a header row.")
            above = RangeRef(r.r0 - 1, r.c0, r.r0 - 1, r.c1)
            for i, rc in enumerate(above.cells()):
                _set(ctx, si, rc, Cell(tt.columns[i].name if i < len(tt.columns) else f"Column{i + 1}", None, sm.style_id(rc)))
            tt.range = RangeRef(r.r0 - 1, r.c0, r.r1, r.c1)
            tt.header_row = True

    return _table_edit(s, p, fn)


@command("table.style", "Table Styles", menu=("Table Design", "Table Styles"), params="{style: TableStyleMedium2, table?}")
def table_style(s, p):
    style = p_str(p, "style", "TableStyleMedium2")
    if style not in TABLE_STYLE_NAMES:
        raise EngineError(f"unknown table style `{style}`")

    def fn(ctx, sm, tt):
        tt.style = style

    return _table_edit(s, p, fn)


@command("table.rename", "Table Name", menu=("Table Design", "Properties"), params="{name, table?}")
def table_rename(s, p):
    from .cmd_formulas import check_name

    d = s.doc()
    new = str(need(p, "name", "table.rename")).strip()
    si, t = _table(s, p)
    if check_name(new) is not None:
        raise EngineError("The name that you entered is not valid.")
    loc = d.wb.table(new)
    if (loc is not None and d.wb.sheets[loc[0]].tables[loc[1]].id != t.id) or any(dn.name.lower() == new.lower() for dn in d.wb.names):
        raise EngineError("The name that you entered already exists. Enter a unique name.")
    old = t.name

    def fn(ctx, sm, tt):
        tt.name = new
        # Structured references follow the new name.
        import re

        pat = re.compile(r"(?<![A-Za-z0-9_.])" + re.escape(old) + r"(?=\[)", re.IGNORECASE)
        for i, shx in enumerate(ctx.wb.sheets):
            changes = []
            for rc, cell in shx.cells.items():
                if cell.formula is not None and old.lower() in cell.formula.text.lower():
                    nt = pat.sub(new, cell.formula.text)
                    if nt != cell.formula.text:
                        changes.append((rc, Cell(cell.value, Formula.new(nt), cell.style)))
            if changes:
                m = ctx.wb.sheet_mut(i)
                for rc, c in changes:
                    m.cells.set(rc, c)

    return _table_edit(s, p, fn)


@command("table.convertToRange", "Convert to Range", menu=("Table Design", "Tools"), params="{table?}")
def convert_to_range(s, p):
    si, t = _table(s, p)
    tid = t.id

    def f(ctx):
        sm = ctx.sheet_mut(si)
        sm.tables = [x for x in sm.tables if x.id != tid]
        ctx.structural = True

    edit(s, f)
    return None


@command("table.resize", "Resize Table", menu=("Table Design", "Properties"), params='{range: "A1:D20", table?}')
def resize_table(s, p):
    si, t = _table(s, p)
    rng = RangeRef.parse(str(need(p, "range", "table.resize")))
    if rng is None or rng.r0 != t.range.r0 or rng.c0 > t.range.c1 or rng.c1 < t.range.c0:
        raise EngineError("The headers must remain in the same row, and the resulting table range must overlap the original table.")

    def fn(ctx, sm, tt):
        cols = list(tt.columns[: rng.width])
        tt.range = rng
        from .tables import new_column_name

        for i in range(len(cols), rng.width):
            cols.append(TableColumn(new_column_name(c.name for c in cols)))
        tt.columns = cols
        if tt.header_row:
            for i in range(rng.width):
                rc = (rng.r0, rng.c0 + i)
                v = sm.value(rc)
                name = value_text(ctx.wb, sm, rc) if v is not None else cols[i].name
                cols[i].name = unique_column_name(tt, name, i)
                _set(ctx, si, rc, Cell(cols[i].name, None, sm.style_id(rc)))

    return _table_edit(s, p, fn)


# ---------------------------------------------------------------- charts


def build_series(wb, si: int, rng: RangeRef, by_rows: Optional[bool], kind: str) -> Tuple[List[Series], bool]:
    sh = wb.sheet(si)
    sheet = quote_sheet(sh.name)

    def is_label(v) -> bool:
        return v is None or isinstance(v, str)

    top = rng.height > 1 and all(is_label(sh.value((rng.r0, c))) for c in range(rng.c0 + (1 if rng.width > 1 else 0), rng.c1 + 1)) and any(sh.value((rng.r0, c)) is not None for c in range(rng.c0, rng.c1 + 1))
    left = rng.width > 1 and all(is_label(sh.value((r, rng.c0))) for r in range(rng.r0 + (1 if top else 0), rng.r1 + 1))
    if kind.startswith("scatter") and rng.width > 1:
        left = True
    data = RangeRef(rng.r0 + (1 if top else 0), rng.c0 + (1 if left else 0), rng.r1, rng.c1)
    if data.r0 > data.r1 or data.c0 > data.c1:
        raise EngineError("Select a range with numbers to chart.")
    if by_rows is None:
        by_rows = data.width > data.height and not kind.startswith("scatter")
    out = []
    if not by_rows:
        cats = f"{sheet}!{RangeRef(data.r0, rng.c0, data.r1, rng.c0).abs_a1()}" if left else None
        for c in range(data.c0, data.c1 + 1):
            name = f"{sheet}!{RangeRef(rng.r0, c, rng.r0, c).abs_a1()}" if top else None
            out.append(Series(name=name, categories=cats, values=f"{sheet}!{RangeRef(data.r0, c, data.r1, c).abs_a1()}"))
    else:
        cats = f"{sheet}!{RangeRef(rng.r0, data.c0, rng.r0, data.c1).abs_a1()}" if top else None
        for r in range(data.r0, data.r1 + 1):
            name = f"{sheet}!{RangeRef(r, rng.c0, r, rng.c0).abs_a1()}" if left else None
            out.append(Series(name=name, categories=cats, values=f"{sheet}!{RangeRef(r, data.c0, r, data.c1).abs_a1()}"))
    if len(out) > 255:
        raise EngineError("A chart can have at most 255 series.")
    if kind in ("pie", "doughnut"):
        out = out[:1] if kind == "pie" else out
    return out, by_rows


@command("insert.chart", "Chart", menu=("Insert", "Charts"), shortcut="Alt+F1",
         params="{kind?: columnClustered|columnStacked|columnStacked100|barClustered|barStacked|barStacked100|line|lineMarkers|lineStacked|pie|doughnut|area|areaStacked|scatter|scatterLines|radar, range?, title?, byRows?: bool}")
def insert_chart(s, p):
    d = s.doc()
    si = param_sheet(s, p)
    kind = p_str(p, "kind", "columnClustered")
    if kind not in CHART_KINDS:
        raise EngineError(f"unknown chart kind `{kind}`")
    rng = _region(s, p)
    sh = d.wb.sheet(si)
    used = sh.used_range()
    if used is None:
        raise EngineError("Select a range with data to chart.")
    rng = rng.intersection(used) or rng
    by_rows = p_bool(p, "byRows") if "byRows" in p else None
    series, by_rows = build_series(d.wb, si, rng, by_rows, kind)
    title = p_str(p, "title")
    if title is None and len(series) == 1 and series[0].name:
        title = None  # the series name shows as the title
    anchor = Anchor(row=rng.r0, col=min(rng.c1 + 2, MAX_COLS - 8), width=480.0, height=288.0)
    cid = d.wb.next_object_id()
    chart = Chart(id=cid, kind=kind, anchor=anchor, title=title, series=series, source=f"{quote_sheet(sh.name)}!{rng.abs_a1()}", by_rows=by_rows)
    chart.legend = "none" if len(series) == 1 and kind not in ("pie", "doughnut") else ("right" if kind in ("pie", "doughnut") else "bottom")

    def f(ctx):
        ctx.sheet_mut(si).charts.append(chart)

    edit(s, f)
    return {"id": cid, "kind": kind, "series": len(series)}


def _object(s: Session, oid: int):
    d = s.doc()
    for si, sh in enumerate(d.wb.sheets):
        for ch in sh.charts:
            if ch.id == oid:
                return si, "chart"
        for im in sh.images:
            if im.id == oid:
                return si, "image"
    raise EngineError(f"no object with id {oid}")


@command("chart.update", "Format Chart", menu=("Chart Design",), params="{id, kind?, title?, legend?: none|bottom|top|left|right, dataLabels?, gridlines?, xTitle?, yTitle?}")
def chart_update(s, p):
    oid = int(need(p, "id", "chart.update"))
    si, what = _object(s, oid)
    if what != "chart":
        raise EngineError("not a chart")
    if p.get("kind") is not None and p["kind"] not in CHART_KINDS:
        raise EngineError(f"unknown chart kind `{p['kind']}`")

    def f(ctx):
        ch = next(c for c in ctx.sheet_mut(si).charts if c.id == oid)
        if p.get("kind") is not None:
            ch.kind = str(p["kind"])
        if "title" in p:
            ch.title = p["title"] or None
        if p.get("legend") is not None:
            ch.legend = str(p["legend"])
        if "dataLabels" in p:
            ch.data_labels = p_bool(p, "dataLabels")
        if "gridlines" in p:
            ch.gridlines = p_bool(p, "gridlines")
        if "xTitle" in p:
            ch.x_title = p["xTitle"] or None
        if "yTitle" in p:
            ch.y_title = p["yTitle"] or None

    edit(s, f)
    return None


@command("chart.switchRowColumn", "Switch Row/Column", menu=("Chart Design", "Data"), params="{id}")
def chart_switch(s, p):
    oid = int(need(p, "id", "chart.switchRowColumn"))
    si, what = _object(s, oid)
    d = s.doc()
    ch = next(c for c in d.wb.sheets[si].charts if c.id == oid)
    if not ch.source:
        raise EngineError("This chart has no source range to switch.")
    tsi, rng = resolve_reference(d.wb, si, ch.source)
    series, by_rows = build_series(d.wb, tsi, rng, not ch.by_rows, ch.kind)

    def f(ctx):
        c = next(c for c in ctx.sheet_mut(si).charts if c.id == oid)
        c.series = series
        c.by_rows = by_rows

    edit(s, f)
    return None


@command("chart.selectData", "Select Data", menu=("Chart Design", "Data"), params='{id, range: "Sheet1!A1:C5", byRows?}')
def chart_select_data(s, p):
    oid = int(need(p, "id", "chart.selectData"))
    si, what = _object(s, oid)
    d = s.doc()
    ch = next(c for c in d.wb.sheets[si].charts if c.id == oid)
    tsi, rng = resolve_reference(d.wb, si, str(need(p, "range", "chart.selectData")))
    by_rows = p_bool(p, "byRows") if "byRows" in p else None
    series, by_rows = build_series(d.wb, tsi, rng, by_rows, ch.kind)
    source = f"{quote_sheet(d.wb.sheets[tsi].name)}!{rng.abs_a1()}"

    def f(ctx):
        c = next(c for c in ctx.sheet_mut(si).charts if c.id == oid)
        c.series = series
        c.by_rows = by_rows
        c.source = source

    edit(s, f)
    return None


@command("object.delete", "Delete Object", params="{id}")
def object_delete(s, p):
    oid = int(need(p, "id", "object.delete"))
    si, what = _object(s, oid)

    def f(ctx):
        sm = ctx.sheet_mut(si)
        sm.charts = [c for c in sm.charts if c.id != oid]
        sm.images = [i for i in sm.images if i.id != oid]

    edit(s, f)
    return None


command("chart.delete", "Delete Chart", params="{id}")(object_delete)


@command("object.move", "Move Object", params="{id, row?, col?, dx?, dy?, width?, height?, x?, y? (px from the sheet's top-left)}")
def object_move(s, p):
    oid = int(need(p, "id", "object.move"))
    si, what = _object(s, oid)
    d = s.doc()
    sh = d.wb.sheets[si]

    def f(ctx):
        sm = ctx.sheet_mut(si)
        obj = next(o for o in (sm.charts + sm.images) if o.id == oid)
        a = obj.anchor
        if p.get("x") is not None or p.get("y") is not None:
            x = max(float(p.get("x", sm.col_left(a.col) + a.dx)), 0.0)
            y = max(float(p.get("y", sm.row_top(a.row) + a.dy)), 0.0)
            col = sm.col_index().at(x)
            row = sm.row_index().at(y)
            a.col, a.row = col, row
            a.dx = x - sm.col_left(col)
            a.dy = y - sm.row_top(row)
        if p.get("row") is not None:
            a.row = max(0, min(int(p["row"]), MAX_ROWS - 1))
        if p.get("col") is not None:
            a.col = max(0, min(int(p["col"]), MAX_COLS - 1))
        if p.get("dx") is not None:
            a.dx = float(p["dx"])
        if p.get("dy") is not None:
            a.dy = float(p["dy"])
        if p.get("width") is not None:
            a.width = max(float(p["width"]), 10.0)
        if p.get("height") is not None:
            a.height = max(float(p["height"]), 10.0)

    edit(s, f)
    return None


# ---------------------------------------------------------------- pictures


def image_size(data: bytes) -> Optional[Tuple[int, int]]:
    """Pixel size of a PNG, JPEG or GIF image."""
    try:
        if data[:8] == b"\x89PNG\r\n\x1a\n":
            w, h = struct.unpack(">II", data[16:24])
            return w, h
        if data[:6] in (b"GIF87a", b"GIF89a"):
            w, h = struct.unpack("<HH", data[6:10])
            return w, h
        if data[:2] == b"\xff\xd8":
            i = 2
            while i < len(data) - 9:
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                    h, w = struct.unpack(">HH", data[i + 5:i + 9])
                    return w, h
                seg = struct.unpack(">H", data[i + 2:i + 4])[0]
                i += 2 + seg
    except struct.error:
        return None
    return None


def image_mime(data: bytes, name: str = "") -> Optional[str]:
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return "image/png"
    if data[:2] == b"\xff\xd8":
        return "image/jpeg"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "image/gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "image/webp"
    if name.lower().endswith(".svg") or data.lstrip()[:5] in (b"<?xml", b"<svg "):
        return "image/svg+xml"
    return None


@command("insert.picture", "Pictures", menu=("Insert", "Illustrations"), params="{data: base64, name?, cell?, width?, height?, alt?}")
def insert_picture(s, p):
    d = s.doc()
    si = param_sheet(s, p)
    raw = need(p, "data", "insert.picture")
    data = raw if isinstance(raw, (bytes, bytearray)) else base64.b64decode(str(raw))
    data = bytes(data)
    if len(data) > 20 * 1024 * 1024:
        raise EngineError("The picture is too large (20 MB at most).")
    mime = image_mime(data, p_str(p, "name", "") or "")
    if mime is None:
        raise EngineError("This picture format isn't supported (use PNG, JPEG, GIF, WebP or SVG).")
    at = param_cell(s, p)
    size = image_size(data) or (320, 240)
    w = p_float(p, "width") or float(size[0])
    h = p_float(p, "height") or float(size[1])
    if not p.get("width") and w > 640:
        h, w = h * 640 / w, 640.0
    oid = d.wb.next_object_id()
    img = Image(oid, Anchor(row=at[0], col=at[1], width=w, height=h, mode="oneCell"), data, mime, p_str(p, "alt", "") or (p_str(p, "name", "") or ""))

    def f(ctx):
        ctx.sheet_mut(si).images.append(img)

    edit(s, f)
    return {"id": oid, "width": w, "height": h}


# ---------------------------------------------------------------- hyperlinks


@command("insert.hyperlink", "Link", menu=("Insert", "Links"), shortcut="Ctrl+K", params='{cell?, target: "https://…" | "mailto:…" | "Sheet2!A1", text?, tooltip?}')
def insert_hyperlink(s, p):
    d = s.doc()
    si = param_sheet(s, p)
    at = param_cell(s, p)
    target = str(need(p, "target", "insert.hyperlink")).strip()
    if not target:
        raise EngineError("Enter an address for the link.")
    if target.startswith("#"):
        target = target[1:]
    text = p.get("text")

    def f(ctx):
        wb = ctx.wb
        sm = ctx.sheet_mut(si)
        sm.hyperlinks[at] = Hyperlink(target, p_str(p, "tooltip"))
        cell = sm.cell(at)
        sid = sm.style_id(at)
        st = wb.styles.get(sid)
        if st.font.color.is_auto and st.font.underline == "none":
            sid = wb.styles.intern(replace(st, font=replace(st.font, color=Color.theme(10), underline="single")))
        if text is not None and str(text) != "":
            _set(ctx, si, at, Cell(str(text), None, sid))
        elif cell is None or (cell.value is None and cell.formula is None):
            _set(ctx, si, at, Cell(target, None, sid))
        else:
            _set(ctx, si, at, cell.with_style(sid))

    edit(s, f)
    return None


@command("insert.removeHyperlink", "Remove Hyperlinks", menu=("Home", "Editing", "Clear"), params="{range?}")
def remove_hyperlink(s, p):
    si = param_sheet(s, p)
    ranges = _ranges(s, p)

    def f(ctx):
        wb = ctx.wb
        sm = ctx.sheet_mut(si)
        for rc in [rc for rc in sm.hyperlinks if any(r.contains(rc) for r in ranges)]:
            del sm.hyperlinks[rc]
            cell = sm.cell(rc)
            if cell is not None:
                st = wb.styles.get(cell.style)
                if st.font.color == Color.theme(10):
                    sid = wb.styles.intern(replace(st, font=replace(st.font, color=Color(), underline="none")))
                    _set(ctx, si, rc, cell.with_style(sid))

    edit(s, f)
    return None


@command("hyperlink.follow", "Open Hyperlink", params="{cell?}", undoable=False, journal=False)
def follow_hyperlink(s, p):
    d = s.doc()
    si = param_sheet(s, p)
    at = param_cell(s, p)
    hl = d.wb.sheet(si).hyperlinks.get(at)
    target = None
    if hl is not None:
        target = hl.target
    else:
        cell = d.wb.sheet(si).cell(at)
        if cell is not None and cell.formula is not None and cell.formula.text.upper().startswith("HYPERLINK("):
            e = cell.formula.get_expr()
            if e is not None and getattr(e, "args", None):
                from ..calc import evaluate_expr

                v = evaluate_expr(d.wb, si, at, e.args[0])
                target = v if isinstance(v, str) else None
    if not target:
        raise EngineError("There's no link in this cell.")
    low = target.lower()
    if low.startswith(("http://", "https://", "mailto:", "ftp://", "file://")) or ("." in target and "!" not in target and "/" in target):
        return {"url": target}
    loc = target[1:] if target.startswith("#") else target
    tsi, rng = resolve_reference(d.wb, si, loc if "!" in loc or d.wb.table(loc) else loc)
    _activate(s, tsi)
    d.selection = Selection.of_range(rng)
    return {"sheet": d.wb.sheets[tsi].name, "range": rng.a1()}


# ---------------------------------------------------------------- notes and comments


def _note_cell(s: Session, p: Dict[str, Any]):
    return param_sheet(s, p), param_cell(s, p)


@command("review.newNote", "New Note", menu=("Review", "Notes"), shortcut="Shift+F2", params="{cell?, text, author?}")
def new_note(s, p):
    si, at = _note_cell(s, p)
    text = str(p.get("text", ""))
    author = p_str(p, "author") or s.prefs.get("userName", "ReflexGrid User")

    def f(ctx):
        sm = ctx.sheet_mut(si)
        old = sm.comments.get(at)
        sm.comments[at] = Comment(author, text, visible=old.visible if old is not None else False)

    edit(s, f)
    return None


command("review.editNote", "Edit Note", menu=("Review", "Notes"), params="{cell?, text}")(new_note)


@command("review.deleteComment", "Delete", menu=("Review", "Comments"), params="{cell?}")
def delete_comment(s, p):
    si, at = _note_cell(s, p)
    if at not in s.doc().wb.sheet(si).comments:
        raise EngineError("There's no note or comment in this cell.")

    def f(ctx):
        sm = ctx.sheet_mut(si)
        sm.comments.pop(at, None)

    edit(s, f)
    return None


command("review.deleteNote", "Delete Note", params="{cell?}")(delete_comment)


@command("review.showNote", "Show/Hide Note", menu=("Review", "Notes"), params="{cell?, on?}")
def show_note(s, p):
    si, at = _note_cell(s, p)
    c = s.doc().wb.sheet(si).comments.get(at)
    if c is None:
        raise EngineError("There's no note in this cell.")
    on = p_bool(p, "on", not c.visible)

    def f(ctx):
        sm = ctx.sheet_mut(si)
        sm.comments[at] = replace(sm.comments[at], visible=on)

    edit(s, f)
    return None


@command("review.showAllNotes", "Show All Notes", menu=("Review", "Notes"), params="{on?}")
def show_all_notes(s, p):
    d = s.doc()
    si = param_sheet(s, p)
    cs = d.wb.sheet(si).comments
    on = p_bool(p, "on", not all(c.visible for c in cs.values()) if cs else True)

    def f(ctx):
        sm = ctx.sheet_mut(si)
        for rc in list(sm.comments):
            if not sm.comments[rc].threaded:
                sm.comments[rc] = replace(sm.comments[rc], visible=on)

    edit(s, f)
    return None


@command("review.newComment", "New Comment", menu=("Review", "Comments"), params="{cell?, text, author?}")
def new_comment(s, p):
    si, at = _note_cell(s, p)
    text = str(need(p, "text", "review.newComment"))
    author = p_str(p, "author") or s.prefs.get("userName", "ReflexGrid User")

    def f(ctx):
        sm = ctx.sheet_mut(si)
        old = sm.comments.get(at)
        if old is not None and old.threaded:
            sm.comments[at] = replace(old, replies=old.replies + ((author, text),))
        else:
            sm.comments[at] = Comment(author, text, threaded=True)

    edit(s, f)
    return None


@command("review.replyComment", "Reply", menu=("Review", "Comments"), params="{cell?, text, author?}")
def reply_comment(s, p):
    si, at = _note_cell(s, p)
    c = s.doc().wb.sheet(si).comments.get(at)
    if c is None or not c.threaded:
        raise EngineError("There's no comment thread in this cell.")
    return new_comment(s, p)


@command("review.resolveComment", "Resolve Thread", menu=("Review", "Comments"), params="{cell?, on?}")
def resolve_comment(s, p):
    si, at = _note_cell(s, p)
    c = s.doc().wb.sheet(si).comments.get(at)
    if c is None or not c.threaded:
        raise EngineError("There's no comment thread in this cell.")
    on = p_bool(p, "on", not c.resolved)

    def f(ctx):
        sm = ctx.sheet_mut(si)
        sm.comments[at] = replace(sm.comments[at], resolved=on)

    edit(s, f)
    return None


@command("review.listComments", "Show Comments", menu=("Review", "Comments"), undoable=False, journal=False)
def list_comments(s, p):
    d = s.doc()
    out = []
    for sh in d.wb.sheets:
        for rc in sorted(sh.comments):
            c = sh.comments[rc]
            out.append({
                "sheet": sh.name,
                "cell": a1(*rc),
                "author": c.author,
                "text": c.text,
                "threaded": c.threaded,
                "resolved": c.resolved,
                "visible": c.visible,
                "replies": [{"author": a, "text": t} for a, t in c.replies],
            })
    return {"comments": out}


# ---------------------------------------------------------------- sparklines


@command("insert.sparkline", "Sparklines", menu=("Insert", "Sparklines"), params='{kind?: line|column|stacked, data: "A2:E5", location: "F2:F5", color?, markers?}')
def insert_sparkline(s, p):
    d = s.doc()
    si = param_sheet(s, p)
    kind = p_str(p, "kind", "line")
    if kind not in ("line", "column", "stacked"):
        raise EngineError(f"unknown sparkline kind `{kind}`")
    tsi, data = resolve_reference(d.wb, si, str(need(p, "data", "insert.sparkline")))
    loc = RangeRef.parse(str(need(p, "location", "insert.sparkline")))
    if loc is None or (loc.width != 1 and loc.height != 1):
        raise EngineError("The location must be a single row or column.")
    sheet = quote_sheet(d.wb.sheets[tsi].name)
    n = loc.count()
    if loc.height > 1 or (loc.is_single() and data.height == 1):
        if data.height != n and not (loc.is_single() and data.height == 1):
            raise EngineError("The data range doesn't match the location range.")
        sources = [RangeRef(data.r0 + i, data.c0, data.r0 + i, data.c1) for i in range(n)]
    else:
        if data.width != n:
            raise EngineError("The data range doesn't match the location range.")
        sources = [RangeRef(data.r0, data.c0 + i, data.r1, data.c0 + i) for i in range(n)]
    from .cmd_format import parse_color

    color = parse_color(p.get("color")) if p.get("color") else Color.theme(4)
    cells = list(loc.cells())

    def f(ctx):
        sm = ctx.sheet_mut(si)
        sm.sparklines = [sp for sp in sm.sparklines if (sp.row, sp.col) not in set(cells)]
        for rc, src in zip(cells, sources):
            sm.sparklines.append(Sparkline(rc[0], rc[1], f"{sheet}!{src.abs_a1()}", kind, color, p_bool(p, "markers")))

    edit(s, f)
    return None


@command("sparkline.clear", "Clear Sparklines", params="{range?}")
def clear_sparklines(s, p):
    si = param_sheet(s, p)
    ranges = _ranges(s, p)

    def f(ctx):
        sm = ctx.sheet_mut(si)
        sm.sparklines = [sp for sp in sm.sparklines if not any(r.contains((sp.row, sp.col)) for r in ranges)]

    edit(s, f)
    return None
