"""Excel tables (ListObjects): header names that follow their cells, automatic expansion,
calculated columns, totals and the look of the built-in table styles."""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef
from ..formula.adjust import shift_relative
from ..model.cell import Cell, Formula
from ..model.features import TOTALS_SUBTOTAL, Table, TableColumn
from ..model.style import Color, Theme, apply_tint


def unique_column_name(t: Table, base: str, skip: Optional[int] = None) -> str:
    used = {c.name.lower() for i, c in enumerate(t.columns) if i != skip}
    name, k = base, 2
    while name.lower() in used:
        name = f"{base}{k}"
        k += 1
    return name


def new_column_name(names) -> str:
    """The header Excel gives a new table column: the first free "ColumnN"."""
    used = {n.lower() for n in names}
    k = 1
    while f"column{k}" in used:
        k += 1
    return f"Column{k}"


def header_text(wb, sh, rc) -> str:
    from .display import value_text

    v = sh.value(rc)
    if v is None:
        return ""
    return value_text(wb, sh, rc)


def sync_headers(wb, changed: List[Tuple[int, int, int]]) -> None:
    """Typing in a header cell renames the table column (and header cells hold text)."""
    by_sheet: Dict[int, List[Tuple[int, int]]] = {}
    for si, r, c in changed:
        by_sheet.setdefault(si, []).append((r, c))
    for si, cells in by_sheet.items():
        sh = wb.sheet(si)
        if sh is None or not sh.tables:
            continue
        for t in sh.tables:
            if not t.header_row:
                continue
            hits = [rc for rc in cells if rc[0] == t.range.r0 and t.range.c0 <= rc[1] <= t.range.c1]
            if not hits:
                continue
            sm = wb.sheet_mut(si)
            tt = next(x for x in sm.tables if x.id == t.id)
            for rc in hits:
                i = rc[1] - tt.range.c0
                while len(tt.columns) <= i:
                    tt.columns.append(TableColumn(f"Column{len(tt.columns) + 1}"))
                # A header keeps spaces at its edges, as in Excel (`Qty ` is not `Qty`).
                text = header_text(wb, sm, rc).replace("\n", " ")
                if not text.strip():
                    text = f"Column{i + 1}"
                name = unique_column_name(tt, text, i)
                tt.columns[i].name = name
                cell = sm.cell(rc)
                if cell is None or cell.formula is not None or cell.value != name:
                    sm.set_cell(rc, Cell(name, None, cell.style if cell is not None else sm.style_id(rc)))


def auto_expand(ctx, si: int, rc: Tuple[int, int]) -> None:
    """Typing directly below or right of a table extends it (AutoExpansion)."""
    sm = ctx.wb.sheet(si)
    cell = sm.cell(rc)
    if cell is None or (cell.value is None and cell.formula is None):
        return
    for t in sm.tables:
        r = t.range
        if not t.totals_row and rc[0] == r.r1 + 1 and r.c0 <= rc[1] <= r.c1:
            if any(o is not t and o.range.intersects(RangeRef(rc[0], r.c0, rc[0], r.c1)) for o in sm.tables):
                return
            m = ctx.sheet_mut(si)
            tt = next(x for x in m.tables if x.id == t.id)
            tt.range = RangeRef(r.r0, r.c0, r.r1 + 1, r.c1)
            # Calculated columns fill the new row.
            for i, col in enumerate(tt.columns):
                c = r.c0 + i
                if col.formula and (rc[0], c) != rc and m.cell((rc[0], c)) is None:
                    src = (r.r1, c)
                    sc = m.cell(src)
                    if sc is not None and sc.formula is not None:
                        e = sc.formula.get_expr()
                        if e is not None:
                            m.set_cell((rc[0], c), Cell(None, Formula.from_expr(shift_relative(e, 1, 0)), sc.style))
                            ctx.touch(si, (rc[0], c))
            ctx.structural = True
            return
        if rc[1] == r.c1 + 1 and r.r0 <= rc[0] <= r.r1 and t.header_row and rc[0] == r.r0:
            if any(o is not t and o.range.intersects(RangeRef(r.r0, rc[1], r.r1, rc[1])) for o in sm.tables):
                return
            m = ctx.sheet_mut(si)
            tt = next(x for x in m.tables if x.id == t.id)
            tt.range = RangeRef(r.r0, r.c0, r.r1, r.c1 + 1)
            tt.columns.append(TableColumn(unique_column_name(tt, header_text(ctx.wb, m, rc) or new_column_name(c.name for c in tt.columns))))
            ctx.structural = True
            return


def calculated_column(ctx, si: int, rc: Tuple[int, int]) -> None:
    """A formula typed into an empty table column fills the whole column."""
    sm = ctx.wb.sheet(si)
    cell = sm.cell(rc)
    if cell is None or cell.formula is None or cell.formula.array is not None:
        return
    t = sm.table_at(rc)
    if t is None:
        return
    data = t.data_range()
    if data is None or not data.contains(rc) or data.height < 2:
        return
    col = rc[1]
    for r in range(data.r0, data.r1 + 1):
        if r == rc[0]:
            continue
        o = sm.cell((r, col))
        if o is not None and (o.value is not None or o.formula is not None):
            return
    e = cell.formula.get_expr()
    if e is None:
        return
    m = ctx.sheet_mut(si)
    for r in range(data.r0, data.r1 + 1):
        if r == rc[0]:
            continue
        style = m.style_id((r, col))
        m.set_cell((r, col), Cell(None, Formula.from_expr(shift_relative(e, r - rc[0], 0)), style))
        ctx.touch(si, (r, col))
    tt = next(x for x in m.tables if x.id == t.id)
    i = col - tt.range.c0
    if i < len(tt.columns):
        tt.columns[i].formula = cell.formula.text


def totals_formula(t: Table, i: int) -> Optional[str]:
    col = t.columns[i]
    if col.totals == "none":
        return None
    if col.totals == "custom":
        return col.totals_label
    fn = TOTALS_SUBTOTAL.get(col.totals)
    if fn is None:
        return None
    name = col.name.replace("'", "''").replace("[", "'[").replace("]", "']").replace("#", "'#")
    return f"SUBTOTAL({fn},{t.name}[{name}])"


# ---------------------------------------------------------------- table styles

def _family(style: str) -> Tuple[str, int]:
    for fam in ("Light", "Medium", "Dark"):
        p = "TableStyle" + fam
        if style.startswith(p):
            try:
                return fam, int(style[len(p):])
            except ValueError:
                return "Medium", 2
    return "Medium", 2


def table_style_colors(style: str, theme: Theme) -> Dict[str, Optional[str]]:
    """Header fill/text, band fill, border and totals colours for a built-in table style name,
    drawn from the theme accents (an approximation of the look, with our own choices)."""
    fam, n = _family(style)
    k = (n - 1) % 7  # 0 = neutral, 1..6 = accents 1..6

    def accent(tint: float = 0.0) -> str:
        if k == 0:
            base = 0x000000
        else:
            base = theme.colors[3 + k]
        rgb = ((base >> 16) & 255, (base >> 8) & 255, base & 255)
        if tint and k == 0:
            g = int(255 * tint) if tint > 0 else 0
            rgb = (g, g, g)
        elif tint:
            rgb = apply_tint(rgb, tint)
        return "#{:02X}{:02X}{:02X}".format(*rgb)

    if style.lower() in ("none", "") :
        return {"header": None, "headerText": None, "band": None, "border": None, "total": None, "headerBold": True}
    if fam == "Light":
        group = (n - 1) // 7
        if group == 0:
            return {"header": None, "headerText": accent(-0.25) if k else "#000000", "band": accent(0.8) if k else "#D9D9D9", "border": accent(), "total": None, "headerBold": True, "rule": "headerBottom"}
        if group == 1:
            return {"header": accent(), "headerText": "#FFFFFF", "band": None, "border": accent(), "total": None, "headerBold": True, "rule": "box"}
        return {"header": None, "headerText": "#000000", "band": accent(0.8) if k else "#D9D9D9", "border": accent(), "total": None, "headerBold": True, "rule": "grid"}
    if fam == "Medium":
        group = (n - 1) // 7
        if group == 0:
            return {"header": accent(), "headerText": "#FFFFFF", "band": accent(0.8) if k else "#D9D9D9", "border": accent(0.4) if k else "#A6A6A6", "total": None, "headerBold": True, "rule": "rows"}
        if group == 1:
            return {"header": accent(), "headerText": "#FFFFFF", "band": accent(0.6) if k else "#BFBFBF", "border": "#FFFFFF", "total": accent(), "headerBold": True, "rule": "white"}
        if group == 2:
            return {"header": accent(-0.25) if k else "#404040", "headerText": "#FFFFFF", "band": accent(0.6) if k else "#BFBFBF", "border": accent(0.4) if k else "#A6A6A6", "total": None, "headerBold": True, "rule": "rows"}
        return {"header": accent(0.6) if k else "#BFBFBF", "headerText": "#000000", "band": accent(0.8) if k else "#D9D9D9", "border": accent(0.4) if k else "#A6A6A6", "total": None, "headerBold": True, "rule": "grid"}
    return {"header": "#000000" if k == 0 else accent(-0.5), "headerText": "#FFFFFF", "band": accent(-0.25) if k else "#595959", "bandText": "#FFFFFF", "border": None, "total": None, "headerBold": True, "rule": "none", "body": accent() if k else "#737373", "bodyText": "#FFFFFF"}


TABLE_STYLE_NAMES = (
    [f"TableStyleLight{i}" for i in range(1, 22)]
    + [f"TableStyleMedium{i}" for i in range(1, 29)]
    + [f"TableStyleDark{i}" for i in range(1, 12)]
)
