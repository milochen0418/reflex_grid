"""What the grid draws: a JSON-ready snapshot of a window of a sheet (formatted texts, styles as
compact CSS-ish dicts, merges, conditional formats, tables, filter buttons, notes, objects) plus
the selection-dependent bits of the UI (name box, formula bar, status bar, ribbon state)."""

from __future__ import annotations

import base64
import math
from typing import Any, Dict, List, Optional, Tuple

from ..calc import evaluate
from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef, a1, col_to_letters
from ..core.value import Array, CellError, display
from ..formula.printer import quote_sheet
from ..model.style import BORDER_WIDTHS, DEFAULT_FONT_SIZE, Style
from ..numfmt import color_hex, format_value, parse as parse_fmt
from .cf import CfCache, CfLook
from .display import edit_text, stats
from .entry import list_items, validation_at
from .tables import table_style_colors

MAX_WINDOW_ROWS = 500
MAX_WINDOW_COLS = 200

_PATTERN_DENSITY = {
    "solid": 1.0,
    "darkGray": 0.75,
    "mediumGray": 0.5,
    "lightGray": 0.25,
    "gray125": 0.125,
    "gray0625": 0.0625,
}


def _blend(bg: Tuple[int, int, int], fg: Tuple[int, int, int], t: float) -> str:
    return "#{:02X}{:02X}{:02X}".format(*(int(round(bg[i] + (fg[i] - bg[i]) * t)) for i in range(3)))


def style_css(st: Style, theme) -> Dict[str, Any]:
    """A style as the compact dict the grid component understands."""
    f = st.font
    out: Dict[str, Any] = {"ff": f.name, "fs": f.size}
    if f.bold:
        out["b"] = 1
    if f.italic:
        out["i"] = 1
    if f.underline != "none":
        out["u"] = f.underline
    if f.strike:
        out["st"] = 1
    if f.vert != "baseline":
        out["va"] = f.vert
    c = f.color.hex(theme)
    if c:
        out["c"] = c
    fill = st.fill
    if fill.pattern != "none":
        fg = fill.fg.resolve(theme)
        bg = fill.bg.resolve(theme)
        if fill.pattern == "solid":
            if fg is not None:
                out["bg"] = "#{:02X}{:02X}{:02X}".format(*fg)
            elif bg is not None:
                out["bg"] = "#{:02X}{:02X}{:02X}".format(*bg)
        else:
            d = _PATTERN_DENSITY.get(fill.pattern, 0.5)
            out["bg"] = _blend(bg or (255, 255, 255), fg or (0, 0, 0), d)
    b = st.borders
    for key, line in (("bt", b.top), ("bb", b.bottom), ("bl", b.left), ("br", b.right), ("dd", b.diag_down), ("du", b.diag_up)):
        if not line.is_none:
            out[key] = [line.style, line.color.hex(theme) or "#000000"]
    a = st.align
    if a.h != "general":
        out["ha"] = a.h
    if a.v != "bottom":
        out["vt"] = a.v
    if a.wrap:
        out["wr"] = 1
    if a.shrink:
        out["sh"] = 1
    if a.indent:
        out["ind"] = a.indent
    if a.rotation:
        out["rot"] = a.rotation
    return out


def _cf_css(look: CfLook, theme) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    st = look.style
    if st is not None:
        if st.fill.pattern != "none":
            c = st.fill.fg.hex(theme) or st.fill.bg.hex(theme)
            if c:
                out["bg"] = c
        c = st.font.color.hex(theme)
        if c:
            out["c"] = c
        if st.font.bold:
            out["b"] = 1
        if st.font.italic:
            out["i"] = 1
        if st.font.strike:
            out["st"] = 1
        if st.font.underline != "none":
            out["u"] = st.font.underline
        bd = st.borders
        for key, line in (("bt", bd.top), ("bb", bd.bottom), ("bl", bd.left), ("br", bd.right)):
            if not line.is_none:
                out[key] = [line.style, line.color.hex(theme) or "#000000"]
    if look.scale_fill is not None:
        out["bg"] = "#{:02X}{:02X}{:02X}".format(*look.scale_fill)
    if look.bar is not None:
        frac, rgb, grad = look.bar
        out["bar"] = [round(frac, 4), "#{:02X}{:02X}{:02X}".format(*rgb), 1 if grad else 0]
    if look.icon is not None:
        out["icon"] = [look.icon[0], look.icon[1]]
    if look.hide_value:
        out["hide"] = 1
    return out


def _width_chars(px: float, size: float) -> int:
    digit = 7.0 * size / DEFAULT_FONT_SIZE
    return max(int((px - 5.0) / max(digit, 1.0)), 1)


def _table_overrides(t, theme) -> Dict[str, Any]:
    return table_style_colors(t.style, theme)


def _cells_for(sh, rng: RangeRef):
    """Stored cells plus spilled values in ``rng``."""
    out = {}
    for rc, cell in sh.cells.iter_range(rng):
        out[rc] = cell
    if sh.spill:
        for anchor, sr in sh.spill_ranges.items():
            inter = sr.intersection(rng)
            if inter is None:
                continue
            for rc in inter.cells():
                if rc not in out and rc in sh.spill:
                    out[rc] = None
    return out


def viewport(wb, si: int, window: RangeRef, show_formulas: Optional[bool] = None) -> Dict[str, Any]:
    sh = wb.sheet(si)
    theme = wb.theme
    window = RangeRef(
        max(window.r0, 0), max(window.c0, 0),
        min(window.r1, window.r0 + MAX_WINDOW_ROWS - 1, MAX_ROWS - 1),
        min(window.c1, window.c0 + MAX_WINDOW_COLS - 1, MAX_COLS - 1),
    )
    fr, fc = sh.freeze or (0, 0)
    areas = [window]
    if fr:
        areas.append(RangeRef(0, window.c0, min(fr - 1, MAX_ROWS - 1), window.c1))
    if fc:
        areas.append(RangeRef(window.r0, 0, window.r1, min(fc - 1, MAX_COLS - 1)))
    if fr and fc:
        areas.append(RangeRef(0, 0, fr - 1, fc - 1))
    formulas = sh.show_formulas if show_formulas is None else show_formulas
    styles_used: Dict[int, Dict[str, Any]] = {}
    cf = CfCache(wb, si) if sh.cond_formats else None
    tables = [t for t in sh.tables if any(t.range.intersects(a) for a in areas)]
    table_looks = {t.id: _table_overrides(t, theme) for t in tables}
    cells_out: List[Dict[str, Any]] = []
    seen = set()

    def style_of(sid: int) -> Style:
        if sid not in styles_used:
            styles_used[sid] = style_css(wb.styles.get(sid), theme)
        return wb.styles.get(sid)

    row_styles = {r: info.style for r, info in sh.rows.items() if info.style is not None}
    col_styles = {c: info.style for c, info in sh.cols.items() if info.style is not None}
    for sid in list(row_styles.values()) + list(col_styles.values()):
        style_of(sid)

    for area in areas:
        cells = _cells_for(sh, area)
        # Cells that only exist through conditional formats or tables still need drawing.
        extra = set()
        if cf is not None:
            used = sh.used_range()
            for c in sh.cond_formats:
                if c.rule.type in ("containsBlanks", "expression", "cellIs"):
                    for r in c.ranges:
                        inter = r.intersection(area)
                        if inter is not None and inter.count() <= 40000:
                            extra.update(inter.cells())
        for t in tables:
            inter = t.range.intersection(area)
            if inter is not None and inter.count() <= 100000:
                extra.update(inter.cells())
        keys = set(cells) | extra
        for rc in keys:
            if rc in seen:
                continue
            seen.add(rc)
            cell = cells.get(rc)
            sid = cell.style if cell is not None else sh.style_id(rc)
            st = style_of(sid)
            v = sh.value(rc)
            entry: Dict[str, Any] = {"r": rc[0], "c": rc[1]}
            if sid:
                entry["s"] = sid
            text = ""
            fmt_code = st.num_fmt
            look = cf.look(rc) if cf is not None else None
            if look is not None and look.style is not None and look.style.num_fmt != "General":
                fmt_code = look.style.num_fmt
            if formulas and cell is not None and cell.formula is not None:
                text = "=" + cell.formula.text
                entry["a"] = "l"
            elif v is not None:
                if isinstance(v, float) and v == 0 and not sh.show_zeros:
                    text = ""
                else:
                    width = sh.col_width(rc[1])
                    m = sh.merge_at(rc)
                    if m is not None and m.start == rc:
                        width = sum(sh.col_width(c) for c in range(m.c0, min(m.c1, m.c0 + 100) + 1))
                    fm = format_value(v, parse_fmt(fmt_code), wb.date_system, _width_chars(width, st.font.size))
                    text = fm.text
                    if fm.color:
                        entry["fc"] = color_hex(fm.color)
                    if fm.fill is not None:
                        entry["fill"] = [fm.fill[0], fm.fill[1]]
                    if isinstance(v, float):
                        entry["n"] = 1
                if st.align.h == "general" and not formulas:
                    if isinstance(v, float):
                        entry["a"] = "r"
                    elif isinstance(v, (bool, CellError)):
                        entry["a"] = "c"
                if isinstance(v, CellError):
                    entry["err"] = 1
            if text:
                entry["t"] = text
            if cell is not None and (cell.value is not None or cell.formula is not None):
                # What editing the cell starts from, when it differs from what it shows.
                inp = edit_text(wb, sh, rc)
                if inp != text:
                    entry["in"] = inp
            if cell is None and v is not None:
                entry["sp"] = 1
            if cell is not None and cell.formula is not None:
                entry["f"] = 1
            if look is not None:
                x = _cf_css(look, theme)
                if x:
                    entry["x"] = x
            for t in tables:
                if t.range.contains(rc):
                    tl = table_looks[t.id]
                    tx = _table_cell_look(t, rc, tl, st)
                    if tx:
                        entry["tx"] = tx
                    break
            if rc in sh.hyperlinks:
                entry["link"] = 1
            if not entry.get("t") and "x" not in entry and "tx" not in entry and not sid and "link" not in entry:
                continue
            cells_out.append(entry)

    merges = [[m.r0, m.c0, m.r1, m.c1] for m in sh.merges if any(m.intersects(a) for a in areas)]
    spills = [[r.r0, r.c0, r.r1, r.c1, anchor[0], anchor[1]] for anchor, r in sh.spill_ranges.items() if any(r.intersects(a) for a in areas)]
    notes = []
    for rc, cm in sh.comments.items():
        if any(a.contains(rc) for a in areas):
            notes.append({"r": rc[0], "c": rc[1], "author": cm.author, "text": cm.text, "threaded": cm.threaded, "visible": cm.visible, "resolved": cm.resolved, "replies": [[x, y] for x, y in cm.replies]})
    # Where each link goes and its ScreenTip, shown when the pointer rests on it.
    links = [{"r": rc[0], "c": rc[1], "target": h.target, "tooltip": h.tooltip} for rc, h in sh.hyperlinks.items() if any(a.contains(rc) for a in areas)]
    buttons = []
    af = sh.autofilter
    if af is not None:
        crit_cols = {k for k, _ in af.criteria}
        for c in range(af.range.c0, af.range.c1 + 1):
            if window.c0 - 1 <= c <= window.c1 + 1:
                buttons.append({"r": af.range.r0, "c": c, "active": (c - af.range.c0) in crit_cols})
    for t in tables:
        if t.header_row and t.filter_button:
            crit_cols = {k for k, _ in t.filter}
            for c in range(t.range.c0, t.range.c1 + 1):
                buttons.append({"r": t.range.r0, "c": c, "active": (c - t.range.c0) in crit_cols, "table": t.name})
    sparks = []
    for sp in sh.sparklines:
        if any(a.contains((sp.row, sp.col)) for a in areas):
            v = evaluate(wb, si, (sp.row, sp.col), sp.source)
            vals = []
            if isinstance(v, Array):
                vals = [x if isinstance(x, float) else None for x in v.data[:500]]
            sparks.append({"r": sp.row, "c": sp.col, "kind": sp.kind, "color": sp.color.hex(theme) or "#156082", "markers": sp.markers, "values": vals})
    return {
        "sheet": si,
        "window": [window.r0, window.c0, window.r1, window.c1],
        "cells": cells_out,
        "styles": {str(k): v for k, v in styles_used.items()},
        "rowStyles": {str(k): v for k, v in row_styles.items()},
        "colStyles": {str(k): v for k, v in col_styles.items()},
        "merges": merges,
        "spills": spills,
        "notes": notes,
        "links": links,
        "buttons": buttons,
        "sparklines": sparks,
        "tables": [{"name": t.name, "range": [t.range.r0, t.range.c0, t.range.r1, t.range.c1]} for t in tables],
    }


def _table_cell_look(t, rc, tl: Dict[str, Any], st: Style) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    has_fill = st.fill.pattern != "none"
    r0 = t.range.r0
    header = t.header_row and rc[0] == r0
    total = t.totals_row and rc[0] == t.range.r1
    if header:
        if tl.get("header") and not has_fill:
            out["bg"] = tl["header"]
        if tl.get("headerText") and st.font.color.is_auto:
            out["c"] = tl["headerText"]
        if tl.get("headerBold"):
            out["b"] = 1
        if tl.get("rule") in ("headerBottom", "rows", "grid", "box") and tl.get("border"):
            out["bb"] = ["thin", tl["border"]]
        return out
    if total:
        out["b"] = 1
        if tl.get("total") and not has_fill:
            out["bg"] = tl["total"]
            out["c"] = "#FFFFFF"
        if tl.get("border"):
            out["bt"] = ["double", tl["border"]]
        return out
    data0 = r0 + (1 if t.header_row else 0)
    i = rc[0] - data0
    band = None
    if t.banded_rows and i % 2 == 0:
        band = tl.get("band")
    if t.banded_cols and (rc[1] - t.range.c0) % 2 == 0:
        band = band or tl.get("band")
    if tl.get("body") and not has_fill:
        out["bg"] = tl["body"]
        if st.font.color.is_auto:
            out["c"] = tl.get("bodyText", "#FFFFFF")
    if band and not has_fill:
        out["bg"] = band
        if tl.get("bandText") and st.font.color.is_auto:
            out["c"] = tl["bandText"]
    rule = tl.get("rule")
    if rule in ("rows", "grid") and tl.get("border"):
        out["bb"] = ["thin", tl["border"]]
    if rule == "grid" and tl.get("border"):
        out["br"] = ["thin", tl["border"]]
    if rule == "white":
        out["bb"] = ["thin", "#FFFFFF"]
        out["br"] = ["thin", "#FFFFFF"]
    if (t.first_col and rc[1] == t.range.c0) or (t.last_col and rc[1] == t.range.c1):
        out["b"] = 1
    return out


def geometry(sh) -> Dict[str, Any]:
    """Column widths and row heights that differ from the defaults (0 = hidden)."""
    cols = {}
    for c, info in sh.cols.items():
        w = 0.0 if info.hidden else (info.size if info.size is not None else sh.default_col_width)
        if w != sh.default_col_width:
            cols[str(c)] = round(w, 2)
    rows = {}
    for r, info in sh.rows.items():
        h = 0.0 if info.hidden else (info.size if info.size is not None else sh.default_row_height)
        if h != sh.default_row_height:
            rows[str(r)] = round(h, 2)
        if len(rows) >= 200_000:
            break
    used = sh.used_range()
    return {
        "defaultColWidth": sh.default_col_width,
        "defaultRowHeight": sh.default_row_height,
        "cols": cols,
        "rows": rows,
        "freeze": list(sh.freeze) if sh.freeze else [0, 0],
        "usedRows": (used.r1 + 1) if used is not None else 0,
        "usedCols": (used.c1 + 1) if used is not None else 0,
        "gridlines": sh.show_gridlines,
        "headings": sh.show_headings,
        "formulas": sh.show_formulas,
        "zoom": sh.zoom,
        "rtl": sh.right_to_left,
    }


def _resolve_series(wb, si: int, text: Optional[str]):
    if not text:
        return None
    v = evaluate(wb, si, (0, 0), text if not text.startswith("=") else text[1:])
    return v


def chart_data(wb, si: int, ch) -> Dict[str, Any]:
    series = []
    cats: List[Any] = []
    for i, s in enumerate(ch.series):
        vals = _resolve_series(wb, si, s.values)
        data = []
        if isinstance(vals, Array):
            data = [x if isinstance(x, float) else None for x in vals.data[:10000]]
        elif isinstance(vals, float):
            data = [vals]
        name = None
        if s.name:
            if s.name.startswith('"'):
                name = s.name.strip('"')
            else:
                nv = _resolve_series(wb, si, s.name)
                if isinstance(nv, Array):
                    nv = nv.get(0, 0)
                name = display(nv) if nv is not None else None
        if not name:
            name = f"Series{i + 1}"
        if s.categories and not cats:
            cv = _resolve_series(wb, si, s.categories)
            if isinstance(cv, Array):
                cats = [(x if isinstance(x, float) and ch.kind.startswith("scatter") else (display(x) if x is not None else "")) for x in cv.data[:10000]]
        color = s.color.hex(wb.theme) if s.color is not None else None
        series.append({"name": name, "values": data, "color": color})
    n = max((len(s["values"]) for s in series), default=0)
    if not cats:
        cats = [str(i + 1) for i in range(n)]
    accents = ["#{:06X}".format(wb.theme.colors[4 + i]) for i in range(6)]
    a = ch.anchor
    sh = wb.sheet(si)
    return {
        "id": ch.id,
        "kind": ch.kind,
        "title": ch.title if ch.title else (series[0]["name"] if len(series) == 1 and ch.kind not in ("scatter", "scatterLines") else None),
        "legend": ch.legend,
        "dataLabels": ch.data_labels,
        "gridlines": ch.gridlines,
        "xTitle": ch.x_title,
        "yTitle": ch.y_title,
        "categories": cats,
        "series": series,
        "palette": accents,
        "x": sh.col_left(a.col) + a.dx if a.mode != "absolute" else a.dx,
        "y": sh.row_top(a.row) + a.dy if a.mode != "absolute" else a.dy,
        "width": a.width,
        "height": a.height,
    }


def objects(wb, si: int, with_images: bool = True) -> Dict[str, Any]:
    sh = wb.sheet(si)
    charts = []
    for ch in sh.charts:
        try:
            charts.append(chart_data(wb, si, ch))
        except Exception:
            continue
    images = []
    for im in sh.images:
        a = im.anchor
        item = {
            "id": im.id,
            "x": sh.col_left(a.col) + a.dx if a.mode != "absolute" else a.dx,
            "y": sh.row_top(a.row) + a.dy if a.mode != "absolute" else a.dy,
            "width": a.width,
            "height": a.height,
            "alt": im.alt,
        }
        if with_images:
            item["src"] = f"data:{im.mime};base64," + base64.b64encode(im.data).decode("ascii")
        images.append(item)
    return {"charts": charts, "images": images}


# ---------------------------------------------------------------- selection-dependent UI state


def name_box(wb, si: int, sel) -> str:
    r = sel.current()
    sh = wb.sheet(si)
    # A defined name or table whose range is exactly the selection shows its name.
    target = f"{quote_sheet(sh.name)}!{r.abs_a1()}".lower()
    for dn in wb.names:
        if not dn.hidden and dn.formula.replace("'", "").lower() == target.replace("'", ""):
            return dn.name
    for t in sh.tables:
        if t.range == r:
            return t.name
    return a1(*sel.active)


def ribbon_state(wb, sh, rc) -> Dict[str, Any]:
    st = wb.styles.get(sh.style_id(rc))
    nf = parse_fmt(st.num_fmt)
    m = sh.merge_at(rc)
    return {
        "bold": st.font.bold,
        "italic": st.font.italic,
        "underline": st.font.underline,
        "strike": st.font.strike,
        "fontName": st.font.name,
        "fontSize": f"{st.font.size:g}",
        "fontColor": st.font.color.hex(wb.theme),
        "fillColor": st.fill.fg.hex(wb.theme) if st.fill.pattern == "solid" else None,
        "hAlign": st.align.h,
        "vAlign": st.align.v,
        "wrap": st.align.wrap,
        "merged": m is not None and not m.is_single(),
        "numFmt": st.num_fmt,
        "numCategory": nf.kind(),
        "locked": st.protection.locked,
    }


def ui_state(session) -> Dict[str, Any]:
    d = session.doc()
    wb = d.wb
    si = wb.active_sheet
    sh = wb.sheet(si)
    sel = d.selection
    rc = sel.active
    cell = sh.cell(rc)
    formula_bar = cell.bar_text() if cell is not None and cell.formula is not None and cell.formula.array is not None else edit_text(wb, sh, rc)
    hidden_formula = sh.protection is not None and wb.styles.get(sh.style_id(rc)).protection.hidden
    st = stats(sh, sel.ranges)
    v = validation_at(sh, rc)
    dropdown = None
    input_msg = None
    if v is not None:
        if v.kind == "list" and v.in_cell_dropdown:
            try:
                dropdown = list_items(wb, si, v, rc)[:1000]
            except Exception:
                dropdown = []
        if v.show_input and (v.input_title or v.input_message):
            input_msg = {"title": v.input_title, "message": v.input_message}
    note = sh.comments.get(rc)
    link = sh.hyperlinks.get(rc)
    t = sh.table_at(rc)
    spill_anchor = None
    for anchor, r in sh.spill_ranges.items():
        if r.contains(rc):
            spill_anchor = [anchor[0], anchor[1], r.r0, r.c0, r.r1, r.c1]
            break
    return {
        "docName": d.name,
        "dirty": d.is_dirty(),
        "revision": d.revision,
        "sheet": si,
        "sheets": [{"name": x.name, "visible": x.visibility == "visible", "color": x.tab_color.hex(wb.theme) if x.tab_color is not None else None, "protected": x.protection is not None} for x in wb.sheets],
        "selection": sel.to_json(),
        "nameBox": name_box(wb, si, sel),
        "formulaBar": "" if hidden_formula else formula_bar,
        "isFormula": cell is not None and cell.formula is not None,
        "stats": st,
        "dropdown": dropdown,
        "inputMessage": input_msg,
        "note": {"author": note.author, "text": note.text, "threaded": note.threaded} if note is not None else None,
        "link": link.target if link is not None else None,
        "table": t.name if t is not None else None,
        "spill": spill_anchor,
        "ribbon": ribbon_state(wb, sh, rc),
        "canUndo": bool(d.undo),
        "canRedo": bool(d.redo),
        "undoLabel": d.undo[-1].label if d.undo else None,
        "redoLabel": d.redo[-1].label if d.redo else None,
        "protected": sh.protection is not None,
        "calcMode": wb.calc.mode,
        "circular": [f"{wb.sheets[k[0]].name}!{a1(k[1], k[2])}" for k in d.calc.circular[:5]],
        "docs": [{"name": x.name, "dirty": x.is_dirty()} for x in session.docs],
        "activeDoc": session.active,
        "clipboard": {"range": [session.clipboard.range.r0, session.clipboard.range.c0, session.clipboard.range.r1, session.clipboard.range.c1], "cut": session.clipboard.cut, "sheet": session.clipboard.sheet}
        if session.clipboard is not None and session.clipboard.doc_uid == d.uid
        else None,
        "formatPainter": session.format_painter is not None,
        "macros": {"project": wb.vba is not None, "enabled": bool(getattr(d, "macros_enabled", False))},
    }
