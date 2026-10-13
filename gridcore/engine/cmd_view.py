"""View and Page Layout commands: freeze panes, gridlines, headings, zoom, formulas view,
print settings."""

from __future__ import annotations

from dataclasses import replace
from typing import Any, Dict

from ..core.addr import RangeRef
from ..xlsx.tables import PAPERS
from .cmd_format import _ranges
from .session import EngineError, Session, command, edit, p_bool, p_float, p_int, p_str, param_sheet


def _sheet_flag(attr: str):
    def run(s: Session, p: Dict[str, Any]):
        si = param_sheet(s, p)
        cur = getattr(s.doc().wb.sheet(si), attr)
        on = p_bool(p, "on", not cur)

        def f(ctx):
            setattr(ctx.sheet_mut(si), attr, on)

        edit(s, f)
        return {attr: on}

    return run


command("view.gridlines", "Gridlines", menu=("View", "Show"), params="{on?: bool}")(_sheet_flag("show_gridlines"))
command("view.headings", "Headings", menu=("View", "Show"), params="{on?: bool}")(_sheet_flag("show_headings"))
command("view.showZeros", "Show Zeros", params="{on?: bool}")(_sheet_flag("show_zeros"))
command("view.rightToLeft", "Sheet Right-to-Left", menu=("Page Layout", "Sheet Options"), params="{on?: bool}")(_sheet_flag("right_to_left"))
command("formulas.showFormulas", "Show Formulas", menu=("Formulas", "Formula Auditing"), shortcut="Ctrl+`", params="{on?: bool}")(_sheet_flag("show_formulas"))


def _freeze(s: Session, p: Dict[str, Any], rows: int, cols: int):
    si = param_sheet(s, p)

    def f(ctx):
        ctx.sheet_mut(si).freeze = (rows, cols) if rows or cols else None

    edit(s, f)
    return {"freeze": [rows, cols]}


@command("view.freezePanes", "Freeze Panes", menu=("View", "Window", "Freeze Panes"), params='{cell?: "B3" (rows above and columns left of it freeze)}')
def freeze_panes(s, p):
    d = s.doc()
    sh = d.sheet
    if sh.freeze is not None and p.get("cell") is None:
        return _freeze(s, p, 0, 0)
    from .session import param_cell

    rows, cols = param_cell(s, p)
    if rows == 0 and cols == 0:
        # Freezing at A1 freezes the middle of the visible area in Excel; we freeze nothing.
        return _freeze(s, p, 0, 0)
    return _freeze(s, p, rows, cols)


@command("view.freezeTopRow", "Freeze Top Row", menu=("View", "Window", "Freeze Panes"))
def freeze_top_row(s, p):
    return _freeze(s, p, 1, 0)


@command("view.freezeFirstColumn", "Freeze First Column", menu=("View", "Window", "Freeze Panes"))
def freeze_first_col(s, p):
    return _freeze(s, p, 0, 1)


@command("view.unfreezePanes", "Unfreeze Panes", menu=("View", "Window", "Freeze Panes"))
def unfreeze(s, p):
    return _freeze(s, p, 0, 0)


@command("view.zoom", "Zoom", menu=("View", "Zoom"), params="{percent: 10..400}", undoable=False)
def zoom(s, p):
    z = p_int(p, "percent", 100)
    if not 10 <= z <= 400:
        raise EngineError("Zoom must be between 10% and 400%.")
    d = s.doc()
    si = param_sheet(s, p)
    sm = d.wb.sheet_mut(si)
    sm.zoom = z
    return {"zoom": z}


@command("view.scroll", "Scroll", params="{top: row, left: col}", undoable=False, journal=False)
def scroll(s, p):
    d = s.doc()
    sm = d.wb.sheet_mut(d.wb.active_sheet)
    sm.view_top_left = (max(p_int(p, "top", 0), 0), max(p_int(p, "left", 0), 0))
    return None


# ---------------------------------------------------------------- page layout


def _print(s: Session, p: Dict[str, Any], **kw):
    si = param_sheet(s, p)

    def f(ctx):
        sm = ctx.sheet_mut(si)
        sm.print = replace(sm.print, **kw)

    edit(s, f)
    return None


@command("pageLayout.orientation", "Orientation", menu=("Page Layout", "Page Setup"), params="{value: portrait|landscape}")
def orientation(s, p):
    v = p_str(p, "value", "portrait")
    if v not in ("portrait", "landscape"):
        raise EngineError("orientation must be portrait or landscape")
    return _print(s, p, orientation=v)


@command("pageLayout.size", "Size", menu=("Page Layout", "Page Setup"), params='{paper: "Letter" | "A4" | …}')
def paper_size(s, p):
    v = p_str(p, "paper", "Letter")
    names = {name.lower(): name for _, name in PAPERS}
    if v.lower() not in names:
        raise EngineError(f"unknown paper size `{v}`")
    return _print(s, p, paper=names[v.lower()])


@command("pageLayout.margins", "Margins", menu=("Page Layout", "Page Setup"), params="{preset?: normal|wide|narrow, left?, right?, top?, bottom?, header?, footer? (inches)}")
def margins(s, p):
    presets = {
        "normal": (0.7, 0.7, 0.75, 0.75, 0.3, 0.3),
        "wide": (1.0, 1.0, 1.0, 1.0, 0.5, 0.5),
        "narrow": (0.25, 0.25, 0.75, 0.75, 0.3, 0.3),
    }
    pr = p_str(p, "preset")
    if pr is not None:
        if pr not in presets:
            raise EngineError(f"unknown margins preset `{pr}`")
        return _print(s, p, margins=presets[pr])
    cur = s.doc().wb.sheet(param_sheet(s, p)).print.margins
    keys = ("left", "right", "top", "bottom", "header", "footer")
    vals = tuple(p_float(p, k, cur[i]) for i, k in enumerate(keys))
    if any(v < 0 or v > 20 for v in vals):
        raise EngineError("Margins must be between 0 and 20 inches.")
    return _print(s, p, margins=vals)


@command("pageLayout.printArea", "Set Print Area", menu=("Page Layout", "Page Setup", "Print Area"), params="{range?}")
def print_area(s, p):
    r = _ranges(s, p)[0]
    return _print(s, p, print_area=r)


@command("pageLayout.clearPrintArea", "Clear Print Area", menu=("Page Layout", "Page Setup", "Print Area"))
def clear_print_area(s, p):
    return _print(s, p, print_area=None)


@command("pageLayout.printTitles", "Print Titles", menu=("Page Layout", "Page Setup"), params='{rows?: "1:2", cols?: "A:A"}')
def print_titles(s, p):
    kw = {}
    if "rows" in p:
        r = RangeRef.parse(str(p["rows"])) if p["rows"] else None
        kw["title_rows"] = (r.r0, r.r1) if r is not None else None
    if "cols" in p:
        r = RangeRef.parse(str(p["cols"])) if p["cols"] else None
        kw["title_cols"] = (r.c0, r.c1) if r is not None else None
    return _print(s, p, **kw)


@command("pageLayout.printGridlines", "Print Gridlines", menu=("Page Layout", "Sheet Options"), params="{on?: bool}")
def print_gridlines(s, p):
    cur = s.doc().wb.sheet(param_sheet(s, p)).print.gridlines
    return _print(s, p, gridlines=p_bool(p, "on", not cur))


@command("pageLayout.printHeadings", "Print Headings", menu=("Page Layout", "Sheet Options"), params="{on?: bool}")
def print_headings(s, p):
    cur = s.doc().wb.sheet(param_sheet(s, p)).print.headings
    return _print(s, p, headings=p_bool(p, "on", not cur))


@command("pageLayout.scale", "Scale to Fit", menu=("Page Layout", "Scale to Fit"), params="{scale?: 10..400, fitWidth?, fitHeight?}")
def scale(s, p):
    if p.get("fitWidth") is not None or p.get("fitHeight") is not None:
        return _print(s, p, fit_to=(p_int(p, "fitWidth", 1), p_int(p, "fitHeight", 0)))
    sc = p_int(p, "scale", 100)
    if not 10 <= sc <= 400:
        raise EngineError("Scale must be between 10% and 400%.")
    return _print(s, p, scale=sc, fit_to=None)


@command("pageLayout.headerFooter", "Header & Footer", menu=("Insert", "Text"), params='{header?: "&C&P", footer?}')
def header_footer(s, p):
    kw = {}
    if "header" in p:
        kw["header"] = str(p["header"] or "")
    if "footer" in p:
        kw["footer"] = str(p["footer"] or "")
    return _print(s, p, **kw)
