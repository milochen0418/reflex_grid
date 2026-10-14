"""The ribbon: tabs of command groups. Every button runs an engine command (``GridState.cmd``)
or opens a dialog (``GridState.open_dialog``)."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

import reflex as rx

from gridcore.engine.cmd_format import CELL_STYLE_NAMES
from gridcore.engine.tables import TABLE_STYLE_NAMES

from ..state import CATEGORIES, FUNCTIONS, GridState

ACCENT = "#156082"

FONTS = [
    "Aptos", "Arial", "Calibri", "Cambria", "Candara", "Consolas", "Courier New", "Georgia", "Segoe UI", "Tahoma",
    "Times New Roman", "Trebuchet MS", "Verdana", "Microsoft JhengHei", "PMingLiU", "DFKai-SB", "Noto Sans TC", "Noto Serif TC",
]
SIZES = ["8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72"]
NUMBER_FORMATS = [
    ("general", "General"), ("number", "Number"), ("currency", "Currency"), ("accounting", "Accounting"),
    ("shortDate", "Short Date"), ("longDate", "Long Date"), ("time", "Time"), ("percent", "Percentage"),
    ("fraction", "Fraction"), ("scientific", "Scientific"), ("text", "Text"),
]
THEME_ROW = ["#FFFFFF", "#000000", "#E8E8E8", "#0E2841", "#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E"]
STANDARD = ["#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"]


def _tint(hex_color: str, t: float) -> str:
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    if t > 0:
        r, g, b = (int(round(x + (255 - x) * t)) for x in (r, g, b))
    else:
        r, g, b = (int(round(x * (1 + t))) for x in (r, g, b))
    return f"#{r:02X}{g:02X}{b:02X}"


PALETTE: List[List[str]] = [THEME_ROW] + [[_tint(c, t) for c in THEME_ROW] for t in (0.8, 0.6, 0.4, -0.25, -0.5)]


def tip(child: rx.Component, text: str) -> rx.Component:
    return rx.tooltip(child, content=text, delay_duration=400)


def rbtn(icon: str, label: str, on_click, tooltip: Optional[str] = None, active=None, big: bool = False, id: Optional[str] = None, show_label: bool = False) -> rx.Component:
    variant = rx.cond(active, "soft", "ghost") if active is not None else "ghost"
    if big:
        content = rx.vstack(rx.icon(icon, size=22), rx.text(label, size="1", class_name="rg-biglabel"), spacing="1", align="center")
        btn = rx.button(content, variant=variant, color_scheme="gray", on_click=on_click, height="62px", padding_x="4px", id=id, class_name="rg-rbtn rg-big")
    else:
        inner = rx.hstack(rx.icon(icon, size=16), rx.text(label, size="1"), spacing="1", align="center") if show_label else rx.icon(icon, size=16)
        btn = rx.button(inner, variant=variant, color_scheme="gray", size="1", on_click=on_click, id=id, class_name="rg-rbtn", padding_x="6px" if show_label else "4px")
    return tip(btn, tooltip or label)


def menu(trigger_icon: str, label: str, items: List[Any], tooltip: Optional[str] = None, big: bool = False, show_label: bool = False, id: Optional[str] = None) -> rx.Component:
    """A dropdown button. ``items``: (label, event) pairs, "-" separators or (label, [submenu items])."""
    if big:
        # A tooltip around a menu trigger swallows the trigger: hint with the native title instead.
        trig = rx.button(rx.vstack(rx.icon(trigger_icon, size=22), rx.text(label, " ▾", size="1", class_name="rg-biglabel"), spacing="1", align="center"), variant="ghost", color_scheme="gray", height="62px", padding_x="4px", id=id, class_name="rg-rbtn rg-big", title=tooltip or label)
    else:
        inner = [rx.icon(trigger_icon, size=16)]
        if show_label:
            inner.append(rx.text(label, size="1"))
        inner.append(rx.icon("chevron_down", size=10))
        trig = rx.button(rx.hstack(*inner, spacing="1", align="center"), variant="ghost", color_scheme="gray", size="1", padding_x="4px", id=id, class_name="rg-rbtn", title=tooltip or label)
    return rx.dropdown_menu.root(
        rx.dropdown_menu.trigger(trig),
        rx.dropdown_menu.content(*_menu_items(items), size="1", max_height="70vh", overflow_y="auto"),
    )


def _menu_items(items: List[Any]) -> List[rx.Component]:
    out = []
    for it in items:
        if it == "-":
            out.append(rx.dropdown_menu.separator())
        elif isinstance(it[1], list):
            out.append(rx.dropdown_menu.sub(rx.dropdown_menu.sub_trigger(it[0]), rx.dropdown_menu.sub_content(*_menu_items(it[1]), max_height="60vh", overflow_y="auto")))
        else:
            out.append(rx.dropdown_menu.item(it[0], on_select=it[1]))
    return out


def group(title: str, *children, spacing: str = "1") -> rx.Component:
    return rx.vstack(
        rx.hstack(*children, spacing=spacing, align="center", wrap="nowrap", flex="1"),
        rx.text(title, size="1", color="#6B7280", align="center", width="100%"),
        spacing="0",
        align="center",
        padding_x="6px",
        border_right="1px solid #E3E6E9",
        height="90px",
        justify="between",
        padding_y="3px",
        flex_shrink="0",
    )


def stack2(*rows) -> rx.Component:
    return rx.vstack(*[rx.hstack(*r, spacing="1", align="center") for r in rows], spacing="1" if len(rows) < 3 else "0", align="start")


def color_menu(icon: str, cid: str, label: str, none_label: str, id: Optional[str] = None) -> rx.Component:
    swatches = []
    for row in PALETTE:
        swatches.append(
            rx.hstack(
                *[
                    rx.box(width="16px", height="16px", background=c, border="1px solid #D0D4D9", cursor="pointer", on_click=GridState.cmd(cid, {"color": c}), title=c)
                    for c in row
                ],
                spacing="1",
            )
        )
    std = rx.hstack(*[rx.box(width="16px", height="16px", background=c, border="1px solid #D0D4D9", cursor="pointer", on_click=GridState.cmd(cid, {"color": c}), title=c) for c in STANDARD], spacing="1")
    return rx.popover.root(
        rx.popover.trigger(rx.button(rx.hstack(rx.icon(icon, size=16), rx.icon("chevron_down", size=10), spacing="0"), variant="ghost", color_scheme="gray", size="1", padding_x="4px", id=id, class_name="rg-rbtn", title=label)),
        rx.popover.content(
            rx.vstack(
                rx.text("Theme Colors", size="1", weight="bold"),
                *swatches,
                rx.text("Standard Colors", size="1", weight="bold"),
                std,
                rx.popover.close(rx.button(none_label, size="1", variant="soft", color_scheme="gray", on_click=GridState.cmd(cid, {"color": None}), width="100%")),
                spacing="1",
            ),
            size="1",
        ),
    )


def c(cid: str, **params):
    return GridState.cmd(cid, params)


def dlg(name: str, **params):
    return GridState.open_dialog(name, params)


# ---------------------------------------------------------------- tabs


def home_tab() -> rx.Component:
    rb = GridState.ribbon
    return rx.hstack(
        group(
            "Clipboard",
            menu("clipboard_paste", "Paste", [
                ("Paste", c("edit.paste")),
                ("Paste Values", c("edit.pasteSpecial", what="values")),
                ("Paste Formulas", c("edit.pasteSpecial", what="formulas")),
                ("Paste Formatting", c("edit.pasteSpecial", what="formats")),
                ("Transpose", c("edit.pasteSpecial", what="all", transpose=True)),
                "-",
                ("Paste Special…", dlg("pasteSpecial")),
            ], big=True, id="rb-paste"),
            stack2([rbtn("scissors", "Cut", c("edit.cut"), "Cut (Ctrl+X)")], [rbtn("copy", "Copy", c("edit.copy"), "Copy (Ctrl+C)")], [rbtn("paintbrush", "Format Painter", c("edit.formatPainter"), "Format Painter", active=GridState.ui["formatPainter"])]),
        ),
        group(
            "Font",
            rx.vstack(
                rx.hstack(
                    rx.select(FONTS, value=rb["fontName"].to(str), on_change=lambda v: GridState.cmd("home.fontName", {"name": v}), size="1", width="150px", id="rb-font"),
                    rx.select(SIZES, value=rb["fontSize"].to(str), on_change=lambda v: GridState.cmd("home.fontSize", {"size": v}), size="1", width="64px", id="rb-size"),
                    rbtn("a_arrow_up", "Increase Font Size", c("home.increaseFontSize")),
                    rbtn("a_arrow_down", "Decrease Font Size", c("home.decreaseFontSize")),
                    spacing="1",
                ),
                rx.hstack(
                    rbtn("bold", "Bold", c("home.bold"), "Bold (Ctrl+B)", active=rb["bold"], id="rb-bold"),
                    rbtn("italic", "Italic", c("home.italic"), "Italic (Ctrl+I)", active=rb["italic"], id="rb-italic"),
                    rbtn("underline", "Underline", c("home.underline"), "Underline (Ctrl+U)", active=rb["underline"] != "none"),
                    rbtn("strikethrough", "Strikethrough", c("home.strikethrough"), "Strikethrough (Ctrl+5)", active=rb["strike"]),
                    menu("grid_3x3", "Borders", [
                        ("Bottom Border", c("home.borders", preset="bottom")),
                        ("Top Border", c("home.borders", preset="top")),
                        ("Left Border", c("home.borders", preset="left")),
                        ("Right Border", c("home.borders", preset="right")),
                        "-",
                        ("No Border", c("home.borders", preset="none")),
                        ("All Borders", c("home.borders", preset="all")),
                        ("Outside Borders", c("home.borders", preset="outside")),
                        ("Thick Outside Borders", c("home.borders", preset="thickOutside")),
                        "-",
                        ("Bottom Double Border", c("home.borders", preset="doubleBottom")),
                        ("Thick Bottom Border", c("home.borders", preset="thickBottom")),
                        ("Top and Bottom Border", c("home.borders", preset="topBottom")),
                        ("Top and Thick Bottom Border", c("home.borders", preset="topThickBottom")),
                        ("Top and Double Bottom Border", c("home.borders", preset="topDoubleBottom")),
                        "-",
                        ("Inside Borders", c("home.borders", preset="inside")),
                        ("More Borders…", dlg("formatCells", tab="border")),
                    ], id="rb-borders"),
                    color_menu("paint_bucket", "home.fillColor", "Fill Color", "No Fill", id="rb-fill"),
                    color_menu("baseline", "home.fontColor", "Font Color", "Automatic", id="rb-fontcolor"),
                    spacing="1",
                ),
                spacing="1",
            ),
        ),
        group(
            "Alignment",
            rx.vstack(
                rx.hstack(
                    rbtn("align_vertical_justify_start", "Top Align", c("home.alignTop"), active=rb["vAlign"] == "top"),
                    rbtn("align_vertical_justify_center", "Middle Align", c("home.alignMiddle"), active=rb["vAlign"] == "center"),
                    rbtn("align_vertical_justify_end", "Bottom Align", c("home.alignBottom"), active=rb["vAlign"] == "bottom"),
                    menu("rotate_ccw", "Orientation", [
                        ("Angle Counterclockwise", c("home.orientation", rotation=45)),
                        ("Angle Clockwise", c("home.orientation", rotation=-45)),
                        ("Vertical Text", c("home.orientation", rotation=255)),
                        ("Rotate Text Up", c("home.orientation", rotation=90)),
                        ("Rotate Text Down", c("home.orientation", rotation=-90)),
                        ("No Rotation", c("home.orientation", rotation=0)),
                        "-",
                        ("Format Cell Alignment…", dlg("formatCells", tab="alignment")),
                    ]),
                    rbtn("wrap_text", "Wrap Text", c("home.wrapText"), active=rb["wrap"], show_label=True, id="rb-wrap"),
                    spacing="1",
                ),
                rx.hstack(
                    rbtn("align_left", "Align Left", c("home.alignLeft"), active=rb["hAlign"] == "left"),
                    rbtn("align_center", "Center", c("home.alignCenter"), active=rb["hAlign"] == "center"),
                    rbtn("align_right", "Align Right", c("home.alignRight"), active=rb["hAlign"] == "right"),
                    rbtn("indent_decrease", "Decrease Indent", c("home.decreaseIndent")),
                    rbtn("indent_increase", "Increase Indent", c("home.increaseIndent")),
                    menu("table_cells_merge", "Merge & Center", [
                        ("Merge & Center", c("home.mergeCenter")),
                        ("Merge Across", c("home.mergeAcross")),
                        ("Merge Cells", c("home.mergeCells")),
                        ("Unmerge Cells", c("home.unmergeCells")),
                    ], show_label=True, id="rb-merge"),
                    spacing="1",
                ),
                spacing="1",
            ),
        ),
        group(
            "Number",
            rx.vstack(
                rx.select.root(
                    rx.select.trigger(placeholder=rb["numCategory"].to(str), width="140px", id="rb-numfmt"),
                    rx.select.content(*[rx.select.item(label, value=key) for key, label in NUMBER_FORMATS], rx.select.separator(), rx.select.item("More Number Formats…", value="__more__")),
                    value="",
                    on_change=GridState.number_format,
                    size="1",
                ),
                rx.hstack(
                    rbtn("dollar_sign", "Accounting Number Format", c("home.accounting")),
                    rbtn("percent", "Percent Style (Ctrl+Shift+%)", c("home.percent")),
                    rbtn("hash", "Comma Style", c("home.comma")),
                    rbtn("decimals_arrow_right", "Increase Decimal", c("home.increaseDecimal")),
                    rbtn("decimals_arrow_left", "Decrease Decimal", c("home.decreaseDecimal")),
                    spacing="1",
                ),
                spacing="1",
            ),
        ),
        group(
            "Styles",
            menu("list_checks", "Conditional Formatting", [
                ("Highlight Cells Rules", [
                    ("Greater Than…", dlg("cfHighlight", kind="greaterThan")),
                    ("Less Than…", dlg("cfHighlight", kind="lessThan")),
                    ("Between…", dlg("cfHighlight", kind="between")),
                    ("Equal To…", dlg("cfHighlight", kind="equal")),
                    ("Text that Contains…", dlg("cfHighlight", kind="text")),
                    ("A Date Occurring…", dlg("cfHighlight", kind="date")),
                    ("Duplicate Values…", dlg("cfHighlight", kind="duplicate")),
                ]),
                ("Top/Bottom Rules", [
                    ("Top 10 Items…", dlg("cfTop", kind="top")),
                    ("Top 10%…", dlg("cfTop", kind="topPercent")),
                    ("Bottom 10 Items…", dlg("cfTop", kind="bottom")),
                    ("Bottom 10%…", dlg("cfTop", kind="bottomPercent")),
                    ("Above Average…", dlg("cfTop", kind="above")),
                    ("Below Average…", dlg("cfTop", kind="below")),
                ]),
                ("Data Bars", [
                    ("Blue Data Bar", c("home.conditionalFormat", rule={"type": "dataBar", "color": "#638EC6"})),
                    ("Green Data Bar", c("home.conditionalFormat", rule={"type": "dataBar", "color": "#63C384"})),
                    ("Red Data Bar", c("home.conditionalFormat", rule={"type": "dataBar", "color": "#FF555A"})),
                    ("Orange Data Bar", c("home.conditionalFormat", rule={"type": "dataBar", "color": "#FFB628"})),
                    ("Purple Data Bar", c("home.conditionalFormat", rule={"type": "dataBar", "color": "#BE6EBC"})),
                ]),
                ("Color Scales", [
                    ("Green - Yellow - Red", c("home.conditionalFormat", rule={"type": "colorScale", "scale": "greenYellowRed"})),
                    ("Red - Yellow - Green", c("home.conditionalFormat", rule={"type": "colorScale", "scale": "redYellowGreen"})),
                    ("Green - White - Red", c("home.conditionalFormat", rule={"type": "colorScale", "scale": "greenWhiteRed"})),
                    ("Blue - White - Red", c("home.conditionalFormat", rule={"type": "colorScale", "scale": "blueWhiteRed"})),
                    ("White - Red", c("home.conditionalFormat", rule={"type": "colorScale", "scale": "whiteRed"})),
                    ("Green - White", c("home.conditionalFormat", rule={"type": "colorScale", "scale": "greenWhite"})),
                ]),
                ("Icon Sets", [
                    ("3 Arrows", c("home.conditionalFormat", rule={"type": "iconSet", "iconSet": "3Arrows"})),
                    ("3 Traffic Lights", c("home.conditionalFormat", rule={"type": "iconSet", "iconSet": "3TrafficLights1"})),
                    ("3 Symbols", c("home.conditionalFormat", rule={"type": "iconSet", "iconSet": "3Symbols"})),
                    ("3 Flags", c("home.conditionalFormat", rule={"type": "iconSet", "iconSet": "3Flags"})),
                    ("3 Stars", c("home.conditionalFormat", rule={"type": "iconSet", "iconSet": "3Stars"})),
                    ("4 Arrows", c("home.conditionalFormat", rule={"type": "iconSet", "iconSet": "4Arrows"})),
                    ("4 Ratings", c("home.conditionalFormat", rule={"type": "iconSet", "iconSet": "4Rating"})),
                    ("5 Arrows", c("home.conditionalFormat", rule={"type": "iconSet", "iconSet": "5Arrows"})),
                    ("5 Quarters", c("home.conditionalFormat", rule={"type": "iconSet", "iconSet": "5Quarters"})),
                ]),
                "-",
                ("New Rule (formula)…", dlg("cfRule")),
                ("Clear Rules from Selected Cells", c("home.clearRules", scope="selection")),
                ("Clear Rules from Entire Sheet", c("home.clearRules", scope="sheet")),
                ("Manage Rules…", dlg("cfManage")),
            ], big=True, id="rb-cf"),
            menu("table", "Format as Table", [(n.replace("TableStyle", ""), GridState.cmd("insert.table", {"style": n})) for n in TABLE_STYLE_NAMES[:7] + TABLE_STYLE_NAMES[21:35] + TABLE_STYLE_NAMES[49:53]], big=True),
            menu("palette", "Cell Styles", [(n, c("home.cellStyle", name=n)) for n in CELL_STYLE_NAMES], big=True),
        ),
        group(
            "Cells",
            menu("square_plus", "Insert", [
                ("Insert Cells…", dlg("insertCells")),
                ("Insert Sheet Rows", c("home.insertRows")),
                ("Insert Sheet Columns", c("home.insertColumns")),
                ("Insert Sheet", c("home.insertSheet")),
            ], big=True, id="rb-insert"),
            menu("square_minus", "Delete", [
                ("Delete Cells…", dlg("deleteCells")),
                ("Delete Sheet Rows", c("home.deleteRows")),
                ("Delete Sheet Columns", c("home.deleteColumns")),
                ("Delete Sheet", c("home.deleteSheet")),
            ], big=True, id="rb-delete"),
            menu("layout_grid", "Format", [
                ("Row Height…", dlg("rowHeight")),
                ("AutoFit Row Height", c("home.autofitRowHeight")),
                ("Column Width…", dlg("columnWidth")),
                ("AutoFit Column Width", c("home.autofitColumnWidth")),
                ("Default Width…", dlg("defaultWidth")),
                "-",
                ("Hide & Unhide", [
                    ("Hide Rows", c("home.hideRows")),
                    ("Hide Columns", c("home.hideColumns")),
                    ("Hide Sheet", c("sheet.hide")),
                    ("Unhide Rows", c("home.unhideRows")),
                    ("Unhide Columns", c("home.unhideColumns")),
                    ("Unhide Sheet…", dlg("unhideSheet")),
                ]),
                "-",
                ("Rename Sheet…", dlg("renameSheet")),
                ("Move or Copy Sheet…", dlg("moveSheet")),
                ("Tab Color…", dlg("tabColor")),
                "-",
                ("Protect Sheet…", dlg("protectSheet")),
                ("Lock Cell", c("home.lockCell")),
                ("Format Cells…", dlg("formatCells")),
            ], big=True),
        ),
        group(
            "Editing",
            stack2(
                [menu("sigma", "AutoSum", [(f, c("formulas.autoSum", function=f)) for f in ("SUM", "AVERAGE", "COUNT", "MAX", "MIN")] + ["-", ("More Functions…", dlg("insertFunction"))], show_label=True, id="rb-autosum")],
                [menu("arrow_down_to_line", "Fill", [("Down", c("edit.fillDown")), ("Right", c("edit.fillRight")), ("Up", c("edit.fillUp")), ("Left", c("edit.fillLeft")), ("Series (linear)", c("edit.fillSeries"))], show_label=True)],
                [menu("eraser", "Clear", [("Clear All", c("edit.clearAll")), ("Clear Formats", c("edit.clearFormats")), ("Clear Contents", c("edit.clearContents")), ("Clear Comments and Notes", c("edit.clearComments")), ("Clear Hyperlinks", c("edit.clearHyperlinks")), ("Remove Hyperlinks", c("insert.removeHyperlink"))], show_label=True)],
            ),
            menu("list_filter", "Sort & Filter", [
                ("Sort A to Z", c("data.sortAscending")),
                ("Sort Z to A", c("data.sortDescending")),
                ("Custom Sort…", dlg("sort")),
                "-",
                ("Filter", c("data.filter")),
                ("Clear", c("data.clearFilter")),
                ("Reapply", c("data.reapply")),
            ], big=True),
            menu("search", "Find & Select", [
                ("Find…", dlg("find")),
                ("Replace…", dlg("replace")),
                ("Go To…", dlg("goTo")),
                "-",
                ("Formulas", c("edit.goToSpecial", kind="formulas")),
                ("Notes", c("edit.goToSpecial", kind="comments")),
                ("Conditional Formatting", c("edit.goToSpecial", kind="conditionalFormats")),
                ("Constants", c("edit.goToSpecial", kind="constants")),
                ("Data Validation", c("edit.goToSpecial", kind="dataValidation")),
                ("Blanks", c("edit.goToSpecial", kind="blanks")),
            ], big=True, id="rb-find"),
        ),
        spacing="0",
        align="stretch",
    )


CHART_KINDS = [
    ("Clustered Column", "columnClustered"), ("Stacked Column", "columnStacked"), ("100% Stacked Column", "columnStacked100"),
    ("Clustered Bar", "barClustered"), ("Stacked Bar", "barStacked"), ("Line", "line"), ("Line with Markers", "lineMarkers"),
    ("Pie", "pie"), ("Doughnut", "doughnut"), ("Area", "area"), ("Stacked Area", "areaStacked"), ("Scatter", "scatter"),
    ("Scatter with Lines", "scatterLines"), ("Radar", "radar"),
]


def insert_tab() -> rx.Component:
    return rx.hstack(
        group("Tables", rbtn("table", "Table", dlg("insertTable"), "Table (Ctrl+T)", big=True, id="rb-table")),
        group("Illustrations", rbtn("image", "Pictures", dlg("insertPicture"), big=True)),
        group(
            "Charts",
            menu("chart_column", "Column", [(l, c("insert.chart", kind=k)) for l, k in CHART_KINDS[:5]], big=True, id="rb-chart-col"),
            menu("chart_line", "Line", [(l, c("insert.chart", kind=k)) for l, k in CHART_KINDS[5:7] + CHART_KINDS[9:11]], big=True),
            menu("chart_pie", "Pie", [(l, c("insert.chart", kind=k)) for l, k in CHART_KINDS[7:9]], big=True),
            menu("chart_scatter", "Scatter", [(l, c("insert.chart", kind=k)) for l, k in CHART_KINDS[11:]], big=True),
            rbtn("chart_bar", "All Charts", dlg("insertChart"), "Insert Chart…", big=True),
        ),
        group("Sparklines", menu("activity", "Sparklines", [
            ("Line (in next column)", GridState.open_dialog("sparkline", {"kind": "line"})),
        ], big=True)),
        group("Links", rbtn("link", "Link", dlg("hyperlink"), "Insert Link (Ctrl+K)", big=True)),
        group("Comments", rbtn("message_square", "Comment", dlg("note", threaded=True), big=True), rbtn("sticky_note", "Note", dlg("note"), big=True)),
        group("Symbols", rbtn("square_function", "Function", dlg("insertFunction"), "Insert Function", big=True)),
        spacing="0",
        align="stretch",
    )


def page_layout_tab() -> rx.Component:
    return rx.hstack(
        group(
            "Page Setup",
            menu("panel_top", "Margins", [("Normal", c("pageLayout.margins", preset="normal")), ("Wide", c("pageLayout.margins", preset="wide")), ("Narrow", c("pageLayout.margins", preset="narrow"))], big=True),
            menu("rotate_cw", "Orientation", [("Portrait", c("pageLayout.orientation", value="portrait")), ("Landscape", c("pageLayout.orientation", value="landscape"))], big=True),
            menu("file", "Size", [(p, c("pageLayout.size", paper=p)) for p in ("Letter", "Legal", "A4", "A3", "A5", "B5", "Executive", "Tabloid")], big=True),
            menu("printer", "Print Area", [("Set Print Area", c("pageLayout.printArea")), ("Clear Print Area", c("pageLayout.clearPrintArea"))], big=True),
            rbtn("printer", "Print", GridState.print_sheet, "Print (Ctrl+P)", big=True),
        ),
        group(
            "Scale to Fit",
            menu("zoom_in", "Scale", [(f"{p}%", c("pageLayout.scale", scale=p)) for p in (50, 75, 85, 100, 125, 150)] + [("Fit to 1 page wide", c("pageLayout.scale", fitWidth=1, fitHeight=0))], big=True),
        ),
        group(
            "Sheet Options",
            stack2(
                [rx.text("Gridlines", size="1", width="64px"), rbtn("eye", "View Gridlines", c("view.gridlines"), active=GridState.geometry["gridlines"]), rbtn("printer", "Print Gridlines", c("pageLayout.printGridlines"))],
                [rx.text("Headings", size="1", width="64px"), rbtn("eye", "View Headings", c("view.headings"), active=GridState.geometry["headings"]), rbtn("printer", "Print Headings", c("pageLayout.printHeadings"))],
            ),
        ),
        group("Arrange", rbtn("arrow_right_left", "Right-to-Left", c("view.rightToLeft"), "Sheet Right-to-Left", big=True)),
        spacing="0",
        align="stretch",
    )


def formulas_tab() -> rx.Component:
    def fn_menu(cat: str):
        items = [(f["name"], GridState.set_edit_request("=" + f["name"] + "(")) for f in FUNCTIONS if f["category"] == cat][:120]
        return items

    lib = [(cat, fn_menu(cat)) for cat in CATEGORIES]
    return rx.hstack(
        group(
            "Function Library",
            rbtn("square_function", "Insert Function", dlg("insertFunction"), big=True, id="rb-insert-fn"),
            menu("sigma", "AutoSum", [(f, c("formulas.autoSum", function=f)) for f in ("SUM", "AVERAGE", "COUNT", "MAX", "MIN")], big=True),
            menu("book_open", "Functions", lib, big=True),
        ),
        group(
            "Defined Names",
            rbtn("tags", "Name Manager", dlg("nameManager"), "Name Manager (Ctrl+F3)", big=True),
            stack2([rbtn("tag", "Define Name", dlg("defineName"), show_label=True)], [rbtn("table_properties", "Create from Selection", c("formulas.createFromSelection"), show_label=True)]),
        ),
        group(
            "Formula Auditing",
            stack2(
                [rbtn("arrow_up_to_line", "Trace Precedents", GridState.audit("precedents"), show_label=True)],
                [rbtn("arrow_down_to_line", "Trace Dependents", GridState.audit("dependents"), show_label=True)],
                [rbtn("triangle_alert", "Error Checking", GridState.audit("errors"), show_label=True)],
            ),
            rbtn("square_function", "Show Formulas", c("formulas.showFormulas"), "Show Formulas (Ctrl+`)", big=True, active=GridState.geometry["formulas"]),
        ),
        group(
            "Calculation",
            menu("calculator", "Calculation Options", [("Automatic", c("formulas.calculationOptions", mode="auto")), ("Automatic Except for Data Tables", c("formulas.calculationOptions", mode="autoNoTable")), ("Manual", c("formulas.calculationOptions", mode="manual")), "-", ("Enable Iterative Calculation", c("formulas.calculationOptions", iterative=True)), ("Disable Iterative Calculation", c("formulas.calculationOptions", iterative=False))], big=True),
            rbtn("refresh_cw", "Calculate Now", c("formulas.calculateNow"), "Calculate Now (F9)", big=True),
        ),
        spacing="0",
        align="stretch",
    )


def data_tab() -> rx.Component:
    return rx.hstack(
        group("Get Data", rbtn("folder_open", "From Text/CSV", dlg("open"), "Import a CSV, TSV or XLSX file", big=True)),
        group(
            "Sort & Filter",
            stack2([rbtn("arrow_down_a_z", "Sort A to Z", c("data.sortAscending"), id="rb-sort-asc")], [rbtn("arrow_up_z_a", "Sort Z to A", c("data.sortDescending"), id="rb-sort-desc")]),
            rbtn("list_ordered", "Sort", dlg("sort"), "Custom Sort…", big=True),
            rbtn("filter", "Filter", c("data.filter"), "Filter (Ctrl+Shift+L)", big=True, id="rb-filter"),
            stack2([rbtn("filter_x", "Clear", c("data.clearFilter"), show_label=True)], [rbtn("refresh_cw", "Reapply", c("data.reapply"), show_label=True)]),
        ),
        group(
            "Data Tools",
            rbtn("split", "Text to Columns", dlg("textToColumns"), big=True),
            rbtn("copy_minus", "Remove Duplicates", dlg("removeDuplicates"), big=True),
            menu("shield_check", "Data Validation", [("Data Validation…", dlg("validation")), ("Circle Invalid Data", GridState.audit("invalid")), ("Clear Validation", c("data.clearValidation"))], big=True, id="rb-validation"),
        ),
        group("Forecast", menu("target", "What-If Analysis", [("Goal Seek…", dlg("goalSeek"))], big=True, id="rb-whatif")),
        spacing="0",
        align="stretch",
    )


def review_tab() -> rx.Component:
    return rx.hstack(
        group(
            "Comments",
            rbtn("message_square", "New Comment", dlg("note", threaded=True), big=True),
            rbtn("trash_2", "Delete", c("review.deleteComment"), "Delete Comment or Note", big=True),
        ),
        group(
            "Notes",
            rbtn("sticky_note", "New Note", dlg("note"), "New Note (Shift+F2)", big=True),
            rbtn("eye", "Show All Notes", c("review.showAllNotes"), big=True),
        ),
        group(
            "Protect",
            rx.cond(
                GridState.ui["protected"],
                rbtn("lock_open", "Unprotect Sheet", dlg("unprotectSheet"), big=True, id="rb-unprotect"),
                rbtn("lock", "Protect Sheet", dlg("protectSheet"), big=True, id="rb-protect"),
            ),
            rbtn("shield", "Protect Workbook", dlg("protectWorkbook"), big=True),
            rbtn("shield_off", "Unprotect Workbook", dlg("unprotectWorkbook"), big=True),
        ),
        spacing="0",
        align="stretch",
    )


def view_tab() -> rx.Component:
    return rx.hstack(
        group(
            "Show",
            stack2(
                [rbtn("grid_3x3", "Gridlines", c("view.gridlines"), active=GridState.geometry["gridlines"], show_label=True)],
                [rbtn("panel_top", "Headings", c("view.headings"), active=GridState.geometry["headings"], show_label=True)],
                [rbtn("square_function", "Formulas", c("formulas.showFormulas"), active=GridState.geometry["formulas"], show_label=True)],
            ),
        ),
        group(
            "Zoom",
            rbtn("zoom_in", "Zoom 100%", GridState.zoom_to(100), big=True),
            stack2([rbtn("plus", "Zoom In", GridState.zoom_step(10))], [rbtn("minus", "Zoom Out", GridState.zoom_step(-10))]),
        ),
        group(
            "Window",
            menu("snowflake", "Freeze Panes", [("Freeze Panes", c("view.freezePanes")), ("Freeze Top Row", c("view.freezeTopRow")), ("Freeze First Column", c("view.freezeFirstColumn")), ("Unfreeze Panes", c("view.unfreezePanes"))], big=True, id="rb-freeze"),
        ),
        group("Help", rbtn("keyboard", "Keyboard Shortcuts", dlg("shortcuts"), big=True)),
        spacing="0",
        align="stretch",
    )


def table_tab() -> rx.Component:
    return rx.hstack(
        group(
            "Properties",
            rx.vstack(rx.text("Table Name:", size="1"), rx.text(GridState.ui["table"].to(str), size="2", weight="bold"), spacing="1"),
        ),
        group(
            "Tools",
            rbtn("rows_3", "Convert to Range", c("table.convertToRange"), big=True),
            rbtn("copy_minus", "Remove Duplicates", dlg("removeDuplicates"), big=True),
        ),
        group(
            "Table Style Options",
            stack2(
                [rbtn("panel_top", "Header Row", c("table.headerRow"), show_label=True), rbtn("columns_3", "First Column", c("table.firstColumn"), show_label=True)],
                [rbtn("sigma", "Total Row", c("table.totalRow"), show_label=True, id="rb-total-row"), rbtn("columns_3", "Last Column", c("table.lastColumn"), show_label=True)],
                [rbtn("rows_3", "Banded Rows", c("table.bandedRows"), show_label=True), rbtn("filter", "Filter Button", c("table.filterButton"), show_label=True)],
            ),
        ),
        group("Table Styles", menu("palette", "Styles", [(n.replace("TableStyle", ""), c("table.style", style=n)) for n in TABLE_STYLE_NAMES], big=True)),
        spacing="0",
        align="stretch",
    )


def chart_tab() -> rx.Component:
    return rx.hstack(
        group("Type", menu("chart_column", "Change Chart Type", [(l, GridState.chart_cmd("kind", k)) for l, k in CHART_KINDS], big=True)),
        group("Data", rbtn("arrow_right_left", "Switch Row/Column", GridState.chart_cmd("switch", ""), big=True), rbtn("table", "Select Data", dlg("selectData"), big=True)),
        group(
            "Chart Layouts",
            rbtn("type", "Titles", dlg("chartTitle"), "Chart and Axis Titles", big=True),
            menu("list", "Legend", [(p.title(), GridState.chart_cmd("legend", p)) for p in ("none", "right", "top", "left", "bottom")], big=True),
            menu("tag", "Data Labels", [("Show", GridState.chart_cmd("dataLabels", True)), ("Hide", GridState.chart_cmd("dataLabels", False))], big=True),
            menu("grid_3x3", "Gridlines", [("Show", GridState.chart_cmd("gridlines", True)), ("Hide", GridState.chart_cmd("gridlines", False))], big=True),
        ),
        group("Delete", rbtn("trash_2", "Delete Chart", GridState.chart_cmd("delete", ""), big=True)),
        spacing="0",
        align="stretch",
    )


TABS = ["File", "Home", "Insert", "Page Layout", "Formulas", "Data", "Review", "View", "Developer"]


def tab_button(name: str) -> rx.Component:
    return rx.box(
        rx.text(name, size="2"),
        on_click=GridState.set_tab(name),
        padding_x="12px",
        padding_y="5px",
        cursor="pointer",
        class_name="rg-tab",
        id=f"tab-{name.lower().replace(' ', '-')}",
        border_bottom=rx.cond(GridState.tab == name, f"3px solid {ACCENT}", "3px solid transparent"),
        color=rx.cond(GridState.tab == name, ACCENT, "#3C4043"),
        font_weight=rx.cond(GridState.tab == name, "600", "400"),
        background=rx.cond(name == "File", "transparent", "transparent"),
        _hover={"background": "#EEF2F5"},
    )


def developer_tab() -> rx.Component:
    from .vba import developer_tab as tab

    return tab()


def ribbon() -> rx.Component:
    return rx.vstack(
        rx.hstack(
            *[tab_button(t) for t in TABS],
            rx.cond(GridState.table_selected, tab_button("Table Design")),
            rx.cond(GridState.selected_kind == "chart", tab_button("Chart Design")),
            spacing="0",
            align="end",
            padding_x="8px",
            width="100%",
            border_bottom="1px solid #E3E6E9",
        ),
        rx.box(
            rx.match(
                GridState.tab,
                ("Home", home_tab()),
                ("Insert", insert_tab()),
                ("Page Layout", page_layout_tab()),
                ("Formulas", formulas_tab()),
                ("Data", data_tab()),
                ("Review", review_tab()),
                ("View", view_tab()),
                ("Developer", developer_tab()),
                ("Table Design", table_tab()),
                ("Chart Design", chart_tab()),
                home_tab(),
            ),
            width="100%",
            overflow_x="auto",
            overflow_y="hidden",
            padding_x="4px",
            background="#FFFFFF",
            class_name="rg-ribbon",
        ),
        spacing="0",
        width="100%",
        background="#FFFFFF",
        border_bottom="1px solid #D5D9DE",
    )
