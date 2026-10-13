"""Dialogs (Format Cells, Find & Replace, Go To, Paste Special, names, conditional formatting,
data validation, sort, filter, functions, links, notes, sheets, protection, charts, tables…)."""

from __future__ import annotations

from typing import Any, List, Optional, Sequence, Tuple

import reflex as rx

from gridcore.engine.cmd_format import FORMAT_PRESETS
from gridcore.engine.tables import TABLE_STYLE_NAMES

from ..state import CATEGORIES, GridState
from .ribbon import CHART_KINDS, PALETTE, STANDARD

D = GridState.dlg


def shell(name: str, title: str, *body, footer: Optional[rx.Component] = None, width: str = "440px") -> rx.Component:
    return rx.dialog.root(
        rx.dialog.content(
            rx.dialog.title(title, size="3"),
            rx.vstack(*body, spacing="2", width="100%"),
            rx.cond(GridState.dlg_error != "", rx.callout(GridState.dlg_error, icon="triangle_alert", color_scheme="red", size="1", margin_top="8px", id="dlg-error")),
            footer if footer is not None else ok_cancel(),
            max_width=width,
            id=f"dlg-{name}",
        ),
        open=GridState.dialog == name,
        on_open_change=GridState.dialog_open_change,
    )


def ok_cancel(ok: str = "OK", action: str = "ok", extra: Sequence[rx.Component] = ()) -> rx.Component:
    return rx.hstack(
        *extra,
        rx.spacer(),
        rx.button("Cancel", variant="soft", color_scheme="gray", on_click=GridState.close_dialog, id="dlg-cancel"),
        rx.button(ok, on_click=GridState.dialog_ok(action), id="dlg-ok"),
        spacing="2",
        width="100%",
        margin_top="12px",
    )


def label(text: str, width: str = "120px") -> rx.Component:
    return rx.text(text, size="2", width=width, flex_shrink="0")


def text_field(text: str, key: str, placeholder: str = "", type_: str = "text", id: Optional[str] = None, width: str = "100%") -> rx.Component:
    return rx.hstack(
        label(text),
        rx.box(
            rx.input(value=D[key].to(str), on_change=lambda v: GridState.set_dlg(key, v), placeholder=placeholder, type=type_, size="2", width=width, id=id or f"dlg-{key}"),
            flex="1",
            min_width="0",
        ),
        align="center",
        width="100%",
    )


def area_field(text: str, key: str, rows: int = 4) -> rx.Component:
    return rx.vstack(
        rx.text(text, size="2"),
        rx.text_area(value=D[key].to(str), on_change=lambda v: GridState.set_dlg(key, v), rows=str(rows), width="100%", id=f"dlg-{key}"),
        width="100%",
        spacing="1",
    )


def check(text: str, key: str) -> rx.Component:
    return rx.checkbox(text, checked=D[key].to(bool), on_change=lambda v: GridState.set_dlg(key, v), size="2", id=f"dlg-{key}")


def select_field(text: str, key: str, options: Sequence[Tuple[str, str]], width: str = "100%") -> rx.Component:
    return rx.hstack(
        label(text),
        rx.box(
            rx.select.root(
                rx.select.trigger(width=width, id=f"dlg-{key}"),
                rx.select.content(*[rx.select.item(lbl, value=val) for val, lbl in options]),
                value=D[key].to(str),
                on_change=lambda v: GridState.set_dlg(key, v),
                size="2",
            ),
            flex="1",
            min_width="0",
        ),
        align="center",
        width="100%",
    )


def radio_field(key: str, options: Sequence[Tuple[str, str]], direction: str = "column") -> rx.Component:
    return rx.radio_group.root(
        rx.flex(*[rx.text(rx.flex(rx.radio_group.item(value=val), lbl, gap="2"), as_="label", size="2") for val, lbl in options], direction=direction, gap="2"),
        value=D[key].to(str),
        on_change=lambda v: GridState.set_dlg(key, v),
    )


def swatches(key: str) -> rx.Component:
    rows = []
    for row in PALETTE[:4] + [STANDARD]:
        rows.append(rx.hstack(*[rx.box(width="18px", height="18px", background=c, border=rx.cond(D[key] == c, "2px solid #156082", "1px solid #D0D4D9"), cursor="pointer", on_click=GridState.set_dlg(key, c)) for c in row], spacing="1"))
    return rx.vstack(*rows, rx.hstack(rx.button("None", size="1", variant="soft", color_scheme="gray", on_click=GridState.set_dlg(key, "")), rx.input(value=D[key].to(str), on_change=lambda v: GridState.set_dlg(key, v), size="1", width="90px", placeholder="#RRGGBB"), spacing="2"), spacing="1")


# ---------------------------------------------------------------- Format Cells

NUMBER_CODES = [
    ("General", "General"), ("0", "Number (0)"), ("0.00", "Number (0.00)"), ("#,##0", "Number (#,##0)"), ("#,##0.00", "Number (#,##0.00)"),
    ('"$"#,##0.00', "Currency"), ('_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', "Accounting"),
    ("m/d/yyyy", "Date (3/14/2026)"), ("yyyy-mm-dd", "Date (2026-03-14)"), ("d-mmm-yy", "Date (14-Mar-26)"), ("dddd, mmmm d, yyyy", "Long Date"),
    ("h:mm AM/PM", "Time (1:30 PM)"), ("h:mm:ss", "Time (13:30:55)"), ("m/d/yyyy h:mm", "Date & Time"), ("0%", "Percentage (0%)"), ("0.00%", "Percentage (0.00%)"),
    ("# ?/?", "Fraction (# ?/?)"), ("0.00E+00", "Scientific"), ("@", "Text"), ("[$NT$-404]#,##0.00", "Currency (NT$)"), ('[$¥-zh-CN]#,##0.00', "Currency (¥)"),
]


def format_cells() -> rx.Component:
    tabs = rx.tabs.root(
        rx.tabs.list(
            rx.tabs.trigger("Number", value="number"),
            rx.tabs.trigger("Alignment", value="alignment"),
            rx.tabs.trigger("Font", value="font"),
            rx.tabs.trigger("Border", value="border"),
            rx.tabs.trigger("Fill", value="fill"),
            rx.tabs.trigger("Protection", value="protection"),
        ),
        rx.tabs.content(
            rx.vstack(
                rx.text("Format code", size="2"),
                rx.input(value=D["numFmt"].to(str), on_change=lambda v: GridState.set_dlg("numFmt", v), id="dlg-numFmt", width="100%"),
                rx.scroll_area(
                    rx.vstack(*[rx.box(rx.text(lbl, size="2"), padding="2px 6px", width="100%", cursor="pointer", background=rx.cond(D["numFmt"] == code, "#D7E3EA", "transparent"), on_click=GridState.set_dlg("numFmt", code)) for code, lbl in NUMBER_CODES], spacing="0", width="100%"),
                    type="always",
                    scrollbars="vertical",
                    style={"height": "200px", "border": "1px solid #E3E6E9"},
                ),
                spacing="2",
                padding_top="8px",
            ),
            value="number",
        ),
        rx.tabs.content(
            rx.vstack(
                select_field("Horizontal", "h", [("general", "General"), ("left", "Left (Indent)"), ("center", "Center"), ("right", "Right (Indent)"), ("fill", "Fill"), ("justify", "Justify"), ("centerContinuous", "Center Across Selection"), ("distributed", "Distributed")]),
                select_field("Vertical", "v", [("top", "Top"), ("center", "Center"), ("bottom", "Bottom"), ("justify", "Justify"), ("distributed", "Distributed")]),
                text_field("Indent", "indent", type_="number"),
                text_field("Rotation (°)", "rotation", type_="number"),
                check("Wrap text", "wrap"),
                check("Shrink to fit", "shrink"),
                spacing="2",
                padding_top="8px",
            ),
            value="alignment",
        ),
        rx.tabs.content(
            rx.vstack(
                text_field("Font", "fontName"),
                text_field("Size", "fontSize", type_="number"),
                rx.hstack(check("Bold", "bold"), check("Italic", "italic"), check("Strikethrough", "strike"), spacing="4"),
                select_field("Underline", "underline", [("none", "None"), ("single", "Single"), ("double", "Double"), ("singleAccounting", "Single Accounting"), ("doubleAccounting", "Double Accounting")]),
                rx.text("Color", size="2"),
                swatches("fontColor"),
                spacing="2",
                padding_top="8px",
            ),
            value="font",
        ),
        rx.tabs.content(
            rx.vstack(
                select_field("Presets", "borderPreset", [("", "(no change)"), ("none", "None"), ("outside", "Outline"), ("inside", "Inside"), ("all", "All"), ("bottom", "Bottom"), ("top", "Top"), ("left", "Left"), ("right", "Right"), ("diagDown", "Diagonal down"), ("diagUp", "Diagonal up")]),
                select_field("Style", "borderStyle", [("thin", "Thin"), ("medium", "Medium"), ("thick", "Thick"), ("double", "Double"), ("dashed", "Dashed"), ("dotted", "Dotted"), ("hair", "Hair")]),
                rx.text("Color", size="2"),
                swatches("borderColor"),
                spacing="2",
                padding_top="8px",
            ),
            value="border",
        ),
        rx.tabs.content(rx.vstack(rx.text("Background color", size="2"), swatches("fillColor"), spacing="2", padding_top="8px"), value="fill"),
        rx.tabs.content(
            rx.vstack(
                check("Locked", "locked"),
                check("Hidden", "hidden"),
                rx.text("Locking cells or hiding formulas has no effect until you protect the worksheet (Review › Protect Sheet).", size="1", color="#6B7280"),
                spacing="2",
                padding_top="8px",
            ),
            value="protection",
        ),
        value=D["tab"].to(str),
        on_change=lambda v: GridState.set_dlg("tab", v),
        width="100%",
    )
    return shell("formatCells", "Format Cells", tabs, width="520px")


# ---------------------------------------------------------------- Find & Replace


def find_replace() -> rx.Component:
    body = [
        text_field("Find what", "what", id="dlg-what"),
        rx.cond(D["mode"] == "replace", text_field("Replace with", "with", id="dlg-with")),
        rx.hstack(check("Match case", "matchCase"), check("Match entire cell contents", "wholeCell"), spacing="4"),
        rx.hstack(
            select_field("Within", "within", [("sheet", "Sheet"), ("workbook", "Workbook")], width="120px"),
            spacing="2",
        ),
        rx.cond(D["mode"] == "find", select_field("Look in", "lookIn", [("formulas", "Formulas"), ("values", "Values"), ("comments", "Notes")], width="120px")),
        rx.text(D["status"].to(str), size="1", color="#6B7280", id="dlg-status"),
        rx.cond(
            GridState.dlg_list.length() > 0,
            rx.scroll_area(
                rx.vstack(
                    rx.foreach(
                        GridState.dlg_list,
                        lambda m: rx.hstack(
                            rx.text(m["sheet"].to(str), size="1", width="80px"),
                            rx.text(m["cell"].to(str), size="1", width="60px"),
                            rx.text(m["value"].to(str), size="1"),
                            cursor="pointer",
                            width="100%",
                            _hover={"background": "#EEF2F5"},
                            on_click=GridState.dialog_ok("goto:" + m["sheet"].to(str) + ":" + m["cell"].to(str)),
                        ),
                    ),
                    spacing="0",
                    width="100%",
                ),
                style={"height": "160px", "border": "1px solid #E3E6E9"},
                type="always",
                scrollbars="vertical",
            ),
        ),
    ]
    footer = rx.hstack(
        rx.cond(D["mode"] == "replace", rx.button("Replace All", variant="soft", on_click=GridState.dialog_ok("replaceAll"), id="dlg-replace-all")),
        rx.cond(D["mode"] == "replace", rx.button("Replace", variant="soft", on_click=GridState.dialog_ok("replace"), id="dlg-replace")),
        rx.button("Find All", variant="soft", on_click=GridState.dialog_ok("findAll"), id="dlg-find-all"),
        rx.button("Find Next", on_click=GridState.dialog_ok("findNext"), id="dlg-find-next"),
        rx.spacer(),
        rx.button("Close", variant="soft", color_scheme="gray", on_click=GridState.close_dialog),
        spacing="2",
        width="100%",
        margin_top="12px",
        wrap="wrap",
    )
    return rx.fragment(
        shell("find", "Find", *body, footer=footer, width="520px"),
        shell("replace", "Find and Replace", *body, footer=footer, width="520px"),
    )


def go_to() -> rx.Component:
    return shell(
        "goTo",
        "Go To",
        text_field("Reference", "reference", placeholder="B5, Sheet2!A1:C3 or a name", id="dlg-reference"),
        rx.divider(),
        select_field("Special", "special", [("blanks", "Blanks"), ("constants", "Constants"), ("formulas", "Formulas"), ("errors", "Errors"), ("numbers", "Numbers"), ("text", "Text"), ("comments", "Notes"), ("currentRegion", "Current region"), ("lastCell", "Last cell"), ("visible", "Visible cells only"), ("conditionalFormats", "Conditional formats"), ("dataValidation", "Data validation")]),
        footer=ok_cancel(extra=[rx.button("Go To Special", variant="soft", on_click=GridState.dialog_ok("special"))]),
    )


def paste_special() -> rx.Component:
    return shell(
        "pasteSpecial",
        "Paste Special",
        rx.text("Paste", size="2", weight="bold"),
        radio_field("what", [("all", "All"), ("formulas", "Formulas"), ("values", "Values"), ("formats", "Formats"), ("comments", "Notes"), ("validation", "Validation"), ("allExceptBorders", "All except borders"), ("columnWidths", "Column widths"), ("formulasAndNumberFormats", "Formulas and number formats"), ("valuesAndNumberFormats", "Values and number formats")]),
        rx.text("Operation", size="2", weight="bold"),
        radio_field("operation", [("none", "None"), ("add", "Add"), ("subtract", "Subtract"), ("multiply", "Multiply"), ("divide", "Divide")], direction="row"),
        rx.hstack(check("Skip blanks", "skipBlanks"), check("Transpose", "transpose"), spacing="4"),
    )


# ---------------------------------------------------------------- names


def define_name() -> rx.Component:
    return shell(
        "defineName",
        "New Name",
        text_field("Name", "name", id="dlg-name"),
        text_field("Scope", "scope"),
        text_field("Refers to", "refersTo", id="dlg-refersTo"),
        text_field("Comment", "comment"),
    )


def name_manager() -> rx.Component:
    return shell(
        "nameManager",
        "Name Manager",
        rx.scroll_area(
            rx.vstack(
                rx.hstack(rx.text("Name", size="1", weight="bold", width="110px"), rx.text("Value", size="1", weight="bold", width="110px"), rx.text("Refers To", size="1", weight="bold", width="180px"), rx.text("Scope", size="1", weight="bold"), width="100%"),
                rx.foreach(
                    GridState.dlg_list,
                    lambda n: rx.hstack(
                        rx.text(n["name"].to(str), size="1", width="110px"),
                        rx.text(n["value"].to(str), size="1", width="110px", overflow="hidden", white_space="nowrap"),
                        rx.text(n["refersTo"].to(str), size="1", width="180px", overflow="hidden", white_space="nowrap"),
                        rx.text(n["scope"].to(str), size="1", width="70px"),
                        rx.icon_button(rx.icon("trash_2", size=12), size="1", variant="ghost", color_scheme="red", on_click=GridState.dialog_ok("delete:" + n["name"].to(str))),
                        width="100%",
                        align="center",
                    ),
                ),
                spacing="1",
                width="100%",
            ),
            style={"height": "240px"},
            type="always",
            scrollbars="vertical",
        ),
        footer=rx.hstack(rx.button("New…", variant="soft", on_click=GridState.open_dialog("defineName", {})), rx.spacer(), rx.button("Close", on_click=GridState.close_dialog), width="100%", margin_top="12px"),
        width="620px",
    )


# ---------------------------------------------------------------- conditional formatting

PRESET_OPTIONS = [("lightRed", "Light Red Fill with Dark Red Text"), ("yellow", "Yellow Fill with Dark Yellow Text"), ("green", "Green Fill with Dark Green Text"), ("lightRedFill", "Light Red Fill"), ("redText", "Red Text"), ("redBorder", "Red Border")]


def cf_highlight() -> rx.Component:
    return shell(
        "cfHighlight",
        "Highlight Cells Rule",
        rx.match(
            D["kind"].to(str),
            ("greaterThan", rx.text("Format cells that are GREATER THAN:", size="2")),
            ("lessThan", rx.text("Format cells that are LESS THAN:", size="2")),
            ("between", rx.text("Format cells that are BETWEEN:", size="2")),
            ("equal", rx.text("Format cells that are EQUAL TO:", size="2")),
            ("text", rx.text("Format cells that contain the text:", size="2")),
            ("date", rx.text("Format cells that contain a date occurring:", size="2")),
            ("duplicate", rx.text("Format cells that contain:", size="2")),
            rx.text("Format cells:", size="2"),
        ),
        rx.match(
            D["kind"].to(str),
            ("date", select_field("Period", "period", [("yesterday", "Yesterday"), ("today", "Today"), ("tomorrow", "Tomorrow"), ("last7Days", "In the last 7 days"), ("lastWeek", "Last week"), ("thisWeek", "This week"), ("nextWeek", "Next week"), ("lastMonth", "Last month"), ("thisMonth", "This month"), ("nextMonth", "Next month")])),
            ("duplicate", select_field("Values", "value", [("duplicate", "Duplicate"), ("unique", "Unique")])),
            text_field("Value", "value", id="dlg-value"),
        ),
        rx.cond(D["kind"] == "between", text_field("and", "value2", id="dlg-value2")),
        select_field("with", "preset", PRESET_OPTIONS),
    )


def cf_top() -> rx.Component:
    return shell(
        "cfTop",
        "Top/Bottom Rule",
        select_field("Rule", "kind", [("top", "Top N items"), ("topPercent", "Top N%"), ("bottom", "Bottom N items"), ("bottomPercent", "Bottom N%"), ("above", "Above average"), ("below", "Below average")]),
        rx.cond((D["kind"] != "above") & (D["kind"] != "below"), text_field("N", "rank", type_="number")),
        select_field("with", "preset", PRESET_OPTIONS),
    )


def cf_rule() -> rx.Component:
    return shell(
        "cfRule",
        "New Formatting Rule",
        rx.text("Use a formula to determine which cells to format. Write it for the top-left cell of the selection.", size="1", color="#6B7280"),
        text_field("Formula", "formula", placeholder="=$B2>100", id="dlg-formula"),
        rx.text("Fill", size="2"),
        swatches("fill"),
        rx.text("Font color", size="2"),
        swatches("font"),
        check("Bold", "bold"),
        width="480px",
    )


def cf_manage() -> rx.Component:
    return shell(
        "cfManage",
        "Conditional Formatting Rules Manager",
        rx.scroll_area(
            rx.vstack(
                rx.foreach(
                    GridState.dlg_list,
                    lambda r: rx.hstack(
                        rx.box(width="14px", height="14px", background=rx.cond(r["fill"], r["fill"].to(str), "#FFFFFF"), border="1px solid #C0C6CC"),
                        rx.text(r["description"].to(str), size="1", width="240px"),
                        rx.text(r["appliesTo"].to(str), size="1", width="120px"),
                        rx.icon_button(rx.icon("chevron_up", size=12), size="1", variant="ghost", on_click=GridState.dialog_ok("up:" + r["index"].to(str))),
                        rx.icon_button(rx.icon("chevron_down", size=12), size="1", variant="ghost", on_click=GridState.dialog_ok("down:" + r["index"].to(str))),
                        rx.icon_button(rx.icon("trash_2", size=12), size="1", variant="ghost", color_scheme="red", on_click=GridState.dialog_ok("delete:" + r["index"].to(str))),
                        align="center",
                        width="100%",
                    ),
                ),
                spacing="1",
                width="100%",
            ),
            style={"height": "220px"},
            type="always",
            scrollbars="vertical",
        ),
        footer=rx.hstack(rx.button("Clear All Rules on Sheet", variant="soft", color_scheme="red", on_click=GridState.dialog_ok("clearSheet")), rx.spacer(), rx.button("Close", on_click=GridState.close_dialog), width="100%", margin_top="12px"),
        width="560px",
    )


# ---------------------------------------------------------------- data


def validation() -> rx.Component:
    return shell(
        "validation",
        "Data Validation",
        rx.tabs.root(
            rx.tabs.list(rx.tabs.trigger("Settings", value="settings"), rx.tabs.trigger("Input Message", value="input"), rx.tabs.trigger("Error Alert", value="error")),
            rx.tabs.content(
                rx.vstack(
                    select_field("Allow", "kind", [("any", "Any value"), ("whole", "Whole number"), ("decimal", "Decimal"), ("list", "List"), ("date", "Date"), ("time", "Time"), ("textLength", "Text length"), ("custom", "Custom")]),
                    rx.cond(
                        (D["kind"] != "any") & (D["kind"] != "list") & (D["kind"] != "custom"),
                        select_field("Data", "operator", [("between", "between"), ("notBetween", "not between"), ("equal", "equal to"), ("notEqual", "not equal to"), ("greaterThan", "greater than"), ("lessThan", "less than"), ("greaterThanOrEqual", "greater than or equal to"), ("lessThanOrEqual", "less than or equal to")]),
                    ),
                    rx.cond(D["kind"] == "list", text_field("Source", "f1", placeholder="Yes,No or =$A$1:$A$5", id="dlg-f1")),
                    rx.cond(D["kind"] == "custom", text_field("Formula", "f1", placeholder="=A1>0", id="dlg-f1")),
                    rx.cond((D["kind"] != "any") & (D["kind"] != "list") & (D["kind"] != "custom"), text_field("Minimum / Value", "f1", id="dlg-f1")),
                    rx.cond(((D["operator"] == "between") | (D["operator"] == "notBetween")) & (D["kind"] != "any") & (D["kind"] != "list") & (D["kind"] != "custom"), text_field("Maximum", "f2", id="dlg-f2")),
                    rx.hstack(check("Ignore blank", "allowBlank"), rx.cond(D["kind"] == "list", check("In-cell dropdown", "inCellDropdown")), spacing="4"),
                    spacing="2",
                    padding_top="8px",
                ),
                value="settings",
            ),
            rx.tabs.content(rx.vstack(text_field("Title", "inputTitle"), area_field("Input message", "inputMessage"), spacing="2", padding_top="8px"), value="input"),
            rx.tabs.content(
                rx.vstack(select_field("Style", "errorStyle", [("stop", "Stop"), ("warning", "Warning"), ("information", "Information")]), text_field("Title", "errorTitle"), area_field("Error message", "errorMessage"), spacing="2", padding_top="8px"),
                value="error",
            ),
            value=D["tab"].to(str),
            on_change=lambda v: GridState.set_dlg("tab", v),
            width="100%",
        ),
        footer=ok_cancel(extra=[rx.button("Clear All", variant="soft", color_scheme="gray", on_click=GridState.dialog_ok("clear"))]),
        width="480px",
    )


def sort_dialog() -> rx.Component:
    def key_row(i: int) -> rx.Component:
        return rx.hstack(
            rx.text("Sort by" if i == 1 else "Then by", size="2", width="70px"),
            rx.select.root(
                rx.select.trigger(width="180px", placeholder="(none)", id=f"dlg-k{i}"),
                rx.select.content(rx.foreach(D["cols"].to(list[dict[str, str]]), lambda c: rx.select.item(c["label"], value=c["value"]))),
                value=D[f"k{i}"].to(str),
                on_change=lambda v: GridState.set_dlg(f"k{i}", v),
                size="2",
            ),
            rx.select.root(
                rx.select.trigger(width="120px"),
                rx.select.content(rx.select.item("A to Z", value="asc"), rx.select.item("Z to A", value="desc")),
                value=D[f"o{i}"].to(str),
                on_change=lambda v: GridState.set_dlg(f"o{i}", v),
                size="2",
            ),
            align="center",
        )

    return shell("sort", "Sort", rx.text("Range: ", D["range"].to(str), size="1", color="#6B7280"), check("My data has headers", "header"), key_row(1), key_row(2), key_row(3), check("Case sensitive", "matchCase"), width="480px")


def filter_dialog() -> rx.Component:
    return shell(
        "filter",
        "Filter",
        rx.hstack(
            rx.button(rx.icon("arrow_down_a_z", size=14), "Sort A to Z", size="1", variant="soft", on_click=GridState.dialog_ok("sortAsc")),
            rx.button(rx.icon("arrow_up_z_a", size=14), "Sort Z to A", size="1", variant="soft", on_click=GridState.dialog_ok("sortDesc")),
            rx.cond(D["active"], rx.button(rx.icon("filter_x", size=14), "Clear Filter", size="1", variant="soft", color_scheme="red", on_click=GridState.dialog_ok("clear"))),
            spacing="2",
            wrap="wrap",
        ),
        rx.segmented_control.root(
            rx.segmented_control.item("Values", value="values"),
            rx.segmented_control.item("Condition", value="custom"),
            value=D["mode"].to(str),
            on_change=lambda v: GridState.set_dlg("mode", v),
            size="1",
        ),
        rx.cond(
            D["mode"] == "custom",
            rx.vstack(
                rx.hstack(
                    rx.select.root(rx.select.trigger(placeholder="condition", width="170px", id="dlg-op"), rx.select.content(*[rx.select.item(l, value=v) for v, l in CUSTOM_OPS]), value=D["op"].to(str), on_change=lambda v: GridState.set_dlg("op", v), size="2"),
                    rx.input(value=D["value"].to(str), on_change=lambda v: GridState.set_dlg("value", v), size="2", id="dlg-filter-value"),
                ),
                rx.radio_group.root(rx.flex(rx.text(rx.flex(rx.radio_group.item(value="and"), "And", gap="2"), as_="label", size="2"), rx.text(rx.flex(rx.radio_group.item(value="or"), "Or", gap="2"), as_="label", size="2"), gap="4"), value=rx.cond(D["and"], "and", "or"), on_change=lambda v: GridState.set_dlg("and", v == "and")),
                rx.hstack(
                    rx.select.root(rx.select.trigger(placeholder="condition", width="170px"), rx.select.content(*[rx.select.item(l, value=v) for v, l in CUSTOM_OPS]), value=D["op2"].to(str), on_change=lambda v: GridState.set_dlg("op2", v), size="2"),
                    rx.input(value=D["value2"].to(str), on_change=lambda v: GridState.set_dlg("value2", v), size="2"),
                ),
                spacing="2",
            ),
            rx.vstack(
                rx.input(value=D["search"].to(str), on_change=lambda v: GridState.set_dlg("search", v), placeholder="Search", size="2", width="100%", id="dlg-filter-search"),
                rx.hstack(rx.button("Select All", size="1", variant="ghost", on_click=GridState.filter_select_all(True)), rx.button("Clear", size="1", variant="ghost", on_click=GridState.filter_select_all(False)), spacing="2"),
                rx.scroll_area(
                    rx.vstack(
                        rx.foreach(GridState.dlg_list, lambda v: rx.checkbox(v["text"].to(str), checked=v["checked"].to(bool), on_change=lambda _: GridState.toggle_filter_value(v["i"].to(int)), size="2", class_name="rg-filter-item")),
                        rx.cond(D["blanks"], rx.checkbox("(Blanks)", checked=D["blanksChecked"].to(bool), on_change=lambda v: GridState.set_dlg("blanksChecked", v), size="2")),
                        spacing="1",
                        width="100%",
                    ),
                    style={"height": "220px", "border": "1px solid #E3E6E9", "padding": "4px"},
                    type="always",
                    scrollbars="vertical",
                ),
                spacing="2",
                width="100%",
            ),
        ),
        width="380px",
    )


CUSTOM_OPS = [
    ("equal", "equals"), ("notEqual", "does not equal"), ("greaterThan", "is greater than"), ("greaterThanOrEqual", "is greater than or equal to"),
    ("lessThan", "is less than"), ("lessThanOrEqual", "is less than or equal to"), ("beginsWith", "begins with"), ("endsWith", "ends with"),
    ("contains", "contains"), ("notContains", "does not contain"),
]


def remove_duplicates() -> rx.Component:
    return shell(
        "removeDuplicates",
        "Remove Duplicates",
        rx.text("Range: ", D["range"].to(str), size="1", color="#6B7280"),
        check("My data has headers", "header"),
        rx.text("Columns", size="2", weight="bold"),
        rx.scroll_area(
            rx.vstack(rx.foreach(GridState.dlg_list, lambda c: rx.checkbox(c["label"].to(str), checked=c["checked"].to(bool), on_change=lambda _: GridState.toggle_list_item("cols", c["i"].to(int)), size="2")), spacing="1"),
            style={"height": "180px"},
            type="always",
            scrollbars="vertical",
        ),
    )


def text_to_columns() -> rx.Component:
    return shell(
        "textToColumns",
        "Convert Text to Columns",
        rx.text("Delimiters", size="2", weight="bold"),
        rx.hstack(check("Tab", "tab"), check("Semicolon", "semicolon"), check("Comma", "comma"), check("Space", "space"), spacing="3", wrap="wrap"),
        text_field("Other", "other"),
        check("Treat consecutive delimiters as one", "consecutive"),
    )


# ---------------------------------------------------------------- functions


def insert_function() -> rx.Component:
    return shell(
        "insertFunction",
        "Insert Function",
        rx.input(value=D["search"].to(str), on_change=lambda v: GridState.set_dlg("search", v), placeholder="Search for a function", size="2", width="100%", id="dlg-fn-search"),
        select_field("Category", "category", [("Most Used", "Most Recently Used"), ("All", "All")] + [(c, c) for c in CATEGORIES]),
        rx.scroll_area(
            rx.vstack(
                rx.foreach(
                    GridState.fn_list,
                    lambda f: rx.box(rx.text(f["name"], size="2"), padding="1px 6px", width="100%", cursor="pointer", background=rx.cond(D["selected"] == f["name"], "#D7E3EA", "transparent"), on_click=GridState.set_dlg("selected", f["name"]), on_double_click=GridState.dialog_ok("ok")),
                ),
                spacing="0",
                width="100%",
            ),
            style={"height": "200px", "border": "1px solid #E3E6E9"},
            type="always",
            scrollbars="vertical",
        ),
        rx.text(GridState.fn_selected["signature"], size="2", weight="bold"),
        rx.text(GridState.fn_selected["description"], size="1", color="#4B5563"),
        width="520px",
    )


# ---------------------------------------------------------------- insert / review


def hyperlink() -> rx.Component:
    return shell(
        "hyperlink",
        "Insert Link",
        text_field("Text to display", "text"),
        text_field("Address", "target", placeholder="https://example.com or Sheet2!A1", id="dlg-target"),
        text_field("ScreenTip", "tooltip"),
        footer=ok_cancel(extra=[rx.button("Remove Link", variant="soft", color_scheme="gray", on_click=GridState.dialog_ok("remove"))]),
    )


def note() -> rx.Component:
    return shell(
        "note",
        rx.cond(D["threaded"], "Comment", "Note"),
        area_field("Text", "text", rows=5),
        footer=ok_cancel(extra=[rx.button("Delete", variant="soft", color_scheme="red", on_click=GridState.dialog_ok("delete"))]),
    )


def size_dialogs() -> rx.Component:
    return rx.fragment(
        shell("rowHeight", "Row Height", text_field("Row height (px)", "value", type_="number", id="dlg-value")),
        shell("columnWidth", "Column Width", text_field("Column width (px)", "value", type_="number", id="dlg-value")),
        shell("defaultWidth", "Standard Width", text_field("Standard column width (px)", "value", type_="number", id="dlg-value")),
    )


def cells_dialogs() -> rx.Component:
    return rx.fragment(
        shell("insertCells", "Insert", radio_field("shift", [("right", "Shift cells right"), ("down", "Shift cells down"), ("row", "Entire row"), ("col", "Entire column")])),
        shell("deleteCells", "Delete", radio_field("shift", [("left", "Shift cells left"), ("up", "Shift cells up"), ("row", "Entire row"), ("col", "Entire column")])),
    )


def sheet_dialogs() -> rx.Component:
    colors = rx.vstack(
        *[rx.hstack(*[rx.box(width="20px", height="20px", background=col, border="1px solid #D0D4D9", cursor="pointer", on_click=GridState.dialog_ok(col)) for col in row], spacing="1") for row in PALETTE[:3] + [STANDARD]],
        spacing="1",
    )
    return rx.fragment(
        shell("renameSheet", "Rename Sheet", text_field("Name", "name", id="dlg-sheet-name")),
        shell(
            "moveSheet",
            "Move or Copy",
            rx.text("Before sheet:", size="2"),
            rx.select.root(
                rx.select.trigger(width="100%"),
                rx.select.content(rx.foreach(GridState.dlg_sheets, lambda n, i: rx.select.item(n, value=i.to(str)))),
                value=D["to"].to(str),
                on_change=lambda v: GridState.set_dlg("to", v),
            ),
            check("Create a copy", "copy"),
        ),
        shell("tabColor", "Tab Color", colors, footer=rx.hstack(rx.button("No Color", variant="soft", on_click=GridState.dialog_ok("ok")), rx.spacer(), rx.button("Cancel", variant="soft", color_scheme="gray", on_click=GridState.close_dialog), width="100%", margin_top="12px")),
        shell(
            "unhideSheet",
            "Unhide",
            rx.select.root(rx.select.trigger(width="100%"), rx.select.content(rx.foreach(GridState.dlg_sheets, lambda n: rx.select.item(n, value=n))), value=D["sheet"].to(str), on_change=lambda v: GridState.set_dlg("sheet", v)),
        ),
    )


def protect_dialogs() -> rx.Component:
    return rx.fragment(
        shell(
            "protectSheet",
            "Protect Sheet",
            text_field("Password (optional)", "password", type_="password", id="dlg-password"),
            rx.text("Allow all users of this worksheet to:", size="2"),
            rx.vstack(
                check("Format cells", "formatCells"), check("Format columns", "formatColumns"), check("Format rows", "formatRows"),
                check("Insert rows", "insertRows"), check("Insert columns", "insertColumns"), check("Delete rows", "deleteRows"),
                check("Delete columns", "deleteColumns"), check("Sort", "sort"), check("Use AutoFilter", "autoFilter"),
                spacing="1",
            ),
        ),
        shell("unprotectSheet", "Unprotect Sheet", text_field("Password", "password", type_="password", id="dlg-password")),
        shell("protectWorkbook", "Protect Structure", text_field("Password (optional)", "password", type_="password")),
        shell("unprotectWorkbook", "Unprotect Workbook", text_field("Password", "password", type_="password")),
    )


def chart_dialogs() -> rx.Component:
    return rx.fragment(
        shell(
            "insertChart",
            "Insert Chart",
            rx.scroll_area(
                rx.vstack(*[rx.box(rx.text(lbl, size="2"), padding="2px 8px", width="100%", cursor="pointer", background=rx.cond(D["kind"] == k, "#D7E3EA", "transparent"), on_click=GridState.set_dlg("kind", k)) for lbl, k in CHART_KINDS], spacing="0", width="100%"),
                style={"height": "240px", "border": "1px solid #E3E6E9"},
                type="always",
                scrollbars="vertical",
            ),
            text_field("Title (optional)", "title"),
        ),
        shell("chartTitle", "Chart Titles", text_field("Chart title", "title"), text_field("Horizontal axis", "xTitle"), text_field("Vertical axis", "yTitle")),
        shell("selectData", "Select Data Source", text_field("Chart data range", "range", placeholder="Sheet1!$A$1:$C$10")),
        shell(
            "sparkline",
            "Create Sparklines",
            select_field("Type", "kind", [("line", "Line"), ("column", "Column"), ("stacked", "Win/Loss")]),
            text_field("Data range", "data"),
            text_field("Location range", "location"),
        ),
    )


def table_dialog() -> rx.Component:
    return shell(
        "insertTable",
        "Create Table",
        text_field("Where is the data?", "range", id="dlg-table-range"),
        check("My table has headers", "header"),
        select_field("Style", "style", [(n, n.replace("TableStyle", "")) for n in TABLE_STYLE_NAMES]),
    )


def properties_dialog() -> rx.Component:
    return shell("properties", "Properties", text_field("Title", "title"), text_field("Subject", "subject"), text_field("Author", "author"), text_field("Company", "company"), text_field("Keywords", "keywords"), area_field("Comments", "description", rows=3))


def upload_dialogs() -> rx.Component:
    return rx.fragment(
        shell(
            "open",
            "Open",
            rx.upload(
                rx.vstack(rx.icon("folder_open", size=28, color="#6B7280"), rx.text("Drop an .xlsx, .csv or .tsv file here, or click to choose one.", size="2", color="#4B5563", text_align="center"), align="center", spacing="2"),
                id="open_upload",
                accept={"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx", ".xlsm"], "text/csv": [".csv"], "text/tab-separated-values": [".tsv", ".txt"]},
                max_files=1,
                multiple=False,
                on_drop=GridState.open_upload(rx.upload_files(upload_id="open_upload")),
                border="1px dashed #A6ADB4",
                padding="28px",
                width="100%",
                cursor="pointer",
            ),
            footer=rx.hstack(rx.spacer(), rx.button("Cancel", variant="soft", color_scheme="gray", on_click=GridState.close_dialog), width="100%", margin_top="12px"),
        ),
        shell(
            "insertPicture",
            "Insert Picture",
            rx.upload(
                rx.vstack(rx.icon("image", size=28, color="#6B7280"), rx.text("Drop a PNG, JPEG, GIF, WebP or SVG picture here, or click to choose one.", size="2", color="#4B5563", text_align="center"), align="center", spacing="2"),
                id="picture_upload",
                accept={"image/*": [".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"]},
                max_files=1,
                multiple=False,
                on_drop=GridState.picture_upload(rx.upload_files(upload_id="picture_upload")),
                border="1px dashed #A6ADB4",
                padding="28px",
                width="100%",
                cursor="pointer",
            ),
            footer=rx.hstack(rx.spacer(), rx.button("Cancel", variant="soft", color_scheme="gray", on_click=GridState.close_dialog), width="100%", margin_top="12px"),
        ),
    )


SHORTCUTS = [
    ("Enter / Shift+Enter", "Commit and move down / up"), ("Tab / Shift+Tab", "Commit and move right / left"), ("Alt+Enter", "Line break in a cell"),
    ("Ctrl+Enter", "Fill the selection with the entry"), ("Ctrl+Shift+Enter", "Array formula"), ("F2", "Edit the active cell"), ("F4", "Toggle $ in a reference while editing"),
    ("Ctrl+Arrow", "Jump to the edge of data"), ("Ctrl+Home / Ctrl+End", "First / last cell"), ("Ctrl+A", "Select region / all"),
    ("Ctrl+Z / Ctrl+Y", "Undo / Redo"), ("Ctrl+C / X / V", "Copy / Cut / Paste"), ("Ctrl+Alt+V", "Paste Special"), ("Ctrl+D / Ctrl+R", "Fill down / right"),
    ("Ctrl+B / I / U / 5", "Bold / Italic / Underline / Strikethrough"), ("Ctrl+1", "Format Cells"), ("Ctrl+F / Ctrl+H", "Find / Replace"), ("Ctrl+G or F5", "Go To"),
    ("Ctrl+K", "Insert link"), ("Ctrl+T", "Create table"), ("Ctrl+Shift+L", "Filter"), ("Alt+=", "AutoSum"), ("Ctrl+;", "Today's date"),
    ("Ctrl+Shift+$ % # @ ! ~", "Currency, percent, date, time, number, general formats"), ("Ctrl+` ", "Show formulas"), ("Ctrl+PageUp/PageDown", "Previous / next sheet"),
    ("Shift+F11", "Insert sheet"), ("Ctrl+- / Ctrl+Shift+=", "Delete / insert cells"), ("Ctrl+9 / Ctrl+0", "Hide rows / columns"), ("F9", "Calculate now"), ("Ctrl+S", "Save (download .xlsx)"),
]


def shortcuts_dialog() -> rx.Component:
    return shell(
        "shortcuts",
        "Keyboard Shortcuts",
        rx.scroll_area(
            rx.vstack(*[rx.hstack(rx.code(k, size="1", width="190px"), rx.text(v, size="1"), width="100%") for k, v in SHORTCUTS], spacing="1"),
            style={"height": "360px"},
            type="always",
            scrollbars="vertical",
        ),
        footer=rx.hstack(rx.spacer(), rx.button("Close", on_click=GridState.close_dialog), width="100%", margin_top="12px"),
        width="560px",
    )


def goal_seek() -> rx.Component:
    return shell(
        "goalSeek",
        "Goal Seek",
        text_field("Set cell", "setCell", id="dlg-setCell"),
        text_field("To value", "toValue", id="dlg-toValue"),
        text_field("By changing cell", "byChangingCell", id="dlg-byChangingCell"),
    )


def all_dialogs() -> rx.Component:
    return rx.fragment(
        goal_seek(),
        format_cells(),
        find_replace(),
        go_to(),
        paste_special(),
        define_name(),
        name_manager(),
        cf_highlight(),
        cf_top(),
        cf_rule(),
        cf_manage(),
        validation(),
        sort_dialog(),
        filter_dialog(),
        remove_duplicates(),
        text_to_columns(),
        insert_function(),
        hyperlink(),
        note(),
        size_dialogs(),
        cells_dialogs(),
        sheet_dialogs(),
        protect_dialogs(),
        chart_dialogs(),
        table_dialog(),
        properties_dialog(),
        upload_dialogs(),
        shortcuts_dialog(),
    )
