"""Window chrome: title bar with quick access, the File backstage, sheet tabs and status bar."""

from __future__ import annotations

import reflex as rx

from ..state import GridState

ACCENT = "#156082"


def app_icon(size: int = 20) -> rx.Component:
    # Our own mark: a 2x2 grid with an accent cell.
    return rx.html(
        f'<svg width="{size}" height="{size}" viewBox="0 0 20 20"><rect x="1" y="1" width="18" height="18" rx="3" fill="#FFFFFF" fill-opacity="0.18"/>'
        '<rect x="3.5" y="3.5" width="6" height="6" rx="1" fill="#FFFFFF"/><rect x="10.5" y="3.5" width="6" height="6" rx="1" fill="#FFFFFF" fill-opacity="0.55"/>'
        '<rect x="3.5" y="10.5" width="6" height="6" rx="1" fill="#FFFFFF" fill-opacity="0.55"/><rect x="10.5" y="10.5" width="6" height="6" rx="1" fill="#F2C12E"/></svg>'
    )


def title_bar() -> rx.Component:
    def qa(icon: str, label: str, on_click, id: str) -> rx.Component:
        return rx.tooltip(rx.icon_button(rx.icon(icon, size=15), variant="ghost", size="1", on_click=on_click, color="#FFFFFF", id=id), content=label)

    return rx.hstack(
        app_icon(),
        rx.text("ReflexGrid", weight="bold", size="2", color="#FFFFFF"),
        rx.divider(orientation="vertical", size="1", color_scheme="gray", height="16px"),
        qa("save", "Save (download .xlsx) — Ctrl+S", GridState.cmd("ui.save", {}), "qa-save"),
        qa("undo_2", "Undo (Ctrl+Z)", GridState.cmd("edit.undo", {}), "qa-undo"),
        qa("redo_2", "Redo (Ctrl+Y)", GridState.cmd("edit.redo", {}), "qa-redo"),
        rx.spacer(),
        rx.text(GridState.doc_title, size="2", color="#FFFFFF", id="doc-title"),
        rx.spacer(),
        rx.text("Excel-compatible (.xlsx)", size="1", color="#CFE3EC"),
        align="center",
        spacing="2",
        padding_x="10px",
        height="34px",
        background=ACCENT,
        width="100%",
        flex_shrink="0",
    )


def backstage() -> rx.Component:
    def item(icon: str, label: str, on_click, id: str = "") -> rx.Component:
        return rx.button(rx.hstack(rx.icon(icon, size=18), rx.text(label, size="3"), spacing="3", align="center"), variant="ghost", color_scheme="gray", width="100%", justify="start", on_click=on_click, id=id or None, padding_y="20px")

    def sample(name: str, desc: str) -> rx.Component:
        return rx.card(
            rx.vstack(rx.text(name, weight="bold", size="2"), rx.text(desc, size="1", color="#6B7280"), spacing="1"),
            on_click=GridState.new_workbook(name),
            cursor="pointer",
            width="190px",
            id=f"sample-{name.lower()}",
            _hover={"border_color": ACCENT},
        )

    return rx.cond(
        GridState.backstage,
        rx.hstack(
            rx.vstack(
                rx.button(rx.icon("arrow_left", size=18), "Back", variant="ghost", color="#FFFFFF", on_click=GridState.close_backstage, width="100%", justify="start", id="bs-back"),
                rx.divider(),
                item("file_plus", "New", GridState.new_workbook(""), "bs-new"),
                item("folder_open", "Open", GridState.open_dialog("open", {}), "bs-open"),
                item("download", "Save a Copy (.xlsx)", GridState.save_as("xlsx"), "bs-save"),
                item("file_text", "Export CSV", GridState.save_as("csv"), "bs-csv"),
                item("printer", "Print", GridState.print_sheet, "bs-print"),
                item("globe", "Save as Web Page", GridState.save_html, "bs-html"),
                item("info", "Properties", GridState.open_dialog("properties", {})),
                item("x", "Close", GridState.close_workbook),
                spacing="1",
                width="220px",
                padding="12px",
                background=ACCENT,
                height="100%",
                class_name="rg-backstage-nav",
            ),
            rx.vstack(
                rx.heading("New", size="5"),
                rx.hstack(
                    rx.card(rx.vstack(rx.text("Blank workbook", weight="bold", size="2"), rx.text("Start from an empty sheet.", size="1", color="#6B7280"), spacing="1"), on_click=GridState.new_workbook(""), cursor="pointer", width="190px", id="sample-blank"),
                    sample("Sales", "A table with a calculated column, data bars and a chart."),
                    sample("Budget", "Monthly budget with totals, cell styles and a colour scale."),
                    sample("Grades", "Scores with IFS, XLOOKUP, validation and icon sets."),
                    spacing="3",
                    wrap="wrap",
                ),
                rx.heading("Open workbooks", size="4", margin_top="16px"),
                rx.vstack(
                    rx.foreach(
                        GridState.open_docs,
                        lambda d: rx.button(rx.hstack(rx.icon("sheet", size=16), rx.text(d["name"]), spacing="2"), variant="soft", color_scheme="gray", on_click=GridState.switch_window(d["index"])),
                    ),
                    spacing="2",
                    align="start",
                ),
                rx.text(
                    "ReflexGrid reads and writes Excel workbooks (.xlsx): formulas, values, styles, merged cells, tables, conditional formats, validation, notes, links, charts, pictures and print settings.",
                    size="1",
                    color="#6B7280",
                    margin_top="24px",
                    max_width="640px",
                ),
                spacing="3",
                padding="24px",
                align="start",
                flex="1",
            ),
            position="absolute",
            inset="0",
            background="#FFFFFF",
            z_index="50",
            spacing="0",
            id="backstage",
        ),
    )


def sheet_tabs() -> rx.Component:
    def tab(t) -> rx.Component:
        return rx.context_menu.root(
            rx.context_menu.trigger(
                rx.box(
                    rx.hstack(rx.cond(t["protected"], rx.icon("lock", size=11)), rx.text(t["name"], size="2", white_space="nowrap"), spacing="1", align="center"),
                    padding_x="14px",
                    padding_y="3px",
                    cursor="pointer",
                    class_name="rg-sheet-tab",
                    data_sheet=t["name"],
                    background=rx.cond(t["active"], "#FFFFFF", "transparent"),
                    color=rx.cond(t["active"], ACCENT, "#3C4043"),
                    font_weight=rx.cond(t["active"], "600", "400"),
                    border_bottom=rx.cond(t["active"], f"2px solid {ACCENT}", "2px solid transparent"),
                    border_top=rx.cond(t["color"] != "", "3px solid " + t["color"].to(str), "3px solid transparent"),
                    on_click=GridState.activate_sheet(t["index"]),
                    on_double_click=GridState.open_dialog("renameSheet", {"index": t["index"]}),
                    _hover={"background": "#EEF2F5"},
                )
            ),
            rx.context_menu.content(
                rx.context_menu.item("Insert Sheet", on_select=GridState.new_sheet),
                rx.context_menu.item("Delete", on_select=GridState.cmd("home.deleteSheet", {"sheet": t["index"]})),
                rx.context_menu.item("Rename…", on_select=GridState.open_dialog("renameSheet", {"index": t["index"]})),
                rx.context_menu.item("Move or Copy…", on_select=GridState.open_dialog("moveSheet", {"index": t["index"]})),
                rx.context_menu.item("Duplicate", on_select=GridState.cmd("sheet.duplicate", {"sheet": t["index"]})),
                rx.context_menu.item("Tab Color…", on_select=GridState.open_dialog("tabColor", {"index": t["index"]})),
                rx.context_menu.separator(),
                rx.context_menu.item("Hide", on_select=GridState.cmd("sheet.hide", {"sheet": t["index"]})),
                rx.context_menu.item("Unhide…", on_select=GridState.open_dialog("unhideSheet", {})),
                rx.context_menu.separator(),
                rx.context_menu.item("Protect Sheet…", on_select=GridState.open_dialog("protectSheet", {})),
                size="1",
            ),
        )

    return rx.hstack(
        rx.icon_button(rx.icon("chevron_left", size=14), variant="ghost", size="1", color_scheme="gray", on_click=GridState.cmd("sheet.previous", {})),
        rx.icon_button(rx.icon("chevron_right", size=14), variant="ghost", size="1", color_scheme="gray", on_click=GridState.cmd("sheet.next", {})),
        rx.hstack(rx.foreach(GridState.sheet_tabs, tab), spacing="0", overflow_x="auto", max_width="70%", class_name="rg-sheet-tabs"),
        rx.tooltip(rx.icon_button(rx.icon("plus", size=14), variant="ghost", size="1", color_scheme="gray", on_click=GridState.new_sheet, id="add-sheet"), content="New sheet (Shift+F11)"),
        spacing="1",
        align="center",
        padding_x="6px",
        height="30px",
        background="#F3F4F6",
        border_top="1px solid #D5D9DE",
        width="100%",
        flex_shrink="0",
    )


def status_bar() -> rx.Component:
    return rx.hstack(
        rx.text(rx.cond(GridState.ui["calcMode"] == "manual", "Calculate", "Ready"), size="1", color="#4B5563", id="status-mode"),
        rx.cond(GridState.circular_ref != "", rx.text("Circular References: ", GridState.circular_ref, size="1", color="#B42318")),
        rx.cond(GridState.ui["protected"], rx.hstack(rx.icon("lock", size=12), rx.text("Protected", size="1"), spacing="1", color="#4B5563")),
        rx.spacer(),
        rx.text(GridState.status_text, size="1", color="#3C4043", id="status-stats", white_space="pre"),
        rx.divider(orientation="vertical", size="1", height="14px"),
        rx.icon_button(rx.icon("minus", size=12), variant="ghost", size="1", color_scheme="gray", on_click=GridState.zoom_step(-10), id="zoom-out"),
        rx.box(rx.slider(value=[GridState.zoom], min=10, max=400, step=10, on_value_commit=GridState.zoom_to, size="1"), width="110px", flex_shrink="0"),
        rx.icon_button(rx.icon("plus", size=12), variant="ghost", size="1", color_scheme="gray", on_click=GridState.zoom_step(10), id="zoom-in"),
        rx.text(GridState.zoom.to(str) + "%", size="1", width="38px", id="zoom-text"),
        spacing="2",
        align="center",
        padding_x="10px",
        height="24px",
        background="#F3F4F6",
        border_top="1px solid #D5D9DE",
        width="100%",
        flex_shrink="0",
    )
