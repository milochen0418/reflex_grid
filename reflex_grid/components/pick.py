"""The page of the intent ``grid.pick`` (:mod:`..pick`): what of a workbook to show in another app."""

from __future__ import annotations

import reflex as rx

from ..pick import PickState

_JS = rx.asset("grid/cell_picker.jsx")
_ROW = "w-full flex items-center gap-3 px-3 py-2 rounded-lg text-left hover:bg-cyan-50 cursor-pointer"
_TITLE = "text-sm font-medium text-gray-900 truncate"
_NOTE = "text-xs text-gray-500 truncate"
_SECTION = "text-xs font-semibold text-gray-500 uppercase tracking-wider mt-3 mb-1 px-3"


class CellPicker(rx.Component):
    library = _JS.importable_path
    tag = "CellPicker"

    rows: rx.Var[list]
    top: rx.Var[int]
    left: rx.Var[int]
    selection: rx.Var[list]
    on_select: rx.EventHandler[lambda s: [s]]


def _row(icon: str, color: str, title, note, on_click, **attrs) -> rx.Component:
    return rx.el.button(
        rx.icon(icon, size=18, color=color, class_name="shrink-0"),
        rx.el.div(rx.el.div(title, class_name=_TITLE), rx.el.div(note, class_name=_NOTE), class_name="min-w-0 flex-1"),
        type="button",
        on_click=on_click,
        class_name=_ROW,
        **attrs,
    )


def _books() -> rx.Component:
    return rx.el.div(
        rx.el.h2("Choose a workbook", class_name="text-lg font-semibold text-gray-900 mb-3"),
        rx.input(default_value="", placeholder="Search workbooks", on_change=PickState.set_query.debounce(150), width="100%", id="pick-search"),
        rx.el.div(
            rx.foreach(PickState.shown_books, lambda b: _row("file_spreadsheet", "#107C41", b["title"], b["note"], PickState.choose(b["id"]), data_book=b["id"])),
            rx.cond(PickState.shown_books.length() == 0, rx.el.p("No workbooks.", class_name="text-sm text-gray-500 p-3")),
            class_name="mt-2 overflow-y-auto",
            style={"maxHeight": "420px"},
        ),
    )


def _named(title: str, items, icon: str, color: str, kind: str, note) -> rx.Component:
    return rx.cond(
        items.length() > 0,
        rx.el.div(
            rx.el.div(title, class_name=_SECTION),
            rx.foreach(items, lambda x: _row(icon, color, x["name"], note(x), PickState.insert(kind, x["name"]), data_pick=x["name"])),
        ),
    )


def _cells() -> rx.Component:
    return rx.el.div(
        rx.el.div("Cells", class_name=_SECTION),
        rx.el.div(
            rx.foreach(
                PickState.sheets,
                lambda name: rx.el.button(
                    name,
                    type="button",
                    on_click=PickState.show_sheet(name),
                    data_sheet=name,
                    class_name=rx.cond(
                        PickState.sheet == name,
                        "px-3 py-1 text-xs rounded-md bg-cyan-700 text-white cursor-pointer",
                        "px-3 py-1 text-xs rounded-md bg-gray-100 text-gray-700 hover:bg-gray-200 cursor-pointer",
                    ),
                ),
            ),
            class_name="flex flex-wrap gap-1 px-3 mb-2",
        ),
        rx.el.div(
            CellPicker.create(rows=PickState.rows, top=PickState.top, left=PickState.left, selection=PickState.selected, on_select=PickState.select),
            class_name="mx-3 overflow-auto border border-gray-200 rounded-md",
            style={"maxHeight": "280px"},
        ),
        rx.cond(PickState.more, rx.el.p("The sheet goes on: this shows its first rows and columns.", class_name="text-xs text-gray-500 px-3 mt-1")),
        rx.el.div(
            rx.el.code(PickState.ref, class_name="text-xs text-gray-800 min-w-0 truncate", id="pick-ref"),
            rx.cond(PickState.can_grow, rx.checkbox("New rows join in", checked=PickState.grow, on_change=PickState.set_grow, size="1", id="pick-grow")),
            rx.el.div(class_name="flex-1"),
            rx.button("Show these cells", size="2", on_click=PickState.insert("range", PickState.ref), id="pick-cells"),
            class_name="flex items-center gap-3 px-3 mt-2",
        ),
    )


def _book() -> rx.Component:
    return rx.el.div(
        rx.el.div(
            rx.el.button(rx.icon("arrow_left", size=18), type="button", on_click=PickState.back, title="Other workbooks", class_name="p-1 rounded hover:bg-gray-100 text-gray-600 cursor-pointer", id="pick-back"),
            rx.el.h2(PickState.title, class_name="text-lg font-semibold text-gray-900 truncate"),
            class_name="flex items-center gap-2 mb-1",
        ),
        rx.el.div(
            _named("UserForms", PickState.forms, "app_window", "#7C3AED", "form", lambda x: x["caption"]),
            _named("Tables", PickState.tables, "table", "#0E7490", "range", lambda x: x["sheet"].to(str) + " · " + x["area"].to(str) + " · grows with the table"),
            _named("Names", PickState.names, "tag", "#B45309", "range", lambda x: x["sheet"].to(str) + " · " + x["area"].to(str)),
            class_name="overflow-y-auto",
            style={"maxHeight": "210px"},
        ),
        _cells(),
    )


def pick_page() -> rx.Component:
    return rx.el.div(
        rx.cond(PickState.error != "", rx.callout(PickState.error, icon="triangle_alert", color_scheme="red", size="1", margin_bottom="8px")),
        rx.cond(PickState.book == "", _books(), _book()),
        style={"padding": "20px 48px 20px 20px", "background": "#FFFFFF", "colorScheme": "light"},
    )
