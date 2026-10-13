"""The workbook list page (``/``): every shared workbook, who is editing which right now, search
and filters, and new / upload / download / delete (like CoDoc's Document Manager)."""

from __future__ import annotations

import reflex as rx

from ..workbooks import SEARCH_INPUT_ID, AuthState, ManagerState
from .chrome import ACCENT, app_icon

_TH = "px-5 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider"
_SELECT = (
    "h-10 pl-3 pr-8 text-sm text-gray-700 bg-white border border-gray-200 rounded-lg "
    "focus:outline-none focus:ring-2 focus:ring-cyan-200 focus:border-cyan-500 cursor-pointer"
)
_CHIP_ON = "flex items-center gap-1.5 h-8 px-3 text-sm font-medium rounded-full text-white border cursor-pointer"
_CHIP_OFF = "flex items-center gap-1.5 h-8 px-3 text-sm rounded-full bg-white text-gray-700 border border-gray-200 hover:border-cyan-400 hover:text-cyan-800 cursor-pointer transition-colors"

_REFRESH_JS = """
if (window.__rgListRefresh) clearInterval(window.__rgListRefresh);
window.__rgListRefresh = setInterval(function () {
    var el = document.getElementById('book-list-refresh');
    if (!el) { clearInterval(window.__rgListRefresh); window.__rgListRefresh = null; return; }
    el.click();
}, 8000);
"""


def live_badge(b) -> rx.Component:
    """Who is on the workbook right now."""
    return rx.cond(
        b["live"] > 0,
        rx.hstack(
            rx.box(width="8px", height="8px", border_radius="9999px", background="#22C55E", class_name="animate-pulse"),
            rx.text(b["live"].to(str) + " editing", size="1", color="#15803D", weight="medium"),
            rx.hstack(
                rx.foreach(
                    b["live_people"],
                    lambda p: rx.tooltip(rx.box(width="10px", height="10px", border_radius="9999px", background=p["color"], border="1px solid #FFFFFF"), content=p["name"]),
                ),
                spacing="0",
            ),
            spacing="2",
            align="center",
            class_name="rg-live",
        ),
        rx.text("—", size="1", color="#9CA3AF"),
    )


def creator_cell(b) -> rx.Component:
    """Creator info; clicking it filters the list to that creator."""
    return rx.el.button(
        rx.cond(
            b["created_by_email"] != "",
            rx.el.div(
                rx.el.span(rx.cond(b["created_by_name"] != "", b["created_by_name"], b["created_by_email"]), class_name="text-gray-700 text-sm font-medium"),
                rx.el.span(b["created_by_email"], class_name="text-gray-400 text-xs block"),
                class_name="text-left",
            ),
            rx.el.div(rx.icon("user", size=13, color="#9CA3AF"), rx.el.span("Guest", class_name="text-gray-400 text-sm italic"), class_name="flex items-center gap-1"),
        ),
        on_click=ManagerState.filter_by_creator_of(b["created_by_email"]),
        title="Show all workbooks by this creator",
        class_name="-mx-2 px-2 py-1 rounded-md hover:bg-cyan-50 transition-colors cursor-pointer",
    )


def delete_button(b) -> rx.Component:
    return rx.alert_dialog.root(
        rx.alert_dialog.trigger(
            rx.icon_button(rx.icon("trash_2", size=15), variant="ghost", color_scheme="red", size="1", title="Delete workbook", class_name="rg-delete"),
        ),
        rx.alert_dialog.content(
            rx.alert_dialog.title("Delete “" + b["title"].to(str) + "”?"),
            rx.alert_dialog.description(
                rx.cond(
                    b["live"] > 0,
                    "People are editing it right now; they will see that it was deleted. This can't be undone.",
                    "The workbook is deleted for everyone. This can't be undone.",
                )
            ),
            rx.hstack(
                rx.alert_dialog.cancel(rx.button("Cancel", variant="soft", color_scheme="gray")),
                rx.alert_dialog.action(rx.button("Delete", color_scheme="red", on_click=ManagerState.delete_book(b["id"]), class_name="rg-delete-ok")),
                spacing="2",
                justify="end",
                margin_top="16px",
            ),
        ),
    )


def book_row(b) -> rx.Component:
    return rx.el.tr(
        rx.el.td(
            rx.el.div(
                rx.icon("sheet", size=16, color=ACCENT),
                rx.link(b["title"], href="/book/" + b["id"].to(str), class_name="rg-book-link font-medium", color=ACCENT),
                class_name="flex items-center gap-2",
            ),
            rx.el.p(
                b["sheet_count"].to(str) + rx.cond(b["sheet_count"] == 1, " sheet", " sheets"),
                rx.cond(b["sheets"] != "", ": " + b["sheets"].to(str), ""),
                rx.cond(b["size"] != "", " · " + b["size"].to(str), ""),
                class_name="mt-0.5 text-[11px] text-gray-400 max-w-md truncate",
            ),
            class_name="px-5 py-3 align-top",
        ),
        rx.el.td(live_badge(b), class_name="px-5 py-3 align-top whitespace-nowrap"),
        rx.el.td(b["id"], class_name="hidden md:table-cell px-5 py-3 text-gray-500 text-sm font-mono align-top"),
        rx.el.td(creator_cell(b), class_name="px-5 py-2 align-top"),
        rx.el.td(
            rx.el.div(b["when"], class_name="text-gray-600 text-sm whitespace-nowrap"),
            rx.cond(b["last_editor"] != "", rx.el.div("by " + b["last_editor"].to(str), class_name="text-gray-400 text-xs"), rx.fragment()),
            class_name="px-5 py-3 align-top",
        ),
        rx.el.td(
            rx.hstack(
                rx.icon_button(rx.icon("download", size=15), variant="ghost", color_scheme="gray", size="1", title="Download .xlsx", on_click=ManagerState.download(b["id"])),
                delete_button(b),
                spacing="1",
                justify="end",
            ),
            class_name="px-5 py-3 align-top",
        ),
        class_name="rg-book-row border-b border-gray-100 hover:bg-gray-50/60 transition-colors",
        data_book=b["id"],
    )


def creator_chip(opt) -> rx.Component:
    active = ManagerState.creator_filter == opt["key"]
    return rx.el.button(
        rx.cond(opt["key"] == "guest", rx.icon("user", size=13), rx.cond(opt["key"] == "mine", rx.icon("star", size=13), rx.fragment())),
        rx.el.span(opt["label"]),
        rx.el.span(opt["count"], class_name=rx.cond(active, "text-[11px] px-1.5 rounded-full bg-white/25", "text-[11px] px-1.5 rounded-full bg-gray-100 text-gray-500")),
        on_click=ManagerState.set_creator_filter(opt["key"]),
        title=opt["sublabel"],
        class_name=rx.cond(active, _CHIP_ON, _CHIP_OFF),
        style=rx.cond(active, {"background": ACCENT, "borderColor": ACCENT}, {}),
    )


def filter_bar() -> rx.Component:
    return rx.el.div(
        rx.el.div(
            rx.el.div(
                rx.icon("search", size=16, color="#9CA3AF", class_name="absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none"),
                # Uncontrolled on purpose: binding `value` to backend state makes every keystroke
                # wait for a round-trip (and debounced events drop characters).
                rx.el.input(
                    id=SEARCH_INPUT_ID,
                    default_value=ManagerState.search,
                    on_change=ManagerState.set_search.debounce(250),
                    placeholder="Search title, sheet names, ID or creator…",
                    class_name="w-full h-10 pl-9 pr-9 text-sm bg-white border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-cyan-200 focus:border-cyan-500",
                ),
                rx.cond(
                    ManagerState.search != "",
                    rx.el.button(rx.icon("x", size=16), on_click=ManagerState.clear_search, title="Clear search", class_name="absolute right-2 top-1/2 -translate-y-1/2 p-1 text-gray-400 hover:text-gray-600 cursor-pointer"),
                ),
                class_name="relative flex-1 min-w-[220px]",
            ),
            rx.el.select(
                rx.el.option("Updated: any time", value="any"),
                rx.el.option("Updated today", value="today"),
                rx.el.option("Last 7 days", value="7d"),
                rx.el.option("Last 30 days", value="30d"),
                rx.el.option("Older than 30 days", value="older"),
                value=ManagerState.time_filter,
                on_change=ManagerState.set_time_filter,
                class_name=_SELECT,
                aria_label="Filter by last update",
            ),
            rx.el.select(
                rx.el.option("Newest first", value="updated_desc"),
                rx.el.option("Oldest first", value="updated_asc"),
                rx.el.option("Title A → Z", value="title_asc"),
                rx.el.option("Title Z → A", value="title_desc"),
                value=ManagerState.sort_by,
                on_change=ManagerState.set_sort_by,
                class_name=_SELECT,
                aria_label="Sort workbooks",
            ),
            class_name="flex flex-wrap items-center gap-2",
        ),
        rx.el.div(
            rx.el.span("Created by", class_name="text-xs font-semibold uppercase tracking-wider text-gray-400 mr-1"),
            rx.foreach(ManagerState.creator_options, creator_chip),
            rx.el.span(class_name="w-px h-5 bg-gray-200 mx-1"),
            rx.el.button(
                rx.box(width="8px", height="8px", border_radius="9999px", background="#22C55E"),
                rx.el.span("Being edited now"),
                rx.el.span(ManagerState.live_count, class_name="text-[11px] px-1.5 rounded-full bg-green-100 text-green-700"),
                on_click=ManagerState.toggle_only_live,
                title="Workbooks someone has open right now",
                class_name=rx.cond(ManagerState.only_live, "flex items-center gap-1.5 h-8 px-3 text-sm font-medium rounded-full bg-green-600 text-white border border-green-600 cursor-pointer", _CHIP_OFF),
                id="filter-live",
            ),
            class_name="flex flex-wrap items-center gap-2 mt-3",
        ),
        class_name="bg-white border border-gray-200 rounded-lg shadow-sm p-3 sm:p-4 mb-4",
    )


def delete_filtered_button() -> rx.Component:
    """Bulk delete of the current filtered view, behind a confirmation."""
    n = ManagerState.books.length().to(str)
    return rx.alert_dialog.root(
        rx.alert_dialog.trigger(rx.button(rx.icon("trash", size=14), "Delete " + n + " shown", variant="outline", color_scheme="red", size="2")),
        rx.alert_dialog.content(
            rx.alert_dialog.title("Delete the workbooks shown?"),
            rx.alert_dialog.description("This permanently deletes the " + n + " workbook(s) matching the current filters, for everyone. Other workbooks are kept."),
            rx.hstack(
                rx.alert_dialog.cancel(rx.button("Cancel", variant="soft", color_scheme="gray")),
                rx.alert_dialog.action(rx.button("Delete", color_scheme="red", on_click=ManagerState.delete_filtered)),
                spacing="2",
                justify="end",
                margin_top="16px",
            ),
        ),
    )


def summary_bar() -> rx.Component:
    return rx.el.div(
        rx.el.div(
            rx.el.p(
                rx.cond(
                    ManagerState.total_count == 0,
                    "No workbooks yet",
                    rx.cond(
                        ManagerState.has_active_filters,
                        "Showing " + ManagerState.books.length().to(str) + " of " + ManagerState.total_count.to(str) + " workbook(s)",
                        ManagerState.total_count.to(str) + " workbook(s)",
                    ),
                ),
                class_name="text-sm text-gray-500",
                id="book-count",
            ),
            rx.cond(
                ManagerState.has_active_filters,
                rx.el.button(rx.icon("rotate_ccw", size=13), "Reset filters", on_click=ManagerState.reset_filters, class_name="flex items-center gap-1 text-sm text-cyan-700 hover:text-cyan-900 cursor-pointer"),
            ),
            class_name="flex items-center gap-3",
        ),
        rx.cond(ManagerState.has_active_filters & (ManagerState.books.length() > 0), delete_filtered_button()),
        class_name="flex items-center justify-between gap-3 flex-wrap mb-3",
    )


def book_table() -> rx.Component:
    return rx.el.div(
        rx.el.table(
            rx.el.thead(
                rx.el.tr(
                    rx.el.th("Workbook", class_name=_TH),
                    rx.el.th("Now", class_name=_TH),
                    rx.el.th("ID", class_name="hidden md:table-cell " + _TH),
                    rx.el.th("Created by", class_name=_TH),
                    rx.el.th("Updated", class_name=_TH),
                    rx.el.th("", class_name="px-5 py-3 w-24"),
                    class_name="border-b border-gray-200",
                ),
                class_name="bg-gray-50",
            ),
            rx.el.tbody(rx.foreach(ManagerState.books, book_row)),
            class_name="min-w-full",
        ),
        class_name="bg-white rounded-lg border border-gray-200 overflow-x-auto shadow-sm",
        id="book-table",
    )


def no_match_state() -> rx.Component:
    return rx.el.div(
        rx.icon("search_x", size=44, color="#D1D5DB"),
        rx.el.p("No workbooks match these filters", class_name="text-base font-medium text-gray-500 my-3"),
        rx.button(rx.icon("rotate_ccw", size=14), "Reset filters", variant="outline", on_click=ManagerState.reset_filters),
        class_name="flex flex-col items-center justify-center py-16 bg-white rounded-lg border border-dashed border-gray-200",
    )


def empty_state() -> rx.Component:
    return rx.el.div(
        rx.icon("sheet", size=56, color="#D1D5DB"),
        rx.el.p("No workbooks yet", class_name="text-lg font-medium text-gray-500 mt-4 mb-1"),
        rx.el.p("Create one, or upload an .xlsx file, then share its link to edit it together.", class_name="text-sm text-gray-400 mb-6"),
        rx.button(rx.icon("plus", size=16), "New workbook", size="3", on_click=ManagerState.new_book(""), id="empty-new"),
        class_name="flex flex-col items-center justify-center py-20",
    )


def new_menu() -> rx.Component:
    return rx.hstack(
        rx.upload(
            rx.button(rx.icon("upload", size=15), "Upload .xlsx", variant="outline", size="2", id="upload-book"),
            id="manager_upload",
            accept={
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx", ".xlsm"],
                "text/csv": [".csv"],
                "text/tab-separated-values": [".tsv", ".txt"],
            },
            max_files=1,
            multiple=False,
            no_drag=True,
            border="0",
            padding="0",
            on_drop=ManagerState.upload(rx.upload_files(upload_id="manager_upload")),
        ),
        rx.menu.root(
            rx.menu.trigger(rx.button(rx.icon("plus", size=15), "New workbook", rx.icon("chevron_down", size=13), size="2", id="new-book")),
            rx.menu.content(
                rx.menu.item("Blank workbook", on_select=ManagerState.new_book(""), id="new-blank"),
                rx.menu.separator(),
                rx.menu.item("Sales sample", on_select=ManagerState.new_book("Sales")),
                rx.menu.item("Budget sample", on_select=ManagerState.new_book("Budget")),
                rx.menu.item("Grades sample", on_select=ManagerState.new_book("Grades")),
                size="2",
            ),
        ),
        spacing="2",
        align="center",
    )


def account() -> rx.Component:
    if AuthState is None:
        return rx.fragment()
    return rx.cond(
        AuthState.is_logged_in,
        rx.hstack(
            rx.cond(AuthState.user_avatar != "", rx.image(src=AuthState.user_avatar, width="28px", height="28px", border_radius="9999px")),
            rx.text(AuthState.user_name, size="2", weight="medium", color="#374151"),
            rx.link("Sign out", href=AuthState.logout_url, size="1", color="#6B7280"),
            spacing="2",
            align="center",
        ),
        rx.hstack(
            rx.icon("user", size=18, color="#9CA3AF"),
            rx.text("Guest", size="2", color="#6B7280"),
            rx.link(rx.button("Sign in", size="1", variant="soft"), href=AuthState.login_url),
            spacing="2",
            align="center",
        ),
    )


def manager_page() -> rx.Component:
    return rx.el.main(
        rx.el.div(
            rx.el.header(
                rx.el.div(
                    rx.el.div(
                        rx.box(app_icon(24), background=ACCENT, border_radius="6px", padding="3px", display="flex"),
                        rx.el.div(
                            rx.el.h1("ReflexGrid", class_name="text-xl font-bold text-gray-900 leading-tight"),
                            rx.el.p("Shared Excel workbooks — edit together, live", class_name="text-xs text-gray-500"),
                        ),
                        class_name="flex items-center gap-3",
                    ),
                    rx.el.div(account(), new_menu(), class_name="flex items-center gap-4"),
                    class_name="flex items-center justify-between w-full",
                ),
                class_name="h-16 px-4 sm:px-6 lg:px-8 flex items-center border-b border-gray-200 bg-white shadow-sm",
            ),
            rx.el.div(
                rx.el.div(
                    rx.cond(ManagerState.total_count > 0, filter_bar()),
                    summary_bar(),
                    rx.cond(
                        ManagerState.books.length() > 0,
                        book_table(),
                        rx.cond(ManagerState.total_count > 0, no_match_state(), empty_state()),
                    ),
                    class_name="max-w-6xl mx-auto w-full",
                ),
                class_name="flex-1 px-4 sm:px-6 lg:px-8 py-8 bg-gray-50 overflow-y-auto",
            ),
            class_name="flex flex-col h-screen w-full bg-white",
        ),
        rx.el.button(id="book-list-refresh", on_click=ManagerState.refresh, on_mount=rx.call_script(_REFRESH_JS), tab_index=-1, aria_hidden="true", class_name="sr-only"),
        id="manager",
    )
