"""Window chrome: the title bar (title, quick access, who is here, call, share, account), the
File backstage, sheet tabs, status bar, and what shows when a workbook is gone."""

from __future__ import annotations

import reflex as rx

from ..room import BookRoom
from ..state import AuthState, GridState
from .call_views import call_button

ACCENT = "#156082"


def app_icon(size: int = 20) -> rx.Component:
    # Our own mark: a 2x2 grid with an accent cell.
    return rx.html(
        f'<svg width="{size}" height="{size}" viewBox="0 0 20 20"><rect x="1" y="1" width="18" height="18" rx="3" fill="#FFFFFF" fill-opacity="0.18"/>'
        '<rect x="3.5" y="3.5" width="6" height="6" rx="1" fill="#FFFFFF"/><rect x="10.5" y="3.5" width="6" height="6" rx="1" fill="#FFFFFF" fill-opacity="0.55"/>'
        '<rect x="3.5" y="10.5" width="6" height="6" rx="1" fill="#FFFFFF" fill-opacity="0.55"/><rect x="10.5" y="10.5" width="6" height="6" rx="1" fill="#F2C12E"/></svg>'
    )


def peer_avatar(p) -> rx.Component:
    """Someone on the workbook, ringed in their selection colour; a click goes to where they are."""
    face = rx.cond(
        p["avatar"] != "",
        rx.image(src=p["avatar"], width="24px", height="24px", border_radius="9999px", object_fit="cover", alt=p["name"]),
        rx.center(rx.text(p["initials"], size="1", weight="bold", color="#FFFFFF"), width="24px", height="24px", border_radius="9999px", background=p["color"]),
    )
    return rx.tooltip(
        rx.box(
            face,
            border="2px solid",
            border_color=p["color"],
            border_radius="9999px",
            background="#FFFFFF",
            margin_left="-6px",
            cursor="pointer",
            on_click=GridState.follow(p["id"]),
            class_name="rg-person",
            data_name=p["name"],
        ),
        content=p["name"],
    )


def presence() -> rx.Component:
    return rx.hstack(
        rx.hstack(rx.foreach(BookRoom.people, peer_avatar), spacing="0", padding_left="6px", align="center"),
        rx.text(BookRoom.people_count.to(str) + " here", size="1", color="#CFE3EC", white_space="nowrap", id="people-count"),
        spacing="2",
        align="center",
        id="presence",
    )


def save_status() -> rx.Component:
    return rx.text(
        rx.cond(
            BookRoom.save_error != "",
            "Couldn't save",
            rx.cond(BookRoom.saved, "All changes saved", "Saving…"),
        ),
        title=BookRoom.save_error,
        size="1",
        color=rx.cond(BookRoom.save_error != "", "#FECACA", "#CFE3EC"),
        white_space="nowrap",
        id="save-status",
    )


def book_title() -> rx.Component:
    """The workbook's title, renamed in place (for everyone on it): Enter or leaving the field
    renames."""
    return rx.el.form(
        rx.el.input(
            name="title",
            default_value=BookRoom.title,
            key=BookRoom.title,
            on_blur=GridState.rename_book,
            id="doc-title",
            title="Rename workbook",
            spell_check=False,
            auto_complete="off",
            class_name="rg-title-input",
        ),
        on_submit=GridState.rename_from_form,
        reset_on_submit=False,
        style={"display": "flex"},
    )


def identity() -> rx.Component:
    guest = rx.hstack(
        rx.text(GridState.me["name"], size="1", color="#FFFFFF", white_space="nowrap", id="me-name"),
        *([rx.link("Sign in", href=AuthState.login_url, size="1", color="#F2C12E", white_space="nowrap")] if AuthState is not None else []),
        spacing="2",
        align="center",
    )
    if AuthState is None:
        return guest
    signed_in = rx.hstack(
        rx.text(AuthState.user_name, size="1", color="#FFFFFF", white_space="nowrap", max_width="140px", overflow="hidden", text_overflow="ellipsis", id="me-name"),
        rx.link("Sign out", href=AuthState.logout_url, size="1", color="#CFE3EC", white_space="nowrap"),
        spacing="2",
        align="center",
    )
    return rx.cond(AuthState.is_logged_in, signed_in, guest)


def title_bar() -> rx.Component:
    def qa(icon: str, label: str, on_click, id: str) -> rx.Component:
        return rx.tooltip(rx.icon_button(rx.icon(icon, size=15), variant="ghost", size="1", on_click=on_click, color="#FFFFFF", id=id), content=label)

    return rx.hstack(
        rx.tooltip(rx.link(app_icon(), href="/", id="home-link", display="flex"), content="All workbooks"),
        rx.text("ReflexGrid", weight="bold", size="2", color="#FFFFFF"),
        rx.divider(orientation="vertical", size="1", color_scheme="gray", height="16px"),
        qa("save", "Download a copy (.xlsx) — Ctrl+S", GridState.cmd("ui.save", {}), "qa-save"),
        qa("undo_2", "Undo your last change (Ctrl+Z)", GridState.cmd("edit.undo", {}), "qa-undo"),
        qa("redo_2", "Redo (Ctrl+Y)", GridState.cmd("edit.redo", {}), "qa-redo"),
        book_title(),
        save_status(),
        rx.spacer(),
        presence(),
        call_button(),
        rx.tooltip(
            rx.icon_button(
                rx.icon("link", size=15),
                variant="ghost",
                size="1",
                color="#FFFFFF",
                id="share-link",
                on_click=[
                    rx.call_script("navigator.clipboard && navigator.clipboard.writeText(window.location.href)"),
                    rx.toast.success("Link copied: anyone with it can edit this workbook with you.", position="bottom-center"),
                ],
            ),
            content="Copy the link to share this workbook",
        ),
        rx.divider(orientation="vertical", size="1", color_scheme="gray", height="16px"),
        identity(),
        align="center",
        spacing="2",
        padding_x="10px",
        height="38px",
        background=ACCENT,
        width="100%",
        flex_shrink="0",
        id="title-bar",
    )


def backstage() -> rx.Component:
    def item(icon: str, label: str, on_click, id: str = "") -> rx.Component:
        return rx.button(rx.hstack(rx.icon(icon, size=18), rx.text(label, size="3"), spacing="3", align="center"), variant="ghost", color_scheme="gray", color="#FFFFFF", width="100%", justify="start", on_click=on_click, id=id or None, padding_y="20px")

    def sample(name: str, desc: str) -> rx.Component:
        return rx.card(
            rx.vstack(rx.text(name, weight="bold", size="2", color="#111827"), rx.text(desc, size="1", color="#6B7280"), spacing="1"),
            on_click=GridState.new_workbook(name),
            cursor="pointer",
            width="190px",
            id=f"sample-{name.lower()}",
            _hover={"border_color": ACCENT},
        )

    def recent(b) -> rx.Component:
        return rx.link(
            rx.hstack(rx.icon("sheet", size=16), rx.text(b["title"], size="2"), rx.text(b["when"], size="1", color="#6B7280"), spacing="2", align="center"),
            href="/book/" + b["id"].to(str),
            class_name="rg-recent",
            color="#1F2937",
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
                item("layout_grid", "All Workbooks", GridState.close_workbook, "bs-close"),
                spacing="1",
                width="220px",
                padding="12px",
                background=ACCENT,
                height="100%",
                class_name="rg-backstage-nav",
            ),
            rx.vstack(
                rx.heading("New", size="5", color="#111827"),
                rx.text("A new workbook is shared: send its link and edit it together.", size="1", color="#6B7280"),
                rx.hstack(
                    rx.card(rx.vstack(rx.text("Blank workbook", weight="bold", size="2", color="#111827"), rx.text("Start from an empty sheet.", size="1", color="#6B7280"), spacing="1"), on_click=GridState.new_workbook(""), cursor="pointer", width="190px", id="sample-blank"),
                    sample("Sales", "A table with a calculated column, data bars and a chart."),
                    sample("Budget", "Monthly budget with totals, cell styles and a colour scale."),
                    sample("Grades", "Scores with IFS, XLOOKUP, validation and icon sets."),
                    spacing="3",
                    wrap="wrap",
                ),
                rx.heading("Other workbooks", size="4", margin_top="16px", color="#111827"),
                rx.cond(
                    GridState.recent.length() > 0,
                    rx.vstack(rx.foreach(GridState.recent, recent), spacing="2", align="start"),
                    rx.text("No other workbooks yet.", size="2", color="#6B7280"),
                ),
                rx.link("All workbooks →", href="/", size="2", id="bs-all"),
                rx.text(
                    "ReflexGrid reads and writes Excel workbooks (.xlsx): formulas, values, styles, merged cells, tables, conditional formats, validation, notes, links, charts, pictures and print settings. Shared workbooks are saved on the server as .xlsx as you edit.",
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


def book_gone() -> rx.Component:
    """The workbook was deleted (or its address isn't one)."""
    return rx.cond(
        GridState.book_gone | BookRoom.deleted,
        rx.center(
            rx.vstack(
                rx.icon("file_x", size=40, color="#9CA3AF"),
                rx.heading("This workbook isn't here any more", size="5", color="#111827"),
                rx.text("Someone deleted it from the workbook list.", size="2", color="#6B7280"),
                rx.link(rx.button("All workbooks", size="2"), href="/"),
                align="center",
                spacing="3",
            ),
            position="absolute",
            inset="0",
            background="rgba(255,255,255,0.94)",
            z_index="70",
            id="book-gone",
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
        rx.text(GridState.zoom.to(str) + "%", size="1", width="38px", color="#3C4043", id="zoom-text"),
        spacing="2",
        align="center",
        padding_x="10px",
        height="24px",
        background="#F3F4F6",
        border_top="1px solid #D5D9DE",
        width="100%",
        flex_shrink="0",
    )
