"""Macros (VBA): the Developer tab, the VBA editor (modules, code, Immediate window, problems),
the Macros dialog, the dialogs a running macro shows (MsgBox, InputBox), its progress, and the
bar that says a workbook's macros are off.

The code editor is a local React component (``assets/grid/vba_editor.jsx``), a plain asset like
the grid (see :mod:`.sheet_view`)."""

from __future__ import annotations

from typing import Any, Dict, List

import reflex as rx

from ..state import GridState
from .ribbon import ACCENT, group, rbtn

_JS = rx.asset("grid/vba_editor.jsx")
MACROS = GridState.ui["macros"].to(dict)


class VbaEditor(rx.Component):
    library = f"$/public{_JS}"
    tag = "VbaEditor"

    value: rx.Var[str]
    doc_key: rx.Var[int]
    error_line: rx.Var[int]
    read_only: rx.Var[bool]
    font_size: rx.Var[int]
    on_change: rx.EventHandler[lambda code: [code]]
    on_save: rx.EventHandler[lambda code: [code]]
    on_run: rx.EventHandler[lambda code, proc: [code, proc]]


vba_code_editor = VbaEditor.create


# ---------------------------------------------------------------- the ribbon tab


def developer_tab() -> rx.Component:
    return rx.hstack(
        group(
            "Code",
            rbtn("square_code", "Visual Basic", GridState.open_vba("", 0), "Visual Basic editor (Alt+F11)", big=True, id="rb-vba"),
            rbtn("list_video", "Macros", GridState.open_dialog("macros", {}), "Macros (Alt+F8)", big=True, id="rb-macros"),
            rx.cond(
                MACROS["enabled"].to(bool),
                rbtn("shield_check", "Macros On", GridState.enable_macros(False), "Event macros run when people edit — click to turn them off", big=True, id="rb-macros-off"),
                rbtn("shield_off", "Macros Off", GridState.enable_macros(True), "Turn on the workbook's event macros (Worksheet_Change…) and run Workbook_Open", big=True, id="rb-macros-on"),
            ),
        ),
        group(
            "Modules",
            rbtn("file_code", "Module", GridState.open_vba_with("standard"), "Insert a module and open it", big=True, id="rb-add-module"),
            rbtn("braces", "Class Module", GridState.open_vba_with("class"), "Insert a class module and open it", big=True, id="rb-add-class"),
        ),
        spacing="0",
        align="stretch",
    )


# ---------------------------------------------------------------- the security bar


def macro_bar() -> rx.Component:
    """Under the ribbon, as in Excel: a workbook's macros stay off until someone turns them on."""
    return rx.cond(
        MACROS["project"].to(bool) & ~MACROS["enabled"].to(bool),
        rx.hstack(
            rx.icon("shield_alert", size=16, color="#8A5A00"),
            rx.text("SECURITY WARNING", size="1", weight="bold", color="#5C3D00"),
            rx.text("Event macros are turned off. Macros run only when someone runs them (Developer › Macros).", size="1", color="#5C3D00"),
            rx.button("Enable Macros", size="1", variant="outline", color_scheme="gray", on_click=GridState.enable_macros(True), id="macro-enable", color="#1F2937", background="#FFFFFF"),
            spacing="2",
            align="center",
            width="100%",
            padding_x="12px",
            padding_y="4px",
            background="#FFF4CE",
            border_bottom="1px solid #E8D9A8",
            id="macro-bar",
        ),
    )


# ---------------------------------------------------------------- the editor


def _module_item(m) -> rx.Component:
    return rx.box(
        rx.hstack(
            rx.icon(rx.match(m["kind"], ("document", "file_spreadsheet"), ("class", "braces"), ("form", "app_window"), "file_code"), size=14, color="#4B5563"),
            rx.text(m["label"], size="1", color="#111827", white_space="nowrap", overflow="hidden", text_overflow="ellipsis"),
            spacing="1",
            align="center",
        ),
        padding="3px 8px",
        cursor="pointer",
        border_radius="3px",
        background=rx.cond(m["name"] == GridState.vba_module, "#D7E3EA", "transparent"),
        on_click=GridState.vba_select(m["name"]),
        class_name="vba-module",
        id="vba-mod-" + m["name"].to(str),
    )


def _module_group(title: str, kind: str) -> rx.Component:
    items = GridState.vba_modules
    return rx.vstack(
        rx.text(title, size="1", weight="bold", color="#6B7280", padding_x="6px", padding_top="6px"),
        rx.foreach(items, lambda m: rx.cond(m["kind"] == kind, _module_item(m))),
        spacing="0",
        width="100%",
    )


def _toolbar() -> rx.Component:
    def tb(icon: str, label: str, on_click, tip: str, id: str, disabled=False) -> rx.Component:
        return rx.tooltip(
            rx.button(rx.icon(icon, size=15), rx.text(label, size="1"), variant="ghost", color_scheme="gray", size="1", on_click=on_click, id=id, disabled=disabled, color="#1F2937"),
            content=tip,
            delay_duration=400,
        )

    return rx.hstack(
        tb("play", "Run", GridState.vba_run(GridState.vba_code, ""), "Run a macro (F5 runs the Sub the cursor is in)", "vba-run", GridState.vba_locked),
        tb("save", "Apply", GridState.vba_save(GridState.vba_code), "Apply the code to the workbook (Ctrl+S)", "vba-save", GridState.vba_locked),
        tb("list_checks", "Compile", GridState.vba_compile, "Check every module for compile errors", "vba-compile", GridState.vba_locked),
        rx.separator(orientation="vertical", size="1"),
        tb("file_code", "Module", GridState.vba_add("standard"), "Insert a module", "vba-add-module", GridState.vba_locked),
        tb("braces", "Class", GridState.vba_add("class"), "Insert a class module", "vba-add-class", GridState.vba_locked),
        rx.upload(
            rx.button(rx.icon("file_up", size=15), rx.text("Import…", size="1"), variant="ghost", color_scheme="gray", size="1", id="vba-import", color="#1F2937"),
            id="vba_import",
            accept={"text/plain": [".bas", ".cls", ".txt"]},
            multiple=True,
            on_drop=GridState.vba_import(rx.upload_files(upload_id="vba_import")),
            border="0",
            padding="0",
        ),
        tb("file_down", "Export", GridState.vba_export, "Export the module as a .bas or .cls file", "vba-export", (GridState.vba_module == "") | GridState.vba_locked),
        rx.popover.root(
            rx.popover.trigger(rx.button(rx.icon("pencil", size=15), rx.text("Rename", size="1"), variant="ghost", color_scheme="gray", size="1", id="vba-rename", color="#1F2937", disabled=(GridState.vba_kind == "document") | GridState.vba_locked)),
            rx.popover.content(
                rx.form(
                    rx.hstack(
                        rx.input(name="name", default_value=GridState.vba_module, size="1", id="vba-rename-input"),
                        rx.popover.close(rx.button("Rename", type="submit", size="1")),
                        spacing="2",
                    ),
                    on_submit=GridState.vba_rename,
                    reset_on_submit=False,
                ),
                size="1",
            ),
        ),
        tb("trash_2", "Remove", GridState.vba_remove, "Remove the module", "vba-remove", (GridState.vba_kind == "document") | (GridState.vba_module == "") | GridState.vba_locked),
        rx.spacer(),
        rx.cond(GridState.vba_dirty, rx.text("● not applied", size="1", color="#B45309", id="vba-dirty")),
        spacing="1",
        align="center",
        width="100%",
        padding_x="4px",
        border_bottom="1px solid #E3E6E9",
        padding_bottom="4px",
    )


def _problems() -> rx.Component:
    return rx.vstack(
        rx.cond(
            GridState.vba_diag.length() == 0,
            rx.text("No compile errors.", size="1", color="#6B7280"),
        ),
        rx.foreach(
            GridState.vba_diag,
            lambda d: rx.hstack(
                rx.icon("circle_x", size=13, color="#C5221F"),
                rx.text(d["module"].to(str) + ", line " + d["line"].to(str) + ": " + d["message"].to(str), size="1", color="#1F2937"),
                spacing="1",
                align="center",
                cursor="pointer",
                on_click=GridState.vba_goto(d["module"].to(str), d["line"].to(int)),
            ),
        ),
        spacing="1",
        width="100%",
        overflow_y="auto",
        height="100%",
        padding="6px",
    )


def _immediate() -> rx.Component:
    return rx.vstack(
        rx.box(
            rx.foreach(GridState.vba_output, lambda line: rx.text(line, size="1", font_family="Menlo, Consolas, monospace", white_space="pre", color="#111827")),
            flex="1",
            min_height="0",
            width="100%",
            overflow_y="auto",
            padding="4px 6px",
            background="#FFFFFF",
            id="vba-output",
        ),
        rx.form(
            rx.hstack(
                rx.text("?", size="1", color="#6B7280", font_family="Menlo, Consolas, monospace"),
                rx.input(name="code", value=GridState.vba_immediate, on_change=GridState.set_vba_immediate, placeholder="? expression, or a statement — Enter runs it", size="1", width="100%", font_family="Menlo, Consolas, monospace", id="vba-immediate"),
                rx.button("Run", type="submit", size="1", variant="soft"),
                rx.button("Clear", size="1", variant="ghost", color_scheme="gray", on_click=GridState.vba_clear_output, type="button"),
                spacing="2",
                align="center",
                width="100%",
            ),
            on_submit=GridState.vba_immediate_run,
            reset_on_submit=False,
            width="100%",
        ),
        spacing="1",
        width="100%",
        height="100%",
        padding="4px",
    )


def vba_editor_dialog() -> rx.Component:
    return rx.dialog.root(
        rx.dialog.content(
            rx.vstack(
                rx.hstack(
                    rx.icon("square_code", size=18, color=ACCENT),
                    rx.dialog.title("VBA Editor", size="3", margin="0"),
                    rx.text(GridState.vba_module, size="2", color="#4B5563"),
                    rx.spacer(),
                    rx.cond(MACROS["enabled"].to(bool), rx.badge("Event macros on", color_scheme="green", size="1"), rx.badge("Event macros off", color_scheme="gray", size="1")),
                    rx.dialog.close(rx.icon_button(rx.icon("x", size=16), variant="ghost", color_scheme="gray", size="1", on_click=GridState.close_vba, id="vba-close")),
                    spacing="2",
                    align="center",
                    width="100%",
                ),
                _toolbar(),
                rx.hstack(
                    rx.vstack(
                        _module_group("Excel Objects", "document"),
                        _module_group("Modules", "standard"),
                        _module_group("Class Modules", "class"),
                        _module_group("Forms (kept, not run)", "form"),
                        spacing="1",
                        width="200px",
                        min_width="200px",
                        height="100%",
                        overflow_y="auto",
                        border_right="1px solid #E3E6E9",
                        background="#FAFBFC",
                        id="vba-modules",
                    ),
                    rx.box(
                        vba_code_editor(
                            value=GridState.vba_code,
                            doc_key=GridState.vba_key,
                            error_line=GridState.vba_error_line,
                            read_only=GridState.vba_locked | (GridState.vba_kind == "form"),
                            font_size=13,
                            on_change=GridState.vba_changed,
                            on_save=GridState.vba_save,
                            on_run=GridState.vba_run,
                        ),
                        flex="1",
                        min_width="0",
                        height="100%",
                    ),
                    spacing="0",
                    width="100%",
                    flex="1",
                    min_height="0",
                ),
                rx.cond(
                    GridState.vba_status != "",
                    rx.text(GridState.vba_status, size="1", color=rx.cond(GridState.vba_status.contains("error"), "#C5221F", "#374151"), id="vba-status", padding_x="4px"),
                ),
                rx.tabs.root(
                    rx.tabs.list(rx.tabs.trigger("Immediate", value="immediate"), rx.tabs.trigger("Problems", value="problems"), size="1"),
                    rx.tabs.content(_immediate(), value="immediate", height="calc(100% - 32px)"),
                    rx.tabs.content(_problems(), value="problems", height="calc(100% - 32px)"),
                    default_value="immediate",
                    height="150px",
                    width="100%",
                    border_top="1px solid #E3E6E9",
                ),
                spacing="2",
                height="100%",
                width="100%",
            ),
            max_width="96vw",
            width="96vw",
            height="90vh",
            padding="12px",
            id="dlg-vba",
            on_escape_key_down=rx.prevent_default,
        ),
        open=GridState.vba_open,
        on_open_change=GridState.vba_open_change,
    )


# ---------------------------------------------------------------- the Macros dialog


def macros_dialog() -> rx.Component:
    selected = GridState.dlg_macro
    return rx.dialog.root(
        rx.dialog.content(
            rx.dialog.title("Macros", size="3"),
            rx.text("Macro name:", size="2"),
            rx.input(value=selected, on_change=lambda v: GridState.set_dlg("selected", v), size="2", width="100%", id="dlg-macro-name"),
            rx.box(
                rx.foreach(
                    GridState.dlg_macros,
                    lambda m: rx.box(
                        rx.text(m["name"].to(str), size="2", color="#111827"),
                        padding="2px 8px",
                        cursor="pointer",
                        background=rx.cond(selected == m["name"].to(str), "#D7E3EA", "transparent"),
                        on_click=GridState.set_dlg("selected", m["name"].to(str)),
                        on_double_click=GridState.run_macro(m["name"].to(str)),
                        class_name="dlg-macro",
                    ),
                ),
                rx.cond(GridState.dlg_macros.length() == 0, rx.text("This workbook has no macros. Type a name and click Create to write one.", size="1", color="#6B7280", padding="8px")),
                border="1px solid #D5D9DE",
                height="200px",
                overflow_y="auto",
                width="100%",
                margin_top="6px",
                id="dlg-macro-list",
            ),
            rx.hstack(
                rx.button("Run", on_click=GridState.run_macro(selected), id="dlg-macro-run", disabled=selected == ""),
                rx.button(rx.cond(GridState.dlg_macros.length() == 0, "Create", "Edit"), variant="soft", on_click=GridState.edit_macro(selected), id="dlg-macro-edit"),
                rx.spacer(),
                rx.button("Cancel", variant="soft", color_scheme="gray", on_click=GridState.close_dialog, id="dlg-cancel"),
                spacing="2",
                width="100%",
                margin_top="12px",
            ),
            max_width="460px",
            id="dlg-macros",
        ),
        open=GridState.dialog == "macros",
        on_open_change=GridState.dialog_open_change,
    )


# ---------------------------------------------------------------- a running macro


def _answer(label: str, value: int, id: str) -> rx.Component:
    return rx.button(label, on_click=GridState.answer_macro(value), id=id, variant="soft", color_scheme="gray", min_width="72px")


def _msgbox_buttons() -> rx.Component:
    kind = GridState.macro_prompt["buttons"].to(int) % 8
    return rx.match(
        kind,
        (1, rx.hstack(_answer("OK", 1, "mb-ok"), _answer("Cancel", 2, "mb-cancel"), spacing="2")),
        (2, rx.hstack(_answer("Abort", 3, "mb-abort"), _answer("Retry", 4, "mb-retry"), _answer("Ignore", 5, "mb-ignore"), spacing="2")),
        (3, rx.hstack(_answer("Yes", 6, "mb-yes"), _answer("No", 7, "mb-no"), _answer("Cancel", 2, "mb-cancel"), spacing="2")),
        (4, rx.hstack(_answer("Yes", 6, "mb-yes"), _answer("No", 7, "mb-no"), spacing="2")),
        (5, rx.hstack(_answer("Retry", 4, "mb-retry"), _answer("Cancel", 2, "mb-cancel"), spacing="2")),
        rx.hstack(_answer("OK", 1, "mb-ok"), spacing="2"),
    )


def _msgbox_icon() -> rx.Component:
    icon = (GridState.macro_prompt["buttons"].to(int) // 16) % 8
    return rx.match(
        icon,
        (1, rx.icon("circle_x", size=30, color="#C5221F")),
        (2, rx.icon("circle_help", size=30, color="#1A73E8")),
        (3, rx.icon("triangle_alert", size=30, color="#E8A400")),
        (4, rx.icon("info", size=30, color="#1A73E8")),
        rx.fragment(),
    )


def macro_prompt_dialog() -> rx.Component:
    P = GridState.macro_prompt
    return rx.dialog.root(
        rx.dialog.content(
            rx.dialog.title(P["title"].to(str), size="3"),
            rx.cond(
                P["kind"].to(str) == "msgbox",
                rx.vstack(
                    rx.hstack(_msgbox_icon(), rx.text(P["prompt"].to(str), size="2", white_space="pre-wrap", color="#111827", id="mb-prompt"), spacing="3", align="start"),
                    rx.hstack(rx.spacer(), _msgbox_buttons(), width="100%", margin_top="12px"),
                    spacing="2",
                    width="100%",
                ),
                rx.form(
                    rx.vstack(
                        rx.text(P["prompt"].to(str), size="2", white_space="pre-wrap", color="#111827", id="ib-prompt"),
                        rx.input(value=GridState.macro_input, on_change=GridState.set_macro_input, size="2", width="100%", auto_focus=True, id="ib-input"),
                        rx.hstack(
                            rx.spacer(),
                            rx.button("OK", type="submit", id="ib-ok"),
                            rx.button("Cancel", type="button", variant="soft", color_scheme="gray", on_click=GridState.answer_macro("cancel"), id="ib-cancel"),
                            spacing="2",
                            width="100%",
                            margin_top="12px",
                        ),
                        spacing="2",
                        width="100%",
                    ),
                    on_submit=lambda _: GridState.answer_macro("ok"),
                    reset_on_submit=False,
                    width="100%",
                ),
            ),
            rx.text("Macro: " + GridState.macro_label, size="1", color="#9CA3AF", margin_top="8px"),
            max_width="460px",
            id="dlg-macro-prompt",
            on_escape_key_down=rx.prevent_default,
            on_pointer_down_outside=rx.prevent_default,
        ),
        open=P.length() > 0,
    )


def macro_running_dialog() -> rx.Component:
    return rx.dialog.root(
        rx.dialog.content(
            rx.hstack(
                rx.spinner(size="2"),
                rx.text("Running ", rx.text.strong(GridState.macro_label), "…", size="2", color="#111827"),
                rx.spacer(),
                rx.button(rx.icon("square", size=12), "Stop", size="1", color_scheme="red", variant="soft", on_click=GridState.stop_macro, id="macro-stop"),
                spacing="3",
                align="center",
                width="100%",
            ),
            max_width="380px",
            id="dlg-macro-running",
            on_escape_key_down=GridState.stop_macro,
            on_pointer_down_outside=rx.prevent_default,
        ),
        open=GridState.macro_busy,
    )


def vba_dialogs() -> rx.Component:
    return rx.fragment(vba_editor_dialog(), macros_dialog(), macro_prompt_dialog(), macro_running_dialog())
