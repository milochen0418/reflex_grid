"""AI controls (through the AI Hub of re_ddns, see ``ai.py``): the title bar's button that opens
the hub's assistant on this workbook, the AI Fill dialog and its progress. Without the AI they
render nothing."""

from __future__ import annotations

import reflex as rx

from .. import ai
from ..room import BookRoom
from ..state import GridState
from .dialogs import D, check, ok_cancel, shell, text_field

try:
    from reflex_ddns_auth.intent import Intent
except ImportError:  # no platform: no assistant
    Intent = None


def ai_button() -> rx.Component:
    """Opens the AI Hub's assistant on this workbook (the ``ai.assist`` intent), in a dialog that
    stays open beside the grid; Alt+A from the grid does the same."""
    if not ai.ENABLED or Intent is None:
        return rx.fragment()
    return rx.tooltip(
        rx.button(
            rx.icon("sparkles", size=14),
            "AI",
            on_click=Intent.start(None, "ai.assist", keep_alive=True, single=True, label="AI assistant", in_app=ai.APP, room=GridState.ai_room),
            size="1",
            variant="outline",
            id="ai-assist",
            custom_attrs={"data-room": GridState.ai_room},
            style={"color": "#FFFFFF", "background": "rgba(255,255,255,0.10)", "boxShadow": "inset 0 0 0 1px rgba(255,255,255,0.45)", "cursor": "pointer"},
        ),
        content="Ask the AI about this workbook (Alt+A)",
    )


def ai_fill_dialog() -> rx.Component:
    if not ai.ENABLED:
        return rx.fragment()
    return shell(
        "aiFill",
        "AI Fill",
        rx.text("Rows: ", D["sheet"].to(str), "!", D["range"].to(str), size="1", color="#6B7280", id="ai-fill-range"),
        rx.text("What should the AI write for each row?", size="2"),
        # Not bound while typing (a value echoed back breaks Chinese and Japanese input): kept
        # when the field is left, which clicking Fill does first.
        rx.text_area(
            default_value=D["instruction"].to(str),
            on_blur=lambda v: GridState.set_dlg("instruction", v),
            placeholder="e.g. Positive, Negative or Neutral · the city in the address · translate to English",
            rows="3",
            width="100%",
            id="dlg-instruction",
        ),
        check("The first row holds the headers", "header"),
        text_field("Results go to column", "target", width="80px"),
        check("Replace what is there", "overwrite"),
        rx.text("The answers are written as values, in one step: Undo takes them back. The AI Hub chooses the model.", size="1", color="#6B7280"),
        footer=ok_cancel("Fill"),
        width="480px",
    )


def ai_job_panel() -> rx.Component:
    """An AI Fill at work: rows done of how many, and Stop."""
    if not ai.ENABLED:
        return rx.fragment()
    job = GridState.ai_job
    return rx.cond(
        job["running"].to(bool),
        rx.hstack(
            rx.icon("sparkles", size=14, color="#7C3AED"),
            rx.text("AI Fill ", job["done"].to(str), " / ", job["total"].to(str), " rows", size="1", weight="medium", white_space="nowrap", id="ai-fill-count"),
            rx.text(job["what"].to(str), size="1", color="#6B7280", max_width="240px", overflow="hidden", text_overflow="ellipsis", white_space="nowrap"),
            rx.cond(
                job["stopping"].to(bool),
                rx.text("Stopping…", size="1", color="#6B7280"),
                rx.button("Stop", size="1", variant="soft", color_scheme="gray", on_click=GridState.ai_fill_stop, id="ai-fill-stop"),
            ),
            spacing="2",
            align="center",
            id="ai-fill-progress",
            style={"position": "fixed", "left": "12px", "bottom": "34px", "zIndex": "30", "background": "#FFFFFF", "border": "1px solid #E5E7EB", "borderRadius": "8px", "padding": "6px 10px", "boxShadow": "0 4px 14px rgba(0,0,0,0.12)"},
        ),
    )


def ai_change_bar() -> rx.Component:
    """What the assistant just changed for this person, with the button that takes it back
    (above the assistant's dialog, which stays open while it works)."""
    if not ai.ENABLED:
        return rx.fragment()

    def note(n) -> rx.Component:
        return rx.cond(
            (n["peer"].to(str) == GridState.me["peer"].to(str)) | (n["peer"].to(str) == ai.AI_USER),
            rx.hstack(
                rx.icon("sparkles", size=14, color=ai.AI_COLOR),
                rx.text("AI: ", n["summary"].to(str), size="1", color="#1F2937", max_width="min(340px, 50vw)", overflow="hidden", text_overflow="ellipsis", white_space="nowrap", id="ai-change-text"),
                rx.button("Undo", size="1", variant="soft", on_click=GridState.ai_undo(n["rev"].to(int)), title="Take back this change, whatever you did after it", id="ai-change-undo"),
                rx.icon_button(rx.icon("x", size=12), size="1", variant="ghost", color_scheme="gray", on_click=GridState.ai_close(n["rev"].to(int)), title="Dismiss (Ctrl+Z still takes it back)", id="ai-change-close"),
                spacing="2",
                align="center",
                id="ai-change-bar",
                style={"position": "fixed", "left": "12px", "bottom": "74px", "zIndex": "1001", "background": "#FFFFFF", "border": "1px solid #DDD6FE", "borderRadius": "8px", "padding": "6px 10px", "boxShadow": "0 4px 14px rgba(0,0,0,0.12)"},
            ),
        )

    return rx.fragment(rx.foreach(BookRoom.ai_notes, note))
