"""Call controls (from CoDoc in MD): call everyone on the workbook, incoming calls, the call's
heartbeat."""

from __future__ import annotations

import reflex as rx

from ..call import CALLS_ENABLED, CallState
from ..room import BookRoom
from ..state import GridState


def call_button() -> rx.Component:
    """Call everyone on the workbook, or join its call while one is in progress."""
    if not CALLS_ENABLED:
        return rx.fragment()
    return rx.cond(
        BookRoom.call_count > 0,
        rx.button(
            rx.icon("phone_call", size=14),
            "Join call · ",
            BookRoom.call_count,
            on_click=CallState.join_call,
            size="1",
            color_scheme="green",
            variant="solid",
            id="call-join",
            class_name="rg-call-btn",
        ),
        rx.button(
            rx.icon("phone", size=14),
            "Call everyone",
            on_click=CallState.join_call,
            title="Call everyone on this workbook",
            size="1",
            variant="outline",
            id="call-everyone",
            class_name="rg-call-btn",
            style={"color": "#FFFFFF", "background": "rgba(255,255,255,0.10)", "boxShadow": "inset 0 0 0 1px rgba(255,255,255,0.45)", "cursor": "pointer"},
        ),
    )


def incoming_call() -> rx.Component:
    """Someone called everyone on the workbook: accept or decline."""
    if not CALLS_ENABLED:
        return rx.fragment()
    me = GridState.me["user"]
    rung = (BookRoom.call_ringing_from != "") & BookRoom.call_invited.contains(me) & ~BookRoom.call_responded.contains(me) & (GridState.book != "")
    # Above the intent dialog (z-index 1000), so a call can be answered from inside another.
    return rx.cond(
        rung,
        rx.el.div(
            rx.el.div(
                rx.icon("phone_incoming", size=20, color="#FFFFFF"),
                style={"width": "40px", "height": "40px", "borderRadius": "9999px", "background": "#22C55E", "display": "flex", "alignItems": "center", "justifyContent": "center", "flexShrink": "0"},
                class_name="animate-pulse",
            ),
            rx.el.div(
                rx.el.span("Incoming call", style={"fontSize": "11px", "fontWeight": "600", "textTransform": "uppercase", "letterSpacing": "0.04em", "color": "#16A34A"}),
                rx.el.span(BookRoom.call_ringing_from, style={"fontSize": "14px", "fontWeight": "700", "color": "#111827", "overflow": "hidden", "textOverflow": "ellipsis", "whiteSpace": "nowrap"}),
                rx.el.span("is calling everyone on this workbook", style={"fontSize": "12px", "color": "#6B7280"}),
                style={"display": "flex", "flexDirection": "column", "minWidth": "0", "flex": "1"},
            ),
            rx.el.button(
                rx.icon("phone_off", size=16),
                on_click=CallState.decline_call,
                title="Decline",
                aria_label="Decline",
                style={"padding": "10px", "borderRadius": "9999px", "background": "#FEE2E2", "color": "#DC2626", "border": "0", "cursor": "pointer", "display": "flex"},
            ),
            rx.el.button(
                rx.icon("phone", size=16),
                on_click=CallState.join_call,
                title="Accept",
                aria_label="Accept",
                style={"padding": "10px", "borderRadius": "9999px", "background": "#16A34A", "color": "#FFFFFF", "border": "0", "cursor": "pointer", "display": "flex"},
            ),
            role="alertdialog",
            aria_label="Incoming call",
            style={
                "position": "fixed",
                "bottom": "40px",
                "right": "24px",
                "zIndex": "1001",
                "width": "320px",
                "display": "flex",
                "alignItems": "center",
                "gap": "12px",
                "padding": "12px",
                "background": "#FFFFFF",
                "borderRadius": "16px",
                "border": "1px solid #F3F4F6",
                "boxShadow": "0 20px 40px rgba(0,0,0,0.22)",
            },
        ),
    )


_CALL_HEARTBEAT_JS = """
if (window.__rgCallHb) clearInterval(window.__rgCallHb);
window.__rgCallHb = setInterval(function () {
    var el = document.getElementById('call-heartbeat-trigger');
    if (!el) { clearInterval(window.__rgCallHb); window.__rgCallHb = null; return; }
    el.click();
}, 3000);
"""


@rx.memo
def call_keepalive() -> rx.Component:
    """While in a call, keep its heartbeat going on every page.

    The call runs in the app-wide intent dialogs, so it outlives the workbook page; the
    heartbeat stops with the tab, which then drops out of the call.
    """
    return rx.cond(
        CallState.current_book != "",
        rx.el.button(
            id="call-heartbeat-trigger",
            on_click=CallState.heartbeat,
            on_mount=rx.call_script(_CALL_HEARTBEAT_JS),
            tab_index=-1,
            aria_hidden="true",
            class_name="sr-only",
        ),
    )
