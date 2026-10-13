"""ReflexGrid: an Excel-compatible spreadsheet in the browser, built with Reflex.

The page is the classic layout: title bar, ribbon, formula bar + grid, sheet tabs, status bar.
All spreadsheet logic lives in the ``gridcore`` engine;
this module only arranges the UI."""

import reflex as rx

from .components.chrome import backstage, sheet_tabs, status_bar, title_bar
from .components.dialogs import all_dialogs
from .components.ribbon import ribbon
from .components.sheet_view import sheet_view
from .state import GridState


def index() -> rx.Component:
    return rx.box(
        rx.vstack(
            title_bar(),
            ribbon(),
            rx.box(
                sheet_view(
                    geometry=GridState.geometry,
                    data=GridState.data,
                    objects=GridState.objects,
                    ui=GridState.ui,
                    functions=GridState.functions,
                    names=GridState.names,
                    focus_key=GridState.focus_key,
                    edit_request=GridState.edit_request,
                    on_event=GridState.on_event,
                ),
                flex="1",
                min_height="0",
                width="100%",
                position="relative",
            ),
            sheet_tabs(),
            status_bar(),
            spacing="0",
            height="100vh",
            width="100%",
            overflow="hidden",
        ),
        backstage(),
        all_dialogs(),
        position="relative",
        height="100vh",
        width="100vw",
        overflow="hidden",
        background="#FFFFFF",
    )


FONT_FACES = "".join(
    f"@font-face {{ font-family: 'Carlito'; font-style: {style}; font-weight: {weight}; font-display: swap; src: url('/fonts/carlito-latin-{weight}-{style}.woff2') format('woff2'); }}\n"
    for weight in (400, 700)
    for style in ("normal", "italic")
) + "".join(
    f"@font-face {{ font-family: 'Carlito'; font-style: normal; font-weight: {weight}; font-display: swap; src: url('/fonts/carlito-latin-ext-{weight}-normal.woff2') format('woff2'); unicode-range: U+0100-02BA, U+02BD-02C5, U+02C7-02CC, U+02CE-02D7, U+02DD-02FF, U+0304, U+0308, U+0329, U+1D00-1DBF, U+1E00-1E9F, U+1EF2-1EFF, U+2020, U+20A0-20AB, U+20AD-20C0, U+2113, U+2C60-2C7F, U+A720-A7FF; }}\n"
    for weight in (400, 700)
)

STYLE = FONT_FACES + """
html, body { margin: 0; padding: 0; overflow: hidden; }
.rg-rbtn { cursor: default; flex-shrink: 0; min-width: 26px; margin: 0 !important; }
.rg-ribbon .rt-Flex { flex-shrink: 0; }
.rg-biglabel { white-space: normal; max-width: 74px; line-height: 1.1; text-align: center; display: block; }
.rg-big { align-items: flex-start; padding-top: 4px; }
.rg-marquee { animation: rg-march 0.6s linear infinite; }
@keyframes rg-march { to { stroke-dashoffset: -8; } }
.rg-ribbon::-webkit-scrollbar { height: 6px; }
.rg-sheet-tabs::-webkit-scrollbar { height: 4px; }
"""

app = rx.App(
    theme=rx.theme(accent_color="cyan", gray_color="slate", radius="small", scaling="95%"),
    head_components=[rx.el.style(STYLE), rx.el.link(rel="icon", href="/favicon.svg", type="image/svg+xml")],
)
app.add_page(index, title="ReflexGrid", description="An Excel-compatible spreadsheet built with Reflex", on_load=GridState.on_load)
