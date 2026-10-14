"""ReflexGrid: Excel-compatible spreadsheets in the browser, shared and edited together live,
built with Reflex.

Pages: ``/`` lists the workbooks; ``/book/<id>`` is a workbook in the classic layout (title
bar, ribbon, formula bar + grid, sheet tabs, status bar), shared by everyone who opens that
address. All spreadsheet logic lives in the ``gridcore`` engine; this module only arranges the UI."""

import reflex as rx
from reflex.style import set_color_mode

from .api import api
from .components.call_views import call_keepalive, incoming_call
from .components.chrome import backstage, book_gone, sheet_tabs, status_bar, title_bar
from .components.dialogs import all_dialogs
from .components.manager import manager_page
from .components.ribbon import ribbon
from .components.sheet_view import sheet_view
from .components.userform import forms_layer
from .components.vba import macro_bar, vba_dialogs
from .room import BookRoom
from .state import AuthState, GridState
from .workbooks import ManagerState


def light_only() -> rx.Component:
    """Keeps the app light. Reflex remembers a colour mode in the browser ("system" or "dark", left
    by an older version) and applies it over the theme's appearance; with the OS in dark mode the
    dialogs then turned dark under our dark text. This puts it back to light (and remembers that)."""
    return rx.box(display="none", on_mount=set_color_mode("light"))


def index() -> rx.Component:
    return rx.box(
        light_only(),
        rx.vstack(
            title_bar(),
            ribbon(),
            macro_bar(),
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
                    peers=BookRoom.peers,
                    rev=BookRoom.rev,
                    changes=BookRoom.changes,
                    me=GridState.me["peer"],
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
        forms_layer(),
        vba_dialogs(),
        book_gone(),
        incoming_call(),
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
/* ReflexGrid is a light UI, like Excel's: an OS in dark mode must not recolour it (the theme is
   light too). color-scheme keeps native inputs and scrollbars light as well. */
:root { color-scheme: light; }
html, body { margin: 0; padding: 0; overflow: hidden; background: #FFFFFF; color: #1F2937; }
/* Ribbon buttons: dark labels and icons, as in Excel (the theme's grey is too faint). */
.rg-ribbon .rg-rbtn { color: #262626; }
/* The File backstage's menu: white on the accent colour, also when hovered. */
.rg-backstage-nav .rt-Button:hover { background: rgba(255, 255, 255, 0.14); }
.rg-rbtn { cursor: default; flex-shrink: 0; min-width: 26px; margin: 0 !important; }
.rg-ribbon .rt-Flex { flex-shrink: 0; }
.rg-biglabel { white-space: normal; max-width: 74px; line-height: 1.1; text-align: center; display: block; }
.rg-big { align-items: flex-start; padding-top: 4px; }
.rg-marquee { animation: rg-march 0.6s linear infinite; }
@keyframes rg-march { to { stroke-dashoffset: -8; } }
.rg-ribbon::-webkit-scrollbar { height: 6px; }
.rg-sheet-tabs::-webkit-scrollbar { height: 4px; }
.rg-flash { opacity: 0; animation: rg-flash 1.6s ease-out; }
@keyframes rg-flash { from { opacity: 0.38; } to { opacity: 0; } }
.rg-title-input { width: 260px; max-width: 30vw; height: 26px; padding: 0 8px; border: 1px solid transparent; border-radius: 4px; background: transparent; color: #FFFFFF; font: 600 14px 'Segoe UI', system-ui, sans-serif; outline: none; text-overflow: ellipsis; }
.rg-title-input:hover { border-color: rgba(255, 255, 255, 0.45); }
.rg-title-input:focus { background: #FFFFFF; color: #111827; border-color: #FFFFFF; }
#manager { height: 100vh; overflow: auto; }
"""

app = rx.App(
    # Always light (as CoDoc in MD): the app paints white sheets, ribbon and dialogs, so the theme's
    # dark colours (light text) would vanish on them when the OS is in dark mode.
    theme=rx.theme(appearance="light", accent_color="cyan", gray_color="slate", radius="small", scaling="95%"),
    head_components=[rx.el.style(STYLE), rx.el.link(rel="icon", href="/favicon.svg", type="image/svg+xml")],
    # Our own backend routes (a closing tab's goodbye), next to Reflex's.
    api_transformer=api,
)

_auth_on_load = [AuthState.load_auth] if AuthState is not None else []

try:
    from reflex_ddns_auth.intent import install_intent_host

    # Intent dialogs (calls) are hosted once, above every page, so a call survives going back
    # to the workbook list.
    install_intent_host(app)
    app.extra_app_wraps[(8, "ReflexGridCallKeepalive")] = lambda stateful: rx.fragment(call_keepalive())
except ImportError:
    pass

def workbook_list() -> rx.Component:
    return rx.fragment(light_only(), manager_page())


app.add_page(workbook_list, route="/", title="ReflexGrid — Workbooks", description="Shared Excel-compatible workbooks, edited together live", on_load=[*_auth_on_load, GridState.leave_book, ManagerState.load])
app.add_page(index, route="/book/[book_id]", title="ReflexGrid", description="An Excel-compatible spreadsheet built with Reflex", on_load=[*_auth_on_load, GridState.on_load])
