"""The grid: a local React component (``assets/grid/sheet_view.jsx``) wrapped for Reflex.

The JSX files are plain assets (Reflex copies ``assets/`` into the frontend), not shared ones:
a shared asset is a symlink Reflex writes into ``assets/external/`` with this machine's absolute
path, which breaks wherever else the same checkout runs (e.g. mounted into an App Store dev
container), and Reflex 0.8 then fails to start."""

import reflex as rx

# Both files are served from the same folder; the view imports the charts module relatively.
_CHARTS = rx.asset("grid/sheet_charts.jsx")
_JS = rx.asset("grid/sheet_view.jsx")


class SheetView(rx.Component):
    library = f"$/public{_JS}"
    tag = "SheetView"

    geometry: rx.Var[dict]
    data: rx.Var[dict]
    objects: rx.Var[dict]
    ui: rx.Var[dict]
    functions: rx.Var[list]
    names: rx.Var[list]
    focus_key: rx.Var[int]
    edit_request: rx.Var[dict]
    # The book's room: everyone on it (with their selections), its revision and recent changes,
    # and which of the peers is this tab.
    peers: rx.Var[list]
    rev: rx.Var[int]
    changes: rx.Var[list]
    me: rx.Var[str]
    on_event: rx.EventHandler[lambda e: [e]]


sheet_view = SheetView.create
