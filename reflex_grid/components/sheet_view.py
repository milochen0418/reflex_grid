"""The grid: a local React component (``sheet_view.jsx``) wrapped for Reflex."""

import reflex as rx

# Both files are served from the same folder; the view imports the charts module relatively.
_CHARTS = rx.asset("sheet_charts.jsx", shared=True)
_JS = rx.asset("sheet_view.jsx", shared=True)


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
    on_event: rx.EventHandler[lambda e: [e]]


sheet_view = SheetView.create
