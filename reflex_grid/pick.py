"""``grid.pick``: choosing, from another app, what of a workbook to show there. A workbook from the
list, then one of its UserForms, tables or names, or cells from a preview of a sheet; what was chosen
goes back as the line to put in the document (:func:`.embed.directive`). This only reads workbooks,
and the preview is plain text."""

from __future__ import annotations

import datetime as dt
from typing import Any, Dict, List

import reflex as rx
from reflex_ddns_auth.intent import IntentPage

from gridcore.core.addr import RangeRef
from gridcore.engine import embed as cells

from .books import STORE, BookMeta
from .embed import GONE, directive, open_book


def _note(m: BookMeta) -> str:
    n = len(m.sheets)
    when = dt.datetime.fromtimestamp(m.updated_at).strftime("%Y-%m-%d %H:%M") if m.updated_at else ""
    return f"{n} sheet{'' if n == 1 else 's'} · {when}" if when else f"{n} sheet{'' if n == 1 else 's'}"


class PickState(rx.State):
    books: List[Dict[str, str]] = []
    query: str = ""
    error: str = ""
    # The workbook chosen, and what it offers.
    book: str = ""
    title: str = ""
    forms: List[Dict[str, str]] = []
    tables: List[Dict[str, Any]] = []
    names: List[Dict[str, Any]] = []
    # The sheet on screen: its preview, from row ``top`` and column ``left``, and what is selected
    # there, as ``embed.suggest`` understood it.
    sheets: List[str] = []
    sheet: str = ""
    rows: List[List[str]] = []
    top: int = 0
    left: int = 0
    more: bool = False
    selected: List[int] = []
    choice: Dict[str, Any] = {}
    grow: bool = False

    @rx.var
    def shown_books(self) -> List[Dict[str, str]]:
        q = self.query.strip().lower()
        return [b for b in self.books if q in b["title"].lower()] if q else self.books

    @rx.var
    def ref(self) -> str:
        """The reference of what is selected."""
        c = self.choice or {}
        if c.get("kind") == "table":
            return str(c.get("ref") or "")
        return str((c.get("growing") if self.grow else c.get("fixed")) or c.get("fixed") or "")

    @rx.var
    def can_grow(self) -> bool:
        return bool((self.choice or {}).get("growing"))

    @rx.event
    def load(self, params: Dict[str, Any]):
        self.book, self.query, self.error = "", "", ""
        STORE.close_idle()
        self.books = [{"id": m.id, "title": m.title, "note": _note(m)} for m in STORE.list()]

    @rx.event
    def set_query(self, text: str):
        self.query = text

    @rx.event
    def back(self):
        self.book, self.error = "", ""

    def _wb(self):
        book = STORE.get(self.book)
        if book is None:
            return None
        with book.lock:
            return book.doc.wb

    def _select(self, wb, si: int, rng: RangeRef) -> None:
        c = cells.suggest(wb, si, rng)
        r = RangeRef.parse(c["area"])
        self.choice, self.grow = c, c["grows"]
        self.selected = [r.r0, r.c0, r.r1, r.c1]

    def _show(self, wb, name: str) -> None:
        si = wb.sheet_index(name)
        p = cells.preview(wb, si)
        self.sheet, self.rows, self.top, self.left, self.more = name, p["rows"], p["r0"], p["c0"], p["more"]
        self._select(wb, si, cells.main_block(wb.sheets[si]))

    @rx.event
    async def choose(self, book_id: str):
        book = await open_book(book_id)
        if book is None:
            self.error = GONE
            return
        with book.lock:
            wb = book.doc.wb
        self.book, self.title, self.error = book_id, book.meta.title, ""
        self.forms, self.tables, self.names = cells.forms(wb), cells.tables(wb), cells.names(wb)
        self.sheets = [s.name for s in wb.sheets if s.visibility == "visible"]
        self._show(wb, book.home_sheet if book.home_sheet in self.sheets else self.sheets[0])

    @rx.event
    def show_sheet(self, name: str):
        wb = self._wb()
        if wb is not None and wb.sheet_index(name) is not None:
            self._show(wb, name)

    @rx.event
    def select(self, s: List[int]):
        wb = self._wb()
        si = wb.sheet_index(self.sheet) if wb is not None else None
        if si is not None and isinstance(s, list) and len(s) == 4:
            self._select(wb, si, RangeRef(*(int(x) for x in s)))

    @rx.event
    def set_grow(self, on: bool):
        self.grow = bool(on)

    @rx.event
    def insert(self, kind: str, value: str):
        return IntentPage.finish({"code": directive(self.book, kind, value)})
