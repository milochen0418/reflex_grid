"""What everybody on a workbook shares, pushed to all of them at once.

``BookRoom`` is a Reflex ``SharedState``: every browser tab on a book links its ``BookRoom`` to
the book (``room_token``), so a change one tab makes to it reaches all the others. It carries
small, broadcast things: the workbook's revision, who is here and what they have selected, the
recent changes (to highlight them), the title, whether it is saved, and the book's call.

The cells themselves are not in it: the workbook lives once on the server, in the shared book
(:mod:`reflex_grid.books`), and each tab draws its own window of it (``GridState``) — a sheet
can hold millions of cells and each person scrolls somewhere else. When ``rev`` moves, a tab
asks for its window again.
"""

from __future__ import annotations

from typing import Any, Dict, List

import reflex as rx


def room_token(book_id: str) -> str:
    """The shared-state token of a book (Reflex link tokens can't contain ``_``)."""
    return f"book-{book_id}"


class BookRoom(rx.SharedState):
    # Revision of the shared workbook: bumped by every change anyone makes.
    rev: int = 0
    title: str = ""
    # One entry per tab on the book: {"id", "user", "name", "color", "avatar", "sheet",
    # "active": [r, c], "ranges": [[r0, c0, r1, c1], …], "editing": [r, c] | None}.
    peers: List[Dict[str, Any]] = []
    # The latest changes: {"rev", "author", "name", "color", "areas": [[sheet, [r0, c0, r1, c1]]]}.
    changes: List[Dict[str, Any]] = []
    saved: bool = True
    save_error: str = ""
    # Someone deleted the book from the workbook list.
    deleted: bool = False
    # The book's call: how many are in it, who is ringing everyone (while it rings), the people
    # it rings (those on the book when it started), and those who joined or declined.
    call_count: int = 0
    call_ringing_from: str = ""
    call_invited: List[str] = []
    call_responded: List[str] = []

    @rx.var
    def people(self) -> List[Dict[str, Any]]:
        """One entry per person: several tabs of one person count once."""
        out: List[Dict[str, Any]] = []
        seen = set()
        for p in self.peers:
            user = p.get("user") or p.get("id")
            if user in seen:
                continue
            seen.add(user)
            name = p.get("name", "") or "?"
            initials = "".join(w[0] for w in name.split()[:2]).upper() or "?"
            out.append({"id": p.get("id", ""), "user": user, "name": name, "initials": initials, "color": p.get("color", ""), "avatar": p.get("avatar", "")})
        return out

    @rx.var
    def people_count(self) -> int:
        return len(self.people)
