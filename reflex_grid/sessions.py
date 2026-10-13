"""Engine sessions live in the backend process, one per browser tab (keyed by Reflex's client
token), outside Reflex state: a workbook can be far larger than what state should carry.

A tab's session is a :class:`CollabSession` (its clipboard, Find, Format Painter, preferences)
bound to the shared book the tab shows, through the tab's :class:`View` of that book."""

from __future__ import annotations

import threading
import time
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from gridcore.engine import CollabSession, View

MAX_SESSIONS = 256
IDLE_SECONDS = 6 * 3600


@dataclass
class Entry:
    session: CollabSession
    last_used: float = field(default_factory=time.time)
    # The book this tab shows and the tab's view of it.
    book_id: str = ""
    view: Optional[View] = None
    # Signature of what geometry/objects were computed from (doc uid, workbook id, sheet, zoom,
    # formulas view).
    sig: tuple = ()
    # Name and position of the sheet the grid shows (its window starts at the top-left on
    # another sheet).
    sheet: Optional[str] = None
    sheet_index: int = -1
    # Picture data URLs by (doc uid, object id, size).
    images: Dict[tuple, str] = field(default_factory=dict)


_LOCK = threading.Lock()
_SESSIONS: Dict[str, Entry] = {}


def get(token: str) -> Entry:
    now = time.time()
    with _LOCK:
        e = _SESSIONS.get(token)
        if e is None:
            e = Entry(CollabSession())
            _SESSIONS[token] = e
            _evict(now)
        e.last_used = now
        return e


def peek(token: str) -> Optional[Entry]:
    with _LOCK:
        return _SESSIONS.get(token)


def tokens_on(book_id: str) -> List[str]:
    """Client tokens of the tabs showing ``book_id``, most recently active first."""
    with _LOCK:
        on = [(e.last_used, k) for k, e in _SESSIONS.items() if e.book_id == book_id and e.view is not None]
    return [k for _, k in sorted(on, reverse=True)]


def _evict(now: float) -> None:
    stale = [k for k, v in _SESSIONS.items() if now - v.last_used > IDLE_SECONDS]
    if len(_SESSIONS) - len(stale) > MAX_SESSIONS:
        stale += [k for k, _ in sorted(_SESSIONS.items(), key=lambda kv: kv[1].last_used) if k not in stale][: len(_SESSIONS) - len(stale) - MAX_SESSIONS]
    for k in stale:
        e = _SESSIONS.pop(k)
        s = e.session
        if s.book is not None and s.view is not None:
            s.book.leave(s.view.key)
