"""Engine sessions live in the backend process, one per browser tab (keyed by Reflex's client
token), outside Reflex state: a workbook can be far larger than what state should carry."""

from __future__ import annotations

import threading
import time
from dataclasses import dataclass, field
from typing import Dict, Optional

from gridcore.engine import Session

MAX_SESSIONS = 64
IDLE_SECONDS = 6 * 3600


@dataclass
class Entry:
    session: Session
    last_used: float = field(default_factory=time.time)
    # Signature of what geometry/objects were computed from (doc uid, workbook id, sheet).
    sig: tuple = ()
    sheet: int = -1
    # Picture data URLs by (doc uid, object id, size).
    images: Dict[tuple, str] = field(default_factory=dict)


_LOCK = threading.Lock()
_SESSIONS: Dict[str, Entry] = {}


def get(token: str) -> Entry:
    now = time.time()
    with _LOCK:
        e = _SESSIONS.get(token)
        if e is None:
            s = Session()
            s.new_workbook()
            e = Entry(s)
            _SESSIONS[token] = e
            _evict(now)
        e.last_used = now
        return e


def peek(token: str) -> Optional[Entry]:
    with _LOCK:
        return _SESSIONS.get(token)


def _evict(now: float) -> None:
    stale = [k for k, v in _SESSIONS.items() if now - v.last_used > IDLE_SECONDS]
    for k in stale:
        del _SESSIONS[k]
    if len(_SESSIONS) > MAX_SESSIONS:
        for k, _ in sorted(_SESSIONS.items(), key=lambda kv: kv[1].last_used)[: len(_SESSIONS) - MAX_SESSIONS]:
            del _SESSIONS[k]
