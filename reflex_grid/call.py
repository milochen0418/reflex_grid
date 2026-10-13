"""Calls on a workbook, carried out by a call app through the DDNS Intent ``call.join``
(ported from CoDoc in MD).

ReflexGrid keeps who is in each book's call and rings the people on the book; the call itself
runs in the call app, opened in the intent dialog:

    Intent.start(<call app>, "call.join", private={"room": <call id>}, keep_alive=True,
                 single=True, title=..., user=..., name=...)

- Anyone on a book can call everyone on it: the others on it get an incoming call (Accept /
  Decline) for RING_SECONDS.
- While the call is on, everyone on the book (also people who open it later) sees "Join call"
  with the number of people in it.
- The call dialog is keep-alive (other dialogs minimize it to the tray) and single (joining
  another call closes it).

Whoever starts a call chooses its app among the installed apps that provide ``call.join`` (the
DDNS Intent registry; ``DDNS_INTENT_PROVIDER_CALL_JOIN`` for local development), through
``Intent.choose_app``, which asks only when there are several. The call keeps that app, so
everyone who answers or joins later opens the same app — the same call — without being asked.
``REFLEX_GRID_CALL_APP`` names the app for every call instead. The call id is private: anyone
holding it can join the call, so it stays out of the iframe URL and out of the shared room.

The members and the call id live in this process's memory (``CallStore``); what everyone on
the book must see (how many are in the call, who is ringing, who answered) goes to the book's
``BookRoom``, so it reaches all of them at once.
"""

from __future__ import annotations

import os
import secrets
import time
from typing import Dict, List, Optional, TypedDict

import reflex as rx

from .books import STORE
from .room import BookRoom

try:
    from reflex_ddns_auth.intent import Intent, IntentState

    CALLS_ENABLED = True
except ImportError:  # the app runs without reflex_ddns_auth: no calls
    CALLS_ENABLED = False

CALL_APP = os.environ.get("REFLEX_GRID_CALL_APP") or None
CALL_ACTION = "call.join"
# How long a new call rings the people on the book.
RING_SECONDS = 45
# A tab in a call refreshes its entry this often...
TOUCH_SECONDS = 15
# ...and is dropped after this long without a heartbeat (closed tab, lost network).
# Generous because background tabs get their timers throttled.
STALE_SECONDS = 150
# The dialog opens one event after join_call; don't mistake that gap for a closed dialog.
OPENING_GRACE_SECONDS = 10


class Member(TypedDict):
    id: str
    name: str


class BookCall(TypedDict):
    call_id: str
    # The call app chosen by whoever started the call: everyone joins through it.
    app: str
    started_by_name: str
    started_at: float
    ringing: bool
    # Client token -> who; a person may be in the call from several tabs.
    members: Dict[str, Member]
    last_seen: Dict[str, float]
    # User ids who joined or declined: they are not rung (again).
    responded: List[str]
    # User ids on the book when the call started: the people it rings (not those who come later).
    invited: List[str]


class CallStore:
    """The call of each book, at most one."""

    calls: Dict[str, BookCall] = {}

    @classmethod
    def cleanup(cls, book_id: str, now: float) -> None:
        """Drop members gone silent, stop ringing after RING_SECONDS, end an empty call."""
        call = cls.calls.get(book_id)
        if call is None:
            return
        for token, seen in list(call["last_seen"].items()):
            if now - seen > STALE_SECONDS:
                call["members"].pop(token, None)
                call["last_seen"].pop(token, None)
        if call["ringing"] and now - call["started_at"] > RING_SECONDS:
            call["ringing"] = False
        if not call["members"]:
            del cls.calls[book_id]

    @classmethod
    def join(cls, book_id: str, token: str, member: Member, now: float, app: str, call_id: str = "", invited: Optional[List[str]] = None) -> BookCall:
        """Join the book's call, starting it in ``app`` (ringing ``invited``) if there is none."""
        call = cls.calls.get(book_id)
        if call is None:
            call = BookCall(
                call_id=call_id or secrets.token_urlsafe(12),
                app=app,
                started_by_name=member["name"],
                started_at=now,
                # A call put back after being dropped (call_id given) doesn't ring again.
                ringing=not call_id,
                members={},
                last_seen={},
                responded=[],
                invited=list(invited or []),
            )
            cls.calls[book_id] = call
        call["members"][token] = member
        call["last_seen"][token] = now
        if member["id"] not in call["responded"]:
            call["responded"].append(member["id"])
        return call

    @classmethod
    def leave(cls, book_id: str, token: str) -> None:
        """Take a tab out of the book's call; the call ends with its last member."""
        call = cls.calls.get(book_id)
        if call is None:
            return
        call["members"].pop(token, None)
        call["last_seen"].pop(token, None)
        if not call["members"]:
            del cls.calls[book_id]

    @classmethod
    def decline(cls, book_id: str, user_id: str) -> None:
        call = cls.calls.get(book_id)
        if call is not None and user_id not in call["responded"]:
            call["responded"].append(user_id)


def sync_room(room: BookRoom, book_id: str) -> None:
    """Shows the book's call as it is now to everyone on the book (only what changed)."""
    CallStore.cleanup(book_id, time.time())
    call = CallStore.calls.get(book_id)
    count = len(call["members"]) if call else 0
    ringing_from = call["started_by_name"] if call and call["ringing"] else ""
    responded = list(call["responded"]) if call else []
    invited = list(call["invited"]) if call and call["ringing"] else []
    if room.call_count != count:
        room.call_count = count
    if room.call_ringing_from != ringing_from:
        room.call_ringing_from = ringing_from
    if list(room.call_responded) != responded:
        room.call_responded = responded
    if list(room.call_invited) != invited:
        room.call_invited = invited


class CallState(rx.State):
    """This tab's side of book calls."""

    # Book whose call this tab is in ("" when in no call).
    current_book: str = ""
    _call_id: str = ""
    # The call's app, to put the call back as it was.
    _call_app: str = ""
    _joined_at: float = 0.0
    # Book of the call this tab starts, while its caller chooses the call app.
    _starting_book: str = ""

    async def _me(self) -> tuple[str, str, str]:
        """(book on screen, user id, display name) of this tab."""
        from .state import GridState

        grid = await self.get_state(GridState)
        me = grid.me or {}
        return grid.book, me.get("user", ""), me.get("name", "")

    async def _room(self, book_id: str) -> None:
        """Pushes the call of ``book_id`` to its room, when that is the book on screen."""
        from .state import GridState

        grid = await self.get_state(GridState)
        if book_id and grid.book == book_id:
            sync_room(await self.get_state(BookRoom), book_id)

    @rx.event
    async def join_call(self):
        """Call everyone on the book, or join its call, and open it in the dialog."""
        return await self._join(CALL_APP or "")

    @rx.event
    async def start_call(self, data: dict):
        """The caller chose the call app: call everyone on the book with it."""
        book_id, self._starting_book = self._starting_book, ""
        on_screen, _, _ = await self._me()
        if book_id and book_id == on_screen and data.get("app"):
            # Joins instead if someone called everyone on the book meanwhile.
            return await self._join(data["app"])

    async def _join(self, app: str):
        """Join the book's call in its app, or start the call in ``app``.

        Without ``app``, a new call first lets the caller choose it (``start_call`` comes back
        here with it).
        """
        if not CALLS_ENABLED:
            return
        book_id, user_id, name = await self._me()
        if not book_id or not user_id:
            return
        dialog = await self.get_state(IntentState)
        if self.current_book == book_id and CALL_ACTION in dialog.open_actions:
            # Already in this call (maybe minimized): bring it back rather than rejoin.
            return Intent.show(CALL_ACTION)
        if book_id not in CallStore.calls and not app:
            # Everyone who joins later opens the app chosen now: other apps can't reach the call.
            self._starting_book = book_id
            return Intent.choose_app(CALL_ACTION, on_result=CallState.start_call)

        now = time.time()
        token = self.router.session.client_token
        if self.current_book and self.current_book != book_id:
            # Switching calls: the old dialog's on_cancel comes too late to know which one.
            CallStore.leave(self.current_book, token)
        book = STORE.get(book_id)
        invited = sorted({v.info.get("user", "") for v in list(book.views.values())} - {""}) if book is not None else []
        call = CallStore.join(book_id, token, Member(id=user_id, name=name), now, app, invited=invited)
        self.current_book = book_id
        self._call_id = call["call_id"]
        self._call_app = call["app"]
        self._joined_at = now
        await self._room(book_id)

        room = await self.get_state(BookRoom)
        title = room.title or "Workbook"
        return Intent.start(
            call["app"],
            CALL_ACTION,
            on_result=CallState.call_closed,
            on_cancel=CallState.call_closed,
            private={"room": call["call_id"]},
            keep_alive=True,
            single=True,
            label=f"Call · {title}",
            title=title,
            user=user_id,
            name=name,
        )

    @rx.event
    async def decline_call(self):
        book_id, user_id, _ = await self._me()
        CallStore.decline(book_id, user_id)
        await self._room(book_id)

    @rx.event
    async def call_closed(self, data: dict):
        """The call dialog closed (hung up or closed): leave the call.

        Other dialogs only minimize it (keep-alive), so being replaced means another call took
        its place (single): join_call already moved us.
        """
        if data.get("reason") == "replaced" and data.get("by") == CALL_ACTION:
            return
        book_id = self.current_book
        self.current_book = ""
        self._call_id = ""
        if book_id:
            CallStore.leave(book_id, self.router.session.client_token)
            await self._room(book_id)

    @rx.event
    async def heartbeat(self):
        """Keep this tab in its call while the call dialog runs (in front or minimized)."""
        if not CALLS_ENABLED or not self.current_book:
            return
        now = time.time()
        token = self.router.session.client_token
        book_id = self.current_book
        dialog = await self.get_state(IntentState)
        if CALL_ACTION not in dialog.open_actions and now - self._joined_at > OPENING_GRACE_SECONDS:
            # The dialog went away without telling us: we are not in the call.
            CallStore.leave(book_id, token)
            self.current_book = ""
            self._call_id = ""
            await self._room(book_id)
            return

        CallStore.cleanup(book_id, now)
        call = CallStore.calls.get(book_id)
        if call is not None and call["call_id"] != self._call_id:
            # Another call took this book's place while ours was dropped.
            self.current_book = ""
            self._call_id = ""
            return
        if call is None or token not in call["members"] or now - call["last_seen"][token] > TOUCH_SECONDS:
            # Touch our entry, or put the call back if it was dropped (e.g. this tab slept past
            # STALE_SECONDS) while the dialog still runs it.
            _, user_id, name = await self._me()
            CallStore.join(book_id, token, Member(id=user_id, name=name), now, self._call_app, call_id=self._call_id)
            if call is None:
                await self._room(book_id)
