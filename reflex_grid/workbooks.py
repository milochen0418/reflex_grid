"""The workbook list (``/``): every shared workbook, who is editing which right now, and creating,
uploading, downloading and deleting them (like CoDoc's Document Manager)."""

from __future__ import annotations

import asyncio
import os
import time
from typing import Dict, List, TypedDict

import reflex as rx

from gridcore.engine import EngineError
from gridcore.engine.cmd_file import MIMES

from .books import STORE
from .state import push

try:
    from reflex_ddns_auth import AuthState
except ImportError:
    AuthState = None

SEARCH_INPUT_ID = "book-search"
_TASKS: set = set()


def _tell(book_id: str) -> None:
    """Tells the people on a deleted book now (outside this event)."""
    task = asyncio.create_task(push(book_id))
    _TASKS.add(task)
    task.add_done_callback(_TASKS.discard)


class BookItem(TypedDict):
    id: str
    title: str
    updated_at: float
    when: str
    created_by_email: str
    created_by_name: str
    last_editor: str
    sheets: str
    sheet_count: int
    size: str
    live: int
    live_people: List[Dict[str, str]]


class CreatorOption(TypedDict):
    key: str
    label: str
    sublabel: str
    count: int


def _size(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    if n < 1024 * 1024:
        return f"{n / 1024:.0f} KB"
    return f"{n / 1024 / 1024:.1f} MB"


class ManagerState(rx.State):
    # The filtered, sorted view shown in the table.
    books: List[BookItem] = []
    total_count: int = 0
    live_count: int = 0
    creator_options: List[CreatorOption] = []
    my_email: str = ""

    # Filters
    search: str = ""
    creator_filter: str = "all"  # all | mine | guest | acct:<email>
    time_filter: str = "any"  # any | today | 7d | 30d | older
    sort_by: str = "updated_desc"  # updated_desc | updated_asc | title_asc | title_desc
    only_live: bool = False

    _all: List[BookItem] = []

    @rx.var
    def has_active_filters(self) -> bool:
        return bool(self.search.strip() or self.creator_filter != "all" or self.time_filter != "any" or self.only_live)

    @staticmethod
    def _creator_key(b: BookItem) -> str:
        email = b["created_by_email"]
        return f"acct:{email}" if email else "guest"

    async def _creator(self) -> Dict[str, str]:
        if AuthState is None:
            return {}
        auth = await self.get_state(AuthState)
        if not auth._auth_loaded:
            await auth.load_auth()
        if auth.is_logged_in:
            return {"email": auth.user_email, "name": auth.user_name or auth.user_email}
        return {}

    @rx.event
    async def load(self):
        """Reads every workbook, with who is on each, then applies the filters."""
        self.my_email = (await self._creator()).get("email", "")
        self._read()

    @rx.event
    def refresh(self):
        """Periodic: who is editing what changes all the time."""
        self._read()

    def _read(self) -> None:
        live = STORE.viewers()
        items: List[BookItem] = []
        for m in STORE.list():
            people = live.get(m.id, [])
            items.append(
                BookItem(
                    id=m.id,
                    title=m.title,
                    updated_at=m.updated_at,
                    when=time.strftime("%Y-%m-%d %H:%M", time.localtime(m.updated_at)) if m.updated_at else "",
                    created_by_email=m.created_by_email,
                    created_by_name=m.created_by_name,
                    last_editor=m.last_editor,
                    sheets=", ".join(m.sheets[:6]) + (" …" if len(m.sheets) > 6 else ""),
                    sheet_count=len(m.sheets),
                    size=_size(m.size) if m.size else "",
                    live=len(people),
                    live_people=[{"name": p["name"], "color": p["color"]} for p in people[:6]],
                )
            )
        self._all = items
        self.total_count = len(items)
        self.live_count = sum(1 for b in items if b["live"])

        counts: Dict[str, int] = {}
        names: Dict[str, str] = {}
        for b in items:
            k = self._creator_key(b)
            counts[k] = counts.get(k, 0) + 1
            if k != "guest" and b["created_by_name"]:
                names.setdefault(k, b["created_by_name"])
        options = [CreatorOption(key="all", label="All", sublabel="", count=len(items))]
        if self.my_email:
            options.append(CreatorOption(key="mine", label="Mine", sublabel=self.my_email, count=counts.get(f"acct:{self.my_email}", 0)))
        if "guest" in counts:
            options.append(CreatorOption(key="guest", label="Guest", sublabel="Not signed in", count=counts["guest"]))
        for k in sorted((k for k in counts if k != "guest"), key=lambda k: (-counts[k], names.get(k, k).lower())):
            email = k.removeprefix("acct:")
            options.append(CreatorOption(key=k, label=names.get(k) or email, sublabel=email, count=counts[k]))
        self.creator_options = options
        if self.creator_filter not in {o["key"] for o in options}:
            self.creator_filter = "all"
        self._apply_filters()

    def _apply_filters(self) -> None:
        now = time.time()
        day = 86400
        start_of_today = time.mktime(time.localtime(now)[:3] + (0, 0, 0, 0, 0, -1))
        query = self.search.strip().lower()
        out: List[BookItem] = []
        for b in self._all:
            k = self._creator_key(b)
            if self.creator_filter == "mine":
                if k != f"acct:{self.my_email}":
                    continue
            elif self.creator_filter != "all" and k != self.creator_filter:
                continue
            ts = b["updated_at"]
            if self.time_filter == "today" and ts < start_of_today:
                continue
            if self.time_filter == "7d" and ts < now - 7 * day:
                continue
            if self.time_filter == "30d" and ts < now - 30 * day:
                continue
            if self.time_filter == "older" and ts >= now - 30 * day:
                continue
            if self.only_live and not b["live"]:
                continue
            if query:
                hay = " ".join((b["title"], b["id"], b["created_by_name"], b["created_by_email"], b["sheets"], b["last_editor"])).lower()
                if query not in hay:
                    continue
            out.append(b)
        if self.sort_by == "updated_asc":
            out.sort(key=lambda b: b["updated_at"])
        elif self.sort_by == "title_asc":
            out.sort(key=lambda b: b["title"].lower())
        elif self.sort_by == "title_desc":
            out.sort(key=lambda b: b["title"].lower(), reverse=True)
        else:
            out.sort(key=lambda b: b["updated_at"], reverse=True)
        self.books = out

    @rx.event
    def set_search(self, value: str):
        self.search = value
        self._apply_filters()

    @rx.event
    def set_creator_filter(self, value: str):
        self.creator_filter = value
        self._apply_filters()

    @rx.event
    def set_time_filter(self, value: str):
        self.time_filter = value
        self._apply_filters()

    @rx.event
    def set_sort_by(self, value: str):
        self.sort_by = value
        self._apply_filters()

    @rx.event
    def toggle_only_live(self):
        self.only_live = not self.only_live
        self._apply_filters()

    @rx.event
    def filter_by_creator_of(self, email: str):
        """Row shortcut: show every workbook from the same creator."""
        self.creator_filter = f"acct:{email}" if email else "guest"
        self._apply_filters()

    @rx.event
    def clear_search(self):
        self.search = ""
        self._apply_filters()
        # The search box is uncontrolled, so clear its DOM value explicitly.
        return rx.set_value(SEARCH_INPUT_ID, "")

    @rx.event
    def reset_filters(self):
        self.search = ""
        self.creator_filter = "all"
        self.time_filter = "any"
        self.only_live = False
        self._apply_filters()
        return rx.set_value(SEARCH_INPUT_ID, "")

    # ------------------------------------------------------------ actions

    @rx.event
    async def new_book(self, sample: str = ""):
        creator = await self._creator()
        try:
            meta = STORE.create_sample(sample, creator) if sample else STORE.create_blank(creator)
        except EngineError as ex:
            return rx.toast.error(str(ex), position="bottom-center")
        return rx.redirect(f"/book/{meta.id}")

    @rx.event
    async def upload(self, files: List[rx.UploadFile]):
        """Opens an uploaded .xlsx/.xlsm/.csv as a new shared workbook."""
        creator = await self._creator()
        for f in files[:1]:
            data = await f.read()
            name = os.path.basename(f.name or f.filename or "Book.xlsx")
            try:
                meta, warnings = STORE.create_from_file(data, name, creator)
            except EngineError as ex:
                return rx.toast.error(str(ex), position="bottom-center")
            return [
                rx.redirect(f"/book/{meta.id}"),
                rx.toast.success(f"Opened {name}" + (f" ({len(warnings)} warnings)" if warnings else ""), position="bottom-center"),
            ]
        return None

    @rx.event
    def download(self, book_id: str):
        meta = STORE.meta(book_id)
        found = STORE.file_bytes(book_id)
        if meta is None or found is None:
            return rx.toast.error("That workbook isn't there any more.", position="bottom-center")
        data, ext = found
        stem = "".join("-" if ch in '\\/:*?"<>|' else ch for ch in meta.title).strip() or "Book"
        return rx.download(data=data, filename=f"{stem}.{ext}", mime_type=MIMES[ext])

    @rx.event
    def delete_book(self, book_id: str):
        """Deletes a workbook (anyone still on it is told it is gone)."""
        STORE.delete(book_id)
        _tell(book_id)
        self._read()

    @rx.event
    def delete_filtered(self):
        """Deletes every workbook in the current filtered view."""
        for b in list(self.books):
            STORE.delete(b["id"])
            _tell(b["id"])
        self._read()
