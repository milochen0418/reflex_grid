"""Shared workbooks ("books"): stored as .xlsx files, opened once per server process, edited by
everyone viewing them at the same time.

Each book is ``<home>/books/<id>.xlsx`` (the workbook, as Excel reads it) plus ``<id>.json``
(title, creator, times, sheet names). ``<home>`` is ``REFLEX_GRID_HOME``, by default
``~/Library/Application Support/reflex_grid`` on macOS and ``$XDG_DATA_HOME/reflex_grid`` (or
``~/.local/share/reflex_grid``) elsewhere: outside the project, so saving never triggers a hot
reload.

An open book is a :class:`Book` (a :class:`gridcore.engine.SharedBook`): everyone on it edits
the same workbook in memory, and changes are written back ``SAVE_DELAY`` seconds after they
happen, in a background thread.
"""

from __future__ import annotations

import asyncio
import atexit
import json
import os
import re
import secrets
import sys
import threading
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Awaitable, Callable, Dict, List, Optional, Tuple

from gridcore.engine import DocState, EngineError, Session, SharedBook, open_bytes
from gridcore.engine.samples import SAMPLES, new_sample
from gridcore.model.workbook import Workbook
from gridcore.xlsx import write_xlsx

SAVE_DELAY = 1.5
# Books nobody has had open for this long are saved and dropped from memory.
IDLE_CLOSE = 15 * 60
DEFAULT_TITLE = "Untitled workbook"
MAX_TITLE = 120

_ID_RE = re.compile(r"^[A-Za-z0-9-]{1,64}$")
# Selection colours of the people on a book (readable on white, distinct from each other);
# blue last: your own selection is blue.
COLORS = ["#E8710A", "#188038", "#9334E6", "#D93025", "#E52592", "#12B5CB", "#B06000", "#3949AB", "#00897B", "#1A73E8"]


def default_home() -> Path:
    if sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share")
    return base / "reflex_grid"


def valid_id(book_id: str) -> bool:
    return bool(book_id) and _ID_RE.match(book_id) is not None


def new_id() -> str:
    return secrets.token_hex(5)


def clean_title(title: str) -> str:
    title = " ".join(str(title or "").split())[:MAX_TITLE]
    return title or DEFAULT_TITLE


@dataclass
class BookMeta:
    id: str
    title: str = DEFAULT_TITLE
    created_at: float = 0.0
    updated_at: float = 0.0
    created_by_email: str = ""
    created_by_name: str = ""
    last_editor: str = ""
    sheets: List[str] = field(default_factory=list)
    size: int = 0

    @staticmethod
    def from_json(d: dict) -> "BookMeta":
        known = {k: d[k] for k in BookMeta.__dataclass_fields__ if k in d}
        return BookMeta(**known)


class Book(SharedBook):
    """A shared workbook and its place in the store."""

    def __init__(self, doc: DocState, meta: BookMeta):
        super().__init__(doc, meta.id)
        self.meta = meta
        self.saved_rev = 0
        self.save_error = ""
        self.deleted = False
        self.closed = False
        self.last_active = time.time()
        self._timer: Optional[threading.Timer] = None
        # The server's event loop, to tell the UI from the saving thread.
        self.loop: Optional[asyncio.AbstractEventLoop] = None

    @property
    def dirty(self) -> bool:
        return self.rev != self.saved_rev

    def color_for(self, user: str) -> str:
        """The person's colour on this book: the one their other tabs use, else a free one."""
        with self.lock:
            used = {}
            for v in self.views.values():
                if v.info.get("color"):
                    used.setdefault(v.info.get("user"), v.info["color"])
            if user in used:
                return used[user]
            taken = set(used.values())
            for c in COLORS:
                if c not in taken:
                    return c
            return COLORS[len(used) % len(COLORS)]


class BookStore:
    """The books on disk and the ones open in this process."""

    def __init__(self, home: Optional[Path] = None):
        self.home = Path(home or os.environ.get("REFLEX_GRID_HOME") or default_home())
        self.dir = self.home / "books"
        self._lock = threading.RLock()
        self._open: Dict[str, Book] = {}
        # Called (on the server's event loop) with a book's id after it was saved.
        self.on_saved: Optional[Callable[[str], Awaitable[Any]]] = None

    # ------------------------------------------------------------ files

    def _xlsx(self, book_id: str) -> Path:
        return self.dir / f"{book_id}.xlsx"

    def _json(self, book_id: str) -> Path:
        return self.dir / f"{book_id}.json"

    def _write_meta(self, meta: BookMeta) -> None:
        self.dir.mkdir(parents=True, exist_ok=True)
        _atomic_write(self._json(meta.id), json.dumps(asdict(meta), ensure_ascii=False, indent=1).encode("utf-8"))

    def meta(self, book_id: str) -> Optional[BookMeta]:
        if not valid_id(book_id):
            return None
        with self._lock:
            book = self._open.get(book_id)
            if book is not None:
                return book.meta
        try:
            return BookMeta.from_json(json.loads(self._json(book_id).read_text("utf-8")))
        except (OSError, ValueError, TypeError):
            return None

    def exists(self, book_id: str) -> bool:
        return valid_id(book_id) and self._json(book_id).exists()

    def list(self) -> List[BookMeta]:
        """Every book, most recently changed first (open ones as they are in memory)."""
        out: Dict[str, BookMeta] = {}
        try:
            paths = list(self.dir.glob("*.json"))
        except OSError:
            paths = []
        for p in paths:
            try:
                m = BookMeta.from_json(json.loads(p.read_text("utf-8")))
            except (OSError, ValueError, TypeError):
                continue
            if valid_id(m.id):
                out[m.id] = m
        with self._lock:
            for book_id, book in self._open.items():
                if not book.deleted and book_id in out:
                    out[book_id] = book.meta
        return sorted(out.values(), key=lambda m: m.updated_at, reverse=True)

    def viewers(self) -> Dict[str, List[dict]]:
        """People on each open book: ``{book id: [{"user", "name", "color", "avatar"}]}``, one per person."""
        out: Dict[str, List[dict]] = {}
        with self._lock:
            books = list(self._open.items())
        for book_id, book in books:
            with book.lock:
                people: Dict[str, dict] = {}
                for v in book.views.values():
                    u = v.info.get("user") or v.key
                    people.setdefault(u, {"user": u, "name": v.name, "color": v.info.get("color", ""), "avatar": v.info.get("avatar", "")})
            if people:
                out[book_id] = list(people.values())
        return out

    # ------------------------------------------------------------ creating

    def create(self, doc: DocState, title: str, creator: Optional[dict] = None, book_id: str = "") -> BookMeta:
        """Stores a new book holding ``doc``'s workbook and returns its metadata."""
        creator = creator or {}
        book_id = book_id or new_id()
        now = time.time()
        meta = BookMeta(
            id=book_id,
            title=clean_title(title),
            created_at=now,
            updated_at=now,
            created_by_email=creator.get("email", ""),
            created_by_name=creator.get("name", "") if creator.get("email") else "Guest",
            sheets=[s.name for s in doc.wb.sheets],
        )
        data = write_xlsx(doc.wb)
        meta.size = len(data)
        self.dir.mkdir(parents=True, exist_ok=True)
        _atomic_write(self._xlsx(book_id), data)
        self._write_meta(meta)
        return meta

    def create_blank(self, creator: Optional[dict] = None, title: str = DEFAULT_TITLE, book_id: str = "") -> BookMeta:
        return self.create(DocState(Workbook(), title), title, creator, book_id)

    def create_sample(self, name: str, creator: Optional[dict] = None) -> BookMeta:
        if name not in SAMPLES:
            raise EngineError(f"no sample named `{name}`")
        d = new_sample(Session(), name)
        return self.create(d, name, creator)

    def create_from_file(self, data: bytes, filename: str, creator: Optional[dict] = None) -> Tuple[BookMeta, List[str]]:
        """A new book from an uploaded .xlsx/.csv/.tsv, with the reader's warnings (raises
        EngineError when the file can't be read)."""
        d = open_bytes(Session(), data, os.path.basename(filename or "Book.xlsx"))
        stem = os.path.splitext(os.path.basename(filename or ""))[0]
        return self.create(d, stem or DEFAULT_TITLE, creator), list(d.warnings)

    # ------------------------------------------------------------ opening

    def open(self, book_id: str, creator: Optional[dict] = None) -> Book:
        """The open book ``book_id``, read from disk, or new and blank when there is none yet
        (a link to a book that doesn't exist creates it, as in CoDoc)."""
        if not valid_id(book_id):
            raise EngineError("That isn't a workbook address.")
        with self._lock:
            book = self._open.get(book_id)
            if book is not None and not book.deleted:
                book.last_active = time.time()
                return book
            meta = self.meta(book_id)
            if meta is None:
                meta = self.create_blank(creator, book_id=book_id)
                doc = DocState(Workbook(), meta.title)
            else:
                try:
                    data = self._xlsx(book_id).read_bytes()
                except OSError as e:
                    raise EngineError(f"The workbook file is missing: {e.strerror}")
                doc = open_bytes(Session(), data, f"{meta.title}.xlsx")
                doc.name = meta.title
            doc.saved = doc.wb
            book = Book(doc, meta)
            book.listeners.append(self._changed)
            self._open[book_id] = book
            return book

    def get(self, book_id: str) -> Optional[Book]:
        with self._lock:
            book = self._open.get(book_id)
            return None if book is None or book.deleted else book

    # ------------------------------------------------------------ changing

    def rename(self, book_id: str, title: str) -> str:
        title = clean_title(title)
        book = self.get(book_id)
        if book is not None:
            # The title isn't part of the workbook: everyone sees it through the book's room.
            with book.lock:
                book.meta.title = title
                book.doc.name = title
                meta = book.meta
        else:
            meta = self.meta(book_id)
            if meta is None:
                raise EngineError("That workbook doesn't exist any more.")
            meta.title = title
        meta.updated_at = time.time()
        with self._lock:
            self._write_meta(meta)
        return title

    def delete(self, book_id: str) -> None:
        with self._lock:
            book = self._open.pop(book_id, None)
            if book is not None:
                with book.lock:
                    book.deleted = True
                    if book._timer is not None:
                        book._timer.cancel()
                        book._timer = None
            for p in (self._xlsx(book_id), self._json(book_id)):
                try:
                    p.unlink()
                except OSError:
                    pass

    def xlsx_bytes(self, book_id: str) -> Optional[bytes]:
        """The workbook as an .xlsx file: as it is now when open, else as stored."""
        book = self.get(book_id)
        if book is not None:
            with book.lock:
                wb = book.doc.wb
            return write_xlsx(wb)
        try:
            return self._xlsx(book_id).read_bytes()
        except OSError:
            return None

    # ------------------------------------------------------------ saving

    def _changed(self, book: Book) -> None:
        """Listener on every open book: note the change and save shortly."""
        try:
            book.loop = asyncio.get_running_loop()
        except RuntimeError:
            pass
        book.last_active = time.time()
        book.meta.updated_at = book.last_active
        if book.changes:
            v = book.views.get(book.changes[-1].author)
            if v is not None and v.name:
                book.meta.last_editor = v.name
        if book._timer is None and not book.deleted:
            t = threading.Timer(SAVE_DELAY, self.save, (book,))
            t.daemon = True
            book._timer = t
            t.start()

    def save(self, book: Book) -> bool:
        """Writes the book's workbook now (the snapshot is read under the book's lock, the file
        is written outside it). Returns whether it was saved."""
        with book.lock:
            book._timer = None
            if book.deleted:
                return False
            wb, rev = book.doc.wb, book.rev
        try:
            data = write_xlsx(wb)
            with self._lock:
                if book.deleted:
                    return False
                _atomic_write(self._xlsx(book.key), data)
                with book.lock:
                    book.meta.size = len(data)
                    book.meta.sheets = [s.name for s in wb.sheets]
                    meta = BookMeta(**asdict(book.meta))
                self._write_meta(meta)
        except Exception as e:  # keep the book open and editable; say so in the UI
            with book.lock:
                book.save_error = str(e) or e.__class__.__name__
            return False
        with book.lock:
            book.save_error = ""
            book.saved_rev = max(book.saved_rev, rev)
            if book.doc.wb is wb:
                book.doc.saved = wb
        hook, loop = self.on_saved, book.loop
        if hook is not None and loop is not None and not loop.is_closed():
            try:
                asyncio.run_coroutine_threadsafe(hook(book.key), loop)
            except RuntimeError:
                pass
        return True

    def flush(self) -> None:
        """Saves every open book with unsaved changes (at exit)."""
        with self._lock:
            books = list(self._open.values())
        for book in books:
            if book.dirty:
                self.save(book)

    def close_idle(self, now: Optional[float] = None) -> List[str]:
        """Saves and drops the books nobody has had open for ``IDLE_CLOSE`` seconds."""
        now = time.time() if now is None else now
        closed = []
        with self._lock:
            candidates = [b for b in self._open.values() if not b.views and now - b.last_active > IDLE_CLOSE]
        for book in candidates:
            if book.dirty and not self.save(book):
                continue
            with self._lock:
                with book.lock:
                    if book.views or book.dirty:
                        continue
                    book.closed = True
                    self._open.pop(book.key, None)
                    closed.append(book.key)
        return closed


def _atomic_write(path: Path, data: bytes) -> None:
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_bytes(data)
    os.replace(tmp, path)


STORE = BookStore()
atexit.register(STORE.flush)
