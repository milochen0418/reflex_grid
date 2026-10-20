"""Shared workbooks ("books"): stored as .xlsx files, opened once per server process, edited by
everyone viewing them at the same time.

Each book is ``<home>/books/<id>.xlsx`` (the workbook, as Excel reads it; ``<id>.xlsm`` when it
has macros) plus ``<id>.json``
(title, creator, times, sheet names). ``<home>`` is ``REFLEX_GRID_HOME``, by default
``~/Library/Application Support/reflex_grid`` on macOS and ``$XDG_DATA_HOME/reflex_grid`` (or
``~/.local/share/reflex_grid``) elsewhere: outside the project, so saving never triggers a hot
reload.

Installed from re_ddns's App Store, the app is given the shared collection ``spreadsheets``
(re_ddns's Owned Collections) and each book is an item of it instead:
``<collection>/by-app/<app>/<id>/book.xlsx`` (``book.xlsm`` with macros), with the title,
creator and sheet names in the item's manifest. See :class:`ItemBookStore`.

An open book is a :class:`Book` (a :class:`gridcore.engine.SharedBook`): everyone on it edits
the same workbook in memory, and changes are written back ``SAVE_DELAY`` seconds after they
happen, in a background thread.
"""

from __future__ import annotations

import asyncio
import atexit
import errno
import json
import os
import re
import secrets
import sys
import threading
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Awaitable, Callable, Dict, List, Optional, Tuple

from gridcore.engine import DocState, EngineError, Session, SharedBook, open_bytes
from gridcore.engine.samples import SAMPLES, new_sample
from gridcore.model.workbook import Workbook
from gridcore.xlsx import write_xlsx

try:  # on re_ddns: the books as items of a shared collection (ItemBookStore)
    from reflex_ddns_auth.collections import Collection, ItemNotFound
except ImportError:
    Collection = ItemNotFound = None

SAVE_DELAY = 1.5
# Books nobody has had open for this long are saved and dropped from memory.
IDLE_CLOSE = 15 * 60
DEFAULT_TITLE = "Untitled workbook"
MAX_TITLE = 120
# The App Store gives the app this collection (its catalog entry declares it, and the install sets
# DDNS_COLLECTIONS_ROOT); every book is then an item of it, of this kind.
COLLECTION = "spreadsheets"
KIND = "spreadsheet.workbook"

# Starts with a letter or digit: a book's id is also its item's id in a collection.
_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9-]{0,63}$")
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


class BookGone(EngineError):
    """The stored book was deleted outside the app (on re_ddns's storage page)."""


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
    """The books on disk and the ones open in this process. How books are stored is up to the
    methods under "files"; :class:`ItemBookStore` keeps them in a shared collection instead."""

    def __init__(self, home: Optional[Path] = None):
        self.home = Path(home or os.environ.get("REFLEX_GRID_HOME") or default_home())
        self.dir = self.home / "books"
        self._lock = threading.RLock()
        self._open: Dict[str, Book] = {}
        # Called (on the server's event loop) with a book's id after it was saved.
        self.on_saved: Optional[Callable[[str], Awaitable[Any]]] = None

    # ------------------------------------------------------------ files

    def _xlsx(self, book_id: str, ext: str = "xlsx") -> Path:
        return self.dir / f"{book_id}.{ext}"

    def _file(self, book_id: str) -> Path:
        """The stored workbook: ``<id>.xlsm`` for a workbook with macros, else ``<id>.xlsx``."""
        xlsm = self._xlsx(book_id, "xlsm")
        return xlsm if xlsm.exists() else self._xlsx(book_id)

    def _store(self, book_id: str, wb: Workbook, data: bytes) -> None:
        """Writes the workbook file (as .xlsm when it has macros) and drops the other one."""
        ext = file_extension(wb)
        _atomic_write(self._xlsx(book_id, ext), data)
        other = self._xlsx(book_id, "xlsx" if ext == "xlsm" else "xlsm")
        try:
            other.unlink()
        except OSError:
            pass

    def _json(self, book_id: str) -> Path:
        return self.dir / f"{book_id}.json"

    def _write_meta(self, meta: BookMeta) -> None:
        self.dir.mkdir(parents=True, exist_ok=True)
        _atomic_write(self._json(meta.id), json.dumps(asdict(meta), ensure_ascii=False, indent=1).encode("utf-8"))

    def _read_meta(self, book_id: str) -> Optional[BookMeta]:
        try:
            return BookMeta.from_json(json.loads(self._json(book_id).read_text("utf-8")))
        except (OSError, ValueError, TypeError):
            return None

    def _stored(self) -> List[BookMeta]:
        """The metadata of every stored book."""
        out = []
        try:
            paths = list(self.dir.glob("*.json"))
        except OSError:
            paths = []
        for p in paths:
            try:
                out.append(BookMeta.from_json(json.loads(p.read_text("utf-8"))))
            except (OSError, ValueError, TypeError):
                continue
        return out

    def _read(self, book_id: str) -> Tuple[bytes, str]:
        """The stored workbook file and its extension (OSError when it is missing)."""
        path = self._file(book_id)
        return path.read_bytes(), path.suffix[1:]

    def _write(self, meta: BookMeta, wb: Workbook, data: bytes, new: bool = False) -> None:
        """Stores a book's workbook file and metadata (``new``: a book not stored before)."""
        self.dir.mkdir(parents=True, exist_ok=True)
        self._store(meta.id, wb, data)
        self._write_meta(meta)

    def _remove(self, book_id: str) -> None:
        for p in (self._xlsx(book_id), self._xlsx(book_id, "xlsm"), self._json(book_id)):
            try:
                p.unlink()
            except OSError:
                pass

    def meta(self, book_id: str) -> Optional[BookMeta]:
        if not valid_id(book_id):
            return None
        with self._lock:
            book = self._open.get(book_id)
            if book is not None:
                return book.meta
        return self._read_meta(book_id)

    def exists(self, book_id: str) -> bool:
        return valid_id(book_id) and self._json(book_id).exists()

    def list(self) -> List[BookMeta]:
        """Every book, most recently changed first (open ones as they are in memory)."""
        out = {m.id: m for m in self._stored() if valid_id(m.id)}
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
        self._write(meta, doc.wb, data, new=True)
        return meta

    def create_blank(self, creator: Optional[dict] = None, title: str = DEFAULT_TITLE, book_id: str = "") -> BookMeta:
        return self.create(DocState(Workbook(), title), title, creator, book_id)

    def create_sample(self, name: str, creator: Optional[dict] = None) -> BookMeta:
        if name not in SAMPLES:
            raise EngineError(f"no sample named `{name}`")
        d = new_sample(Session(), name)
        return self.create(d, name, creator)

    def create_from_file(self, data: bytes, filename: str, creator: Optional[dict] = None) -> Tuple[BookMeta, List[str]]:
        """A new book from an uploaded .xlsx/.xlsm/.csv/.tsv, with the reader's warnings (raises
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
                    data, ext = self._read(book_id)
                except OSError as e:
                    raise EngineError(f"The workbook file is missing: {e.strerror}")
                doc = open_bytes(Session(), data, f"{meta.title}.{ext}")
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
            try:
                self._write_meta(meta)
            except BookGone:
                self._drop(book_id)
                if book is not None:
                    self._notify(book)
                raise
        return title

    def delete(self, book_id: str) -> None:
        with self._lock:
            self._drop(book_id)
            self._remove(book_id)

    def _drop(self, book_id: str) -> Optional[Book]:
        """Takes a deleted book out of memory for good."""
        with self._lock:
            book = self._open.pop(book_id, None)
            if book is not None:
                with book.lock:
                    book.deleted = True
                    if book._timer is not None:
                        book._timer.cancel()
                        book._timer = None
        return book

    def file_bytes(self, book_id: str) -> Optional[Tuple[bytes, str]]:
        """The workbook as a file and its extension (xlsx, or xlsm with macros): as it is now
        when open, else as stored."""
        book = self.get(book_id)
        if book is not None:
            with book.lock:
                wb = book.doc.wb
            return write_xlsx(wb), file_extension(wb)
        try:
            return self._read(book_id)
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
                with book.lock:
                    book.meta.size = len(data)
                    book.meta.sheets = [s.name for s in wb.sheets]
                    meta = BookMeta(**asdict(book.meta))
                self._write(meta, wb, data)
        except BookGone:  # as if deleted from the workbook list: everyone on it is told
            self._drop(book.key)
            self._notify(book)
            return False
        except Exception as e:  # keep the book open and editable; say so in the UI
            with book.lock:
                book.save_error = str(e) or e.__class__.__name__
            return False
        with book.lock:
            book.save_error = ""
            book.saved_rev = max(book.saved_rev, rev)
            if book.doc.wb is wb:
                book.doc.saved = wb
        self._notify(book)
        return True

    def _notify(self, book: Book) -> None:
        """Shows everyone on the book how it is now (on the server's event loop)."""
        hook, loop = self.on_saved, book.loop
        if hook is not None and loop is not None and not loop.is_closed():
            try:
                asyncio.run_coroutine_threadsafe(hook(book.key), loop)
            except RuntimeError:
                pass

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


class ItemBookStore(BookStore):
    """The books as items of a shared collection (re_ddns's Owned Collections): book ``<id>`` is the
    item ``spreadsheets:<app>/<id>``, a folder holding ``book.xlsx`` (``book.xlsm`` with macros),
    with the title, creator and sheet names in its manifest. Every app given the collection reads
    the workbooks, and the App Store's storage page lists them. A book is published once and then
    revised in place each time it is saved, so its item always holds the workbook as it is now."""

    def __init__(self, items: Any):
        super().__init__()
        self.items = items
        self.dir = items.own_dir

    def _item(self, book_id: str) -> Any:
        if not valid_id(book_id):
            return None
        item = self.items.get(self.items.ref(book_id), remote=False)
        return item if item is not None and item.kind == KIND else None

    def _read_meta(self, book_id: str) -> Optional[BookMeta]:
        item = self._item(book_id)
        return _book_meta(item) if item is not None else None

    def _stored(self) -> List[BookMeta]:
        # Copies fetched from another host are that host's books, not ours.
        return [_book_meta(i) for i in self.items.items(kind=KIND, remote=False)
                if i.owner == self.items.app_id and not i.replica]

    def exists(self, book_id: str) -> bool:
        return self._item(book_id) is not None

    def _read(self, book_id: str) -> Tuple[bytes, str]:
        item = self._item(book_id)
        if item is None or item.primary is None:
            raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), book_id)
        return item.primary.read_bytes(), item.primary.suffix[1:]

    def _write(self, meta: BookMeta, wb: Workbook, data: bytes, new: bool = False) -> None:
        name = f"book.{file_extension(wb)}"
        fields = dict(title=meta.title, meta=_item_meta(meta), primary=name)
        try:
            if new:
                with self.items.publish(meta.id, kind=KIND, **fields) as item:
                    with item.open(name) as fh:
                        fh.write(data)
            else:
                with self.items.revise(meta.id, **fields) as rev:
                    with rev.open(name) as fh:
                        fh.write(data)
                    rev.remove("book.xlsm" if name == "book.xlsx" else "book.xlsx")
        except ItemNotFound:
            raise BookGone("That workbook doesn't exist any more.")

    def _write_meta(self, meta: BookMeta) -> None:
        try:
            with self.items.revise(meta.id, title=meta.title, meta=_item_meta(meta)):
                pass
        except ItemNotFound:
            raise BookGone("That workbook doesn't exist any more.")

    def _remove(self, book_id: str) -> None:
        self.items.remove(book_id)


def _item_meta(meta: BookMeta) -> dict:
    """What an item's manifest keeps of a book's metadata in its ``meta`` (the title, the times
    and the size are the manifest's own)."""
    return {k: v for k, v in asdict(meta).items() if k not in ("id", "title", "created_at", "updated_at", "size")}


def _book_meta(item: Any) -> BookMeta:
    meta = BookMeta.from_json({**item.meta, "id": item.id, "title": item.title})
    meta.created_at = _timestamp(item.created_at)
    meta.updated_at = _timestamp(item.updated_at)
    meta.size = item.bytes
    return meta


def _timestamp(iso: str) -> float:
    try:
        return datetime.fromisoformat(iso).timestamp()
    except (TypeError, ValueError):
        return 0.0


def file_extension(wb: Workbook) -> str:
    """xlsm for a workbook with a VBA project (Excel won't open macros in an .xlsx), else xlsx."""
    return "xlsm" if getattr(wb, "vba", None) is not None else "xlsx"


def _atomic_write(path: Path, data: bytes) -> None:
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_bytes(data)
    os.replace(tmp, path)


def _make_store() -> BookStore:
    """The books as items of the shared collection when re_ddns gave the app one, else as files in
    ``REFLEX_GRID_HOME``."""
    if os.environ.get("DDNS_COLLECTIONS_ROOT") and Collection is not None:
        if hasattr(Collection, "revise"):
            return ItemBookStore(Collection(COLLECTION))
        print("ReflexGrid: this reflex_ddns_auth can't revise items; workbooks stay in REFLEX_GRID_HOME", file=sys.stderr)
    return BookStore()


STORE = _make_store()
atexit.register(STORE.flush)
