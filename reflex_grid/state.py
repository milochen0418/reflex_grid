"""Reflex state: the bridge between the browser (grid component, ribbon, dialogs) and the
gridcore engine.

Each tab shows one shared book (``/book/<id>``): the workbook lives once on the server
(:mod:`.books`) and everyone on it edits it at the same time. ``GridState`` holds what this tab
draws — its window of cells, its selection, ribbon and dialogs — and runs the tab's commands
on the shared book through the tab's view of it (``book.act(view)``). ``BookRoom``, a Reflex
SharedState linked to the book, tells every tab on it about each change and shows who is where;
a tab that drew an older revision than the room's asks for its window again (``sync``)."""

from __future__ import annotations

import base64
import hashlib
import os
import random
import re
import secrets
import time
from typing import Any, Dict, List, Optional

import reflex as rx

from gridcore.core.addr import MAX_COLS, MAX_ROWS, RangeRef, a1, col_to_letters
from gridcore.engine import EngineError, Session, View, geometry, objects, ui_state, viewport
from gridcore.engine.cmd_file import MIMES, encode, workbook_format
from gridcore.engine.cmd_format import CELL_STYLE_NAMES, FORMAT_PRESETS  # noqa: F401
from gridcore.engine.samples import SAMPLES
from gridcore.functions import all_functions
from gridcore.numfmt import parse as parse_fmt

from . import call, sessions
from .books import STORE, Book, clean_title, valid_id
from .room import BookRoom, room_token

try:
    from reflex_ddns_auth import AuthState
except ImportError:  # runs without sign-in: everyone is a guest
    AuthState = None

FUNCTIONS = sorted(
    ({"name": f.name, "signature": f.signature, "description": f.description, "category": f.category} for f in all_functions() if not f.name.startswith("_")),
    key=lambda x: x["name"],
)
CATEGORIES = sorted({f["category"] for f in FUNCTIONS})
# A tab that hasn't pinged for this long has left (it pings every 5 s; hidden tabs are throttled).
PEER_IDLE = 30
# Events that only redraw this tab: the server's answer never moves its selection.
_PASSIVE = frozenset({"viewport", "sync", "ping", "presence", "objectSelect"})


# The event the VBA editor opens when a control is double-clicked in the form view, and the
# arguments each event procedure takes.
_DEFAULT_EVENTS = {
    "TextBox": "Change", "ComboBox": "Change", "ScrollBar": "Change", "SpinButton": "Change",
    "MultiPage": "Change", "TabStrip": "Change",
}
_EVENT_ARGS = {
    "BeforeUpdate": "ByVal Cancel As MSForms.ReturnBoolean",
    "Exit": "ByVal Cancel As MSForms.ReturnBoolean",
    "DblClick": "ByVal Cancel As MSForms.ReturnBoolean",
    "KeyPress": "ByVal KeyAscii As MSForms.ReturnInteger",
    "QueryClose": "Cancel As Integer, CloseMode As Integer",
}


def _rng(g) -> str:
    r0, c0, r1, c1 = (int(x) for x in g)
    return RangeRef(r0, c0, r1, c1).a1()


def _num(v, default=None):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _file_stem(name: str) -> str:
    stem = re.sub(r"\.(xlsx|xlsm|csv|tsv|txt)$", "", name or "", flags=re.I)
    stem = re.sub(r'[\\/:*?"<>|]+', "-", stem).strip()
    return stem or "Book"


def _peer_json(v: View) -> Dict[str, Any]:
    sel = v.selection
    return {
        "id": v.key,
        "user": v.info.get("user", v.key),
        "name": v.name,
        "color": v.info.get("color", "#5F6368"),
        "avatar": v.info.get("avatar", ""),
        "sheet": v.sheet,
        "active": [sel.active[0], sel.active[1]],
        "ranges": [[r.r0, r.c0, r.r1, r.c1] for r in sel.ranges[-8:]],
        "editing": v.info.get("editing"),
    }


class GridState(rx.State):
    # Grid component props.
    geometry: Dict[str, Any] = {}
    data: Dict[str, Any] = {}
    objects: Dict[str, Any] = {}
    ui: Dict[str, Any] = {}
    functions: List[Dict[str, str]] = []
    names: List[str] = []
    focus_key: int = 0
    edit_request: Dict[str, Any] = {}
    window: List[int] = [0, 0, 80, 30]
    # Chrome.
    tab: str = "Home"
    backstage: bool = False
    selected_object: int = -1
    selected_kind: str = ""
    # Dialogs: which one is open and its fields.
    dialog: str = ""
    dlg: Dict[str, Any] = {}
    dlg_error: str = ""
    # The book on screen (``/book/<id>``), who this tab is on it, and whether the book is gone.
    book: str = ""
    me: Dict[str, str] = {}
    book_gone: bool = False
    # Other books, for the File backstage.
    recent: List[Dict[str, str]] = []
    # Macros: the VBA editor (its module list, the module on screen, its code and problems,
    # the Immediate window) and a running macro's dialog (MsgBox/InputBox) or progress.
    vba_open: bool = False
    vba_modules: List[Dict[str, Any]] = []
    vba_macros: List[Dict[str, Any]] = []
    vba_module: str = ""
    vba_kind: str = ""
    vba_code: str = ""
    vba_key: int = 0
    vba_version: str = ""
    vba_dirty: bool = False
    vba_locked: bool = False
    vba_diag: List[Dict[str, Any]] = []
    vba_error_line: int = 0
    vba_status: str = ""
    vba_output: List[str] = []
    vba_immediate: str = ""
    macro_prompt: Dict[str, Any] = {}
    macro_input: str = ""
    macro_busy: bool = False
    macro_label: str = ""
    # UserForms: the person's forms on screen (the engine's snapshots), the modal one a macro
    # waits on, and the editor's form view (the design of the form module on screen).
    vba_forms: List[Dict[str, Any]] = []
    vba_form_pictures: Dict[str, str] = {}
    vba_form_waiting: int = 0
    vba_form_view: bool = False
    # The form designer: the form module's design (a snapshot), the control selected ("" is the
    # form), the toolbox tool in hand, the Properties window's rows, whether it can be edited.
    vba_form_snapshot: Dict[str, Any] = {}
    vba_form_selected: str = ""
    vba_form_tool: str = ""
    vba_form_props: List[Dict[str, Any]] = []
    vba_form_editable: bool = False
    vba_form_warnings: List[str] = []
    vba_goto_line: int = 0
    vba_goto_key: int = 0
    vba_reopen: bool = False
    _peer: str = ""
    _guest: str = ""

    # ------------------------------------------------------------ computed

    @rx.var
    def ribbon(self) -> Dict[str, Any]:
        return self.ui.get("ribbon", {}) if self.ui else {}

    @rx.var
    def sheet_tabs(self) -> List[Dict[str, Any]]:
        out = []
        for i, s in enumerate(self.ui.get("sheets", []) if self.ui else []):
            if s.get("visible", True):
                out.append({"index": i, "name": s["name"], "color": s.get("color") or "", "active": i == self.ui.get("sheet", 0), "protected": s.get("protected", False)})
        return out

    @rx.var
    def hidden_sheets(self) -> List[str]:
        return [s["name"] for s in (self.ui.get("sheets", []) if self.ui else []) if not s.get("visible", True)]

    @rx.var
    def status_text(self) -> str:
        st = (self.ui or {}).get("stats") or {}
        if not st or st.get("count", 0) < 2:
            return ""
        parts = []
        if st.get("numeric"):
            avg = st.get("average")
            if avg is not None:
                parts.append(f"Average: {_fmt_stat(avg)}")
        parts.append(f"Count: {int(st.get('count', 0))}")
        if st.get("numeric"):
            parts.append(f"Sum: {_fmt_stat(st.get('sum', 0))}")
        return "    ".join(parts)

    @rx.var
    def circular_ref(self) -> str:
        c = (self.ui or {}).get("circular") or []
        return c[0] if c else ""

    @rx.var
    def zoom(self) -> int:
        return int((self.geometry or {}).get("zoom", 100))

    @rx.var
    def table_selected(self) -> bool:
        return bool((self.ui or {}).get("table"))

    @rx.var
    def dlg_list(self) -> List[Dict[str, Any]]:
        """The list a dialog shows (filter values, names, rules, find results, columns), indexed."""
        d = self.dlg or {}
        key = {"filter": "values", "nameManager": "names", "cfManage": "rules", "find": "results", "replace": "results", "removeDuplicates": "cols"}.get(self.dialog)
        if key is None:
            return []
        items = [{**x, "i": i} for i, x in enumerate(d.get(key, []) or [])]
        if self.dialog == "filter":
            q = str(d.get("search", "") or "").lower()
            if q:
                items = [x for x in items if q in str(x.get("text", "")).lower()]
        return items[:2000]

    @rx.var
    def fn_list(self) -> List[Dict[str, str]]:
        if self.dialog != "insertFunction":
            return []
        d = self.dlg or {}
        q = str(d.get("search", "") or "").lower()
        cat = d.get("category", "Most Used")
        most = ["SUM", "AVERAGE", "IF", "COUNT", "MAX", "MIN", "XLOOKUP", "VLOOKUP", "SUMIFS", "COUNTIF", "INDEX", "MATCH", "TEXT", "ROUND", "IFERROR", "FILTER", "SORT", "UNIQUE", "LET", "TODAY"]
        out = []
        for f in FUNCTIONS:
            if q and q not in f["name"].lower() and q not in f["description"].lower():
                continue
            if not q and cat == "Most Used" and f["name"] not in most:
                continue
            if not q and cat not in ("Most Used", "All") and f["category"] != cat:
                continue
            out.append(f)
        return out[:400]

    @rx.var
    def fn_selected(self) -> Dict[str, str]:
        if self.dialog != "insertFunction":
            return {}
        name = (self.dlg or {}).get("selected", "SUM")
        shown = self.fn_list
        if shown and not any(f["name"] == name for f in shown):
            return shown[0]
        return next((f for f in FUNCTIONS if f["name"] == name), {"name": name, "signature": "", "description": ""})

    @rx.var
    def dlg_macros(self) -> List[Dict[str, Any]]:
        return list((self.dlg or {}).get("macros", []) or []) if self.dialog == "macros" else []

    @rx.var
    def dlg_macro(self) -> str:
        return str((self.dlg or {}).get("selected", "") or "") if self.dialog == "macros" else ""

    @rx.var
    def dlg_sheets(self) -> List[str]:
        return list((self.dlg or {}).get("sheets", []) or [])

    @rx.var
    def vba_forms_blocked(self) -> bool:
        """A MsgBox, an InputBox or a busy macro is in front of the forms."""
        kind = (self.macro_prompt or {}).get("kind")
        return bool(self.macro_busy or (kind and kind != "form"))

    # ------------------------------------------------------------ plumbing

    def _entry(self) -> sessions.Entry:
        return sessions.get(self.router.session.client_token)

    def _bound(self):
        """This tab's ``(entry, book, view)``, joining the book again when the server lost it (a
        restart, a book closed while idle); ``(entry, None, None)`` without a book."""
        entry = self._entry()
        if not self.book or not self.me:
            return entry, None, None
        book = STORE.get(self.book)
        if book is None:
            try:
                if not STORE.exists(self.book):
                    raise EngineError("deleted")
                book = STORE.open(self.book)
            except EngineError:
                self.book_gone = True
                return entry, None, None
        view = entry.view
        if entry.book_id != book.key or view is None or view.session is not entry.session:
            view = self._join(entry, book)
        return entry, book, view

    def _join(self, entry: sessions.Entry, book: Book) -> View:
        """Puts this tab on ``book`` (leaving the book it was on)."""
        me = self.me
        if entry.view is not None and entry.book_id and entry.book_id != book.key:
            old = STORE.get(entry.book_id)
            if old is not None:
                old.leave(entry.view.key)
        # A tab that reloaded (and so left) comes back with its own view and undo history.
        before = entry.view if entry.book_id == book.key else None
        view = book.join(me["peer"], entry.session, me["name"], view=before)
        view.info.update(user=me["user"], email=me.get("email", ""), avatar=me.get("avatar", ""))
        if not view.info.get("color"):
            view.info["color"] = book.color_for(me["user"])
        entry.book_id, entry.view = book.key, view
        entry.sig, entry.sheet = (), None
        if me.get("color") != view.info["color"]:
            self.me = {**me, "color": view.info["color"]}
        return view

    async def _identify(self) -> None:
        """Who this tab is: the signed-in account (reflex_ddns_auth), else a guest."""
        email = name = avatar = ""
        if AuthState is not None:
            try:
                auth = await self.get_state(AuthState)
                if not auth._auth_loaded:
                    await auth.load_auth()
                if auth.is_logged_in:
                    email, name, avatar = auth.user_email, auth.user_name or auth.user_email, auth.user_avatar
            except Exception:
                pass
        if not self._peer:
            self._peer = secrets.token_hex(6)
        if email:
            user = email
        else:
            if not self._guest:
                self._guest = str(random.randint(1000, 9999))
            user, name = f"guest-{self._guest}", f"Guest {self._guest}"
        me = {"peer": self._peer, "user": user, "name": name, "email": email, "avatar": avatar, "color": (self.me or {}).get("color", "")}
        if me != self.me:
            self.me = me

    def _refresh(self, entry: sessions.Entry, book: Book, view: View, seq: Optional[int] = None, extra_ui: Optional[Dict[str, Any]] = None, passive: bool = False) -> None:
        """Recomputes what this tab draws (inside ``book.act(view)``)."""
        s = entry.session
        d = s.doc()
        wb = d.wb
        si = wb.active_sheet
        sh = wb.sheet(si)
        zoom, formulas = view.zoom_of(sh), view.formulas_of(sh)
        if entry.sheet != sh.name:
            # Another sheet: start at its top-left (not when someone renamed this one).
            renamed = entry.sheet_index == si and entry.sheet is not None and wb.sheet_index(entry.sheet) is None
            if entry.sheet is not None and not renamed:
                self.window = [0, 0, 80, 30]
            entry.sheet = sh.name
        entry.sheet_index = si
        w = self.window or [0, 0, 80, 30]
        vp = viewport(wb, si, RangeRef(int(w[0]), int(w[1]), int(w[2]), int(w[3])), show_formulas=formulas)
        if vp != self.data:
            self.data = vp
        sig = (d.uid, id(wb), si, zoom, formulas)
        if sig != entry.sig:
            entry.sig = sig
            geo = geometry(sh)
            geo["zoom"] = zoom
            geo["formulas"] = formulas
            if geo != self.geometry:
                self.geometry = geo
            ob = objects(wb, si, with_images=False)
            for im in ob["images"]:
                img = next((x for x in sh.images if x.id == im["id"]), None)
                if img is None:
                    continue
                k = (d.uid, img.id, len(img.data))
                src = entry.images.get(k)
                if src is None:
                    src = f"data:{img.mime};base64," + base64.b64encode(img.data).decode("ascii")
                    entry.images[k] = src
                im["src"] = src
            if ob != self.objects:
                self.objects = ob
            names = sorted({n.name for n in wb.names if not n.hidden} | {t.name for x in wb.sheets for t in x.tables}, key=str.lower)
            if names != self.names:
                self.names = names
        ui = ui_state(s)
        ui.update(
            seq=seq if seq is not None else (self.ui or {}).get("seq", 0),
            # A redraw this tab didn't ask for by moving: it keeps its own selection.
            passive=passive,
            docName=d.name,
            dirty=book.dirty,
            docs=[],
            activeDoc=0,
            canUndo=bool(view.undo),
            canRedo=bool(view.redo),
            undoLabel=view.undo[-1].label if view.undo else None,
            redoLabel=view.redo[-1].label if view.redo else None,
            bookRev=book.rev,
            sheetName=sh.name,
        )
        if extra_ui:
            ui.update(extra_ui)
        if ui != self.ui:
            self.ui = ui
        if self.selected_object >= 0 and not any(o["id"] == self.selected_object for o in (self.objects.get("charts", []) + self.objects.get("images", []))):
            self.selected_object = -1
            self.selected_kind = ""

    def _publish(self, room: BookRoom, book: Book) -> None:
        """Shows the book as it is now to everyone on it: revision, title, who is where, recent
        changes, saving and the call. Only what changed is assigned, so only that is sent."""
        with book.lock:
            rev, title = book.rev, book.meta.title
            peers = [_peer_json(v) for v in book.views.values()]
            who = {v.key: (v.name, v.info.get("color", "#5F6368")) for v in book.views.values()}
            changes = [
                {"rev": c.rev, "author": c.author, "name": who.get(c.author, ("", "#5F6368"))[0], "color": who.get(c.author, ("", "#5F6368"))[1], "areas": [[n, list(r)] for n, r in c.areas]}
                for c in book.changes[-12:]
            ]
            saved, error = not book.dirty, book.save_error
        if room.rev != rev:
            room.rev = rev
        if room.title != title:
            room.title = title
        if list(room.peers) != peers:
            room.peers = peers
        if changes and (not room.changes or room.changes[-1].get("rev") != changes[-1]["rev"]):
            room.changes = changes
        if room.saved != saved:
            room.saved = saved
        if room.save_error != error:
            room.save_error = error
        if room.deleted:
            room.deleted = False
        call.sync_room(room, book.key)

    def _run(self, s: Session, cid: str, params: Optional[Dict[str, Any]] = None):
        return s.execute(cid, params or {})

    def _focus(self) -> None:
        self.focus_key += 1

    # ------------------------------------------------------------ lifecycle

    @rx.event
    async def on_load(self):
        """Opens the book in the URL (a new, blank one when the address is new) and links this
        tab to its room."""
        path = str(self.router.url.path or "")
        book_id = path.rstrip("/").rsplit("/", 1)[-1] if path else ""
        if not valid_id(book_id):
            return rx.redirect("/")
        await self._identify()
        if not self.functions:
            self.functions = FUNCTIONS
        try:
            book = STORE.open(book_id, creator=self.me)
        except EngineError as ex:
            self.book, self.book_gone = "", True
            return rx.toast.error(str(ex).split("\n", 1)[0], position="bottom-center")
        entry = self._entry()
        if self.book != book_id:
            self.window = [0, 0, 80, 30]
            self.tab, self.backstage = "Home", False
            self.dialog, self.dlg, self.dlg_error = "", {}, ""
            self.selected_object, self.selected_kind = -1, ""
        self.book, self.book_gone = book_id, False
        view = self._join(entry, book)
        room = await self.get_state(BookRoom)
        room = await room._link_to(room_token(book_id))
        with book.act(view):
            entry.sig, entry.sheet = (), None
            self._refresh(entry, book, view)
        self._publish(room, book)
        self._focus()

    @rx.event
    async def leave_book(self):
        """Leaves the book on screen (going back to the workbook list): the others stop seeing
        this tab, and it stops hearing about the book."""
        entry = self._entry()
        book = STORE.get(entry.book_id) if entry.book_id else None
        if book is not None and entry.view is not None:
            book.leave(entry.view.key)
        entry.book_id, entry.view = "", None
        if not self.book:
            return None
        self.book = ""
        room = await self.get_state(BookRoom)
        if book is not None:
            self._publish(room, book)
        try:
            return await room._unlink()
        except Exception:  # not linked (e.g. after a server restart)
            return None

    # ------------------------------------------------------------ grid events

    @rx.event
    async def on_event(self, e: Dict[str, Any]):
        room = await self.get_state(BookRoom)
        entry, book, view = self._bound()
        if book is None:
            return None
        t = e.get("type")
        seq = int(e.get("seq") or 0)
        out = None
        extra: Dict[str, Any] = {}
        with book.act(view) as s:
            try:
                if t == "viewport":
                    w = e.get("window") or [0, 0, 80, 30]
                    self.window = [max(0, int(w[0])), max(0, int(w[1])), min(MAX_ROWS - 1, int(w[2])), min(MAX_COLS - 1, int(w[3]))]
                elif t == "presence":
                    ed = e.get("editing")
                    view.info["editing"] = [int(ed[0]), int(ed[1])] if isinstance(ed, (list, tuple)) and len(ed) == 2 else None
                elif t in ("sync", "ping"):
                    pass
                elif t == "select":
                    sel = e.get("sel") or {}
                    ranges = sel.get("ranges") or []
                    if ranges:
                        act = sel.get("active") or ranges[-1][:2]
                        anc = sel.get("anchor") or act
                        self._run(s, "selection.set", {"range": ",".join(_rng(g) for g in ranges[:256]), "active": a1(int(act[0]), int(act[1])), "anchor": a1(int(anc[0]), int(anc[1]))})
                elif t == "commit":
                    view.info["editing"] = None
                    out = self._commit(s, e, seq, extra)
                elif t == "command":
                    out = self._command(entry, str(e.get("id")), e.get("params") or {})
                elif t == "copy":
                    r = self._run(s, "edit.cut" if e.get("cut") else "edit.copy", {"text": e.get("text")} if isinstance(e.get("text"), str) else {})
                    if not isinstance(e.get("text"), str):
                        extra["clipboardText"] = {"seq": seq, "text": r["text"]}
                elif t == "paste":
                    text = e.get("text")
                    self._run(s, "edit.paste", {"text": text} if isinstance(text, str) else {})
                elif t == "resize":
                    axis = e.get("axis")
                    idx = [int(i) for i in (e.get("indices") or [])][:16384]
                    if idx:
                        spans = _spans(idx, axis)
                        if e.get("autofit"):
                            self._run(s, "home.autofitColumnWidth" if axis == "cols" else "home.autofitRowHeight", {axis: spans})
                        else:
                            size = float(e.get("size") or 0)
                            if axis == "cols":
                                self._run(s, "home.columnWidth", {"cols": spans, "width": size})
                            else:
                                self._run(s, "home.rowHeight", {"rows": spans, "height": size})
                elif t == "fill":
                    self._run(s, "edit.autoFill", {"source": e["source"], "target": e["target"], "mode": e.get("mode", "series")})
                elif t == "move":
                    self._run(s, "edit.moveCells", {"from": e["from"], "to": e["to"], "copy": bool(e.get("copy"))})
                elif t == "filter":
                    self._open_filter(s, int(e["r"]), int(e["c"]), e.get("table"))
                elif t == "link":
                    r = self._run(s, "hyperlink.follow", {"cell": a1(int(e["r"]), int(e["c"]))})
                    if r and r.get("url"):
                        out = rx.call_script(f"window.open({_js_str(r['url'])}, '_blank', 'noopener')")
                elif t == "object":
                    self._run(s, "object.move", {k: e[k] for k in ("id", "x", "y", "width", "height") if k in e})
                elif t == "objectSelect":
                    oid = e.get("id")
                    self.selected_object = int(oid) if oid is not None else -1
                    self.selected_kind = str(e.get("kind") or "")
                    if self.selected_kind == "chart":
                        self.tab = "Chart Design"
                    elif self.tab == "Chart Design":
                        self.tab = "Home"
                elif t == "dialog":
                    self._open_dialog(entry, str(e.get("name")), e.get("params") or {})
                elif t == "namebox":
                    out = self._namebox(s, str(e.get("text", "")))
            except EngineError as ex:
                msg = str(ex).split("\n", 1)[0]
                extra["lastError"] = {"seq": seq, "message": msg, "kind": t, **({"r": e.get("r"), "c": e.get("c"), "text": e.get("text")} if t == "commit" else {})}
                out = rx.toast.error(msg, position="bottom-center")
            except Exception as ex:  # never let a UI event kill the session
                extra["lastError"] = {"seq": seq, "message": str(ex), "kind": t}
                out = rx.toast.error(f"Something went wrong: {ex}", position="bottom-center")
            pending = self._macro_pending(s)
            self._refresh(entry, book, view, seq, extra, passive=t in _PASSIVE)
        if t == "ping":
            book.prune(PEER_IDLE)
            STORE.close_idle()
        self._publish(room, book)
        return _both(out, pending)

    def _commit(self, s: Session, e: Dict[str, Any], seq: int, extra: Dict[str, Any]):
        r, c = int(e["r"]), int(e["c"])
        text = str(e.get("text") or "")
        mode = e.get("mode", "normal")
        if mode == "fill":
            self._run(s, "range.fill", {"input": text})
            return None
        if mode == "array":
            self._run(s, "cell.set", {"cell": a1(r, c), "input": text, "array": True})
            return None
        self._run(s, "cell.set", {"cell": a1(r, c), "input": text})
        move = e.get("move", "down")
        if move != "none":
            self._run(s, "selection.set", {"range": ",".join(_rng(g) for g in s.doc().selection.ranges), "active": a1(r, c)})
            self._run(s, "selection.next", {"forward": move in ("down", "right"), "byRow": move in ("left", "right")})
        msgs = s.messages[:]
        s.messages.clear()
        if msgs:
            return rx.toast.warning(msgs[-1], position="bottom-center")
        return None

    def _command(self, entry: sessions.Entry, cid: str, params: Dict[str, Any]):
        s = entry.session
        if cid == "ui.save":
            return self._download(s, "auto")
        if cid == "ui.saveCsv":
            return self._download(s, "csv")
        if cid == "ui.print":
            return self._print(s)
        if cid.startswith("ui."):
            return None
        r = self._run(s, cid, params)
        msgs = s.messages[:]
        s.messages.clear()
        if cid == "data.removeDuplicates" and isinstance(r, dict):
            return rx.toast.info(f"{r.get('removed', 0)} duplicate values found and removed; {r.get('remaining', 0)} unique values remain.", position="bottom-center")
        if cid == "edit.find" and isinstance(r, dict) and "count" in r:
            return None
        if msgs:
            return rx.toast.warning(msgs[-1], position="bottom-center")
        return None

    def _download(self, s: Session, fmt: str):
        """A copy of the workbook as a file (the shared book itself is saved on the server).
        ``auto``: .xlsm for a workbook with macros, else .xlsx."""
        if fmt == "auto":
            fmt = workbook_format(s)
        data = encode(s, fmt)
        return rx.download(data=data, filename=f"{_file_stem(s.doc().name)}.{fmt}", mime_type=MIMES[fmt])

    def _print(self, s: Session):
        r = self._run(s, "file.exportHtml", {"print": True})
        page = _js_str(r["html"])
        js = (
            "(() => { let f = document.getElementById('rg-print-frame'); if (f) f.remove();"
            " f = document.createElement('iframe'); f.id = 'rg-print-frame';"
            " Object.assign(f.style, {position: 'fixed', right: '0', bottom: '0', width: '0', height: '0', border: '0'});"
            " document.body.appendChild(f); const d = f.contentWindow.document; d.open(); d.write(" + page + "); d.close();"
            " setTimeout(() => { f.contentWindow.focus(); f.contentWindow.print(); }, 300); })()"
        )
        return rx.call_script(js)

    def _namebox(self, s: Session, text: str):
        if not text:
            return None
        try:
            self._run(s, "edit.goTo", {"reference": text})
            return None
        except EngineError:
            pass
        from gridcore.engine.cmd_formulas import check_name

        if check_name(text) is None:
            self._run(s, "formulas.defineName", {"name": text})
            return rx.toast.success(f"Defined the name {text}.", position="bottom-center")
        raise EngineError("Reference isn't valid.")

    # ------------------------------------------------------------ ribbon actions

    @rx.event
    async def cmd(self, cid: str, params: Dict[str, Any] = {}):
        room = await self.get_state(BookRoom)
        entry, book, view = self._bound()
        if book is None:
            return None
        out = None
        with book.act(view) as s:
            try:
                out = self._command(entry, cid, dict(params or {}))
            except EngineError as ex:
                out = rx.toast.error(str(ex).split("\n", 1)[0], position="bottom-center")
            pending = self._macro_pending(s)
            self._refresh(entry, book, view)
        self._publish(room, book)
        self._focus()
        return _both(out, pending)

    @rx.event
    def set_tab(self, tab: str):
        self.tab = tab
        self.backstage = tab == "File"
        if tab == "File":
            self.tab = "Home"
            self.recent = [
                {"id": m.id, "title": m.title, "when": time.strftime("%Y-%m-%d %H:%M", time.localtime(m.updated_at)) if m.updated_at else ""}
                for m in STORE.list()
                if m.id != self.book
            ][:8]

    @rx.event
    def close_backstage(self):
        self.backstage = False
        self._focus()

    @rx.event
    async def zoom_to(self, z):
        room = await self.get_state(BookRoom)
        entry, book, view = self._bound()
        if book is None:
            return None
        with book.act(view) as s:
            try:
                self._run(s, "view.zoom", {"percent": int(float(z if not isinstance(z, list) else z[0]))})
            except (EngineError, TypeError, ValueError):
                pass
            self._refresh(entry, book, view)
        self._publish(room, book)

    @rx.event
    def zoom_step(self, delta: int):
        z = max(10, min(400, int(self.zoom) + int(delta)))
        return GridState.zoom_to(z)

    @rx.event
    async def activate_sheet(self, index: int):
        room = await self.get_state(BookRoom)
        entry, book, view = self._bound()
        if book is None:
            return None
        out = None
        with book.act(view) as s:
            try:
                self._run(s, "sheet.activate", {"sheet": int(index)})
            except EngineError as ex:
                out = rx.toast.error(str(ex))
            self._refresh(entry, book, view)
        self._publish(room, book)
        self._focus()
        return out

    @rx.event
    async def new_sheet(self):
        room = await self.get_state(BookRoom)
        entry, book, view = self._bound()
        if book is None:
            return None
        out = None
        with book.act(view) as s:
            try:
                self._run(s, "home.insertSheet", {"index": len(s.doc().wb.sheets)})
            except EngineError as ex:
                out = rx.toast.error(str(ex))
            self._refresh(entry, book, view)
        self._publish(room, book)
        self._focus()
        return out

    @rx.event
    def new_workbook(self, sample: str = ""):
        """A new shared workbook (blank or a sample), opened in this tab."""
        try:
            meta = STORE.create_sample(sample, self.me) if sample and sample in SAMPLES else STORE.create_blank(self.me)
        except EngineError as ex:
            return rx.toast.error(str(ex).split("\n", 1)[0], position="bottom-center")
        self.backstage = False
        return rx.redirect(f"/book/{meta.id}")

    @rx.event
    def close_workbook(self):
        self.backstage = False
        return rx.redirect("/")

    async def _rename(self, title: str):
        if not self.book:
            return None
        room = await self.get_state(BookRoom)
        meta = STORE.meta(self.book)
        if meta is not None and clean_title(title) == meta.title:
            return None
        try:
            STORE.rename(self.book, title)
        except EngineError as ex:
            return rx.toast.error(str(ex), position="bottom-center")
        book = STORE.get(self.book)
        if book is not None:
            self._publish(room, book)
        return None

    @rx.event
    async def rename_book(self, title: str):
        """Renames the workbook for everyone on it (and in the workbook list)."""
        return await self._rename(title)

    @rx.event
    async def rename_from_form(self, form: Dict[str, Any]):
        """Enter in the title: rename, and give the keyboard back to the grid."""
        out = await self._rename(str(form.get("title", "")))
        self._focus()
        blur = rx.call_script("document.activeElement && document.activeElement.blur()")
        return [out, blur] if out is not None else blur

    @rx.event
    async def follow(self, peer: str):
        """Jumps to where someone on the book is (their sheet and selection)."""
        room = await self.get_state(BookRoom)
        entry, book, view = self._bound()
        if book is None:
            return None
        target = book.views.get(peer)
        if target is None or target is view:
            return None
        with book.act(view) as s:
            try:
                self._run(s, "sheet.activate", {"sheet": target.sheet})
                sel = target.selection
                self._run(s, "selection.set", {"range": sel.a1(), "active": a1(*sel.active)})
            except EngineError:
                pass
            self._refresh(entry, book, view)
        self._publish(room, book)
        self._focus()

    @rx.event
    async def save_as(self, fmt: str):
        entry, book, view = self._bound()
        self.backstage = False
        if book is None:
            return None
        with book.act(view) as s:
            try:
                return self._download(s, fmt)
            except EngineError as ex:
                return rx.toast.error(str(ex))

    @rx.event
    async def print_sheet(self):
        entry, book, view = self._bound()
        self.backstage = False
        if book is None:
            return None
        with book.act(view) as s:
            try:
                return self._print(s)
            except EngineError as ex:
                return rx.toast.error(str(ex))

    @rx.event
    async def save_html(self):
        entry, book, view = self._bound()
        self.backstage = False
        if book is None:
            return None
        with book.act(view) as s:
            try:
                r = self._run(s, "file.exportHtml", {"print": False})
            except EngineError as ex:
                return rx.toast.error(str(ex))
        return rx.download(data=r["html"].encode("utf-8"), filename=r["name"], mime_type="text/html")

    @rx.event
    async def open_upload(self, files: List[rx.UploadFile]):
        """Opens an uploaded file as a new shared workbook."""
        for f in files[:1]:
            data = await f.read()
            name = os.path.basename(f.name or f.filename or "Book.xlsx")
            try:
                meta, warnings = STORE.create_from_file(data, name, self.me)
            except EngineError as ex:
                return rx.toast.error(str(ex), position="bottom-center")
            self.backstage = False
            self.dialog = ""
            return [
                rx.redirect(f"/book/{meta.id}"),
                rx.toast.success(f"Opened {name}" + (f" ({len(warnings)} warnings)" if warnings else ""), position="bottom-center"),
            ]
        return None

    @rx.event
    async def picture_upload(self, files: List[rx.UploadFile]):
        room = await self.get_state(BookRoom)
        entry, book, view = self._bound()
        if book is None:
            return None
        out = None
        pictures = [(await f.read(), f.name or f.filename or "picture") for f in files[:1]]
        with book.act(view) as s:
            for data, name in pictures:
                try:
                    self._run(s, "insert.picture", {"data": data, "name": name})
                except EngineError as ex:
                    out = rx.toast.error(str(ex), position="bottom-center")
            self.dialog = ""
            self._refresh(entry, book, view)
        self._publish(room, book)
        self._focus()
        return out

    # ------------------------------------------------------------ dialogs

    @rx.event
    async def open_dialog(self, name: str, params: Dict[str, Any] = {}):
        entry, book, view = self._bound()
        if book is None:
            return None
        with book.act(view):
            try:
                self._open_dialog(entry, name, dict(params or {}))
            except EngineError as ex:
                return rx.toast.error(str(ex).split("\n", 1)[0])

    @rx.event
    def dialog_open_change(self, is_open: bool):
        if not is_open:
            self.dialog = ""
            self.dlg = {}
            self.dlg_error = ""
            self._focus()

    @rx.event
    def toggle_list_item(self, field: str, index: int):
        d = dict(self.dlg)
        items = [dict(x) for x in d.get(field, [])]
        if 0 <= index < len(items):
            items[index]["checked"] = not items[index].get("checked", False)
        d[field] = items
        self.dlg = d

    @rx.event
    def close_dialog(self):
        self.dialog = ""
        self.dlg = {}
        self.dlg_error = ""
        self._focus()

    @rx.event
    def set_dlg(self, field: str, value: Any):
        d = dict(self.dlg)
        d[field] = value
        self.dlg = d

    @rx.event
    def toggle_filter_value(self, index: int):
        d = dict(self.dlg)
        vals = [dict(v) for v in d.get("values", [])]
        if 0 <= index < len(vals):
            vals[index]["checked"] = not vals[index]["checked"]
        d["values"] = vals
        self.dlg = d

    @rx.event
    def filter_select_all(self, on: bool):
        d = dict(self.dlg)
        d["values"] = [{**v, "checked": bool(on)} for v in d.get("values", [])]
        d["blanksChecked"] = bool(on)
        self.dlg = d

    def _open_dialog(self, entry: sessions.Entry, name: str, params: Dict[str, Any]) -> None:
        s = entry.session
        d = s.doc()
        sh = d.sheet
        sel = d.selection
        ui = self.ui or {}
        rb = ui.get("ribbon", {})
        dlg: Dict[str, Any] = {}
        if name == "vba":
            # Alt+F11: the VBA editor isn't one of the small dialogs.
            self._vba_load(s, str(params.get("module") or ""))
            self.vba_open = True
            return
        if name == "macros":
            macros = s.execute("macro.list", {}).get("macros", [])
            module = str(params.get("module") or "")
            if module:
                macros = [m for m in macros if m["module"].lower() == module.lower()] or macros
            self.dialog = "macros"
            self.dlg = {"macros": macros, "selected": macros[0]["name"] if macros else ""}
            self.dlg_error = ""
            return
        if name == "formatCells":
            st = d.wb.styles.get(sh.style_id(sel.active))
            nf = parse_fmt(st.num_fmt)
            b = st.borders
            dlg = {
                "tab": params.get("tab", "number"),
                "numFmt": st.num_fmt,
                "category": nf.kind(),
                "fontName": st.font.name,
                "fontSize": str(st.font.size).rstrip("0").rstrip("."),
                "bold": st.font.bold,
                "italic": st.font.italic,
                "underline": st.font.underline,
                "strike": st.font.strike,
                "fontColor": st.font.color.hex(d.wb.theme) or "#000000",
                "fillColor": (st.fill.fg.hex(d.wb.theme) or "") if st.fill.pattern != "none" else "",
                "h": st.align.h,
                "v": st.align.v,
                "wrap": st.align.wrap,
                "shrink": st.align.shrink,
                "indent": str(st.align.indent),
                "rotation": str(st.align.rotation),
                "locked": st.protection.locked,
                "hidden": st.protection.hidden,
                "borderStyle": "thin",
                "borderColor": "#000000",
                "borderPreset": "",
                "sample": "",
            }
            dlg["orig"] = dict(dlg)
        elif name in ("find", "replace"):
            last = s.last_find or {}
            dlg = {"mode": name, "what": last.get("what", ""), "with": "", "matchCase": bool(last.get("matchCase", False)), "wholeCell": bool(last.get("wholeCell", False)), "within": last.get("within", "sheet"), "lookIn": last.get("lookIn", "formulas"), "results": [], "status": ""}
        elif name == "goTo":
            dlg = {"reference": "", "special": "blanks"}
        elif name == "pasteSpecial":
            if s.clipboard is None:
                raise EngineError("Copy something first: Paste Special pastes what you copied in ReflexGrid.")
            dlg = {"what": "all", "operation": "none", "skipBlanks": False, "transpose": False}
        elif name == "defineName":
            from gridcore.formula.printer import quote_sheet

            dlg = {"name": "", "refersTo": "=" + f"{quote_sheet(sh.name)}!{sel.current().abs_a1()}", "scope": "Workbook", "comment": ""}
        elif name == "nameManager":
            dlg = {"names": self._run(s, "formulas.nameManager")["names"]}
        elif name == "cfHighlight":
            dlg = {"kind": params.get("kind", "greaterThan"), "value": "", "value2": "", "preset": "lightRed", "period": "today"}
        elif name == "cfTop":
            dlg = {"kind": params.get("kind", "top"), "rank": "10", "preset": "lightRed"}
        elif name == "cfRule":
            dlg = {"formula": "=", "fill": "#FFC7CE", "font": "#9C0006", "bold": False}
        elif name == "cfManage":
            dlg = {"rules": self._run(s, "home.manageRules", {"scope": "sheet"})["rules"]}
        elif name == "validation":
            v = s.execute("data.validationAt", {})
            dlg = {
                "kind": (v or {}).get("kind", "any"),
                "operator": (v or {}).get("operator", "between"),
                "f1": ((v or {}).get("f1") or "").strip('"') if (v or {}).get("kind") == "list" and ((v or {}).get("f1") or "").startswith('"') else ("=" + v["f1"] if v and v.get("f1") and v.get("kind") == "list" else (v or {}).get("f1", "") or ""),
                "f2": (v or {}).get("f2") or "",
                "allowBlank": (v or {}).get("allowBlank", True),
                "inCellDropdown": (v or {}).get("inCellDropdown", True),
                "inputTitle": (v or {}).get("inputTitle", ""),
                "inputMessage": (v or {}).get("inputMessage", ""),
                "errorStyle": (v or {}).get("errorStyle", "stop"),
                "errorTitle": (v or {}).get("errorTitle", ""),
                "errorMessage": (v or {}).get("errorMessage", ""),
                "tab": "settings",
            }
        elif name == "sort":
            from gridcore.engine.cmd_data import _data_region, guess_header

            si, rng, owner = _data_region(s, {})
            header = owner.header_row if hasattr(owner, "header_row") else (True if owner is not None else guess_header(sh, rng, d.wb))
            cols = []
            for c in range(rng.c0, min(rng.c1, rng.c0 + 200) + 1):
                label = f"Column {col_to_letters(c)}"
                if header:
                    v = sh.value((rng.r0, c))
                    if v is not None:
                        label = str(v) if not isinstance(v, float) else f"{v:g}"
                cols.append({"value": col_to_letters(c), "label": label})
            dlg = {"range": rng.a1(), "header": header, "cols": cols, "k1": cols[0]["value"] if cols else "A", "o1": "asc", "k2": "", "o2": "asc", "k3": "", "o3": "asc", "matchCase": False}
        elif name == "insertFunction":
            dlg = {"search": "", "category": "Most Used", "selected": "SUM"}
        elif name == "hyperlink":
            link = ui.get("link")
            dlg = {"target": link or "", "text": "", "tooltip": ""}
        elif name == "note":
            note = ui.get("note")
            dlg = {"text": (note or {}).get("text", ""), "threaded": bool(params.get("threaded", False))}
        elif name in ("rowHeight", "columnWidth", "defaultWidth"):
            if name == "rowHeight":
                v = sh.row_height(sel.active[0])
            elif name == "columnWidth":
                v = sh.col_width(sel.active[1])
            else:
                v = sh.default_col_width
            dlg = {"value": f"{v:g}"}
        elif name in ("insertCells", "deleteCells"):
            dlg = {"shift": "down" if name == "insertCells" else "up"}
        elif name == "renameSheet":
            idx = int(params.get("index", d.wb.active_sheet))
            dlg = {"index": idx, "name": d.wb.sheets[idx].name}
        elif name == "moveSheet":
            idx = int(params.get("index", d.wb.active_sheet))
            dlg = {"index": idx, "to": str(idx), "copy": False, "sheets": [x.name for x in d.wb.sheets]}
        elif name == "tabColor":
            dlg = {"index": int(params.get("index", d.wb.active_sheet))}
        elif name == "unhideSheet":
            hidden = [x.name for x in d.wb.sheets if x.visibility != "visible"]
            if not hidden:
                raise EngineError("No sheets are hidden.")
            dlg = {"sheet": hidden[0], "sheets": hidden}
        elif name == "protectSheet":
            dlg = {"password": "", "formatCells": False, "formatColumns": False, "formatRows": False, "insertRows": False, "insertColumns": False, "deleteRows": False, "deleteColumns": False, "sort": False, "autoFilter": False}
        elif name in ("unprotectSheet", "protectWorkbook", "unprotectWorkbook"):
            dlg = {"password": ""}
        elif name == "insertChart":
            dlg = {"kind": params.get("kind", "columnClustered"), "title": ""}
        elif name == "insertTable":
            from gridcore.engine.cmd_data import guess_header
            from gridcore.engine.selection import current_region

            rng = sel.current() if not sel.current().is_single() else current_region(sh, sel.active)
            dlg = {"range": rng.a1(), "header": guess_header(sh, rng, d.wb) or all(isinstance(sh.value((rng.r0, c)), str) for c in range(rng.c0, min(rng.c1, rng.c0 + 50) + 1)), "style": "TableStyleMedium2"}
        elif name == "removeDuplicates":
            from gridcore.engine.cmd_data import _data_region, guess_header

            si, rng, owner = _data_region(s, {})
            header = owner.header_row if hasattr(owner, "header_row") else guess_header(sh, rng, d.wb)
            cols = []
            for i, c in enumerate(range(rng.c0, min(rng.c1, rng.c0 + 100) + 1)):
                v = sh.value((rng.r0, c)) if header else None
                cols.append({"index": i, "label": (str(v) if v is not None else f"Column {col_to_letters(c)}"), "checked": True})
            dlg = {"range": rng.a1(), "header": header, "cols": cols}
        elif name == "textToColumns":
            dlg = {"tab": True, "comma": False, "semicolon": False, "space": False, "other": "", "consecutive": False}
        elif name == "properties":
            p = d.wb.props
            dlg = {"title": p.title, "subject": p.subject, "author": p.author, "company": p.company, "keywords": p.keywords, "description": p.description}
        elif name == "open":
            dlg = {}
        elif name == "insertPicture":
            dlg = {}
        elif name == "chartTitle":
            ch = next((c for c in sh.charts if c.id == self.selected_object), None)
            if ch is None:
                raise EngineError("Select a chart first.")
            dlg = {"id": ch.id, "title": ch.title or "", "xTitle": ch.x_title or "", "yTitle": ch.y_title or ""}
        elif name == "shortcuts":
            dlg = {}
        elif name == "goalSeek":
            dlg = {"setCell": a1(*sel.active), "toValue": "", "byChangingCell": ""}
        elif name == "sparkline":
            r = sel.current()
            loc = RangeRef(r.r0, r.c1 + 1, r.r1, r.c1 + 1) if r.c1 + 1 < MAX_COLS else r
            dlg = {"kind": params.get("kind", "line"), "data": r.a1(), "location": loc.a1()}
        elif name == "selectData":
            ch = next((c for c in sh.charts if c.id == self.selected_object), None)
            if ch is None:
                raise EngineError("Select a chart first.")
            dlg = {"id": ch.id, "range": ch.source or ""}
        else:
            raise EngineError(f"unknown dialog `{name}`")
        self.dlg = dlg
        self.dlg_error = ""
        self.dialog = name

    def _open_filter(self, s: Session, r: int, c: int, table: Optional[str]) -> None:
        sel_cell = a1(r, c)
        s.execute("selection.set", {"range": sel_cell})
        p = {"col": col_to_letters(c)}
        if table:
            p["table"] = table
        res = s.execute("data.filterValues", p)
        vals = [{"text": v["text"], "checked": v["checked"]} for v in res["values"]]
        self.dlg = {
            "col": col_to_letters(c),
            "table": table or "",
            "values": vals,
            "blanks": res["blanks"],
            "blanksChecked": res["blanksChecked"],
            "active": res["active"],
            "numeric": res["numeric"],
            "search": "",
            "op": "",
            "value": "",
            "op2": "",
            "value2": "",
            "and": True,
            "mode": "values",
        }
        self.dlg_error = ""
        self.dialog = "filter"

    @rx.event
    async def dialog_ok(self, action: str = "ok"):
        room = await self.get_state(BookRoom)
        entry, book, view = self._bound()
        if book is None:
            return None
        name = self.dialog
        d = dict(self.dlg)
        out = None
        failed = False
        with book.act(view) as s:
            try:
                out, keep_open = self._dialog_action(s, name, action, d)
            except EngineError as ex:
                self.dlg_error = str(ex).split("\n", 1)[0]
                failed = True
            else:
                if not keep_open:
                    self.dialog = ""
                    self.dlg = {}
                    self.dlg_error = ""
                    self._focus()
            pending = self._macro_pending(s)
            self._refresh(entry, book, view)
        self._publish(room, book)
        return _both(None if failed else out, pending)

    def _dialog_action(self, s: Session, name: str, action: str, d: Dict[str, Any]):
        run = lambda cid, **p: s.execute(cid, p)
        if name == "formatCells":
            o = d.get("orig", {})
            patch: Dict[str, Any] = {}
            if d.get("numFmt") != o.get("numFmt"):
                patch["numFmt"] = d.get("numFmt") or "General"
            font = {}
            for k_ui, k in (("fontName", "name"), ("bold", "bold"), ("italic", "italic"), ("underline", "underline"), ("strike", "strike"), ("fontColor", "color")):
                if d.get(k_ui) != o.get(k_ui):
                    font[k] = d.get(k_ui)
            if d.get("fontSize") != o.get("fontSize"):
                size = _num(d.get("fontSize"))
                if size is None or not 1 <= size <= 409:
                    raise EngineError("The font size must be a number between 1 and 409.")
                font["size"] = size
            if font:
                patch["font"] = font
            if d.get("fillColor") != o.get("fillColor"):
                patch["fill"] = {"color": d.get("fillColor") or None}
            align = {}
            for k in ("h", "v", "wrap", "shrink"):
                if d.get(k) != o.get(k):
                    align[k] = d.get(k)
            for k in ("indent", "rotation"):
                if d.get(k) != o.get(k):
                    v = _num(d.get(k))
                    if v is None:
                        raise EngineError(f"{k.title()} must be a number.")
                    align[k] = int(v)
            if align:
                patch["align"] = align
            prot = {k: d.get(k) for k in ("locked", "hidden") if d.get(k) != o.get(k)}
            if prot:
                patch["protection"] = prot
            if patch:
                run("home.formatCells", **patch)
            if d.get("borderPreset"):
                run("home.borders", preset=d["borderPreset"], style=d.get("borderStyle", "thin"), color=d.get("borderColor") or "#000000")
            return None, False
        if name in ("find", "replace"):
            common = {"what": d.get("what", ""), "matchCase": d.get("matchCase"), "wholeCell": d.get("wholeCell"), "within": d.get("within", "sheet"), "lookIn": d.get("lookIn", "formulas")}
            if action == "findAll":
                r = s.execute("edit.find", {**common, "all": True})
                d["results"] = r["matches"][:500]
                d["status"] = f"{len(r['matches'])} cell(s) found"
                self.dlg = d
                return None, True
            if action == "findNext":
                r = s.execute("edit.find", common)
                d["status"] = f"{r['sheet']}!{r['cell']}"
                self.dlg = d
                return None, True
            if action == "replace":
                s.execute("edit.replace", {**common, "with": d.get("with", "")})
                return None, True
            if action == "replaceAll":
                r = s.execute("edit.replace", {**common, "with": d.get("with", ""), "all": True})
                d["status"] = f"All done. We made {r['replaced']} replacements."
                self.dlg = d
                return None, True
            if action.startswith("goto:"):
                _, sheet, cell = action.split(":", 2)
                s.execute("edit.goTo", {"reference": f"'{sheet}'!{cell}"})
                return None, True
            return None, False
        if name == "goTo":
            if action == "special":
                run("edit.goToSpecial", kind=d.get("special", "blanks"))
            else:
                ref = (d.get("reference") or "").strip()
                if not ref:
                    raise EngineError("Enter a reference.")
                run("edit.goTo", reference=ref)
            return None, False
        if name == "pasteSpecial":
            run("edit.pasteSpecial", what=d.get("what", "all"), operation=d.get("operation", "none"), skipBlanks=bool(d.get("skipBlanks")), transpose=bool(d.get("transpose")))
            return None, False
        if name == "defineName":
            p = {"name": d.get("name", ""), "refersTo": d.get("refersTo", ""), "comment": d.get("comment", "")}
            if d.get("scope") and d["scope"] != "Workbook":
                p["scope"] = d["scope"]
            run("formulas.defineName", **p)
            return None, False
        if name == "nameManager":
            if action.startswith("delete:"):
                nm = action.split(":", 1)[1]
                entry = next((x for x in d.get("names", []) if x["name"] == nm), None)
                if entry and entry.get("table"):
                    raise EngineError("A table name can't be deleted here; use Convert to Range.")
                p = {"name": nm}
                if entry and entry.get("scope") not in (None, "Workbook"):
                    p["scope"] = entry["scope"]
                s.execute("formulas.deleteName", p)
                d["names"] = s.execute("formulas.nameManager", {})["names"]
                self.dlg = d
                return None, True
            return None, False
        if name == "cfHighlight":
            k = d.get("kind")
            rule: Dict[str, Any] = {"preset": d.get("preset", "lightRed")}
            if k in ("greaterThan", "lessThan", "equal"):
                rule.update({"type": "cellIs", "operator": k, "value": d.get("value", "")})
            elif k == "between":
                rule.update({"type": "cellIs", "operator": "between", "value": d.get("value", ""), "value2": d.get("value2", "")})
            elif k == "text":
                rule.update({"type": "containsText", "text": d.get("value", "")})
            elif k == "date":
                rule.update({"type": "timePeriod", "period": d.get("period", "today")})
            elif k == "duplicate":
                rule.update({"type": "duplicateValues" if d.get("value", "duplicate") != "unique" else "uniqueValues"})
            run("home.conditionalFormat", rule=rule)
            return None, False
        if name == "cfTop":
            k = d.get("kind", "top")
            rank = int(_num(d.get("rank"), 10) or 10)
            if k in ("top", "bottom", "topPercent", "bottomPercent"):
                rule = {"type": "top10", "rank": rank, "bottom": k.startswith("bottom"), "percent": k.endswith("Percent"), "preset": d.get("preset", "lightRed")}
            else:
                rule = {"type": "aboveAverage", "below": k == "below", "preset": d.get("preset", "lightRed")}
            run("home.conditionalFormat", rule=rule)
            return None, False
        if name == "cfRule":
            style = {"fill": {"color": d.get("fill") or None}, "font": {"color": d.get("font") or None, "bold": bool(d.get("bold"))}}
            run("home.conditionalFormat", rule={"type": "expression", "formula": d.get("formula", ""), "style": style})
            return None, False
        if name == "cfManage":
            if action.startswith("delete:"):
                s.execute("home.deleteRule", {"index": int(action.split(":", 1)[1])})
            elif action.startswith("up:") or action.startswith("down:"):
                kind, idx = action.split(":", 1)
                s.execute("home.moveRule", {"index": int(idx), "up": kind == "up"})
            elif action == "clearSheet":
                s.execute("home.clearRules", {"scope": "sheet"})
            else:
                return None, False
            d["rules"] = s.execute("home.manageRules", {"scope": "sheet"})["rules"]
            self.dlg = d
            return None, True
        if name == "validation":
            if action == "clear":
                run("data.clearValidation")
                return None, False
            p = {k: d.get(k) for k in ("kind", "operator", "f1", "f2", "allowBlank", "inCellDropdown", "inputTitle", "inputMessage", "errorStyle", "errorTitle", "errorMessage")}
            if p["kind"] == "any" and not (p.get("inputTitle") or p.get("inputMessage")):
                run("data.clearValidation")
                return None, False
            run("data.validation", **p)
            return None, False
        if name == "sort":
            keys = []
            for i in (1, 2, 3):
                col = d.get(f"k{i}")
                if col:
                    keys.append({"col": col, "order": d.get(f"o{i}", "asc")})
            if not keys:
                raise EngineError("Choose a column to sort by.")
            run("data.sort", range=d.get("range"), keys=keys, header=bool(d.get("header")), matchCase=bool(d.get("matchCase")))
            return None, False
        if name == "filter":
            p: Dict[str, Any] = {"col": d.get("col")}
            if d.get("table"):
                p["table"] = d["table"]
            if action in ("sortAsc", "sortDesc"):
                s.execute("selection.set", {"range": f"{d.get('col')}1"} if False else {"range": s.doc().selection.current().a1()})
                s.execute("data.sortAscending" if action == "sortAsc" else "data.sortDescending", {"col": d.get("col")})
                return None, False
            if action == "clear":
                p["criterion"] = None
                s.execute("data.filterColumn", p)
                return None, False
            if d.get("mode") == "custom" and d.get("op"):
                conds = [{"op": d["op"], "value": d.get("value", "")}]
                if d.get("op2"):
                    conds.append({"op": d["op2"], "value": d.get("value2", "")})
                p["criterion"] = {"kind": "custom", "conditions": conds, "and": bool(d.get("and", True))}
            else:
                vals = d.get("values", [])
                if all(v["checked"] for v in vals) and (d.get("blanksChecked") or not d.get("blanks")):
                    p["criterion"] = None
                else:
                    p["criterion"] = {"kind": "values", "values": [v["text"] for v in vals if v["checked"]], "blanks": bool(d.get("blanksChecked"))}
            s.execute("data.filterColumn", p)
            return None, False
        if name == "insertFunction":
            fn = self.fn_selected.get("name") or d.get("selected") or "SUM"
            self.edit_request = {"seq": (self.edit_request or {}).get("seq", 0) + 1, "text": f"={fn}("}
            return None, False
        if name == "hyperlink":
            if action == "remove":
                run("insert.removeHyperlink")
                return None, False
            p = {"target": d.get("target", "")}
            if d.get("text"):
                p["text"] = d["text"]
            if d.get("tooltip"):
                p["tooltip"] = d["tooltip"]
            run("insert.hyperlink", **p)
            return None, False
        if name == "note":
            if action == "delete":
                run("review.deleteComment")
                return None, False
            if d.get("threaded"):
                run("review.newComment", text=d.get("text", ""))
            else:
                run("review.newNote", text=d.get("text", ""))
            return None, False
        if name == "rowHeight":
            run("home.rowHeight", height=_num(d.get("value"), -1))
            return None, False
        if name == "columnWidth":
            run("home.columnWidth", width=_num(d.get("value"), -1))
            return None, False
        if name == "defaultWidth":
            run("home.defaultWidth", width=_num(d.get("value"), -1))
            return None, False
        if name == "insertCells":
            run("home.insertCells", shift=d.get("shift", "down"))
            return None, False
        if name == "deleteCells":
            run("home.deleteCells", shift=d.get("shift", "up"))
            return None, False
        if name == "renameSheet":
            run("sheet.rename", sheet=int(d["index"]), name=d.get("name", ""))
            return None, False
        if name == "moveSheet":
            run("sheet.move", sheet=int(d["index"]), to=int(_num(d.get("to"), 0)), copy=bool(d.get("copy")))
            return None, False
        if name == "tabColor":
            run("sheet.tabColor", sheet=int(d["index"]), color=action if action != "ok" else None)
            return None, False
        if name == "unhideSheet":
            run("sheet.unhide", sheet=d.get("sheet"))
            return None, False
        if name == "protectSheet":
            p = {k: bool(d.get(k)) for k in ("formatCells", "formatColumns", "formatRows", "insertRows", "insertColumns", "deleteRows", "deleteColumns", "sort", "autoFilter")}
            if d.get("password"):
                p["password"] = d["password"]
            run("review.protectSheet", **p)
            return None, False
        if name == "unprotectSheet":
            run("review.unprotectSheet", password=d.get("password", ""))
            return None, False
        if name == "protectWorkbook":
            run("review.protectWorkbook", password=d.get("password", ""))
            return None, False
        if name == "unprotectWorkbook":
            run("review.unprotectWorkbook", password=d.get("password", ""))
            return None, False
        if name == "insertChart":
            p = {"kind": d.get("kind", "columnClustered")}
            if d.get("title"):
                p["title"] = d["title"]
            run("insert.chart", **p)
            return None, False
        if name == "insertTable":
            run("insert.table", range=d.get("range"), header=bool(d.get("header")), style=d.get("style", "TableStyleMedium2"))
            return None, False
        if name == "removeDuplicates":
            cols = [c["index"] for c in d.get("cols", []) if c.get("checked")]
            if not cols:
                raise EngineError("Select at least one column.")
            r = s.execute("data.removeDuplicates", {"range": d.get("range"), "columns": cols, "header": bool(d.get("header"))})
            return rx.toast.info(f"{r.get('removed', 0)} duplicate values found and removed; {r.get('remaining', 0)} unique values remain.", position="bottom-center"), False
        if name == "textToColumns":
            delims = [ch for k, ch in (("tab", "\t"), ("comma", ","), ("semicolon", ";"), ("space", " ")) if d.get(k)]
            if not delims and not d.get("other"):
                raise EngineError("Choose at least one delimiter.")
            run("data.textToColumns", delimiters=delims or [str(d.get("other"))[:1]], other=d.get("other") or None, consecutive=bool(d.get("consecutive")))
            return None, False
        if name == "properties":
            run("file.properties", **{k: d.get(k, "") for k in ("title", "subject", "author", "company", "keywords", "description")})
            return None, False
        if name == "chartTitle":
            run("chart.update", id=int(d["id"]), title=d.get("title") or None, xTitle=d.get("xTitle") or None, yTitle=d.get("yTitle") or None)
            return None, False
        if name == "goalSeek":
            goal = _num(d.get("toValue"))
            if goal is None:
                raise EngineError("To value must be a number.")
            r = s.execute("data.goalSeek", {"setCell": d.get("setCell", ""), "toValue": goal, "byChangingCell": d.get("byChangingCell", "")})
            msg = f"Goal Seeking with Cell {r['target']} " + ("found a solution." if r["found"] else "may not have found a solution.") + f" Target value: {goal:g}. Current value: {r['result']}"
            return (rx.toast.success(msg, position="bottom-center") if r["found"] else rx.toast.warning(msg, position="bottom-center")), False
        if name == "sparkline":
            run("insert.sparkline", kind=d.get("kind", "line"), data=d.get("data", ""), location=d.get("location", ""))
            return None, False
        if name == "selectData":
            run("chart.selectData", id=int(d["id"]), range=d.get("range", ""))
            return None, False
        return None, False

    @rx.event
    def number_format(self, value: str):
        if value == "__more__":
            return GridState.open_dialog("formatCells", {"tab": "number"})
        return GridState.cmd("home.numberFormat", {"format": value})

    @rx.event
    def set_edit_request(self, text: str):
        self.edit_request = {"seq": (self.edit_request or {}).get("seq", 0) + 1, "text": text}

    @rx.event
    async def audit(self, what: str):
        entry, book, view = self._bound()
        if book is None:
            return None
        with book.act(view) as s:
            try:
                if what == "precedents":
                    r = s.execute("formulas.tracePrecedents", {})
                    items = [f"{p['sheet']}!{p['range']}" for p in r["precedents"]]
                    return rx.toast.info("Precedents: " + (", ".join(items) if items else "none"), position="bottom-center")
                if what == "dependents":
                    r = s.execute("formulas.traceDependents", {})
                    items = [f"{p['sheet']}!{p['cell']}" for p in r["dependents"]]
                    return rx.toast.info("Dependents: " + (", ".join(items[:30]) if items else "none"), position="bottom-center")
                if what == "errors":
                    r = s.execute("formulas.errorChecking", {})
                    items = [f"{e['sheet']}!{e['cell']} {e['error']}" for e in r["errors"]]
                    circ = r.get("circular") or []
                    msg = ("Errors: " + ", ".join(items[:20])) if items else "No formula errors were found."
                    if circ:
                        msg += " Circular references: " + ", ".join(circ)
                    return rx.toast.info(msg, position="bottom-center")
                if what == "invalid":
                    r = s.execute("data.circleInvalid", {})
                    items = r["invalid"]
                    return rx.toast.info(("Invalid data in: " + ", ".join(items[:40])) if items else "All data is valid.", position="bottom-center")
            except EngineError as ex:
                return rx.toast.error(str(ex).split("\n", 1)[0], position="bottom-center")

    # ------------------------------------------------------------ chart ribbon

    @rx.event
    async def chart_cmd(self, field: str, value: Any):
        room = await self.get_state(BookRoom)
        entry, book, view = self._bound()
        if book is None:
            return None
        oid = self.selected_object
        if oid < 0:
            return rx.toast.error("Select a chart first.")
        out = None
        with book.act(view) as s:
            try:
                if field == "switch":
                    self._run(s, "chart.switchRowColumn", {"id": oid})
                elif field == "delete":
                    self._run(s, "object.delete", {"id": oid})
                    self.selected_object = -1
                    self.tab = "Home"
                else:
                    self._run(s, "chart.update", {"id": oid, field: value})
            except EngineError as ex:
                out = rx.toast.error(str(ex).split("\n", 1)[0])
            self._refresh(entry, book, view)
        self._publish(room, book)
        return out


    # ------------------------------------------------------------ macros (VBA)

    def _vba_load(self, s: Session, module: str = "", line: int = 0) -> None:
        """Fills the editor from the workbook's project: the module list, and ``module`` (or
        the one on screen, or the first code module) in the editor."""
        r = s.execute("macro.list", {})
        mods = [{"name": m["name"], "kind": m["kind"], "sheet": m.get("sheet", ""), "label": m["name"] + (f" ({m['sheet']})" if m.get("sheet") and m["sheet"] != "(workbook)" else "")} for m in r.get("modules", [])]
        self.vba_modules = mods
        self.vba_macros = r.get("macros", [])
        self.vba_diag = r.get("diagnostics", [])
        self.vba_locked = bool(r.get("locked")) or bool(r.get("unreadable"))
        self.vba_status = r.get("unreadable") or ("The VBA project is locked: its code can't be viewed." if r.get("locked") else "")
        names = [m["name"] for m in mods]
        pick = next((n for n in names if n.lower() == (module or self.vba_module).lower()), "")
        if not pick:
            pick = next((m["name"] for m in mods if m["kind"] == "standard"), names[0] if names else "")
        self._vba_show(s, pick, line)

    def _vba_show(self, s: Session, module: str, line: int = 0) -> None:
        other = module.lower() != (self.vba_module or "").lower()
        self.vba_module = module
        self.vba_error_line = line
        self.vba_dirty = False
        if not module or self.vba_locked:
            self.vba_code, self.vba_version, self.vba_kind = "", "", ""
        else:
            m = s.execute("macro.getModule", {"name": module})
            self.vba_code, self.vba_version, self.vba_kind = m["code"], m["version"], m["kind"]
            if not line:
                d = next((x for x in self.vba_diag if x["module"].lower() == module.lower()), None)
                self.vba_error_line = int(d["line"]) if d else 0
        if self.vba_kind == "form":
            # A form opens on its design, as in the VBA editor (unless an error is to be shown).
            if other:
                self.vba_form_view = not line
            if other:
                self.vba_form_selected, self.vba_form_tool = "", ""
            try:
                self._vba_form_design(s)
            except EngineError:
                self.vba_form_snapshot = {}
        else:
            self.vba_form_view = False
            self.vba_form_snapshot, self.vba_form_props, self.vba_form_warnings = {}, [], []
        self.vba_key += 1

    def _vba_apply(self, s: Session) -> None:
        """Puts the editor's code into the workbook (when it changed). Code typed while there is
        no module (a workbook without macros) becomes a new module."""
        if not self.vba_dirty or self.vba_locked:
            return
        if not self.vba_module:
            if not (self.vba_code or "").strip():
                return
            taken = {m["name"].lower() for m in self.vba_modules}
            n = 1
            while f"module{n}" in taken:
                n += 1
            r = s.execute("macro.setModule", {"name": f"Module{n}", "code": self.vba_code})
            self._vba_load(s, r["name"])
            return
        r = s.execute("macro.setModule", {"name": self.vba_module, "code": self.vba_code, "version": self.vba_version})
        self.vba_version = r["version"]
        self.vba_diag = r.get("diagnostics", [])
        self.vba_dirty = False
        d = next((x for x in self.vba_diag if x["module"].lower() == self.vba_module.lower()), None)
        self.vba_error_line = int(d["line"]) if d else 0
        self.vba_status = f"Compile error in {d['module']}, line {d['line']}: {d['message']}" if d else ""
        self.vba_macros = s.execute("macro.list", {}).get("macros", [])

    def _macro_after(self, s: Session, r: Optional[Dict[str, Any]]):
        """What a macro step left: a dialog to show, more to run, or its end (output, errors)."""
        if not isinstance(r, dict):
            return None
        status = r.get("status")
        self.macro_label = r.get("label") or self.macro_label
        if isinstance(r.get("forms"), list):
            self._forms_shown(r["forms"])
        self.vba_form_waiting = 0
        if status == "prompt":
            p = dict(r.get("prompt") or {})
            if p.get("kind") == "form":
                # A modal UserForm: it is the dialog; what the person does on it answers.
                self.macro_prompt = {}
                self.vba_form_waiting = int(p.get("formId") or 0)
                self.macro_busy = False
                return None
            self.macro_prompt = p
            self.macro_input = str(p.get("default") or "")
            self.macro_busy = False
            return None
        if status == "busy":
            self.macro_prompt = {}
            self.macro_busy = True
            return GridState.macro_pump
        self.macro_prompt = {}
        self.macro_busy = False
        out = [str(x) for x in r.get("output") or []]
        if out:
            self.vba_output = (self.vba_output + out)[-500:]
        err = r.get("error")
        if err and status in ("error", "cancelled") or (err and err.get("kind") in ("timeout", "steps")):
            where = f" ({err['module']}.{err.get('proc') or ''}, line {err['line']})" if err.get("module") else ""
            text = (err.get("text") or "The macro stopped.") + where
            if self.vba_open and err.get("module"):
                try:
                    self._vba_apply(s)
                    self._vba_show(s, err["module"], int(err.get("line") or 0))
                except EngineError:
                    pass
                self.vba_status = text
                return None
            if status == "cancelled":
                return rx.toast.info("The macro was stopped.", position="bottom-center")
            return rx.toast.error(text, position="bottom-center")
        msgs = [m.get("prompt", "") for m in r.get("messages") or [] if m.get("prompt")]
        if msgs:
            return rx.toast.info(msgs[-1], position="bottom-center")
        if self.vba_open and status == "done":
            self.vba_status = f"{self.macro_label}: done." if self.macro_label else ""
        return None

    def _forms_shown(self, forms: List[Dict[str, Any]]) -> None:
        """The forms now on the person's screen. The VBA editor makes way for them (it would
        keep the keyboard to itself) and comes back when they are gone. Pictures go apart, by
        key, so that a form's events don't send them again."""
        pics: Dict[str, str] = {}

        def strip(c: Dict[str, Any]) -> None:
            pic = c.get("picture")
            if isinstance(pic, str) and pic.startswith("data:"):
                key = "p" + hashlib.sha1(pic.encode("ascii", "replace")).hexdigest()[:16]
                pics[key] = pic
                c["picture"] = key
            for ch in c.get("children") or c.get("controls") or []:
                strip(ch)

        for f in forms:
            strip(f)
        if set(pics) != set(self.vba_form_pictures):
            self.vba_form_pictures = pics
        self.vba_forms = forms
        if forms and self.vba_open:
            self.vba_open = False
            self.vba_reopen = True
        elif not forms and self.vba_reopen:
            self.vba_reopen = False
            self.vba_open = True

    def _macro_pending(self, s: Session):
        """A run that an edit started (an event macro) and that waits for this tab."""
        run = getattr(s, "macro_run", None)
        if run is None or not run.paused or self.macro_prompt or self.macro_busy or self.vba_form_waiting:
            return None
        self.macro_label = run.label
        if run.state == "prompt":
            prompt = dict(run.prompt or {})
            if prompt.get("kind") == "form":
                try:
                    self._forms_shown(s.execute("macro.forms", {})["forms"])
                except EngineError:
                    pass
                self.vba_form_waiting = int(prompt.get("formId") or 0)
                return None
            self.macro_prompt = prompt
            self.macro_input = str((run.prompt or {}).get("default") or "")
            return None
        self.macro_busy = True
        return GridState.macro_pump

    async def _macro_step(self, cid: str, params: Dict[str, Any], apply_code: bool = False):
        room = await self.get_state(BookRoom)
        entry, book, view = self._bound()
        if book is None:
            return None
        out = None
        with book.act(view) as s:
            try:
                if apply_code:
                    self._vba_apply(s)
                r = s.execute(cid, params)
                out = self._macro_after(s, r)
            except EngineError as ex:
                msg = str(ex).split("\n", 1)[0]
                if self.vba_open:
                    self.vba_status = msg
                out = rx.toast.error(msg, position="bottom-center")
            self._refresh(entry, book, view)
        self._publish(room, book)
        return out

    @rx.event
    async def run_macro(self, name: str):
        self.dialog, self.dlg, self.dlg_error = "", {}, ""
        self.macro_label = name
        return await self._macro_step("macro.run", {"name": name}, apply_code=self.vba_open)

    @rx.event
    async def answer_macro(self, value: Any):
        p = self.macro_prompt or {}
        if p.get("kind") == "inputbox" and value == "ok":
            value = self.macro_input
        elif p.get("kind") == "inputbox" and value == "cancel":
            value = None
        self.macro_prompt = {}
        return await self._macro_step("macro.answer", {"value": value})

    @rx.event
    def set_macro_input(self, text: str):
        self.macro_input = text

    @rx.event
    async def stop_macro(self):
        self.macro_prompt = {}
        self.macro_busy = False
        return await self._macro_step("macro.stop", {})

    @rx.event(background=True)
    async def macro_pump(self):
        """Runs a long macro in time slices, letting everyone else work in between."""
        import asyncio

        while True:
            await asyncio.sleep(0.03)
            async with self:
                if not self.macro_busy:
                    return
                room = await self.get_state(BookRoom)
                entry, book, view = self._bound()
                if book is None:
                    self.macro_busy = False
                    return
                out = None
                with book.act(view) as s:
                    run = getattr(s, "macro_run", None)
                    if run is None or run.state != "busy":
                        self.macro_busy = False
                    else:
                        try:
                            out = self._macro_after(s, s.execute("macro.continue", {}))
                        except EngineError as ex:
                            self.macro_busy = False
                            out = rx.toast.error(str(ex).split("\n", 1)[0], position="bottom-center")
                    self._refresh(entry, book, view)
                self._publish(room, book)
                if not self.macro_busy:
                    return None if out is GridState.macro_pump else out

    @rx.event
    async def open_vba(self, module: str = "", line: int = 0):
        entry, book, view = self._bound()
        if book is None:
            return None
        self.dialog, self.dlg, self.dlg_error = "", {}, ""
        with book.act(view) as s:
            try:
                self._vba_load(s, module, int(line or 0))
            except EngineError as ex:
                return rx.toast.error(str(ex).split("\n", 1)[0], position="bottom-center")
        self.vba_open = True
        return None

    @rx.event
    async def close_vba(self):
        out = None
        if self.vba_dirty:
            out = await self._vba_edit(lambda s: None)
        self.vba_open = False
        self._focus()
        return out

    @rx.event
    def vba_open_change(self, is_open: bool):
        if not is_open:
            return GridState.close_vba

    @rx.event
    def vba_changed(self, code: str):
        if code != self.vba_code:
            self.vba_code = code
            self.vba_dirty = True

    async def _vba_edit(self, f, apply_first: bool = True):
        """Runs ``f(session)`` on the project (after applying the editor's code) and refreshes
        the editor and everyone's view."""
        room = await self.get_state(BookRoom)
        entry, book, view = self._bound()
        if book is None:
            return None
        out = None
        with book.act(view) as s:
            try:
                if apply_first:
                    self._vba_apply(s)
                out = f(s)
            except EngineError as ex:
                self.vba_status = str(ex).split("\n", 1)[0]
            self._refresh(entry, book, view)
        self._publish(room, book)
        return out

    @rx.event
    async def vba_save(self, code: str):
        self.vba_changed(code)
        out = await self._vba_edit(lambda s: None)
        if not self.vba_status:
            self.vba_status = "Saved."
        return out

    @rx.event
    async def vba_select(self, name: str):
        def f(s):
            self._vba_load(s, name)

        return await self._vba_edit(f)

    @rx.event
    async def vba_add(self, kind: str):
        def f(s):
            r = s.execute("macro.addModule", {"kind": kind})
            self._vba_load(s, r["name"])
            self.vba_status = f"Inserted {r['name']}."

        return await self._vba_edit(f)

    @rx.event
    async def vba_remove(self):
        name = self.vba_module

        def f(s):
            s.execute("macro.deleteModule", {"name": name})
            self.vba_dirty = False
            self._vba_load(s, "")
            self.vba_status = f"Removed {name}."

        self.vba_dirty = False
        return await self._vba_edit(f, apply_first=False)

    @rx.event
    async def vba_rename(self, form: Dict[str, Any]):
        new = str(form.get("name") or "").strip()
        old = self.vba_module

        def f(s):
            s.execute("macro.renameModule", {"name": old, "newName": new})
            self._vba_load(s, new)

        return await self._vba_edit(f)

    @rx.event
    async def vba_compile(self):
        def f(s):
            d = s.execute("macro.check", {})["diagnostics"]
            self.vba_diag = d
            if d:
                self._vba_load(s, d[0]["module"], int(d[0]["line"] or 0))
                self.vba_status = f"Compile error in {d[0]['module']}, line {d[0]['line']}: {d[0]['message']}"
            else:
                self.vba_status = "No compile errors."

        return await self._vba_edit(f)

    @rx.event
    async def vba_goto(self, module: str, line: int):
        def f(s):
            self._vba_load(s, module, int(line or 0))

        return await self._vba_edit(f)

    @rx.event
    async def vba_run(self, code: str, proc: str):
        """F5 or Run: the Sub the cursor is in (a macro), else the Macros dialog; in a form's
        module, the form."""
        self.vba_changed(code)
        if self.vba_kind == "form":
            return await self.vba_show_form(self.vba_module)
        target = None
        for m in self.vba_macros:
            if m["module"].lower() == self.vba_module.lower() and m["proc"].lower() == (proc or "").lower():
                target = m["name"]
        if target is None and proc:
            # Not in the list yet (just typed): apply the code first, then look again.
            await self._vba_edit(lambda s: None)
            for m in self.vba_macros:
                if m["module"].lower() == self.vba_module.lower() and m["proc"].lower() == proc.lower():
                    target = m["name"]
        if target is None:
            return GridState.open_dialog("macros", {"module": self.vba_module})
        self.macro_label = target
        return await self._macro_step("macro.run", {"name": target}, apply_code=True)

    @rx.event
    async def vba_show_form(self, name: str):
        self.macro_label = name
        return await self._macro_step("macro.showForm", {"name": name}, apply_code=True)

    @rx.event
    async def form_event(self, e: Dict[str, Any]):
        """Something the person did on a form (or, in the editor's form view, a double-click)."""
        if not isinstance(e, dict):
            return None
        ev = e.get("event") if isinstance(e.get("event"), dict) else {}
        if ev.get("event") == "design":
            return await self.vba_form_handler(str(ev.get("control") or ""))
        try:
            fid = int(e.get("form") or 0)
        except (TypeError, ValueError):
            return None
        return await self._macro_step("macro.formEvent", {"form": fid, "event": ev})

    @rx.event
    async def vba_form_view_set(self, on: bool):
        """The editor's View Object / View Code for a form module."""

        def f(s):
            self._vba_form_design(s)

        self.vba_form_view = bool(on)
        if on:
            return await self._vba_edit(f)
        return None

    def _vba_form_design(self, s: Session) -> None:
        if self.vba_kind != "form" or not self.vba_module:
            self.vba_form_snapshot, self.vba_form_props, self.vba_form_warnings = {}, [], []
            return
        self._vba_form_shown(s.execute("macro.formDesign", {"name": self.vba_module, "control": self.vba_form_selected}))

    def _vba_form_shown(self, r: Dict[str, Any]) -> None:
        self.vba_form_snapshot = r.get("form") or {}
        self.vba_form_selected = str(r.get("control") or "")
        self.vba_form_props = list(r.get("props") or [])
        self.vba_form_editable = bool(r.get("editable"))
        self.vba_form_warnings = list(r.get("warnings") or [])

    async def _vba_form_edit(self, ops: List[Dict[str, Any]], control: Optional[str] = None):
        name = self.vba_module

        def f(s):
            r = s.execute("macro.editForm", {"name": name, "ops": ops, "control": self.vba_form_selected if control is None else control})
            self._vba_form_shown(r)
            self.vba_status = ""

        return await self._vba_edit(f)

    @rx.event
    async def form_designer_event(self, e: Dict[str, Any]):
        """What the person did in the form designer: select, move or resize, add from the toolbox,
        remove, bring up a page, or double-click for code."""
        if not isinstance(e, dict) or self.vba_kind != "form":
            return None
        kind = e.get("event")
        control = str(e.get("control") or "")
        if kind == "select":
            self.vba_form_selected = control

            def f(s):
                self._vba_form_design(s)

            return await self._vba_edit(f, apply_first=False)
        if kind == "code":
            return await self.vba_form_handler(control)
        if not self.vba_form_editable:
            return None
        if kind == "geometry":
            op = {"op": "geometry", "control": control}
            for k in ("left", "top", "width", "height"):
                if k in e:
                    op[k] = e[k]
            return await self._vba_form_edit([op], control)
        if kind == "page":
            return await self._vba_form_edit([{"op": "set", "control": control, "key": "position", "value": e.get("value")}], control)
        if kind == "add":
            self.vba_form_tool = ""
            op = {"op": "add", "kind": str(e.get("kind") or ""), "parent": str(e.get("parent") or "")}
            for k in ("left", "top", "width", "height"):
                if e.get(k) is not None:
                    op[k] = e[k]
            return await self._vba_form_edit([op])
        if kind == "delete" and control:
            return await self._vba_form_edit([{"op": "delete", "control": control}], "")
        return None

    @rx.event
    async def vba_form_set(self, key: str, value: Any):
        """A property changed in the Properties window (of the selected control, or the form)."""
        if not self.vba_form_editable:
            return None
        current = next((r for r in self.vba_form_props if r.get("key") == key), None)
        if current is not None and str(current.get("value")) == str(value):
            return None
        return await self._vba_form_edit([{"op": "set", "control": self.vba_form_selected, "key": key, "value": value}])

    @rx.event
    async def vba_form_submit(self, form: Dict[str, Any]):
        return await self.vba_form_set(str(form.get("key") or ""), form.get("value"))

    @rx.event
    def vba_form_tool_set(self, kind: str):
        self.vba_form_tool = "" if kind == self.vba_form_tool else kind

    @rx.event
    async def vba_form_handler(self, control: str):
        """Double-clicking a control in the form view opens its event procedure, written when the
        module doesn't have it yet (the VBA editor's own behaviour)."""

        def find(cs):
            for c in cs or []:
                if c.get("name") == control:
                    return c.get("kind", "")
                k = find(c.get("children"))
                if k:
                    return k
            return ""

        kind = find((self.vba_form_snapshot or {}).get("controls"))
        name = control or "UserForm"
        event = _DEFAULT_EVENTS.get(kind, "Click") if control else "Click"
        code = self.vba_code or ""
        lines = code.split("\n")
        want = re.compile(rf"^\s*(Private\s+|Public\s+)?Sub\s+{re.escape(name)}_{event}\s*\(", re.I)
        at = next((i for i, ln in enumerate(lines) if want.match(ln)), -1)
        if at < 0:
            body = code.rstrip("\n")
            stub = f"Private Sub {name}_{event}({_EVENT_ARGS.get(event, '')})"
            code = (body + "\n\n" if body else "") + stub + "\n    \nEnd Sub\n"
            self.vba_code = code
            self.vba_dirty = True
            self.vba_key += 1
            at = code.split("\n").index(stub)
        self.vba_form_view = False
        self.vba_goto_line = at + 2
        self.vba_goto_key += 1
        return None

    @rx.event
    async def vba_export(self):
        name = self.vba_module

        def f(s):
            r = s.execute("macro.exportModule", {"name": name})
            files = [rx.download(data=r["text"].encode("cp1252", "replace"), filename=r["fileName"], mime_type="text/plain")]
            if r.get("frx"):
                files.append(rx.download(data=base64.b64decode(r["frx"]), filename=r["frxName"], mime_type="application/octet-stream"))
            return files

        return await self._vba_edit(f)

    @rx.event
    async def vba_import(self, files: List[rx.UploadFile]):
        got: Dict[str, bytes] = {}
        for up in files[:10]:
            got[os.path.basename(up.name or up.filename or "Module1.bas")] = await up.read()
        frx = {n.lower()[:-4]: data for n, data in got.items() if n.lower().endswith(".frx")}
        for name, raw in got.items():
            if name.lower().endswith(".frx"):
                continue
            try:
                text = raw.decode("utf-8")
            except UnicodeDecodeError:
                text = raw.decode("cp1252", "replace")
            params: Dict[str, Any] = {"fileName": name, "text": text}
            if name.lower().endswith(".frm"):
                blob = frx.get(name.lower()[:-4])
                if blob is None:
                    self.vba_status = f"{name} is a UserForm: choose it together with its .frx file."
                    continue
                params["frx"] = base64.b64encode(blob).decode("ascii")

            def f(s, params=params, name=name):
                r = s.execute("macro.importModule", params)
                self._vba_load(s, r["name"])
                self.vba_status = f"Imported {name} as {r['name']}."

            out = await self._vba_edit(f)
            if out is not None:
                return out
        return None

    @rx.event
    async def open_vba_with(self, kind: str):
        """Developer › Module / Class Module: a new module, open in the editor."""
        entry, book, view = self._bound()
        if book is None:
            return None
        self.vba_open = True
        out = None
        room = await self.get_state(BookRoom)
        with book.act(view) as s:
            try:
                r = s.execute("macro.addModule", {"kind": kind})
                self._vba_load(s, r["name"])
            except EngineError as ex:
                out = rx.toast.error(str(ex).split("\n", 1)[0], position="bottom-center")
            self._refresh(entry, book, view)
        self._publish(room, book)
        return out

    @rx.event
    async def edit_macro(self, name: str):
        """The Macros dialog's Edit: the macro's module, at the macro (a new module when the
        name isn't a macro yet, as Excel's Create)."""
        entry, book, view = self._bound()
        if book is None:
            return None
        self.dialog, self.dlg, self.dlg_error = "", {}, ""
        room = await self.get_state(BookRoom)
        out = None
        with book.act(view) as s:
            try:
                macros = s.execute("macro.list", {}).get("macros", [])
                hit = next((m for m in macros if m["name"].lower() == (name or "").lower()), None)
                if hit is not None:
                    self._vba_load(s, hit["module"], int(hit.get("line") or 0))
                    self.vba_error_line = 0
                elif name and re.match(r"^[A-Za-z][A-Za-z0-9_]*$", name):
                    lst = s.execute("macro.list", {})
                    mod = next((m["name"] for m in lst.get("modules", []) if m["kind"] == "standard"), "")
                    if not mod:
                        mod = s.execute("macro.addModule", {"kind": "standard"})["name"]
                    code = s.execute("macro.getModule", {"name": mod})["code"]
                    code = (code.rstrip("\n") + "\n\n" if code.strip() else "") + f"Sub {name}()\n\nEnd Sub\n"
                    s.execute("macro.setModule", {"name": mod, "code": code})
                    self._vba_load(s, mod)
                else:
                    self._vba_load(s, "")
                self.vba_open = True
            except EngineError as ex:
                out = rx.toast.error(str(ex).split("\n", 1)[0], position="bottom-center")
            self._refresh(entry, book, view)
        self._publish(room, book)
        return out

    @rx.event
    def set_vba_immediate(self, text: str):
        self.vba_immediate = text

    @rx.event
    async def vba_immediate_run(self, form: Optional[Dict[str, Any]] = None):
        code = str((form or {}).get("code", self.vba_immediate) or "").strip()
        if not code:
            return None
        self.vba_output = (self.vba_output + ["> " + code])[-500:]
        self.vba_immediate = ""
        self.macro_label = "Immediate"
        return await self._macro_step("macro.immediate", {"code": code}, apply_code=True)

    @rx.event
    def vba_clear_output(self):
        self.vba_output = []

    @rx.event
    async def enable_macros(self, on: bool = True):
        """Event macros on (and Workbook_Open, as Excel runs it when content is enabled)."""
        room = await self.get_state(BookRoom)
        entry, book, view = self._bound()
        if book is None:
            return None
        out = None
        with book.act(view) as s:
            try:
                s.execute("macro.enableEvents", {"on": bool(on)})
                if on:
                    out = self._macro_after(s, s.execute("macro.open", {}))
            except EngineError as ex:
                out = rx.toast.error(str(ex).split("\n", 1)[0], position="bottom-center")
            self._refresh(entry, book, view)
        self._publish(room, book)
        return out

async def push(book_id: str) -> None:
    """Shows a book as it is now to everyone on it, from outside their own events (a tab that
    closed, a deletion, a save that finished): one of them publishes it to the room, and the
    room — a SharedState — carries it to all the others."""
    from reflex.state import _substate_key
    from reflex.utils.prerequisites import get_app

    app = get_app().app
    for token in sessions.tokens_on(book_id):
        try:
            async with app.modify_state(_substate_key(token, GridState)) as root:
                grid = await root.get_state(GridState)
                if grid.book != book_id:
                    continue
                room = await root.get_state(BookRoom)
                book = STORE.get(book_id)
                if book is None:
                    if not room.deleted:
                        room.deleted = True
                else:
                    grid._publish(room, book)
            return
        except Exception:  # that tab went away meanwhile: try the next one
            continue


STORE.on_saved = push


def _both(a, b):
    """Two event results as one (either may be None)."""
    if a is None:
        return b
    if b is None:
        return a
    return [*(a if isinstance(a, list) else [a]), b]


def _spans(idx: List[int], axis: str) -> str:
    idx = sorted(set(idx))
    out = []
    start = prev = idx[0]
    for i in idx[1:]:
        if i == prev + 1:
            prev = i
            continue
        out.append((start, prev))
        start = prev = i
    out.append((start, prev))
    if axis == "cols":
        return ",".join(f"{col_to_letters(a)}:{col_to_letters(b)}" for a, b in out)
    return ",".join(f"{a + 1}:{b + 1}" for a, b in out)


def _fmt_stat(v: float) -> str:
    if v is None:
        return ""
    if abs(v) >= 1e15 or (v != 0 and abs(v) < 1e-9):
        return f"{v:.6g}"
    s = f"{v:,.10f}".rstrip("0").rstrip(".")
    return s


def _js_str(s: str) -> str:
    import json

    return json.dumps(s)
