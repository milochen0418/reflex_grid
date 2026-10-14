"""The engine façade: documents, the command registry, undo/redo and the edit/commit cycle.

Every user-visible action is a command with a stable id (``cell.set``, ``home.bold``,
``data.sortAscending``…) and JSON-like parameters. The Reflex UI, tests, the CLI and agents all
go through :meth:`Session.execute`, so every action is scriptable and journaled.
"""

from __future__ import annotations

import itertools
import traceback
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Set, Tuple

from ..calc import Calc
from ..core.addr import RangeRef, parse_cell
from ..model.workbook import Workbook
from .selection import Selection

UNDO_LIMIT = 100
_uid = itertools.count(1)


class EngineError(Exception):
    """A command could not run (bad parameters, disabled, protected sheet…)."""


@dataclass
class HistoryEntry:
    label: str
    wb: Workbook
    selection: Selection
    sheet_index: int
    # Steps of one macro run share a group and undo as one.
    group: Optional[str] = None


@dataclass
class Clipboard:
    wb: Workbook
    sheet: int
    range: RangeRef
    cut: bool
    doc_uid: int
    text: str


class DocState:
    def __init__(self, wb: Workbook, name: str = "Book1", path: Optional[str] = None):
        self.wb = wb
        self.calc = Calc()
        self.calc.recalc_all(self.wb)
        self.saved = self.wb
        self.name = name
        self.path = path
        self.uid = next(_uid)
        self.revision = 1
        self.undo: List[HistoryEntry] = []
        self.redo: List[HistoryEntry] = []
        sh = wb.active()
        self.selection = Selection.at(sh.view_active if sh is not None else (0, 0))
        self.sheet_selections: Dict[int, Selection] = {}
        self.warnings: List[str] = []
        # Set by an edit that moved cells to other addresses (see ``Ctx.moved``); whoever needs
        # to know resets it (the collaboration layer, before each command).
        self.cells_moved = False
        # Whether the workbook's event macros (Worksheet_Change…) run when people edit it.
        self.macros_enabled = False
        # Sheets a macro protected with UserInterfaceOnly: protected from people, while macros
        # may still change them (lower-case names; like Excel, not saved with the file).
        self.ui_only_protection: Set[str] = set()

    def is_dirty(self) -> bool:
        return self.wb is not self.saved

    @property
    def sheet_index(self) -> int:
        return self.wb.active_sheet

    @property
    def sheet(self):
        return self.wb.sheets[self.wb.active_sheet]


@dataclass
class CommandSpec:
    id: str
    label: str
    run: Callable[["Session", Dict[str, Any]], Any]
    menu: Tuple[str, ...] = ()
    shortcut: Optional[str] = None
    params: str = ""
    undoable: bool = True
    journal: bool = True


COMMANDS: Dict[str, CommandSpec] = {}

# Called after every command that ran, as hook(session, command id, workbook before, selection
# before, result): gridcore.engine.macro runs the workbook's event macros there.
EVENT_HOOK: Optional[Callable[["Session", str, Optional[Workbook], Optional[Selection], Any], None]] = None
# Commands that carry on a paused macro run: what they do joins the run's undo step.
_RESUMING = ("macro.answer", "macro.continue", "macro.stop")


def command(cid: str, label: str, menu: Tuple[str, ...] = (), shortcut: Optional[str] = None, params: str = "", undoable: bool = True, journal: bool = True):
    def deco(fn):
        COMMANDS[cid] = CommandSpec(cid, label, fn, tuple(menu), shortcut, params, undoable, journal)
        return fn

    return deco


class Ctx:
    """An edit in progress: a copy of the workbook, the cells whose content changed, whether
    the edit was structural (rows, columns, sheets: full recalculation) and whether it moved
    cells to other addresses (rows or columns inserted or deleted, cells shifted or moved,
    sorted, sheets added, removed or reordered)."""

    def __init__(self, wb: Workbook, sel: Selection):
        self.wb = wb
        self.changed: List[Tuple[int, int, int]] = []
        self.structural = False
        self.recalc = False  # recalculate everything, even in manual mode (F9)
        self.unprotected: Set[str] = set()  # protected sheets this edit may change anyway (macros)
        self.moved = False
        self.sel = sel
        self.fit_rows: List[Tuple[int, int]] = []

    @property
    def sheet_index(self) -> int:
        return self.wb.active_sheet

    def sheet_mut(self, i: Optional[int] = None):
        if i is None:
            i = self.wb.active_sheet
        s = self.wb.sheet_mut(i)
        if s is None:
            raise EngineError("no such sheet")
        return s

    def touch(self, sheet: int, rc: Tuple[int, int]) -> None:
        self.changed.append((sheet, rc[0], rc[1]))


class Session:
    def __init__(self) -> None:
        self.docs: List[DocState] = []
        self.active = 0
        self.clipboard: Optional[Clipboard] = None
        self.journal: List[Tuple[str, Dict[str, Any]]] = []
        self.prefs: Dict[str, Any] = {"enterDirection": "down", "userName": "ReflexGrid User"}
        # Messages for the UI (errors, info), drained by the UI layer.
        self.messages: List[str] = []
        # Format painter: (workbook, sheet, range) of the copied formats.
        self.format_painter: Optional[Tuple[Workbook, int, RangeRef]] = None
        # Last Find query (for Find Next).
        self.last_find: Optional[Dict[str, Any]] = None
        # Macros: the run waiting for an answer or its next time slice, whether runs may pause
        # for the person (a UI drives them), and the undo group of the run in progress.
        self.macro_run = None
        self.interactive_macros = False
        self.undo_group: Optional[str] = None
        # The UserForms this person's macros loaded (gridcore.vba.msforms.FormRegistry), and
        # the text macros put on the clipboard (MSForms.DataObject).
        self.vba_forms = None
        self.vba_clipboard: Optional[str] = None

    # ------------------------------------------------------------ documents

    def doc(self) -> DocState:
        if not self.docs:
            raise EngineError("no workbook open")
        return self.docs[self.active]

    def add_document(self, d: DocState) -> int:
        self.docs.append(d)
        self.active = len(self.docs) - 1
        return self.active

    def new_workbook(self, sheets: int = 1) -> DocState:
        n = 1
        names = {d.name for d in self.docs}
        while f"Book{n}" in names:
            n += 1
        wb = Workbook()
        for _ in range(1, max(sheets, 1)):
            from ..model.sheet import Sheet

            wb.add_sheet(Sheet(wb.next_sheet_name()))
        d = DocState(wb, f"Book{n}")
        self.add_document(d)
        return d

    # ------------------------------------------------------------ commands

    def execute(self, cid: str, params: Optional[Dict[str, Any]] = None) -> Any:
        spec = COMMANDS.get(cid)
        if spec is None:
            raise EngineError(f"unknown command `{cid}`")
        params = params or {}
        if cid not in _RESUMING:
            self.undo_group = None
        before = None
        if self.docs:
            d = self.doc()
            before = (d.uid, d.wb, d.selection.copy())
        undoable, label = spec.undoable, spec.label
        try:
            result = spec.run(self, params)
            hook = EVENT_HOOK
            if hook is not None and before is not None and self.docs:
                d = self.doc()
                mid = (d.uid, d.wb, d.selection.copy())
                hook(self, cid, before[1], before[2], result)
                if not undoable and mid[0] == before[0] and d.wb is not mid[1]:
                    # An event macro changed the workbook after a step that isn't an edit (a
                    # new selection, another sheet): undo takes back what the macro did.
                    undoable, label, before = True, "Macro", mid
        except EngineError:
            raise
        except Exception as e:  # never crash: restore the workbook as it was
            if before is not None:
                for d in self.docs:
                    if d.uid == before[0]:
                        d.wb = before[1]
                        d.selection = before[2]
                        d.calc.rebuild(d.wb)
            raise EngineError(f"internal error in `{cid}` (the workbook was kept as it was): {e}\n{traceback.format_exc(limit=6)}") from e
        if spec.journal:
            self.journal.append((cid, {k: v for k, v in params.items() if "password" not in k.lower()}))
            if len(self.journal) > 10_000:
                del self.journal[:5_000]
        if undoable and before is not None:
            for d in self.docs:
                if d.uid == before[0] and d.wb is not before[1]:
                    group = self.undo_group
                    if group is not None and d.undo and d.undo[-1].group == group:
                        # The next step of the same macro run: one undo takes back the whole run.
                        d.redo.clear()
                        d.revision += 1
                        continue
                    d.undo.append(HistoryEntry(label.rstrip("…"), before[1], before[2], before[1].active_sheet, group))
                    if len(d.undo) > UNDO_LIMIT:
                        del d.undo[0]
                    d.redo.clear()
                    d.revision += 1
        return result

    def run(self, cid: str, **params) -> Any:
        return self.execute(cid, params)


def edit(s: Session, f: Callable[[Ctx], Any]) -> Any:
    """Applies an edit to a copy of the active workbook, then recalculates and commits it."""
    d = s.doc()
    return commit(d, f)


def commit(d: DocState, f: Callable[[Ctx], Any]) -> Any:
    ctx = begin_edit(d)
    r = f(ctx)
    finish_edit(d, ctx)
    return r


def begin_edit(d: DocState) -> Ctx:
    """An edit on a copy of the document's workbook; :func:`finish_edit` recalculates and
    commits it. (Macros keep one open across many cell writes.)"""
    return Ctx(d.wb.clone(), d.selection.copy())


def finish_edit(d: DocState, ctx: Ctx) -> None:
    wb = ctx.wb
    if ctx.changed and len(ctx.changed) <= 200_000:
        from .tables import sync_headers

        sync_headers(wb, ctx.changed)
    check_protection(d.wb, wb, ctx.unprotected)
    if ctx.structural or ctx.recalc:
        if wb.calc.mode == "manual" and not ctx.recalc:
            d.calc.rebuild(wb)
        else:
            d.calc.recalc_all(wb)
    elif ctx.changed:
        d.calc.cells_changed(wb, ctx.changed)
    rows = set(ctx.fit_rows)
    if len(ctx.changed) <= 20000:
        rows.update((k[0], k[1]) for k in ctx.changed)
    from .layout import auto_row_height

    for si, row in sorted(rows):
        auto_row_height(wb, si, row)
    d.wb = wb
    d.selection = ctx.sel
    if ctx.moved:
        d.cells_moved = True


PROTECTED = "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet (Review › Unprotect Sheet)."


def same_content(a, b) -> bool:
    """Two cells hold the same value or formula (formats aside)."""

    def empty(c) -> bool:
        return c is None or (c.formula is None and c.value is None)

    if empty(a) and empty(b):
        return True
    if a is None or b is None:
        return False
    if (a.formula is None) != (b.formula is None):
        return False
    if a.formula is not None:
        return a.formula == b.formula
    return type(a.value) is type(b.value) and a.value == b.value


def check_protection(old: Workbook, new: Workbook, exempt: Optional[Set[str]] = None) -> None:
    """Refuses an edit that changed what a locked cell on a protected sheet holds (``exempt``:
    lower-case names of sheets the edit may change anyway)."""
    for sh in old.sheets:
        if sh.protection is None or (exempt and sh.name.lower() in exempt):
            continue
        i = new.sheet_index(sh.name)
        if i is None:
            continue
        after = new.sheets[i]
        if after is sh or after.protection is None:
            continue
        for rc in sh.cells.diff_keys(after.cells):
            if not same_content(sh.cell(rc), after.cell(rc)) and old.styles.get(sh.style_id(rc)).protection.locked:
                raise EngineError(PROTECTED)


# ---------------------------------------------------------------- param helpers


def param_ranges(s: Session, p: Dict[str, Any], key: str = "range") -> List[RangeRef]:
    """A range parameter (``"A1:B2"``, ``"B:B"``, several comma-separated), defaulting to the
    current selection."""
    t = p.get(key) or p.get("cell")
    if t:
        body = t.rsplit("!", 1)[-1] if isinstance(t, str) else t
        out = []
        for part in str(body).split(","):
            r = RangeRef.parse(part)
            if r is None:
                raise EngineError(f"not a range: `{part}`")
            out.append(r)
        return out
    return list(s.doc().selection.ranges)


def param_range(s: Session, p: Dict[str, Any], key: str = "range") -> RangeRef:
    return param_ranges(s, p, key)[0]


def param_cell(s: Session, p: Dict[str, Any], key: str = "cell") -> Tuple[int, int]:
    t = p.get(key)
    if t is None:
        return s.doc().selection.active
    if isinstance(t, (list, tuple)) and len(t) == 2:
        return int(t[0]), int(t[1])
    c = parse_cell(str(t).rsplit("!", 1)[-1])
    if c is None:
        raise EngineError(f"invalid cell reference `{t}`")
    return c


def param_sheet(s: Session, p: Dict[str, Any]) -> int:
    d = s.doc()
    v = p.get("sheet")
    if v is None:
        t = p.get("range") or p.get("cell")
        if isinstance(t, str) and "!" in t:
            name = t.rsplit("!", 1)[0].strip("'").replace("''", "'")
            i = d.wb.sheet_index(name)
            if i is None:
                raise EngineError(f"no sheet named `{name}`")
            return i
        return d.wb.active_sheet
    if isinstance(v, int):
        if not 0 <= v < len(d.wb.sheets):
            raise EngineError("no such sheet")
        return v
    i = d.wb.sheet_index(str(v))
    if i is None:
        raise EngineError(f"no sheet named `{v}`")
    return i


def p_bool(p: Dict[str, Any], key: str, default: bool = False) -> bool:
    v = p.get(key)
    if v is None:
        return default
    if isinstance(v, str):
        return v.strip().lower() in ("1", "true", "yes", "on")
    return bool(v)


def p_int(p: Dict[str, Any], key: str, default: Optional[int] = None) -> Optional[int]:
    v = p.get(key)
    if v is None or v == "":
        return default
    try:
        return int(float(v))
    except (TypeError, ValueError):
        raise EngineError(f"`{key}` must be a number")


def p_float(p: Dict[str, Any], key: str, default: Optional[float] = None) -> Optional[float]:
    v = p.get(key)
    if v is None or v == "":
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        raise EngineError(f"`{key}` must be a number")


def p_str(p: Dict[str, Any], key: str, default: Optional[str] = None) -> Optional[str]:
    v = p.get(key)
    if v is None:
        return default
    return str(v)


def need(p: Dict[str, Any], key: str, cid: str) -> Any:
    if p.get(key) is None:
        raise EngineError(f"{cid}: missing `{key}`")
    return p[key]
