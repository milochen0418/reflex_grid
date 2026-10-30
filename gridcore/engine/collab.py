"""Several people editing one workbook at once.

A :class:`SharedBook` holds the workbook everybody edits — one :class:`DocState` with its
snapshots, calculation and names — and a :class:`View` per person (per browser tab): the sheet
they look at, their selection, zoom and formulas view, and their own undo history. Each
person's engine work runs against the shared workbook inside :meth:`SharedBook.act`, one person
at a time:

    view = book.join("tab-1", CollabSession(), name="Amy")
    with book.act(view) as s:          # s: Amy's session, bound to the shared workbook
        s.run("cell.set", cell="B2", input="=A1*2")

Each change bumps :attr:`SharedBook.rev` and is described in :attr:`SharedBook.changes`, so the
UI knows when, and where, to redraw everybody's view.

Undo is per person, as in Google Sheets or Excel co-authoring: it takes back your own changes,
not the ones other people made since. When nobody has changed the workbook after you, undo puts
back the snapshot from before your change, whatever the change was. Otherwise it puts back only
what you changed and nobody has touched since — cell by cell, and feature by feature for the
rest (row heights, merges, rules, tables…); what others changed meanwhile is kept. That matching
goes by address, so once someone has moved cells (rows or columns inserted or deleted, cells
shifted or moved, a sort, sheets added, removed or reordered), changes made before can't be
undone that way any more, and undo says so.
"""

from __future__ import annotations

import copy
import itertools
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Iterable, Iterator, List, Optional, Tuple

from ..model.cell import Cell
from ..model.workbook import Workbook
from .selection import Selection
from .session import DocState, EngineError, Session, commit, p_bool, p_int, param_sheet

UNDO_LIMIT = 100
# Changes kept for the UI (who changed what, for highlighting).
CHANGE_LOG = 50
# How ``changed_areas`` tells changes apart: past these, one box around them all.
AREA_BOXES, AREA_CELLS = 8, 20_000

# Commands on the server's files or on the session's own list of workbooks: a shared workbook
# is created, opened, saved and closed by the app around it.
BLOCKED = frozenset({"file.new", "file.open", "file.close", "file.save", "file.saveAs", "window.activate"})

# Sheet attributes undo leaves alone: the cells (merged one by one), the calculation's output,
# caches, and what each person sees for themselves.
_SHEET_SKIP = frozenset({"cells", "spill", "spill_ranges", "_row_index", "_col_index", "view_active", "view_top_left", "zoom"})
# Sheet attributes merged entry by entry: dicts keyed by row, column or cell.
_SHEET_KEYED = ("rows", "cols", "comments", "hyperlinks")
# Workbook attributes undo puts back as a whole.
_BOOK_ATTRS = ("names", "props", "calc", "protected_structure", "structure_password", "cell_styles", "date_system", "theme", "csv_delimiter", "vba", "code_name")


@dataclass
class Transition:
    """One undoable step of one person: the workbook went from ``src`` to ``dst``. Selections
    and sheets are where that person was at each end; epochs tell whether cells moved since."""

    label: str
    src: Workbook
    dst: Workbook
    sel_src: Selection
    sel_dst: Selection
    sheet_src: str
    sheet_dst: str
    epoch_src: int
    epoch_dst: int
    moved: bool
    # Steps of one macro run share a group: they undo as one.
    group: Optional[str] = None


@dataclass
class Change:
    """What one action changed: ``areas`` are ``(sheet name, [r0, c0, r1, c1])`` of the cells whose
    content or format changed (empty when cells moved, or only rows, columns or features did)."""

    rev: int
    author: str
    label: str
    areas: List[Tuple[str, List[int]]]
    moved: bool
    at: float = field(default_factory=time.time)

    def to_json(self) -> Dict[str, Any]:
        return {"rev": self.rev, "author": self.author, "label": self.label, "areas": [[s, list(r)] for s, r in self.areas], "moved": self.moved}


class View:
    """One person's window into a shared book. ``info`` is free for the app (colour, avatar…)."""

    def __init__(self, key: str, session: "CollabSession", name: str = ""):
        self.key = key
        self.session = session
        self.name = name
        self.info: Dict[str, Any] = {}
        # The sheet on screen, by name, and its position (to land nearby when it is deleted).
        self.sheet = ""
        self.sheet_index = 0
        self.selection = Selection.at((0, 0))
        # The last selection on the other sheets, by sheet name.
        self.selections: Dict[str, Selection] = {}
        # Zoom and formulas view per sheet name, overriding what the file says.
        self.zoom: Dict[str, int] = {}
        self.formulas: Dict[str, bool] = {}
        self.undo: List[Transition] = []
        self.redo: List[Transition] = []
        self.last_seen = time.time()

    def zoom_of(self, sheet) -> int:
        return int(self.zoom.get(sheet.name, sheet.zoom))

    def formulas_of(self, sheet) -> bool:
        return bool(self.formulas.get(sheet.name, sheet.show_formulas))


class CollabSession(Session):
    """A person's session: clipboard, Find, Format Painter and preferences of their own, with a
    shared book as its document. Commands run inside :meth:`SharedBook.act`; undo, redo, zoom and
    the formulas view belong to the person."""

    def __init__(self) -> None:
        super().__init__()
        self.book: Optional[SharedBook] = None
        self.view: Optional[View] = None

    def execute(self, cid: str, params: Optional[Dict[str, Any]] = None) -> Any:
        book, view = self.book, self.view
        if book is None or view is None:
            return super().execute(cid, params)
        if cid in BLOCKED:
            raise EngineError(f"`{cid}` isn't available in a shared workbook.")
        if book._act is None or book._act.view is not view:
            raise EngineError("internal error: commands on a shared workbook run inside SharedBook.act()")
        p = params or {}
        if cid == "edit.undo":
            return book._step(view, undo=True, steps=max(p_int(p, "steps", 1) or 1, 1))
        if cid == "edit.redo":
            return book._step(view, undo=False, steps=max(p_int(p, "steps", 1) or 1, 1))
        if cid == "view.zoom":
            z = p_int(p, "percent", 100)
            if not 10 <= z <= 400:
                raise EngineError("Zoom must be between 10% and 400%.")
            view.zoom[self.doc().wb.sheets[param_sheet(self, p)].name] = z
            return {"zoom": z}
        if cid == "view.scroll":
            return None
        if cid == "formulas.showFormulas":
            sh = self.doc().wb.sheets[param_sheet(self, p)]
            on = p_bool(p, "on", not view.formulas_of(sh))
            view.formulas[sh.name] = on
            return {"show_formulas": on}
        return super().execute(cid, params)


class _Act:
    """One ``act`` in progress: whose, and the workbook and epoch it started from."""

    def __init__(self, view: View, before: Workbook, epoch: int):
        self.view = view
        self.before = before
        self.epoch = epoch


class SharedBook:
    """The workbook several people edit, with everyone's :class:`View`."""

    def __init__(self, doc: DocState, key: str = "") -> None:
        doc.undo.clear()
        doc.redo.clear()
        self.key = key
        self.doc = doc
        self.lock = threading.RLock()
        # Bumped by every change; the UI compares it to what it drew.
        self.rev = 0
        # Changes when cells move to other addresses (undo can't match them across that).
        self.epoch = 0
        self._epochs = itertools.count(1)
        self.views: Dict[str, View] = {}
        self.changes: List[Change] = []
        # Called (under the lock) after every change, e.g. to schedule saving.
        self.listeners: List[Callable[["SharedBook"], None]] = []
        sh = doc.wb.active() or doc.wb.sheets[0]
        # Where newcomers start: the sheet the file was saved on.
        self.home_sheet = sh.name
        self._act: Optional[_Act] = None

    # ------------------------------------------------------------ people

    def join(self, key: str, session: CollabSession, name: str = "", view: Optional[View] = None) -> View:
        """Adds a person (or finds them again) and binds their session to this book. ``view`` is
        the view they had before leaving (a tab that reloads), taken back with its history."""
        with self.lock:
            v = self.views.get(key)
            if v is None and view is not None and view.key == key:
                v = self.views[key] = view
            if v is None:
                v = View(key, session, name)
                wb = self.doc.wb
                v.sheet = self.home_sheet
                i = self._resolve(wb, v)
                v.sheet, v.sheet_index = wb.sheets[i].name, i
                v.selection = Selection.at(wb.sheets[i].view_active)
                self.views[key] = v
            if v.session is not session and v.session.view is v:
                _unbind(v.session)
            v.session = session
            if name:
                v.name = name
                session.prefs["userName"] = name
            v.last_seen = time.time()
            self._bind(v)
            return v

    def _bind(self, view: View) -> None:
        s = view.session
        if s.book is not self or s.view is not view:
            s.book, s.view = self, view
            s.docs = [self.doc]
            s.active = 0

    def leave(self, key: str) -> None:
        with self.lock:
            v = self.views.pop(key, None)
            if v is not None and v.session.view is v:
                _unbind(v.session)

    def prune(self, idle: float, now: Optional[float] = None) -> List[str]:
        """Drops the people not seen for ``idle`` seconds; returns their keys."""
        now = time.time() if now is None else now
        with self.lock:
            gone = [k for k, v in self.views.items() if now - v.last_seen > idle]
            for k in gone:
                self.leave(k)
            return gone

    # ------------------------------------------------------------ acting

    @contextmanager
    def act(self, view: View) -> Iterator[CollabSession]:
        """Runs a person's engine work on the shared workbook: their sheet and selection are put
        in place, and whatever changed is recorded (revision, change log, their undo history)."""
        with self.lock:
            if self._act is not None:
                raise RuntimeError("SharedBook.act() can't be nested")
            # Someone dropped as idle who comes back is simply in again.
            if self.views.get(view.key) is not view:
                self.views[view.key] = view
            self._bind(view)
            d = self.doc
            self._enter(view)
            view.last_seen = time.time()
            act = self._act = _Act(view, d.wb, self.epoch)
            d.undo, d.redo = [], []
            d.cells_moved = False
            try:
                yield view.session
            finally:
                self._act = None
                hist = d.undo
                d.undo, d.redo = [], []
                moved = d.cells_moved
                d.cells_moved = False
                self._leave(view)
                if d.wb is not act.before:
                    self._record(view, act, hist, moved)

    def discard(self) -> None:
        """Drops what the current act did so far: the workbook, and its person's place on it, are
        as the act found them."""
        act, d = self._act, self.doc
        d.undo, d.redo, d.cells_moved = [], [], False
        if d.wb is not act.before:
            d.wb = act.before
            d.calc.rebuild(d.wb)
        self._enter(act.view)

    def _flush(self) -> None:
        """Records what the current act changed so far (undo and redo record their own steps)."""
        act = self._act
        d = self.doc
        if d.wb is not act.before:
            hist = d.undo
            d.undo = []
            moved = d.cells_moved
            d.cells_moved = False
            self._leave(act.view)
            self._record(act.view, act, hist, moved)
        act.before = d.wb
        act.epoch = self.epoch

    def _record(self, view: View, act: _Act, hist: list, moved: bool) -> None:
        before, after = act.before, self.doc.wb
        if not moved:
            self._follow_renames(before, after)
        if moved:
            self.epoch = next(self._epochs)
        self.rev += 1
        label = hist[-1].label if hist else "Edit"
        self._log(view, before, after, label, moved)
        if hist:
            h = hist[0]
            group = hist[-1].group
            last = view.undo[-1] if view.undo else None
            if group is not None and last is not None and last.group == group and last.dst is before and not moved and not last.moved and last.epoch_dst == act.epoch:
                # The next step of a macro run nobody interrupted: one transition for the run.
                view.undo[-1] = Transition(last.label, last.src, after, last.sel_src, view.selection.copy(), last.sheet_src, view.sheet, last.epoch_src, self.epoch, False, group)
            else:
                sheet_src = before.sheets[h.sheet_index].name if 0 <= h.sheet_index < len(before.sheets) else view.sheet
                view.undo.append(Transition(label, before, after, h.selection, view.selection.copy(), sheet_src, view.sheet, act.epoch, self.epoch, moved, group))
            del view.undo[:-UNDO_LIMIT]
            view.redo.clear()
        self._notify()

    def _follow_renames(self, before: Workbook, after: Workbook) -> None:
        """A renamed sheet keeps everyone on it (views find their sheet by name)."""
        if len(before.sheets) != len(after.sheets):
            return
        renames = {a.name: b.name for a, b in zip(before.sheets, after.sheets) if a.name != b.name}
        if not renames:
            return
        for v in self.views.values():
            v.sheet = renames.get(v.sheet, v.sheet)
            for m in (v.selections, v.zoom, v.formulas):
                for old, new in renames.items():
                    if old in m:
                        m[new] = m.pop(old)

    def _log(self, view: View, before: Workbook, after: Workbook, label: str, moved: bool) -> None:
        areas: List[Tuple[str, List[int]]] = []
        if not moved and len(before.sheets) == len(after.sheets):
            for sb, sa in zip(before.sheets, after.sheets):
                if sb is sa:
                    continue
                box = _changed_box(sb, sa)
                if box is not None:
                    areas.append((sa.name, box))
        self.changes.append(Change(self.rev, view.key, label, areas, moved))
        del self.changes[:-CHANGE_LOG]

    def _notify(self) -> None:
        for fn in list(self.listeners):
            try:
                fn(self)
            except Exception:  # a listener must not break an edit
                pass

    # ------------------------------------------------------------ views

    @staticmethod
    def _resolve(wb: Workbook, view: View) -> int:
        """The sheet a view is on: by name, else at the same position (it was deleted), and
        never a hidden one."""
        i = wb.sheet_index(view.sheet) if view.sheet else None
        if i is None:
            i = min(max(view.sheet_index, 0), len(wb.sheets) - 1)
        if wb.sheets[i].visibility != "visible":
            shown = [j for j, s in enumerate(wb.sheets) if s.visibility == "visible"]
            if shown:
                i = min(shown, key=lambda j: (abs(j - i), j))
        return i

    def _enter(self, view: View) -> None:
        d = self.doc
        wb = d.wb
        i = self._resolve(wb, view)
        sh = wb.sheets[i]
        if sh.name != view.sheet:
            # The sheet on screen is gone (or hidden): carry on with this one.
            view.selection = view.selections.pop(sh.name, None) or Selection.at(sh.view_active)
            view.sheet = sh.name
        view.sheet_index = i
        wb.active_sheet = i
        d.selection = view.selection
        d.sheet_selections = {j: view.selections[s.name] for j, s in enumerate(wb.sheets) if j != i and s.name in view.selections}

    def _leave(self, view: View) -> None:
        d = self.doc
        wb = d.wb
        i = min(max(wb.active_sheet, 0), len(wb.sheets) - 1)
        view.sheet_index = i
        view.sheet = wb.sheets[i].name
        view.selection = d.selection
        view.selections = {wb.sheets[j].name: s for j, s in d.sheet_selections.items() if 0 <= j < len(wb.sheets) and j != i}

    # ------------------------------------------------------------ undo, redo

    def _step(self, view: View, undo: bool, steps: int) -> Dict[str, Any]:
        stack, other = (view.undo, view.redo) if undo else (view.redo, view.undo)
        verb = "undo" if undo else "redo"
        if not stack:
            raise EngineError(f"Nothing to {verb}.")
        self._flush()
        label = ""
        kept = 0
        for _ in range(steps):
            if not stack:
                break
            t = stack[-1]
            if self.doc.wb is t.dst:
                back = self._exact(t)
            elif t.moved or t.epoch_dst != self.epoch:
                if label:
                    break
                raise EngineError(
                    f"Can't {verb} “{t.label}”: someone has moved cells since "
                    "(inserted or deleted rows or columns, sorted, or added, removed or moved sheets)."
                )
            else:
                back, k = self._selective(t)
                kept += k
            stack.pop()
            other.append(back)
            del other[:-UNDO_LIMIT]
            label = t.label
        if kept:
            view.session.messages.append(
                f"{kept} of the changes you {'undid' if undo else 'redid'} had been changed again by someone else since; theirs were kept."
            )
        return {"undone" if undo else "redone": label, "kept": kept}

    def take_back(self, view: View, step: Transition) -> Dict[str, Any]:
        """Undoes one of ``view``'s steps, not only the latest (inside :meth:`act`): what the
        person did after it stays. As with any undo, what was changed again since is kept, and
        once cells have moved a step can't be matched any more."""
        if not any(t is step for t in view.undo):
            raise EngineError("That was undone already.")
        self._flush()
        kept = 0
        if self.doc.wb is step.dst:
            back = self._exact(step)
        elif step.moved or step.epoch_dst != self.epoch:
            raise EngineError(f"Can't undo “{step.label}”: cells have moved since (rows or columns inserted or deleted, a sort, sheets added or removed).")
        else:
            back, kept = self._selective(step)
        view.undo[:] = [t for t in view.undo if t is not step]
        view.redo.append(back)
        del view.redo[:-UNDO_LIMIT]
        if kept:
            view.session.messages.append(f"{kept} of the changes you undid had been changed again since; those were kept.")
        return {"undone": step.label, "kept": kept}

    def _exact(self, t: Transition) -> Transition:
        """Nobody changed the workbook since ``t``: go back to its snapshot."""
        d = self.doc
        cur, epoch0 = d.wb, self.epoch
        d.wb = t.src
        d.calc.rebuild(d.wb)
        self.epoch = t.epoch_src
        return self._reverted(t, cur, epoch0)

    def _selective(self, t: Transition) -> Tuple[Transition, int]:
        """Others changed the workbook since ``t``: put back only what nobody touched since."""
        d = self.doc
        cur, epoch0 = d.wb, self.epoch
        kept: List[int] = []
        commit(d, lambda ctx: kept.append(_put_back(ctx, t, cur)))
        return self._reverted(t, cur, epoch0), kept[0] if kept else 0

    def _reverted(self, t: Transition, cur: Workbook, epoch0: int) -> Transition:
        d = self.doc
        act = self._act
        view = act.view
        i = d.wb.sheet_index(t.sheet_src)
        if i is not None and d.wb.sheets[i].visibility == "visible":
            if i != d.wb.active_sheet:
                d.sheet_selections[d.wb.active_sheet] = d.selection
                d.sheet_selections.pop(i, None)
            d.wb.active_sheet = i
            d.selection = t.sel_src.copy()
        self._leave(view)
        if not t.moved:
            self._follow_renames(cur, d.wb)
        self.rev += 1
        self._log(view, cur, d.wb, t.label, t.moved)
        self._notify()
        act.before = d.wb
        act.epoch = self.epoch
        return Transition(t.label, cur, d.wb, t.sel_dst, t.sel_src, t.sheet_dst, t.sheet_src, epoch0, self.epoch, t.moved)

    def can_undo(self, view: View) -> bool:
        return bool(view.undo)

    def can_redo(self, view: View) -> bool:
        return bool(view.redo)


def _unbind(session: CollabSession) -> None:
    session.book = None
    session.view = None
    session.docs = []
    session.active = 0


# ---------------------------------------------------------------- comparing snapshots


def _same_input(a: Optional[Cell], b: Optional[Cell]) -> bool:
    """Same formula or constant and the same format: the cells differ at most in a cached result."""
    if a is None or b is None:
        return (a is None or a.is_blank()) and (b is None or b.is_blank())
    if a.style != b.style or a.formula != b.formula:
        return False
    return a.formula is not None or (type(a.value) is type(b.value) and a.value == b.value)


def _eq(a: Any, b: Any) -> bool:
    if a is b:
        return True
    try:
        return bool(a == b)
    except Exception:
        return False


def _changed_cells(before, after) -> Iterator[Tuple[int, int]]:
    """The cells whose content or format differs between two versions of a sheet."""
    for k in before.cells.diff_keys(after.cells):
        if not _same_input(before.cell(k), after.cell(k)):
            yield k


def _changed_box(before, after) -> Optional[List[int]]:
    """Bounding box of the cells that differ between two versions of a sheet."""
    r0 = c0 = r1 = c1 = -1
    for r, c in _changed_cells(before, after):
        if r0 < 0:
            r0, c0, r1, c1 = r, c, r, c
        else:
            r0, c0, r1, c1 = min(r0, r), min(c0, c), max(r1, r), max(c1, c)
    return None if r0 < 0 else [r0, c0, r1, c1]


def _rectangles(cells: Iterable[Tuple[int, int]]) -> List[List[int]]:
    """Cells as rectangles, neighbours together; as one when there are many or they scatter."""
    cells = list(cells)
    if not cells:
        return []
    rows, cols = [r for r, _ in cells], [c for _, c in cells]
    whole = [min(rows), min(cols), max(rows), max(cols)]
    if len(cells) > AREA_CELLS:
        return [whole]
    boxes: List[List[int]] = []
    for r, c in cells:
        box = [r, c, r, c]
        for b in [b for b in boxes if b[0] - 1 <= r <= b[2] + 1 and b[1] - 1 <= c <= b[3] + 1]:
            boxes.remove(b)
            box = [min(box[0], b[0]), min(box[1], b[1]), max(box[2], b[2]), max(box[3], b[3])]
        boxes.append(box)
        if len(boxes) > AREA_BOXES:
            return [whole]
    return boxes


def changed_areas(before: Workbook, after: Workbook) -> List[Tuple[str, List[int]]]:
    """Where two versions of a workbook differ, moved cells and what a changed formula spills
    included: ``(sheet name, [r0, c0, r1, c1])``, several for a sheet when the changes lie apart,
    sheets matched by name (a new sheet counts as a whole)."""
    was = {s.name: s for s in before.sheets}
    out: List[Tuple[str, List[int]]] = []
    for sheet in after.sheets:
        old = was.get(sheet.name)
        if old is sheet:
            continue
        if old is None:
            used = sheet.used_range()
            if used is not None:
                out.append((sheet.name, [used.r0, used.c0, used.r1, used.c1]))
            continue
        for box in _rectangles(_changed_cells(old, sheet)):
            for (r, c), sr in sheet.spill_ranges.items():
                if box[0] <= r <= box[2] and box[1] <= c <= box[3]:
                    box = [min(box[0], sr.r0), min(box[1], sr.c0), max(box[2], sr.r1), max(box[3], sr.c1)]
            out.append((sheet.name, box))
    return out


def _put_back(ctx, t: Transition, cur: Workbook) -> int:
    """In ``ctx.wb`` (a copy of ``cur``), undoes what went from ``t.src`` to ``t.dst`` wherever
    ``cur`` still holds what ``t.dst`` had. Returns how many parts were kept (changed since)."""
    kept = 0
    src, dst = t.src, t.dst
    for i, (ss, sd) in enumerate(zip(src.sheets, dst.sheets)):
        if ss is sd or i >= len(cur.sheets):
            continue
        sc = cur.sheets[i]
        cells: List[Tuple[Tuple[int, int], Optional[Cell]]] = []
        for k in ss.cells.diff_keys(sd.cells):
            a, b = ss.cell(k), sd.cell(k)
            if _same_input(a, b):
                continue
            if _same_input(sc.cell(k), b):
                cells.append((k, a))
            else:
                kept += 1
        keyed: Dict[str, List[Tuple[Any, Any]]] = {}
        for attr in _SHEET_KEYED:
            ms, md, mc = getattr(ss, attr), getattr(sd, attr), getattr(sc, attr)
            if ms is md:
                continue
            puts = []
            for k in set(ms) | set(md):
                vs, vd = ms.get(k), md.get(k)
                if _eq(vs, vd):
                    continue
                if _eq(mc.get(k), vd):
                    puts.append((k, vs))
                else:
                    kept += 1
            if puts:
                keyed[attr] = puts
        whole: Dict[str, Any] = {}
        for attr, vs in vars(ss).items():
            if attr in _SHEET_SKIP or attr in _SHEET_KEYED:
                continue
            vd = getattr(sd, attr, None)
            if _eq(vs, vd):
                continue
            if _eq(getattr(sc, attr, None), vd):
                whole[attr] = copy.deepcopy(vs)
            else:
                kept += 1
        if not (cells or keyed or whole):
            continue
        sm = ctx.sheet_mut(i)
        for k, a in cells:
            if a is None:
                sm.cells.remove(k)
            else:
                sm.cells.set(k, a)
            ctx.touch(i, k)
        for attr, puts in keyed.items():
            m = getattr(sm, attr)
            for k, v in puts:
                if v is None:
                    m.pop(k, None)
                else:
                    m[k] = v
        for attr, v in whole.items():
            setattr(sm, attr, v)
        if keyed.get("rows") or keyed.get("cols") or whole:
            sm._invalidate_geometry()
        if whole:
            ctx.structural = True
    for attr in _BOOK_ATTRS:
        vs, vd = getattr(src, attr, None), getattr(dst, attr, None)
        if _eq(vs, vd):
            continue
        if _eq(getattr(cur, attr, None), vd):
            setattr(ctx.wb, attr, copy.deepcopy(vs))
            ctx.structural = True
        else:
            kept += 1
    return kept
