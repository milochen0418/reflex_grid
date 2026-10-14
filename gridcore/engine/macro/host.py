"""The host VBA code runs against: Excel's object model over an engine session.

Macros change the workbook through the engine: values, formulas and formats are written into one
open edit (``Ctx``) that is committed — recalculating — when the code reads a formula's result,
runs a command, shows a dialog or ends, so a loop writing thousands of cells costs one commit;
everything else (sorting, inserting rows, sheets, names…) runs as the engine's own commands. The
macro thus becomes ordinary edits: everyone on the workbook sees them, and undo takes them back.

:class:`MacroHost` is for macros (and event handlers); :class:`UdfHost` serves worksheet
functions written in VBA while a formula is calculated: it can read cells but not change them,
as in Excel.
"""

from __future__ import annotations

import datetime as _dt
import math
import time
from dataclasses import replace
from typing import Any, Callable, Dict, Iterator, List, Optional, Sequence, Tuple

from ...core.addr import MAX_COLS, MAX_ROWS, RangeRef, a1, col_to_letters
from ...core.value import Array, CellError
from ...model.cell import Cell
from ...numfmt import parse as parse_fmt
from ...vba import values as V
from ...vba.errors import MacroStopped, NotSupported, VbaError
from ...vba.runtime import Execution, Host
from ...vba.values import EMPTY, MISSING, NOTHING, NULL, Currency, Date, ErrorValue, VArray, is_object
from ..session import Ctx, EngineError, Session
from .constants import CONSTANTS

# Excel's cell errors as VBA error values (CVErr) and back.
ERROR_CODES = {
    CellError.NULL: 2000, CellError.DIV0: 2007, CellError.VALUE: 2015, CellError.REF: 2023, CellError.NAME: 2029,
    CellError.NUM: 2036, CellError.NA: 2042, CellError.GETTING_DATA: 2043, CellError.SPILL: 2045, CellError.CALC: 2050,
    CellError.CIRC: 2015,
}
CELL_ERRORS = {v: k for k, v in ERROR_CODES.items() if k is not CellError.CIRC}

# Application members reachable without "Application." in front.
GLOBAL_NAMES = {
    "application", "activecell", "activechart", "activesheet", "activewindow", "activeworkbook", "cells", "columns",
    "range", "rows", "selection", "sheets", "thisworkbook", "workbooks", "worksheets", "worksheetfunction", "names",
    "calculate", "evaluate", "intersect", "union", "run", "windows", "charts", "commandbars", "addins", "sendkeys",
    "executeexcel4macro", "dialogs", "excel4macrosheets", "excel4intlmacrosheets", "activeprinter", "ddeinitiate",
}

MAX_CELLS_PER_OP = 2_000_000


def excel_to_ole(serial: float, d1904: bool) -> float:
    """An Excel date serial as a VBA Date (they differ before 3/1/1900 and in the 1904 system)."""
    if d1904:
        return serial + 1462
    if 0 < serial < 60:
        return serial + 1
    return serial


def ole_to_excel(ole: float, d1904: bool) -> float:
    if d1904:
        return ole - 1462
    if 1 < ole < 61:
        return ole - 1
    return ole


_CURRENCY_CHARS = "$€£¥₩₹"


def is_currency_format(code: str) -> bool:
    if not code or code == "General":
        return False
    bare = []
    i = 0
    while i < len(code):
        c = code[i]
        if c == '"':
            j = code.find('"', i + 1)
            seg = code[i + 1:j if j >= 0 else len(code)]
            if any(ch in _CURRENCY_CHARS for ch in seg):
                return True
            i = len(code) if j < 0 else j + 1
            continue
        bare.append(c)
        i += 1
    text = "".join(bare)
    return any(ch in text for ch in _CURRENCY_CHARS) or "[$" in text


def cell_to_vba(v: Any, fmt_code: str, d1904: bool, value2: bool = False) -> Any:
    """A cell's value as ``Range.Value`` gives it (``Value2``: no Date or Currency)."""
    if v is None:
        return EMPTY
    t = type(v)
    if t is float:
        if not value2 and fmt_code and fmt_code != "General":
            nf = parse_fmt(fmt_code)
            # A date (with a day, month or year) is a Date; a time alone stays a Double.
            if nf.is_date() and not nf.is_time_only():
                ole = excel_to_ole(v, d1904)
                if V.DATE_MIN <= ole < 2958466.0:
                    return Date(ole)
            elif is_currency_format(fmt_code):
                try:
                    return Currency.of(v)
                except VbaError:
                    return v
        return v
    if t is str or t is bool:
        return v
    if t is CellError:
        return ErrorValue(ERROR_CODES.get(v, 2015))
    if isinstance(v, Array):
        return array_to_vba(v, "General", d1904, value2)
    return v


def array_to_vba(a: Array, fmt_code: str, d1904: bool, value2: bool) -> VArray:
    rows = [[cell_to_vba(a.get(r, c), fmt_code, d1904, value2) for c in range(a.cols)] for r in range(a.rows)]
    return VArray.from_rows(rows, 1)


def vba_to_cell_value(v: Any, d1904: bool) -> Tuple[Any, Optional[str]]:
    """(value or input text, automatic number format) for a VBA value written to a cell.
    Strings are typed input (numbers, dates and formulas are recognised, as in Excel)."""
    if is_object(v) and v is not NOTHING:
        v = V.value_of(v)
    t = type(v)
    if v is EMPTY or v is NULL:
        return None, None
    if t is bool:
        return v, None
    if t is str:
        return ("text", v), None
    if t is Date:
        x = float(v)
        if x < 0:
            # Before 1900: Excel can't hold it.
            raise VbaError(1004, "Application-defined or object-defined error")
        serial = ole_to_excel(x, d1904)
        if x == int(x):
            fmt = "m/d/yyyy"
        elif abs(x) < 1:
            fmt = "h:mm:ss AM/PM"
        else:
            fmt = "m/d/yyyy h:mm"
        return serial, fmt
    if t is Currency:
        return float(v), "$#,##0.00_);($#,##0.00)"
    if t in (int, V.Integer, V.Byte, V.LongLong, float, V.Single):
        f = float(v)
        if not math.isfinite(f):
            raise VbaError(1004)
        return f, None
    if t is ErrorValue:
        e = CELL_ERRORS.get(v.code)
        if e is None:
            return CellError.VALUE, None
        return e, None
    import decimal

    if t is decimal.Decimal:
        return float(v), None
    if v is NOTHING:
        raise VbaError(91)
    if t is VArray:
        raise VbaError(13)
    if v is MISSING:
        raise VbaError(449)
    raise VbaError(13)


class MacroError(Exception):
    """An engine refusal while a macro ran, as VBA's run-time error 1004."""


def engine_error(e: EngineError, what: str = "") -> VbaError:
    msg = str(e).split("\n", 1)[0]
    return VbaError(1004, msg if not what else f"{what}: {msg}")


# ---------------------------------------------------------------- hosts


class XlHost(Host):
    """What macros and worksheet functions share: constants, the Application object and how
    sheets are found."""

    globals = frozenset(GLOBAL_NAMES)
    writable = True

    def __init__(self) -> None:
        from .excel import Application, Workbook

        self.app = Application(self)
        self.workbook = Workbook(self)
        self.sheet_objects: Dict[str, Any] = {}
        self.runtime = None

    # ------------------------------------------------------------ to override

    def wb(self):
        raise NotImplementedError

    def read_value(self, si: int, rc: Tuple[int, int]) -> Any:
        raise NotImplementedError

    def read_block(self, si: int, rng: RangeRef) -> List[List[Any]]:
        return [[self.read_value(si, (r, c)) for c in range(rng.c0, rng.c1 + 1)] for r in range(rng.r0, rng.r1 + 1)]

    def active_index(self) -> int:
        return self.wb().active_sheet

    def selection(self):
        raise NotImplementedError

    # ------------------------------------------------------------ Host

    def constant(self, name: str) -> Any:
        return CONSTANTS.get(name.lower())

    def global_get(self, ex: Execution, name: str, args: list, named: Optional[dict]) -> Any:
        return self.app.vba_get(ex, name, args, named)

    def global_let(self, ex: Execution, name: str, args: list, named: Optional[dict], value: Any) -> None:
        self.app.vba_let(ex, name, args, named, value)

    def global_set(self, ex: Execution, name: str, args: list, named: Optional[dict], value: Any) -> None:
        self.app.vba_set(ex, name, args, named, value)

    def document(self, ex: Execution, module) -> Any:
        """``Me`` of a document module: the workbook, or the sheet with that code name."""
        wb = self.wb()
        if module.name.lower() == (wb.code_name or "ThisWorkbook").lower() or self._is_workbook_module(module):
            return self.workbook
        for i, sh in enumerate(wb.sheets):
            if sh.code_name and sh.code_name.lower() == module.name.lower():
                return self.sheet_object(sh.name)
        # Code names follow the sheets' order when the file had none.
        from ...model.vba import reconcile

        _, _, names = reconcile(wb.vba, wb.code_name, [s.code_name for s in wb.sheets])
        for i, n in enumerate(names):
            if n.lower() == module.name.lower():
                return self.sheet_object(wb.sheets[i].name)
        raise VbaError(424, f"No sheet has the code name {module.name}")

    def _is_workbook_module(self, module) -> bool:
        rt = self.runtime
        return rt is not None and rt.doc_kinds.get(module.name.lower()) == "workbook"

    def evaluate(self, ex: Execution, text: str) -> Any:
        return self.app.g_evaluate(ex, text)

    # ------------------------------------------------------------ sheets

    def sheet_object(self, name: str):
        from .excel import Worksheet

        key = name.lower()
        ws = self.sheet_objects.get(key)
        if ws is None or ws.name.lower() != key:
            ws = Worksheet(self, name)
            self.sheet_objects[key] = ws
        return ws

    def sheet_renamed(self, old: str, new: str) -> None:
        ws = self.sheet_objects.pop(old.lower(), None)
        if ws is not None:
            ws.name = new
            self.sheet_objects[new.lower()] = ws

    def index_of(self, name: str) -> int:
        i = self.wb().sheet_index(name)
        if i is None:
            raise VbaError(424, f"The sheet {name} no longer exists")
        return i

    def d1904(self) -> bool:
        from ...core.date import DateSystem

        return self.wb().date_system is DateSystem.D1904

    def style_code(self, si: int, rc: Tuple[int, int]) -> str:
        wb = self.wb()
        return wb.styles.get(wb.sheets[si].style_id(rc)).num_fmt

    def cell_vba_value(self, si: int, rc: Tuple[int, int], value2: bool = False) -> Any:
        v = self.read_value(si, rc)
        if isinstance(v, Array):
            v = v.get(0, 0)
        return cell_to_vba(v, self.style_code(si, rc) if not value2 else "", self.d1904(), value2)

    # ------------------------------------------------------------ changes (macros only)

    def sort_state(self, ws):
        from .excel import SortState

        states = self.__dict__.setdefault("sort_states", {})
        key = ws.name.lower()
        st = states.get(key)
        if st is None:
            st = states[key] = SortState(ws)
        return st

    def user_name(self) -> str:
        return ""

    def flush(self) -> None:
        pass

    def require_writable(self) -> None:
        if not self.writable:
            raise VbaError(1004, "A worksheet function can't change cells or the workbook")

    def write_values(self, si: int, cells: List[Tuple[Tuple[int, int], Any]], kind: str = "value") -> None:
        raise VbaError(1004, "A worksheet function can't change cells or the workbook")

    def command(self, cid: str, params: Optional[Dict[str, Any]] = None, sheet: Optional[int] = None) -> Any:
        raise VbaError(1004, "A worksheet function can't change cells or the workbook")


class MacroHost(XlHost):
    """Excel's object model for a macro run on ``session`` (a person's session on a shared book,
    or a plain one). ``ask`` answers MsgBox and InputBox (blocking); without it every dialog
    gets its default answer."""

    def __init__(self, session: Session, ask: Optional[Callable[..., Any]] = None, user_name: str = ""):
        super().__init__()
        self.s = session
        self.pending: Optional[Ctx] = None
        self.ask = ask
        self.user = user_name or session.prefs.get("userName", "")
        # Application settings for this run.
        self.screen_updating = True
        self.enable_events = True
        self.display_alerts = True
        self.calc_manual = False
        self.status = ""
        self.cut_copy_mode = False
        self.event_depth = 0
        self.messages: List[Dict[str, Any]] = []
        self.last_find: Optional[Dict[str, Any]] = None
        self.events_allowed = True
        self.on_checkpoint: Optional[Callable[[Execution], None]] = None
        self.changed_ranges: List[Tuple[int, RangeRef]] = []
        self.sort_states: Dict[str, Any] = {}
        self.status_text = ""
        self.base_calc_mode = session.doc().wb.calc.mode if session.docs else "auto"

    # ------------------------------------------------------------ documents

    @property
    def doc(self):
        return self.s.doc()

    def wb(self):
        p = self.pending
        return p.wb if p is not None else self.s.doc().wb

    def user_name(self) -> str:  # type: ignore[override]
        return self.user

    def selection(self):
        if self.pending is not None:
            return self.pending.sel
        return self.s.doc().selection

    # ------------------------------------------------------------ reading

    def _needs_flush(self, si: int, rc: Tuple[int, int]) -> bool:
        p = self.pending
        if p is None or not p.changed:
            return False
        sh = p.wb.sheets[si]
        cell = sh.cells.get(rc)
        if cell is not None and cell.formula is not None:
            return True
        return rc in sh.spill or bool(sh.spill_ranges)

    def read_value(self, si: int, rc: Tuple[int, int]) -> Any:
        if self._needs_flush(si, rc):
            self.flush()
        return self.wb().sheets[si].value(rc)

    def read_block(self, si: int, rng: RangeRef) -> List[List[Any]]:
        p = self.pending
        if p is not None and p.changed:
            sh = p.wb.sheets[si]
            if sh.spill_ranges or any(c.formula is not None for _, c in sh.cells.iter_range(rng)):
                self.flush()
        sh = self.wb().sheets[si]
        out = [[None] * rng.width for _ in range(rng.height)]
        for (r, c), cell in sh.cells.iter_range(rng):
            out[r - rng.r0][c - rng.c0] = cell.value
        if sh.spill:
            for (r, c), v in sh.spill.items():
                if rng.r0 <= r <= rng.r1 and rng.c0 <= c <= rng.c1 and out[r - rng.r0][c - rng.c0] is None:
                    out[r - rng.r0][c - rng.c0] = v
        return out

    # ------------------------------------------------------------ writing

    def edit(self) -> Ctx:
        """The open edit macros write into (opened on first use)."""
        if self.pending is None:
            from ..session import begin_edit

            d = self.s.doc()
            self.pending = begin_edit(d)
            self.pending.unprotected = set(d.ui_only_protection)
            if self.calc_manual:
                self.pending.wb.calc = replace(self.pending.wb.calc, mode="manual")
        return self.pending

    def flush(self) -> None:  # type: ignore[override]
        """Commits the open edit (recalculating what it changed)."""
        p = self.pending
        if p is None:
            return
        self.pending = None
        from ..session import finish_edit

        try:
            finish_edit(self.s.doc(), p)
        except EngineError as e:
            raise engine_error(e)

    def write_values(self, si: int, cells: List[Tuple[Tuple[int, int], Any]], kind: str = "value") -> None:
        """Writes VBA values (``kind`` value/formula/formulaR1C1/value2) like Excel's Range.Value."""
        from ..cmd_edit import _set
        from ..entry import input_to_cell
        from ..tables import auto_expand, calculated_column

        if len(cells) > MAX_CELLS_PER_OP:
            raise VbaError(1004, "Too many cells for one operation")
        ctx = self.edit()
        wb = ctx.wb
        sheet = wb.sheets[si]
        if sheet.protection is not None and sheet.name.lower() not in ctx.unprotected:
            locked = [rc for rc, _ in cells[:100_000] if wb.styles.get(sheet.style_id(rc)).protection.locked]
            if locked:
                raise VbaError(1004, "The cell or chart you're trying to change is on a protected sheet.")
        d1904 = self.d1904()
        tables = bool(sheet.tables)
        for rc, v in cells:
            if kind == "formulaR1C1" and isinstance(v, str) and v.startswith("="):
                v = r1c1_formula_to_a1(v, rc)
            val, fmt = vba_to_cell_value(v, d1904)
            style = sheet.style_id(rc)
            if isinstance(val, tuple) and val and val[0] == "text":
                text = val[1]
                if len(text) > 32767:
                    raise VbaError(1004, "A cell can hold at most 32,767 characters")
                try:
                    cell = input_to_cell(wb, sheet, rc, text, strict=True)
                except EngineError as e:
                    raise engine_error(e)
            elif val is None:
                cell = Cell(None, None, style)
                if cell.is_blank():
                    cell = None
            else:
                if fmt is not None and wb.styles.get(style).num_fmt == "General":
                    style = wb.styles.derive(style, lambda s, fmt=fmt: replace(s, num_fmt=fmt))
                cell = Cell(val, None, style)
            old = sheet.cells.get(rc)
            if old is not None and old.formula is not None and old.formula.array is not None and not old.formula.array.is_single():
                raise VbaError(1004, "You can't change part of an array.")
            _set(ctx, si, rc, cell)
            sheet = wb.sheets[si]
            if tables and (cell is None or cell.formula is None):
                auto_expand(ctx, si, rc)
                calculated_column(ctx, si, rc)
                sheet = wb.sheets[si]
        if cells:
            rs = [rc for rc, _ in cells]
            r0 = min(r for r, _ in rs)
            r1 = max(r for r, _ in rs)
            c0 = min(c for _, c in rs)
            c1 = max(c for _, c in rs)
            self.changed_ranges.append((si, RangeRef(r0, c0, r1, c1)))
            if self.calc_manual and any(isinstance(v, str) and v.startswith("=") for _, v in cells):
                # With manual calculation a formula is calculated as it is entered, from the
                # cells as they are then; later changes don't reach it until Calculate.
                self.flush()

    def style(self, si: int, ranges: List[RangeRef], f: Callable) -> None:
        """Changes the format of ranges (batched with the values)."""
        from ..cmd_format import apply_style_ctx, check_format_allowed

        try:
            check_format_allowed(self.wb().sheets[si])
        except EngineError as e:
            raise engine_error(e)
        for r in ranges:
            if r.count() > MAX_CELLS_PER_OP and not (r.is_full_cols() or r.is_full_rows()):
                raise VbaError(1004, "Too many cells for one operation")
        ctx = self.edit()
        try:
            apply_style_ctx(ctx, si, ranges, f)
        except EngineError as e:
            raise engine_error(e)

    def command(self, cid: str, params: Optional[Dict[str, Any]] = None, sheet: Optional[int] = None) -> Any:
        """Runs an engine command (after committing what's pending), on ``sheet`` if given: the
        sheet is made active for the command and the person's own sheet and selection come back."""
        from ..session import COMMANDS

        self.flush()
        spec = COMMANDS.get(cid)
        if spec is None:
            raise VbaError(1004, f"unknown command {cid}")
        s = self.s
        d = s.doc()
        p = dict(params or {})
        restore = None
        if sheet is not None and sheet != d.wb.active_sheet:
            restore = (d.wb.active_sheet, d.selection, dict(d.sheet_selections))
            w = d.wb.clone()
            w.active_sheet = sheet
            d.wb = w
            d.selection = d.sheet_selections.get(sheet) or _home_selection(w.sheets[sheet])
        # Sheets protected with UserInterfaceOnly don't stop macros: the command runs on them
        # unprotected and their protection is put back after.
        shielded = {}
        if d.ui_only_protection and not cid.startswith(("review.protect", "review.unprotect")):
            shielded = {sh.name.lower(): sh.protection for sh in d.wb.sheets if sh.protection is not None and sh.name.lower() in d.ui_only_protection}
            if shielded:
                w = d.wb.clone()
                for i, sh in enumerate(w.sheets):
                    if sh.name.lower() in shielded:
                        w.sheet_mut(i).protection = None
                d.wb = w
        try:
            r = spec.run(s, p)
        except EngineError as e:
            raise engine_error(e)
        finally:
            if shielded:
                d = s.doc()
                w = d.wb.clone()
                for i, sh in enumerate(w.sheets):
                    if sh.name.lower() in shielded:
                        w.sheet_mut(i).protection = shielded[sh.name.lower()]
                d.wb = w
            if restore is not None:
                d = s.doc()
                active, sel, sels = restore
                if 0 <= active < len(d.wb.sheets):
                    w = d.wb.clone()
                    w.active_sheet = active
                    d.wb = w
                d.selection = sel
                d.sheet_selections = sels
        return r

    # ------------------------------------------------------------ dialogs and output

    def prompt(self, ex: Execution, kind: str, prompt: str, title: str, buttons: int, default: Any = None, **extra: Any) -> Any:
        from ...vba.stdlib import msgbox_default

        self.flush()
        if kind == "msgbox" and (buttons & 7) == 0 and self.ask is None:
            self.messages.append({"kind": "msgbox", "prompt": prompt, "title": title, "buttons": buttons})
            return 1
        if self.ask is None:
            return msgbox_default(buttons) if kind == "msgbox" else (default if default is not None else "")
        return self.ask(ex, kind, prompt, title, buttons, default, **extra)

    def output(self, ex: Execution, text: str) -> None:
        ex.output.append(text)
        if len(ex.output) > 5000:
            del ex.output[:1000]

    def checkpoint(self, ex: Execution) -> None:
        if self.on_checkpoint is not None:
            self.on_checkpoint(ex)

    def status(self, ex: Execution, text: str) -> None:
        self.status_text = text

    def finish(self) -> None:
        """The run is over: commit, and put back automatic calculation if the macro left it off."""
        self.flush()
        if self.calc_manual:
            self.calc_manual = False
            d = self.s.doc()
            if d.wb.calc.mode != self.base_calc_mode:
                w = d.wb.clone()
                w.calc = replace(w.calc, mode=self.base_calc_mode)
                d.wb = w
            if self.base_calc_mode != "manual":
                d.calc.recalc_all(d.wb)

    # ------------------------------------------------------------ events

    def fire(self, ex: Execution, sheet_index: Optional[int], event: str, args: List[Any]) -> bool:
        """Runs an event handler (``Worksheet_Change`` in the sheet's module, ``Workbook_…`` in
        ThisWorkbook) if the project has one and events are on. Returns whether one ran."""
        if not self.enable_events or not self.events_allowed or self.runtime is None:
            return False
        rt = self.runtime
        wb = self.wb()
        if event.startswith("Workbook_"):
            module = wb.code_name or _workbook_module(rt)
        else:
            if sheet_index is None or not 0 <= sheet_index < len(wb.sheets):
                return False
            module = sheet_code_name(wb, sheet_index)
        if not module:
            return False
        proc = rt.handler(module, event)
        if proc is None:
            return False
        if self.event_depth >= 40:
            raise VbaError(28, f"Out of stack space: {event} kept triggering itself")
        self.flush()
        self.event_depth += 1
        try:
            rt.call(ex, proc, args)
        finally:
            self.event_depth -= 1
        return True

    def changed(self, ex: Execution, rng) -> None:
        """A macro changed cells: Worksheet_Change and Workbook_SheetChange, as in Excel."""
        if not self.enable_events or self.runtime is None:
            return
        si = self.index_of(rng.ws.name)
        self.fire(ex, si, "Worksheet_Change", [rng])
        self.fire(ex, si, "Workbook_SheetChange", [rng.ws, rng])

    def selection_changed(self, ex: Execution, si: int, before) -> None:
        """A macro selected other cells: Worksheet_SelectionChange and
        Workbook_SheetSelectionChange."""
        if not self.enable_events or self.runtime is None:
            return
        sel = self.selection()
        if before is not None and list(before.ranges) == list(sel.ranges):
            return
        from .range import Range

        ws = self.sheet_object(self.wb().sheets[si].name)
        target = Range(self, ws, list(sel.ranges))
        self.fire(ex, si, "Worksheet_SelectionChange", [target])
        self.fire(ex, si, "Workbook_SheetSelectionChange", [ws, target])


def _workbook_module(rt) -> str:
    for name, kind in rt.doc_kinds.items():
        if kind == "workbook":
            m = rt.module(name)
            return m.name if m is not None else name
    return "ThisWorkbook"


def sheet_code_name(wb, si: int) -> str:
    sh = wb.sheets[si]
    if sh.code_name:
        return sh.code_name
    from ...model.vba import reconcile

    _, _, names = reconcile(wb.vba, wb.code_name, [s.code_name for s in wb.sheets])
    return names[si]


class UdfHost(XlHost):
    """Worksheet functions written in VBA: they read cells through the calculation in progress
    and can't change anything."""

    writable = False

    def __init__(self, ev, caller: Tuple[int, Tuple[int, int]]):
        super().__init__()
        self.ev = ev
        self.caller = caller

    def wb(self):
        return self.ev.host.wb

    def read_value(self, si: int, rc: Tuple[int, int]) -> Any:
        return self.ev.host.cell_value(si, rc)

    def read_block(self, si: int, rng: RangeRef) -> List[List[Any]]:
        vals = self.ev.host.range_values(si, rng)
        w = rng.width
        return [vals[i * w:(i + 1) * w] for i in range(rng.height)]

    def active_index(self) -> int:
        return self.ev.host.wb.active_sheet

    def selection(self):
        from ..selection import Selection

        return Selection.at(self.caller[1])

    def prompt(self, ex: Execution, kind: str, prompt: str, title: str, buttons: int, default: Any = None, **extra: Any) -> Any:
        from ...vba.stdlib import msgbox_default

        return msgbox_default(buttons) if kind == "msgbox" else (default if default is not None else "")

    def output(self, ex: Execution, text: str) -> None:
        ex.output.append(text)
        del ex.output[:-200]


def _home_selection(sheet):
    from ..selection import Selection

    return Selection.at(sheet.view_active)


# ---------------------------------------------------------------- R1C1

_R1C1 = None


def _r1c1_regex():
    global _R1C1
    if _R1C1 is None:
        import re

        _R1C1 = re.compile(r"([RC])(\[-?\d+\]|\d+)?(?:([C])(\[-?\d+\]|\d+)?)?", re.I)
    return _R1C1


def _ident_char(ch: str) -> bool:
    return ch.isalnum() or ch in "_.\\"


def r1c1_formula_to_a1(text: str, at: Tuple[int, int]) -> str:
    """``=R[1]C[-1]*2`` (``FormulaR1C1``) written to the cell ``at``, as A1 text. Strings and
    quoted sheet names are left alone, and so are names and functions (``ROUND``, ``Rate``…)."""
    rx = _r1c1_regex()
    out: List[str] = []
    i = 0
    n = len(text)

    def num(spec: Optional[str], base: int) -> Tuple[int, bool]:
        if spec is None:
            return base, False
        if spec.startswith("["):
            return base + int(spec[1:-1]), False
        return int(spec) - 1, True

    def token(j: int):
        """An R1C1 reference starting at ``j``: (end, kind, a, a_abs, b, b_abs)."""
        if j > 0 and _ident_char(text[j - 1]):
            return None
        m = rx.match(text, j)
        if m is None:
            return None
        e = m.end()
        if e < n and (_ident_char(text[e]) or text[e] in "(["):
            return None
        first = m.group(1).upper()
        if first == "R" and m.group(3) is not None:
            r, ra = num(m.group(2), at[0])
            c, ca = num(m.group(4), at[1])
            return e, "cell", r, ra, c, ca
        if m.group(3) is not None:
            return None
        if first == "R":
            r, ra = num(m.group(2), at[0])
            return e, "row", r, ra, 0, False
        c, ca = num(m.group(2), at[1])
        return e, "col", c, ca, 0, False

    while i < n:
        ch = text[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            out.append(text[i:j + 1])
            i = j + 1
            continue
        if ch == "'":
            j = text.find("'", i + 1)
            j = n - 1 if j < 0 else j
            out.append(text[i:j + 1])
            i = j + 1
            continue
        t = token(i)
        if t is None:
            out.append(ch)
            i += 1
            continue
        end, kind, a, aa, b, ba = t
        if kind == "cell":
            if not (0 <= a < MAX_ROWS and 0 <= b < MAX_COLS):
                out.append("#REF!")
            else:
                out.append(("$" if ba else "") + col_to_letters(b) + ("$" if aa else "") + str(a + 1))
            i = end
            continue
        # A whole row or column, alone (R2) or as a range (R2:R5).
        last, last_abs = a, aa
        if end < n and text[end] == ":":
            t2 = token(end + 1)
            if t2 is not None and t2[1] == kind:
                end, _, last, last_abs, _, _ = t2
        lim = MAX_ROWS if kind == "row" else MAX_COLS
        if not (0 <= a < lim and 0 <= last < lim):
            out.append("#REF!")
        elif kind == "row":
            out.append(f"{'$' if aa else ''}{a + 1}:{'$' if last_abs else ''}{last + 1}")
        else:
            out.append(f"{'$' if aa else ''}{col_to_letters(a)}:{'$' if last_abs else ''}{col_to_letters(last)}")
        i = end
    return "".join(out)


def a1_formula_to_r1c1(text: str, at: Tuple[int, int]) -> str:
    """A formula's text (no ``=``) as ``FormulaR1C1`` shows it, relative to ``at``."""
    from ...formula import ParseError, parse
    from ...formula.printer import to_text

    try:
        return to_text(parse(text), r1c1=at)
    except ParseError:
        return text
