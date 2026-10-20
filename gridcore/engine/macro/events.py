"""Workbook events and worksheet functions written in VBA.

:func:`after_command` is called by :meth:`Session.execute` after each command: when the
workbook has macros and its events are enabled, a person's edit runs ``Worksheet_Change``
(and ``Workbook_SheetChange``), a new selection ``Worksheet_SelectionChange``, switching sheets
``Worksheet_Deactivate``/``Worksheet_Activate`` (and the Workbook_ ones), as Excel does. Their
changes become part of the same edit, so one undo takes back both.

:func:`udf` is the formula evaluator's hook: a formula calling a public Function of a standard
module (``=Twice(A1)``) runs it with a read-only Excel object model.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence, Tuple

from ...core.addr import RangeRef
from ...core.value import Array, CellError
from ...vba import values as V
from ...vba.errors import MacroStopped, VbaError, VbaSyntaxError
from ...vba.runtime import Execution
from ...vba.values import EMPTY, MISSING, NOTHING, NULL, ErrorValue, VArray, is_object
from .host import CELL_ERRORS, UdfHost, cell_to_vba, ole_to_excel, sheet_code_name
from .project import runtime_for

UDF_TIME_LIMIT = 5.0
UDF_DEPTH = 200
# Commands that never fire events: undo and redo, files, macros themselves, reading.
_QUIET = ("macro.", "file.", "window.", "cell.get", "range.get", "edit.undo", "edit.redo", "formulas.evaluate", "data.validationAt", "data.filterValues", "edit.find", "formulas.nameManager", "formulas.functions", "formulas.trace", "formulas.errorChecking", "home.manageRules", "review.listComments", "edit.copy", "edit.clearClipboard", "edit.formatPainter", "edit.cancelFormatPainter", "view.zoom", "view.scroll")
_EVENT_NAMES = ("worksheet_change", "worksheet_selectionchange", "worksheet_activate", "worksheet_deactivate", "workbook_sheetchange", "workbook_sheetselectionchange", "workbook_sheetactivate", "workbook_sheetdeactivate", "workbook_newsheet")


def handlers_present(rt) -> bool:
    cached = rt.data.get("has_events")
    if cached is None:
        cached = any(m.kind == "document" and any(n in m.procs for n in _EVENT_NAMES) for m in rt.order)
        rt.data["has_events"] = cached
    return cached


def after_command(s, cid: str, before_wb, before_sel, result: Any) -> None:
    d = s.doc() if s.docs else None
    if d is None or not getattr(d, "macros_enabled", False):
        return
    p = d.wb.vba
    if p is None or p.unreadable or cid.startswith(_QUIET):
        return
    rt = runtime_for(p)
    if not handlers_present(rt):
        return
    events = _what_happened(d, before_wb, before_sel)
    if not events:
        return
    from .runner import MacroRun

    def job(ex: Execution, host) -> None:
        for sheet_index, name, args in events:
            # Objects for the arguments are made now, for this run's host.
            real = [a(host) if callable(a) else a for a in args]
            host.fire(ex, sheet_index, name, real)

    run = MacroRun(s, job, "Event", interactive=bool(getattr(s, "interactive_macros", False)), time_limit=30.0, user_name=s.prefs.get("userName", ""))
    # What the handlers do is part of the person's edit: one undo step, even across the
    # dialogs they show.
    s.undo_group = f"macro-{run.id}"
    r = run.step()
    if run.paused:
        s.macro_run = run
        return
    _report(s, r)


def _report(s, r) -> None:
    for m in r.get("messages") or []:
        s.messages.append(m.get("prompt", ""))
    err = r.get("error")
    if err:
        s.messages.append(f"Macro: {err.get('text', '')}" + (f" ({err['module']}, line {err['line']})" if err.get("module") else ""))


def _what_happened(d, before_wb, before_sel) -> List[Tuple[Optional[int], str, List[Any]]]:
    from ..session import same_content

    wb = d.wb
    out: List[Tuple[Optional[int], str, List[Any]]] = []
    if before_wb is None:
        return out
    # Cells whose content changed.
    if len(before_wb.sheets) == len(wb.sheets):
        for si, (sb, sa) in enumerate(zip(before_wb.sheets, wb.sheets)):
            if sb is sa:
                continue
            cells = []
            for rc in sb.cells.diff_keys(sa.cells):
                if not same_content(sb.cell(rc), sa.cell(rc)):
                    cells.append(rc)
                    if len(cells) > 100_000:
                        break
            if not cells and getattr(d, "cells_moved", False) and si == wb.active_sheet:
                cells = [rc for r in d.selection.ranges for rc in ([r.start] if r.count() > 1 else r.cells())]
            if cells:
                from ..ranges import cells_to_ranges

                areas = cells_to_ranges(sorted(set(cells)))[:200]
                name = sa.name
                out.append((si, "Worksheet_Change", [_range_maker(name, areas)]))
                out.append((si, "Workbook_SheetChange", [_sheet_maker(name), _range_maker(name, areas)]))
    # Another sheet.
    if before_wb.active_sheet != wb.active_sheet or (len(before_wb.sheets) != len(wb.sheets)):
        old = before_wb.sheets[before_wb.active_sheet].name if 0 <= before_wb.active_sheet < len(before_wb.sheets) else None
        new = wb.sheets[wb.active_sheet].name
        if old is not None and old != new:
            oi = wb.sheet_index(old)
            if oi is not None:
                out.append((oi, "Worksheet_Deactivate", []))
                out.append((None, "Workbook_SheetDeactivate", [_sheet_maker(old)]))
            out.append((wb.active_sheet, "Worksheet_Activate", []))
            out.append((None, "Workbook_SheetActivate", [_sheet_maker(new)]))
        return out
    # A new selection on the same sheet.
    sel = d.selection
    if before_sel is not None and (before_sel.ranges != sel.ranges):
        name = wb.sheets[wb.active_sheet].name
        areas = list(sel.ranges)
        out.append((wb.active_sheet, "Worksheet_SelectionChange", [_range_maker(name, areas)]))
        out.append((None, "Workbook_SheetSelectionChange", [_sheet_maker(name), _range_maker(name, areas)]))
    return out


def _range_maker(sheet: str, areas: List[RangeRef]):
    def make(host):
        from .range import Range

        return Range(host, host.sheet_object(sheet), list(areas))

    return make


def _sheet_maker(sheet: str):
    def make(host):
        return host.sheet_object(sheet)

    return make


# ---------------------------------------------------------------- worksheet functions


def udf(ev, name: str, args: Sequence[Any]) -> Any:
    """Evaluates ``name(args…)`` when ``name`` is a public Function of the workbook's VBA project
    (None when it isn't, so the formula gets #NAME?)."""
    from ...calc.eval import RefVal
    from ...formula import Missing

    wb = ev.host.wb
    p = getattr(wb, "vba", None)
    if p is None or p.unreadable:
        return None
    rt = runtime_for(p)
    if name.upper().startswith("_XLUDF."):
        name = name[7:]
    proc = rt.udf(name)
    if proc is None:
        return None
    if proc.module.name.lower() in rt.errors:
        return CellError.VALUE
    host = UdfHost(ev, (ev.sheet, ev.at))
    host.runtime = rt
    d1904 = host.d1904()
    vargs: List[Any] = []
    for a in args:
        if isinstance(a, Missing):
            vargs.append(MISSING)
            continue
        x = ev.eval(a)
        vargs.append(_to_vba(host, x, d1904))
    ex = Execution(rt, host, time_limit=UDF_TIME_LIMIT, udf=True, max_depth=UDF_DEPTH)
    with ex:
        try:
            r = rt.call(ex, proc, vargs)
        except (VbaError, VbaSyntaxError, MacroStopped):
            return CellError.VALUE
        except RecursionError:
            return CellError.VALUE
        try:
            return _to_cell(host, r, d1904)
        except VbaError:
            return CellError.VALUE


def _to_vba(host, x: Any, d1904: bool) -> Any:
    from ...calc.eval import Lambda, RefVal
    from .range import Range

    if isinstance(x, RefVal):
        if not x.areas:
            return V.ErrorValue(2023)
        s0 = x.areas[0].sheet
        wb = host.wb()
        ws = host.sheet_object(wb.sheets[s0].name)
        return Range(host, ws, [a.rng for a in x.areas if a.sheet == s0])
    if isinstance(x, Array):
        return VArray.from_rows([[cell_to_vba(x.get(r, c), "", d1904, True) for c in range(x.cols)] for r in range(x.rows)], 1)
    if isinstance(x, Lambda):
        return V.ErrorValue(2015)
    return cell_to_vba(x, "", d1904, True)


def _to_cell(host, v: Any, d1904: bool) -> Any:
    from .range import Range

    if isinstance(v, Range):
        vals = v._values(None, "value2")
        v = vals
    t = type(v)
    if v is EMPTY:
        return 0.0
    if v is NULL:
        return CellError.NA
    if v is NOTHING or v is MISSING:
        return CellError.VALUE
    if t is bool or t is str:
        if t is str and len(v) > 32767:
            return CellError.VALUE
        return v
    if t is V.Date:
        return ole_to_excel(float(v), d1904)
    if t is ErrorValue:
        return CELL_ERRORS.get(v.code, CellError.VALUE)
    if t in (int, V.Integer, V.Byte, V.LongLong, float, V.Single, V.Currency):
        return float(v)
    import decimal

    if t is decimal.Decimal:
        return float(v)
    if t is VArray:
        if v.dims == 1:
            if not v.data:
                return CellError.CALC
            return Array(1, len(v.data), [_to_cell(host, x, d1904) for x in v.data])
        if v.dims == 2:
            rows = v.rows()
            if not rows or not rows[0]:
                return CellError.CALC
            return Array(len(rows), len(rows[0]), [_to_cell(host, x, d1904) for row in rows for x in row])
        return CellError.VALUE
    if is_object(v):
        return _to_cell(host, V.value_of(v), d1904)
    return CellError.VALUE
