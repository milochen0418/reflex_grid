"""Excel's object model above Range: Application, Workbook(s), Sheets, Worksheet, Window,
Names, ListObjects, Hyperlinks, Validation, Sort, AutoFilter, PageSetup…

The workbook a macro sees is the one it belongs to (``ThisWorkbook`` = ``ActiveWorkbook``);
ReflexGrid macros can't open, create, save as or close other files, nor reach the VBA project
itself. Members missing here raise error 438 like an unknown member in Excel.
"""

from __future__ import annotations

import time
from dataclasses import replace
from typing import Any, Callable, Dict, Iterator, List, Optional, Tuple

from ...core.addr import MAX_COLS, MAX_ROWS, RangeRef, a1, col_to_letters
from ...vba import values as V
from ...vba.errors import MacroStopped, NotSupported, VbaError
from ...vba.runtime import Execution, HostObject
from ...vba.values import EMPTY, MISSING, NOTHING, NULL, VArray, is_object
from ..session import EngineError
from .host import engine_error
from .range import Range, _as_range, _bool, _long, _m, _plain, _str, bgr_to_hex, color_value, make_range, palette_hex, parse_reference


def _all_cells(h, ws) -> Range:
    return Range(h, ws, [RangeRef(0, 0, MAX_ROWS - 1, MAX_COLS - 1)])


# ---------------------------------------------------------------- Application


class Application(HostObject):
    TYPE_NAME = "Application"
    PARAMS = {
        "range": ("Cell1", "Cell2"),
        "cells": ("RowIndex", "ColumnIndex"),
        "inputbox": ("Prompt", "Title", "Default", "Left", "Top", "HelpFile", "HelpContextID", "Type"),
        "run": ("Macro",) + tuple(f"Arg{i}" for i in range(1, 31)),
        "goto": ("Reference", "Scroll"),
        "intersect": tuple(f"Arg{i}" for i in range(1, 31)),
        "union": tuple(f"Arg{i}" for i in range(1, 31)),
        "wait": ("Time",),
        "ontime": ("EarliestTime", "Procedure", "LatestTime", "Schedule"),
        "convertformula": ("Formula", "FromReferenceStyle", "ToReferenceStyle", "ToAbsolute", "RelativeTo"),
    }

    def __init__(self, h):
        self.h = h

    def vba_same(self, other) -> bool:
        return other is self

    # -- objects
    def g_application(self, ex):
        return self

    g_parent = g_application

    def g_activeworkbook(self, ex):
        return self.h.workbook

    g_thisworkbook = g_activeworkbook

    def g_workbooks(self, ex, index=MISSING):
        w = Workbooks(self.h)
        return w if _m(index) else w.g_item(ex, index)

    def g_activesheet(self, ex):
        wb = self.h.wb()
        return self.h.sheet_object(wb.sheets[self.h.active_index()].name)

    def g_worksheets(self, ex, index=MISSING):
        return self.h.workbook.g_worksheets(ex, index)

    def g_sheets(self, ex, index=MISSING):
        return self.h.workbook.g_sheets(ex, index)

    def g_charts(self, ex, index=MISSING):
        return self.h.workbook.g_charts(ex, index)

    def g_names(self, ex, index=MISSING):
        return self.h.workbook.g_names(ex, index)

    def g_range(self, ex, cell1=MISSING, cell2=MISSING):
        return make_range(self.h, self.g_activesheet(ex), cell1, cell2)

    def g_cells(self, ex, row=MISSING, col=MISSING):
        return self.g_activesheet(ex).g_cells(ex, row, col)

    def g_rows(self, ex, index=MISSING):
        return self.g_activesheet(ex).g_rows(ex, index)

    def g_columns(self, ex, index=MISSING):
        return self.g_activesheet(ex).g_columns(ex, index)

    def g_activecell(self, ex):
        sel = self.h.selection()
        return Range(self.h, self.g_activesheet(ex), [RangeRef.of(sel.active)])

    def g_selection(self, ex):
        sel = self.h.selection()
        return Range(self.h, self.g_activesheet(ex), list(sel.ranges))

    def g_activewindow(self, ex):
        return Window(self.h)

    def g_windows(self, ex, index=MISSING):
        return Window(self.h) if not _m(index) else Windows(self.h)

    def g_activechart(self, ex):
        return NOTHING

    def g_worksheetfunction(self, ex):
        from .wsfunction import WorksheetFunction

        return WorksheetFunction(self.h, raise_errors=True)

    # -- settings
    def g_screenupdating(self, ex):
        return getattr(self.h, "screen_updating", True)

    def l_screenupdating(self, ex, v):
        if hasattr(self.h, "screen_updating"):
            self.h.screen_updating = _bool(v)

    def g_enableevents(self, ex):
        return getattr(self.h, "enable_events", True)

    def l_enableevents(self, ex, v):
        if hasattr(self.h, "enable_events"):
            self.h.enable_events = _bool(v)

    def g_displayalerts(self, ex):
        return getattr(self.h, "display_alerts", True)

    def l_displayalerts(self, ex, v):
        if hasattr(self.h, "display_alerts"):
            self.h.display_alerts = _bool(v)

    def g_calculation(self, ex):
        if getattr(self.h, "calc_manual", False):
            return -4135
        return -4135 if self.h.wb().calc.mode == "manual" else -4105

    def l_calculation(self, ex, v):
        n = _long(v)
        if n not in (-4105, -4135, 2):
            raise VbaError(1004, "Unable to set the Calculation property of the Application class")
        self.h.require_writable()
        manual = n == -4135
        if manual == getattr(self.h, "calc_manual", False):
            return
        self.h.flush()
        self.h.calc_manual = manual
        if not manual:
            # Back to automatic: what changed meanwhile is calculated now.
            from dataclasses import replace as _r

            d = self.h.s.doc()
            mode = self.h.base_calc_mode
            if d.wb.calc.mode != mode:
                w = d.wb.clone()
                w.calc = _r(w.calc, mode=mode)
                d.wb = w
            if mode != "manual":
                d.calc.recalc_all(d.wb)

    def g_calculationstate(self, ex):
        return 0

    def g_calculate(self, ex):
        self.h.require_writable()
        self.h.command("formulas.calculateNow", {})
        return EMPTY

    g_calculatefull = g_calculate
    g_calculatefullrebuild = g_calculate

    def g_statusbar(self, ex):
        return getattr(self.h, "status", "") or False

    def l_statusbar(self, ex, v):
        pv = _plain(v)
        self.h.status = "" if (type(pv) is bool and not pv) else V.text(pv)

    def g_cutcopymode(self, ex):
        return bool(getattr(self.h, "cut_copy_mode", False))

    def l_cutcopymode(self, ex, v):
        if not _bool(v) and getattr(self.h, "s", None) is not None:
            self.h.s.clipboard = None
            self.h.cut_copy_mode = False

    def g_interactive(self, ex):
        return True

    def l_interactive(self, ex, v):
        pass

    def g_cursor(self, ex):
        return -4143

    def l_cursor(self, ex, v):
        pass

    def g_enablecancelkey(self, ex):
        return 1

    def l_enablecancelkey(self, ex, v):
        pass

    def g_displaystatusbar(self, ex):
        return True

    def l_displaystatusbar(self, ex, v):
        pass

    def g_displayformulabar(self, ex):
        return True

    def l_displayformulabar(self, ex, v):
        pass

    def g_displayfullscreen(self, ex):
        return False

    def l_displayfullscreen(self, ex, v):
        pass

    def g_windowstate(self, ex):
        return -4137

    def l_windowstate(self, ex, v):
        pass

    def g_visible(self, ex):
        return True

    def l_visible(self, ex, v):
        pass

    def g_asktoupdatelinks(self, ex):
        return False

    def g_iteration(self, ex):
        return self.h.wb().calc.iterative

    def l_iteration(self, ex, v):
        self.h.require_writable()
        self.h.command("formulas.calculationOptions", {"mode": self.h.wb().calc.mode, "iterative": _bool(v)})

    def g_maxiterations(self, ex):
        return self.h.wb().calc.max_iterations

    def l_maxiterations(self, ex, v):
        self.h.require_writable()
        self.h.command("formulas.calculationOptions", {"mode": self.h.wb().calc.mode, "maxIterations": _long(v)})

    def g_maxchange(self, ex):
        return self.h.wb().calc.max_change

    def l_maxchange(self, ex, v):
        self.h.require_writable()
        self.h.command("formulas.calculationOptions", {"mode": self.h.wb().calc.mode, "maxChange": V.to_double(_plain(v))})

    def g_referencestyle(self, ex):
        return 1

    def l_referencestyle(self, ex, v):
        pass

    def g_automationsecurity(self, ex):
        return 1

    # -- information
    def g_name(self, ex):
        return "Microsoft Excel"

    def g_version(self, ex):
        return "16.0"

    def g_build(self, ex):
        return 0

    def g_operatingsystem(self, ex):
        return "ReflexGrid"

    def g_username(self, ex):
        n = self.h.user_name() if hasattr(self.h, "user_name") else ""
        return n or "ReflexGrid User"

    def l_username(self, ex, v):
        pass

    def g_organizationname(self, ex):
        return ""

    def g_pathseparator(self, ex):
        return "/"

    def g_path(self, ex):
        return ""

    g_defaultfilepath = g_path
    g_startuppath = g_path
    g_templatespath = g_path
    g_libraryPath = g_path

    def g_decimalseparator(self, ex):
        return "."

    def g_thousandsseparator(self, ex):
        return ","

    def g_usesystemseparators(self, ex):
        return True

    def g_international(self, ex, index=MISSING):
        n = _long(index)
        return {1: 1, 2: 1, 3: ".", 4: ",", 5: ",", 32: 0, 44: 0}.get(n, 0)

    def g_hwnd(self, ex):
        return 0

    def g_caller(self, ex, index=MISSING):
        caller = getattr(self.h, "caller", None)
        if caller is None:
            raise VbaError(2023, "Caller is only known in a worksheet function")
        si, rc = caller
        wb = self.h.wb()
        return Range(self.h, self.h.sheet_object(wb.sheets[si].name), [RangeRef.of(rc)])

    def g_thiscell(self, ex):
        return self.g_caller(ex)

    def g_volatile(self, ex, on=MISSING):
        return EMPTY

    def g_mouseavailable(self, ex):
        return True

    # -- methods
    def g_evaluate(self, ex, text=MISSING):
        return evaluate(self.h, self.g_activesheet(ex), _str(text))

    g__evaluate = g_evaluate

    def g_intersect(self, ex, *args):
        rngs = [a for a in args if a is not MISSING]
        if len(rngs) < 2:
            raise VbaError(449)
        for r in rngs:
            if not isinstance(r, Range):
                raise VbaError(13 if r is not NOTHING else 91)
        ws = rngs[0].ws
        if any(r.ws.name.lower() != ws.name.lower() for r in rngs):
            raise VbaError(1004, "Method 'Intersect' of object '_Application' failed")
        cur = list(rngs[0].areas)
        for r in rngs[1:]:
            nxt = []
            for a in cur:
                for b in r.areas:
                    i = a.intersection(b)
                    if i is not None:
                        nxt.append(i)
            cur = nxt
            if not cur:
                return NOTHING
        return Range(self.h, ws, cur)

    def g_union(self, ex, *args):
        rngs = [a for a in args if a is not MISSING]
        if len(rngs) < 2:
            raise VbaError(449)
        for r in rngs:
            if not isinstance(r, Range):
                raise VbaError(13 if r is not NOTHING else 91)
        ws = rngs[0].ws
        if any(r.ws.name.lower() != ws.name.lower() for r in rngs):
            raise VbaError(1004, "Method 'Union' of object '_Application' failed")
        areas = []
        for r in rngs:
            areas.extend(r.areas)
        return Range(self.h, ws, _merge_areas(areas))

    def g_run(self, ex, macro=MISSING, *args):
        if _m(macro):
            raise VbaError(449)
        name = _str(macro)
        rt = ex.rt
        p = rt.find(name)
        if p is None:
            raise VbaError(1004, f"Cannot run the macro '{name}'. The macro may not be available in this workbook or all macros may be disabled.")
        return rt.call(ex, p, [a for a in args if a is not MISSING] if args else [])

    def g_inputbox(self, ex, prompt=MISSING, title=MISSING, default=MISSING, left=MISSING, top=MISSING, helpfile=MISSING, context=MISSING, type_=MISSING):
        t = _long(type_, 2)
        default_text = ""
        if not _m(default):
            d = default if not isinstance(default, Range) else default
            default_text = d.address() if isinstance(d, Range) else V.text(_plain(d))
        r = self.h.prompt(ex, "inputbox", _str(prompt), _str(title, "Input"), 0, default_text, input_type=t, cancel_false=True)
        if r is None or r is False:
            return False
        text = V.text(r)
        if t & 8:
            try:
                return parse_reference(self.h, self.g_activesheet(ex), text.lstrip("="))
            except VbaError:
                raise VbaError(424, "The reference isn't valid")
        if t & 1:
            n = V.parse_number(text)
            if n is None:
                raise VbaError(13, "Number is not valid")
            return float(n)
        if t & 4:
            low = text.strip().lower()
            if low in ("true", "false"):
                return low == "true"
            raise VbaError(13)
        if t == 0:
            return text
        return text

    def g_wait(self, ex, when=MISSING):
        target = float(V.to_date(_plain(when)))
        from ...vba.stdlib import now_serial

        delay = (target - now_serial()) * 86400.0
        end = time.monotonic() + max(0.0, min(delay, 10.0))
        while time.monotonic() < end:
            ex.check()
            time.sleep(min(0.1, max(end - time.monotonic(), 0)))
        return True

    def g_ontime(self, ex, *args):
        raise NotSupported("Application.OnTime (scheduling macros)")

    def g_onkey(self, ex, key=MISSING, procedure=MISSING):
        raise NotSupported("Application.OnKey (keyboard shortcuts for macros)")

    def g_goto(self, ex, ref=MISSING, scroll=MISSING):
        if _m(ref):
            return EMPTY
        rng = ref if isinstance(ref, Range) else parse_reference(self.h, self.g_activesheet(ex), _str(ref))
        rng.ws.g_activate(ex)
        rng.g_select(ex)
        return EMPTY

    def g_quit(self, ex):
        raise MacroStopped("end")

    def g_convertformula(self, ex, formula=MISSING, fromstyle=MISSING, tostyle=MISSING, toabs=MISSING, relto=MISSING):
        from .host import a1_formula_to_r1c1, r1c1_formula_to_a1

        text = _str(formula)
        f, t = _long(fromstyle), _long(tostyle, _long(fromstyle))
        base = (0, 0)
        if isinstance(relto, Range):
            base = (relto.first.r0, relto.first.c0)
        if f == -4150 and t == 1:
            return r1c1_formula_to_a1(text, base)
        if f == 1 and t == -4150:
            body = text[1:] if text.startswith("=") else text
            return ("=" if text.startswith("=") else "") + a1_formula_to_r1c1(body, base)
        return text

    def g_undo(self, ex):
        raise NotSupported("Application.Undo")

    def g_sendkeys(self, ex, *args):
        raise NotSupported("SendKeys")

    def g_executeexcel4macro(self, ex, *args):
        raise NotSupported("Excel 4 macros")

    def g_getopenfilename(self, ex, *args):
        raise NotSupported("File dialogs (GetOpenFilename)")

    def g_getsaveasfilename(self, ex, *args):
        raise NotSupported("File dialogs (GetSaveAsFilename)")

    def g_filedialog(self, ex, *args):
        raise NotSupported("File dialogs (FileDialog)")

    def g_dialogs(self, ex, *args):
        raise NotSupported("Built-in dialogs")

    def g_commandbars(self, ex, *args):
        raise NotSupported("CommandBars (menus and toolbars)")

    def g_speech(self, ex):
        raise NotSupported("Speech")

    def g_vbe(self, ex):
        raise NotSupported("The VBA editor's object model (VBE)")

    def g_addins(self, ex, *args):
        raise NotSupported("Add-ins")

    def vba_get_missing(self, ex, name, args, named):
        # Application.Sum(…), Application.VLookup(…): worksheet functions that return errors
        # as values instead of raising them.
        from .wsfunction import WorksheetFunction, function_spec

        if function_spec(name) is not None:
            return WorksheetFunction(self.h, raise_errors=False).vba_get(ex, name, args, named)
        raise VbaError(438, f"Object doesn't support this property or method: Application.{name}")


def evaluate(h, ws, text: str) -> Any:
    """``Evaluate("…")`` / ``[…]``: a formula, a reference (as a Range) or a name."""
    from ...calc.eval import Evaluator, RefVal
    from ...calc.recalc import Calc, _PassHost
    from ...core.date import now_serial
    from ...core.value import Array
    from ...formula import ParseError, parse
    from .host import cell_to_vba

    t = text.strip()
    if t.startswith("="):
        t = t[1:]
    if not t:
        raise VbaError(13)
    flush = getattr(h, "flush", None)
    if flush is not None and getattr(h, "pending", None) is not None:
        flush()
    try:
        e = parse(t)
    except ParseError:
        return V.ErrorValue(2029)
    wb = h.wb()
    si = h.index_of(ws.name)
    host = _PassHost(wb, {}, Calc(), now_serial())
    r = Evaluator(host, si, (0, 0)).eval(e)
    if isinstance(r, RefVal):
        if len(r.areas) >= 1 and all(a.sheet == r.areas[0].sheet for a in r.areas):
            return Range(h, h.sheet_object(wb.sheets[r.areas[0].sheet].name), [a.rng for a in r.areas])
        return V.ErrorValue(2015)
    if isinstance(r, Array):
        if r.rows == 1 and r.cols == 1:
            return cell_to_vba(r.get(0, 0), "", h.d1904(), True)
        return VArray.from_rows([[cell_to_vba(r.get(i, j), "", h.d1904(), True) for j in range(r.cols)] for i in range(r.rows)], 1)
    return cell_to_vba(r, "", h.d1904(), True)


# ---------------------------------------------------------------- workbooks


class Workbooks(HostObject):
    TYPE_NAME = "Workbooks"
    DEFAULT = "item"

    def __init__(self, h):
        self.h = h

    def g_count(self, ex):
        return 1

    def g_item(self, ex, index=MISSING):
        v = _plain(index)
        if isinstance(v, str):
            names = {self.h.workbook.g_name(ex).lower(), self.h.workbook.name().lower()}
            stem = v.lower().rsplit(".", 1)[0]
            if v.lower() in names or stem in {n.rsplit(".", 1)[0] for n in names}:
                return self.h.workbook
            raise VbaError(9)
        if V.to_long(v) != 1:
            raise VbaError(9)
        return self.h.workbook

    def vba_enum(self, ex):
        return iter([self.h.workbook])

    def g_open(self, ex, *args):
        raise NotSupported("Opening other workbooks (Workbooks.Open)")

    def g_add(self, ex, *args):
        raise NotSupported("Creating workbooks (Workbooks.Add)")

    def g_close(self, ex, *args):
        raise NotSupported("Closing workbooks")


class Workbook(HostObject):
    TYPE_NAME = "Workbook"
    TYPE_NAMES = {"workbook", "_workbook"}
    PARAMS = {"protect": ("Password", "Structure", "Windows"), "unprotect": ("Password",), "saveas": ("Filename", "FileFormat")}

    def __init__(self, h):
        self.h = h

    def vba_same(self, other) -> bool:
        return other is self

    def name(self) -> str:
        d = getattr(self.h, "s", None)
        name = d.doc().name if d is not None else "Book1"
        if not name.lower().endswith((".xlsm", ".xlsx", ".xlsb", ".xls")):
            name += ".xlsm"
        return name

    def vba_get_missing(self, ex, name, args, named):
        found, v = ex.rt.document_member(ex, self.h.wb().code_name or "ThisWorkbook", name, args, named, self)
        if found:
            return v
        raise VbaError(438, f"Object doesn't support this property or method: Workbook.{name}")

    def g_name(self, ex):
        return self.name()

    def g_fullname(self, ex):
        return self.name()

    def g_path(self, ex):
        return ""

    def g_codename(self, ex):
        return self.h.wb().code_name or "ThisWorkbook"

    def g_application(self, ex):
        return self.h.app

    g_parent = g_application

    def g_worksheets(self, ex, index=MISSING):
        c = Sheets(self.h, worksheets_only=True)
        return c if _m(index) else c.g_item(ex, index)

    def g_sheets(self, ex, index=MISSING):
        c = Sheets(self.h, worksheets_only=False)
        return c if _m(index) else c.g_item(ex, index)

    def g_charts(self, ex, index=MISSING):
        if not _m(index):
            raise VbaError(9)
        return EmptyCollection("Sheets")

    def g_activesheet(self, ex):
        return self.h.app.g_activesheet(ex)

    def g_names(self, ex, index=MISSING):
        n = Names(self.h)
        return n if _m(index) else n.g_item(ex, index)

    def g_windows(self, ex, index=MISSING):
        return Window(self.h) if not _m(index) else Windows(self.h)

    def g_activate(self, ex):
        return EMPTY

    def g_save(self, ex):
        flush = getattr(self.h, "flush", None)
        if flush is not None:
            flush()
        return EMPTY

    def g_saveas(self, ex, *args):
        raise NotSupported("Saving to other files (SaveAs)")

    def g_savecopyas(self, ex, *args):
        raise NotSupported("Saving copies to files (SaveCopyAs)")

    def g_close(self, ex, *args):
        raise NotSupported("Closing the workbook from a macro")

    def g_saved(self, ex):
        return True

    def l_saved(self, ex, v):
        pass

    def g_readonly(self, ex):
        return False

    def g_fileformat(self, ex):
        return 52

    def g_protectstructure(self, ex):
        return bool(self.h.wb().protected_structure)

    def g_protectwindows(self, ex):
        return False

    def g_protect(self, ex, password=MISSING, structure=MISSING, windows=MISSING):
        self.h.require_writable()
        if _m(structure) or _bool(structure):
            self.h.command("review.protectWorkbook", {"password": _str(password)})
        return EMPTY

    def g_unprotect(self, ex, password=MISSING):
        self.h.require_writable()
        self.h.command("review.unprotectWorkbook", {"password": _str(password)})
        return EMPTY

    def g_builtindocumentproperties(self, ex, index=MISSING):
        p = DocumentProperties(self.h)
        return p if _m(index) else p.g_item(ex, index)

    def g_customdocumentproperties(self, ex, index=MISSING):
        if not _m(index):
            raise VbaError(5)
        return EmptyCollection("DocumentProperties")

    def g_refreshall(self, ex):
        return EMPTY

    def g_vbproject(self, ex):
        raise NotSupported("Changing VBA code from a macro (VBProject)")

    def g_styles(self, ex, index=MISSING):
        raise NotSupported("Workbook.Styles")

    def g_colors(self, ex, index=MISSING):
        from ...numfmt import PALETTE56

        n = _long(index)
        if not 1 <= n <= 56:
            raise VbaError(9)
        c = PALETTE56[n - 1]
        return float(((c & 0xFF) << 16) | (c & 0xFF00) | ((c >> 16) & 0xFF))

    def g_date1904(self, ex):
        return self.h.d1904()

    def g_followhyperlink(self, ex, *args):
        raise NotSupported("Opening links from a macro (FollowHyperlink)")


class DocumentProperties(HostObject):
    TYPE_NAME = "DocumentProperties"
    DEFAULT = "item"
    _KEYS = {"title": "title", "subject": "subject", "author": "author", "keywords": "keywords", "comments": "description", "company": "company"}

    def __init__(self, h):
        self.h = h

    def g_item(self, ex, index=MISSING):
        v = _plain(index)
        name = V.text(v).lower() if isinstance(v, str) else ["title", "subject", "author", "keywords", "comments"][max(0, min(V.to_long(v) - 1, 4))]
        if name not in self._KEYS:
            raise VbaError(5, f"Unknown document property {V.text(v)}")
        return DocumentProperty(self.h, name, self._KEYS[name])

    def g_count(self, ex):
        return len(self._KEYS)


class DocumentProperty(HostObject):
    TYPE_NAME = "DocumentProperty"
    DEFAULT = "value"

    def __init__(self, h, name: str, key: str):
        self.h = h
        self.name = name
        self.key = key

    def g_name(self, ex):
        return self.name.title()

    def g_value(self, ex):
        return getattr(self.h.wb().props, self.key)

    def l_value(self, ex, v):
        self.h.require_writable()
        self.h.command("file.properties", {self.key: V.text(_plain(v))})


class EmptyCollection(HostObject):
    DEFAULT = "item"

    def __init__(self, name: str):
        self.TYPE_NAME = name  # type: ignore[misc]

    def vba_type_name(self) -> str:
        return self.TYPE_NAME

    def g_count(self, ex):
        return 0

    def g_item(self, ex, index=MISSING):
        raise VbaError(9)

    def vba_enum(self, ex):
        return iter([])


# ---------------------------------------------------------------- sheets


class Sheets(HostObject):
    TYPE_NAME = "Sheets"
    DEFAULT = "item"
    PARAMS = {"add": ("Before", "After", "Count", "Type"), "item": ("Index",)}

    def __init__(self, h, worksheets_only: bool):
        self.h = h
        if worksheets_only:
            self.TYPE_NAME = "Sheets"  # type: ignore[misc]

    def g_count(self, ex):
        return len(self.h.wb().sheets)

    def g_item(self, ex, index=MISSING):
        wb = self.h.wb()
        v = _plain(index)
        if type(v) is VArray:
            raise NotSupported("Several sheets at once (Sheets(Array(…)))")
        if isinstance(v, str):
            i = wb.sheet_index(v)
            if i is None:
                raise VbaError(9)
        else:
            i = V.to_long(v) - 1
            if not 0 <= i < len(wb.sheets):
                raise VbaError(9)
        return self.h.sheet_object(wb.sheets[i].name)

    def vba_enum(self, ex):
        wb = self.h.wb()
        return iter([self.h.sheet_object(s.name) for s in wb.sheets])

    def g_add(self, ex, before=MISSING, after=MISSING, count=MISSING, type_=MISSING):
        self.h.require_writable()
        n = _long(count, 1)
        wb = self.h.wb()
        if not _m(before):
            index = self.h.index_of(before.name)
        elif not _m(after):
            index = self.h.index_of(after.name) + 1
        else:
            index = wb.active_sheet
        last = None
        for _ in range(max(1, n)):
            r = self.h.command("home.insertSheet", {"index": index})
            last = r.get("sheet") if isinstance(r, dict) else None
        # A new sheet gets a code name in a workbook with macros.
        _assign_code_names(self.h)
        return self.h.sheet_object(last) if last else self.h.app.g_activesheet(ex)

    def g_visible(self, ex):
        return -1

    def g_select(self, ex, *args):
        return EMPTY

    def g_application(self, ex):
        return self.h.app

    def g_parent(self, ex):
        return self.h.workbook


def _assign_code_names(h) -> None:
    from dataclasses import replace

    from ...model.vba import reconcile

    wb = h.wb()
    if wb.vba is None or all(s.code_name for s in wb.sheets):
        return
    _, _, names = reconcile(wb.vba, wb.code_name, [s.code_name for s in wb.sheets])
    d = h.s.doc()
    w = d.wb.clone()
    for i, n in enumerate(names):
        if not w.sheets[i].code_name:
            sm = w.sheet_mut(i)
            sm.code_name = n
    d.wb = w


class Worksheet(HostObject):
    TYPE_NAME = "Worksheet"
    TYPE_NAMES = {"worksheet", "_worksheet", "object"}
    PARAMS = {
        "range": ("Cell1", "Cell2"),
        "cells": ("RowIndex", "ColumnIndex"),
        "protect": ("Password", "DrawingObjects", "Contents", "Scenarios", "UserInterfaceOnly", "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows", "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks", "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting", "AllowFiltering", "AllowUsingPivotTables"),
        "unprotect": ("Password",),
        "copy": ("Before", "After"),
        "move": ("Before", "After"),
        "paste": ("Destination", "Link"),
        "select": ("Replace",),
        "delete": (),
    }

    def __init__(self, h, name: str):
        self.h = h
        self.name = name

    def vba_same(self, other) -> bool:
        return isinstance(other, Worksheet) and other.name.lower() == self.name.lower()

    @property
    def si(self) -> int:
        return self.h.index_of(self.name)

    def _sheet(self):
        return self.h.wb().sheets[self.si]

    def vba_get_missing(self, ex, name, args, named):
        sh = self._sheet()
        code = sh.code_name
        if not code:
            from ...model.vba import reconcile

            wb = self.h.wb()
            _, _, names = reconcile(wb.vba, wb.code_name, [s.code_name for s in wb.sheets])
            code = names[self.si]
        found, v = ex.rt.document_member(ex, code, name, args, named, self)
        if found:
            return v
        raise VbaError(438, f"Object doesn't support this property or method: Worksheet.{name}")

    def has_dynamic_member(self, name: str) -> bool:
        return False

    # -- identity
    def g_name(self, ex):
        return self.name

    def l_name(self, ex, v):
        self.h.require_writable()
        new = _str(v)
        if new == self.name:
            return
        old = self.name
        self.h.command("sheet.rename", {"sheet": self.si, "name": new})
        self.h.sheet_renamed(old, new)

    def g_codename(self, ex):
        sh = self._sheet()
        if sh.code_name:
            return sh.code_name
        from ...model.vba import reconcile

        wb = self.h.wb()
        _, _, names = reconcile(wb.vba, wb.code_name, [s.code_name for s in wb.sheets])
        return names[self.si]

    g__codename = g_codename

    def g_index(self, ex):
        return self.si + 1

    def g_type(self, ex):
        return -4167

    def g_application(self, ex):
        return self.h.app

    def g_parent(self, ex):
        return self.h.workbook

    def g_next(self, ex):
        wb = self.h.wb()
        i = self.si + 1
        return self.h.sheet_object(wb.sheets[i].name) if i < len(wb.sheets) else NOTHING

    def g_previous(self, ex):
        wb = self.h.wb()
        i = self.si - 1
        return self.h.sheet_object(wb.sheets[i].name) if i >= 0 else NOTHING

    # -- ranges
    def g_range(self, ex, cell1=MISSING, cell2=MISSING):
        r = make_range(self.h, self, cell1, cell2)
        if r.ws.name.lower() != self.name.lower() and isinstance(_plain(cell1) if not isinstance(cell1, Range) else None, str):
            raise VbaError(1004, "Method 'Range' of object '_Worksheet' failed")
        return r

    def g_cells(self, ex, row=MISSING, col=MISSING):
        allr = _all_cells(self.h, self)
        if _m(row) and _m(col):
            return allr
        return allr.g_item(ex, row, col)

    def g_rows(self, ex, index=MISSING):
        r = Range(self.h, self, [RangeRef(0, 0, MAX_ROWS - 1, MAX_COLS - 1)], "rows")
        if _m(index):
            return r
        v = _plain(index)
        if isinstance(v, str):
            rr = RangeRef.parse(v.replace("$", ""))
            if rr is None:
                raise VbaError(13)
            return Range(self.h, self, [rr], "rows")
        return r.g_item(ex, v)

    def g_columns(self, ex, index=MISSING):
        r = Range(self.h, self, [RangeRef(0, 0, MAX_ROWS - 1, MAX_COLS - 1)], "cols")
        if _m(index):
            return r
        v = _plain(index)
        if isinstance(v, str):
            t = v.replace("$", "").strip()
            rr = RangeRef.parse(t if ":" in t else f"{t}:{t}")
            if rr is None:
                raise VbaError(13)
            return Range(self.h, self, [rr], "cols")
        return r.g_item(ex, v)

    def g_usedrange(self, ex):
        flush = getattr(self.h, "flush", None)
        if flush is not None and getattr(self.h, "pending", None) is not None:
            flush()
        sh = self._sheet()
        u = sh.used_range()
        if u is None:
            return Range(self.h, self, [RangeRef(0, 0, 0, 0)])
        return Range(self.h, self, [u])

    def g_evaluate(self, ex, text=MISSING):
        return evaluate(self.h, self, _str(text))

    # -- activation and structure
    def g_activate(self, ex):
        self.h.require_writable()
        if self._sheet().visibility != "visible":
            raise VbaError(1004, "Activate method of Worksheet class failed (the sheet is hidden)")
        self.h.command("sheet.activate", {"sheet": self.si})
        return EMPTY

    def g_select(self, ex, replace_=MISSING):
        return self.g_activate(ex)

    def g_delete(self, ex):
        self.h.require_writable()
        wb = self.h.wb()
        if len(wb.sheets) == 1:
            raise VbaError(1004, "Delete method of Worksheet class failed (a workbook needs at least one sheet)")
        if getattr(self.h, "display_alerts", True) and getattr(self.h, "ask", None) is not None:
            r = self.h.prompt(ex, "msgbox", f"Microsoft Excel will permanently delete this sheet. Do you want to continue?", "Microsoft Excel", 1)
            if r != 1:
                return False
        self.h.command("home.deleteSheet", {"sheet": self.si})
        return True

    def g_copy(self, ex, before=MISSING, after=MISSING):
        self.h.require_writable()
        if _m(before) and _m(after):
            raise NotSupported("Copying a sheet to a new workbook")
        to = self.h.index_of(before.name) if not _m(before) else self.h.index_of(after.name) + 1
        self.h.command("sheet.move", {"sheet": self.si, "to": to, "copy": True})
        _assign_code_names(self.h)
        return EMPTY

    def g_move(self, ex, before=MISSING, after=MISSING):
        self.h.require_writable()
        if _m(before) and _m(after):
            raise NotSupported("Moving a sheet to a new workbook")
        target = self.h.index_of(before.name) if not _m(before) else self.h.index_of(after.name) + 1
        me = self.si
        if target > me:
            target -= 1
        self.h.command("sheet.move", {"sheet": me, "to": target})
        return EMPTY

    def g_visible(self, ex):
        return {"visible": -1, "hidden": 0, "veryHidden": 2}[self._sheet().visibility]

    def l_visible(self, ex, v):
        self.h.require_writable()
        pv = _plain(v)
        n = (-1 if pv else 0) if type(pv) is bool else V.to_long(pv)
        if n in (-1, 1):
            if self._sheet().visibility != "visible":
                self.h.command("sheet.unhide", {"sheet": self.si})
        elif n == 0:
            if self._sheet().visibility == "visible":
                self.h.command("sheet.hide", {"sheet": self.si})
        elif n == 2:
            if self._sheet().visibility == "visible":
                self.h.command("sheet.hide", {"sheet": self.si})
            d = self.h.s.doc()
            w = d.wb.clone()
            w.sheet_mut(self.si).visibility = "veryHidden"
            d.wb = w
        else:
            raise VbaError(1004, "Unable to set the Visible property of the Worksheet class")

    def g_tab(self, ex):
        return Tab(self)

    # -- protection
    def g_protect(self, ex, password=MISSING, drawing=MISSING, contents=MISSING, scenarios=MISSING, uionly=MISSING, *allow):
        self.h.require_writable()
        names = ("formatCells", "formatColumns", "formatRows", "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns", "deleteRows", "sort", "autoFilter")
        p: Dict[str, Any] = {"sheet": self.si}
        if not _m(password):
            p["password"] = _str(password)
        for k, v in zip(names, allow):
            if v is not MISSING and k != "insertHyperlinks":
                p[k] = _bool(v)
        if self._sheet().protection is None:
            self.h.command("review.protectSheet", p, self.si)
        ui = self.h.s.doc().ui_only_protection
        if not _m(uionly) and _bool(uionly):
            ui.add(self.name.lower())
        else:
            ui.discard(self.name.lower())
        return EMPTY

    def g_unprotect(self, ex, password=MISSING):
        self.h.require_writable()
        self.h.command("review.unprotectSheet", {"password": _str(password)}, self.si)
        self.h.s.doc().ui_only_protection.discard(self.name.lower())
        return EMPTY

    def g_protectcontents(self, ex):
        return self._sheet().protection is not None

    def g_protectionmode(self, ex):
        """Protected with UserInterfaceOnly (macros may change the sheet)."""
        return self._sheet().protection is not None and self.name.lower() in self.h.s.doc().ui_only_protection

    g_protectdrawingobjects = g_protectcontents
    g_protectscenarios = g_protectcontents

    # -- data
    def g_calculate(self, ex):
        self.h.require_writable()
        self.h.command("formulas.calculateNow", {})
        return EMPTY

    def g_enablecalculation(self, ex):
        return True

    def l_enablecalculation(self, ex, v):
        pass

    def g_paste(self, ex, destination=MISSING, link=MISSING):
        self.h.require_writable()
        if _m(destination):
            sel = self.h.selection()
            at = a1(*sel.active)
        else:
            at = a1(destination.first.r0, destination.first.c0)
        self.h.command("edit.paste", {"at": at}, self.si)
        return EMPTY

    def g_listobjects(self, ex, index=MISSING):
        c = ListObjects(self.h, self)
        return c if _m(index) else c.g_item(ex, index)

    def g_names(self, ex, index=MISSING):
        n = Names(self.h, scope=self)
        return n if _m(index) else n.g_item(ex, index)

    def g_hyperlinks(self, ex, index=MISSING):
        hl = Hyperlinks(self.h, self, None)
        return hl if _m(index) else hl.g_item(ex, index)

    def g_comments(self, ex, index=MISSING):
        c = Comments(self.h, self)
        return c if _m(index) else c.g_item(ex, index)

    def g_autofiltermode(self, ex):
        return self._sheet().autofilter is not None

    def l_autofiltermode(self, ex, v):
        if not _bool(v) and self._sheet().autofilter is not None:
            self.h.require_writable()
            self.h.command("data.filter", {"on": False}, self.si)

    def g_filtermode(self, ex):
        af = self._sheet().autofilter
        return af is not None and bool(af.criteria)

    def g_showalldata(self, ex):
        self.h.require_writable()
        if self._sheet().autofilter is not None:
            self.h.command("data.clearFilter", {}, self.si)
        return EMPTY

    def g_autofilter(self, ex):
        if self._sheet().autofilter is None:
            return NOTHING
        return AutoFilter(self)

    def g_sort(self, ex):
        return self.h.sort_state(self)

    def g_pagesetup(self, ex):
        return PageSetup(self)

    def g_standardwidth(self, ex):
        from .range import px_to_chars

        return round(px_to_chars(self._sheet().default_col_width), 2)

    def l_standardwidth(self, ex, v):
        from .range import chars_to_px

        self.h.require_writable()
        self.h.command("home.defaultWidth", {"width": chars_to_px(V.to_double(_plain(v)))}, self.si)

    def g_standardheight(self, ex):
        return round(self._sheet().default_row_height * 0.75, 2)

    def g_scrollarea(self, ex):
        return ""

    def l_scrollarea(self, ex, v):
        pass

    def g_displaypagebreaks(self, ex):
        return False

    def l_displaypagebreaks(self, ex, v):
        pass

    def g_displayrighttoleft(self, ex):
        return self._sheet().right_to_left

    def l_displayrighttoleft(self, ex, v):
        self.h.require_writable()
        self.h.command("view.rightToLeft", {"on": _bool(v)}, self.si)

    def g_enableselection(self, ex):
        return 0

    def l_enableselection(self, ex, v):
        pass

    def g_circularreference(self, ex):
        return NOTHING

    def g_shapes(self, ex, index=MISSING):
        s = Shapes(self)
        return s if _m(index) else s.g_item(ex, index)

    def g_chartobjects(self, ex, index=MISSING):
        s = ChartObjects(self)
        return s if _m(index) else s.g_item(ex, index)

    def g_printout(self, ex, *args):
        raise NotSupported("Printing from a macro")

    g_printpreview = g_printout

    def g_outline(self, ex):
        raise NotSupported("Outlines (grouping rows and columns)")

    def g_pivottables(self, ex, *args):
        raise NotSupported("PivotTables")

    def g_oleobjects(self, ex, *args):
        raise NotSupported("ActiveX controls (OLEObjects)")


class Tab(HostObject):
    TYPE_NAME = "Tab"

    def __init__(self, ws: Worksheet):
        self.ws = ws

    def g_color(self, ex):
        c = self.ws._sheet().tab_color
        if c is None or c.is_auto:
            return False
        rgb = c.resolve(self.ws.h.wb().theme)
        from .range import rgb_to_bgr

        return float(rgb_to_bgr(rgb)) if rgb else False

    def l_color(self, ex, v):
        self.ws.h.require_writable()
        pv = _plain(v)
        color = None if (type(pv) is bool and not pv) or (not isinstance(pv, bool) and V.to_long(pv) < 0) else color_value(pv)
        self.ws.h.command("sheet.tabColor", {"sheet": self.ws.si, "color": color})

    def g_colorindex(self, ex):
        return -4142 if self.ws._sheet().tab_color is None else 1

    def l_colorindex(self, ex, v):
        self.ws.h.require_writable()
        n = _long(v)
        self.ws.h.command("sheet.tabColor", {"sheet": self.ws.si, "color": None if n in (-4142, -4105) else palette_hex(n)})

    def g_themecolor(self, ex):
        return 0

    def l_themecolor(self, ex, v):
        self.ws.h.require_writable()
        n = _long(v)
        self.ws.h.command("sheet.tabColor", {"sheet": self.ws.si, "color": f"theme:{max(n - 1, 0)}:0"})

    def g_tintandshade(self, ex):
        return 0.0

    def l_tintandshade(self, ex, v):
        pass


# ---------------------------------------------------------------- window


class Window(HostObject):
    TYPE_NAME = "Window"
    PARAMS = {"smallscroll": ("Down", "Up", "ToRight", "ToLeft"), "largescroll": ("Down", "Up", "ToRight", "ToLeft"), "scrollworkbooktabs": ("Sheets", "Position")}

    def __init__(self, h):
        self.h = h

    def _ws(self):
        return self.h.app.g_activesheet(None)

    def g_activesheet(self, ex):
        return self._ws()

    def g_activecell(self, ex):
        return self.h.app.g_activecell(ex)

    def g_selection(self, ex):
        return self.h.app.g_selection(ex)

    g_rangeselection = g_selection

    def g_selectedsheets(self, ex):
        return Sheets(self.h, False)

    def g_caption(self, ex):
        return self.h.workbook.name()

    def l_caption(self, ex, v):
        pass

    def g_freezepanes(self, ex):
        return self._ws()._sheet().freeze is not None

    def l_freezepanes(self, ex, v):
        self.h.require_writable()
        if _bool(v):
            sel = self.h.selection()
            self.h.command("view.freezePanes", {"cell": a1(*sel.active)})
        else:
            self.h.command("view.unfreezePanes", {})

    def g_splitrow(self, ex):
        f = self._ws()._sheet().freeze
        return f[0] if f else 0

    def l_splitrow(self, ex, v):
        n = _long(v)
        f = self._ws()._sheet().freeze or (0, 0)
        self._freeze(n, f[1])

    def g_splitcolumn(self, ex):
        f = self._ws()._sheet().freeze
        return f[1] if f else 0

    def l_splitcolumn(self, ex, v):
        n = _long(v)
        f = self._ws()._sheet().freeze or (0, 0)
        self._freeze(f[0], n)

    def _freeze(self, rows: int, cols: int) -> None:
        self.h.require_writable()
        if rows <= 0 and cols <= 0:
            self.h.command("view.unfreezePanes", {})
        else:
            self.h.command("view.freezePanes", {"cell": a1(max(rows, 0), max(cols, 0))})

    def g_split(self, ex):
        return self.g_freezepanes(ex)

    def l_split(self, ex, v):
        if not _bool(v):
            self._freeze(0, 0)

    def g_zoom(self, ex):
        return self._ws()._sheet().zoom

    def l_zoom(self, ex, v):
        self.h.require_writable()
        pv = _plain(v)
        if type(pv) is bool:
            return
        self.h.command("view.zoom", {"percent": max(10, min(V.to_long(pv), 400))})

    def g_scrollrow(self, ex):
        return self._ws()._sheet().view_top_left[0] + 1

    def l_scrollrow(self, ex, v):
        pass

    def g_scrollcolumn(self, ex):
        return self._ws()._sheet().view_top_left[1] + 1

    def l_scrollcolumn(self, ex, v):
        pass

    def g_smallscroll(self, ex, *args):
        return EMPTY

    g_largescroll = g_smallscroll
    g_scrollworkbooktabs = g_smallscroll

    def g_displaygridlines(self, ex):
        return self._ws()._sheet().show_gridlines

    def l_displaygridlines(self, ex, v):
        self.h.require_writable()
        self.h.command("view.gridlines", {"on": _bool(v)})

    def g_displayheadings(self, ex):
        return self._ws()._sheet().show_headings

    def l_displayheadings(self, ex, v):
        self.h.require_writable()
        self.h.command("view.headings", {"on": _bool(v)})

    def g_displayformulas(self, ex):
        return self._ws()._sheet().show_formulas

    def l_displayformulas(self, ex, v):
        self.h.require_writable()
        self.h.command("formulas.showFormulas", {"on": _bool(v)})

    def g_displayzeros(self, ex):
        return self._ws()._sheet().show_zeros

    def l_displayzeros(self, ex, v):
        self.h.require_writable()
        self.h.command("view.showZeros", {"on": _bool(v)})

    def g_displayworkbooktabs(self, ex):
        return True

    def l_displayworkbooktabs(self, ex, v):
        pass

    def g_displayhorizontalscrollbar(self, ex):
        return True

    def l_displayhorizontalscrollbar(self, ex, v):
        pass

    g_displayverticalscrollbar = g_displayhorizontalscrollbar
    l_displayverticalscrollbar = l_displayhorizontalscrollbar

    def g_windowstate(self, ex):
        return -4137

    def l_windowstate(self, ex, v):
        pass

    def g_view(self, ex):
        return 1

    def l_view(self, ex, v):
        pass

    def g_visible(self, ex):
        return True

    def l_visible(self, ex, v):
        pass

    def g_activate(self, ex):
        return EMPTY

    def g_width(self, ex):
        return 1000.0

    def g_height(self, ex):
        return 600.0

    def g_visiblerange(self, ex):
        ws = self._ws()
        sh = ws._sheet()
        r0, c0 = sh.view_top_left
        return Range(self.h, ws, [RangeRef(r0, c0, min(r0 + 30, MAX_ROWS - 1), min(c0 + 15, MAX_COLS - 1))])


class Windows(HostObject):
    TYPE_NAME = "Windows"
    DEFAULT = "item"

    def __init__(self, h):
        self.h = h

    def g_count(self, ex):
        return 1

    def g_item(self, ex, index=MISSING):
        return Window(self.h)

    def vba_enum(self, ex):
        return iter([Window(self.h)])


# ---------------------------------------------------------------- names


class Names(HostObject):
    TYPE_NAME = "Names"
    DEFAULT = "item"
    PARAMS = {"add": ("Name", "RefersTo", "Visible", "MacroType", "ShortcutKey", "Category", "NameLocal", "RefersToLocal", "CategoryLocal", "RefersToR1C1", "RefersToR1C1Local"), "item": ("Index", "IndexLocal", "RefersTo")}

    def __init__(self, h, scope: Optional[Worksheet] = None):
        self.h = h
        self.scope = scope

    def _list(self):
        wb = self.h.wb()
        if self.scope is None:
            return [n for n in wb.names]
        si = self.scope.si
        return [n for n in wb.names if n.scope == si]

    def g_count(self, ex):
        return len(self._list())

    def g_item(self, ex, index=MISSING, local=MISSING, refersto=MISSING):
        v = _plain(index) if not _m(index) else _plain(local)
        names = self._list()
        if isinstance(v, str):
            low = v.lower()
            bare = low.split("!", 1)[-1]
            for n in names:
                if n.name.lower() == bare:
                    return Name(self.h, n.name, n.scope)
            raise VbaError(1004, f"No name {v}")
        i = V.to_long(v)
        if not 1 <= i <= len(names):
            raise VbaError(9)
        n = names[i - 1]
        return Name(self.h, n.name, n.scope)

    def vba_enum(self, ex):
        return iter([Name(self.h, n.name, n.scope) for n in self._list()])

    def g_add(self, ex, name=MISSING, refersto=MISSING, visible=MISSING, macrotype=MISSING, shortcut=MISSING, category=MISSING, namelocal=MISSING, referstolocal=MISSING, categorylocal=MISSING, refersr1c1=MISSING, refersr1c1local=MISSING):
        self.h.require_writable()
        nm = _str(name if not _m(name) else namelocal)
        if _m(refersto) and _m(referstolocal) and _m(refersr1c1):
            raise VbaError(449)
        if not _m(refersr1c1):
            from .host import r1c1_formula_to_a1

            ref = r1c1_formula_to_a1(_str(refersr1c1), (0, 0))
        else:
            rt = refersto if not _m(refersto) else referstolocal
            if isinstance(rt, Range):
                from ...formula.printer import quote_sheet

                ref = f"={quote_sheet(rt.ws.name)}!{rt.address()}"
            else:
                ref = V.text(_plain(rt))
                if not ref.startswith("="):
                    ref = "=" + (ref if not _looks_number(ref) else ref)
        scope = None
        if "!" in nm:
            sheet, nm = nm.rsplit("!", 1)
            scope = sheet.strip("'")
        elif self.scope is not None:
            scope = self.scope.name
        p: Dict[str, Any] = {"name": nm, "refersTo": ref, "replace": True}
        if scope:
            p["scope"] = scope
        self.h.command("formulas.defineName", p)
        if not _m(visible) and not _bool(visible):
            _set_name_hidden(self.h, nm, scope, True)
        wb = self.h.wb()
        si = wb.sheet_index(scope) if scope else None
        return Name(self.h, nm, si)


def _looks_number(t: str) -> bool:
    return V.parse_number(t) is not None


def _set_name_hidden(h, name: str, scope: Optional[str], hidden: bool) -> None:
    d = h.s.doc()
    w = d.wb.clone()
    si = w.sheet_index(scope) if scope else None
    for n in w.names:
        if n.name.lower() == name.lower() and n.scope == si:
            n.hidden = hidden
    d.wb = w


class Name(HostObject):
    TYPE_NAME = "Name"
    DEFAULT = "_default"

    def __init__(self, h, name: str, scope: Optional[int]):
        self.h = h
        self.name = name
        self.scope = scope

    def _def(self):
        for n in self.h.wb().names:
            if n.name.lower() == self.name.lower() and n.scope == self.scope:
                return n
        raise VbaError(1004, "The name was deleted")

    def g_name(self, ex):
        if self.scope is not None:
            from ...formula.printer import quote_sheet

            return f"{quote_sheet(self.h.wb().sheets[self.scope].name)}!{self.name}"
        return self.name

    g_namelocal = g_name

    def l_name(self, ex, v):
        self.h.require_writable()
        d = self._def()
        new = _str(v)
        p: Dict[str, Any] = {"name": new, "refersTo": "=" + d.formula.lstrip("="), "comment": d.comment}
        if self.scope is not None:
            p["scope"] = self.h.wb().sheets[self.scope].name
        self.h.command("formulas.defineName", p)
        self.h.command("formulas.deleteName", {k: v for k, v in (("name", self.name), ("scope", p.get("scope"))) if v})
        self.name = new

    def g_refersto(self, ex):
        return "=" + self._def().formula.lstrip("=")

    g_referstolocal = g_refersto

    def g__default(self, ex):
        return self.g_refersto(ex)

    def g_value(self, ex):
        return self.g_refersto(ex)

    def l_refersto(self, ex, v):
        self.h.require_writable()
        ref = v if not isinstance(v, Range) else v
        if isinstance(ref, Range):
            from ...formula.printer import quote_sheet

            text = f"={quote_sheet(ref.ws.name)}!{ref.address()}"
        else:
            text = V.text(_plain(ref))
            if not text.startswith("="):
                text = "=" + text
        p: Dict[str, Any] = {"name": self.name, "refersTo": text, "replace": True}
        if self.scope is not None:
            p["scope"] = self.h.wb().sheets[self.scope].name
        self.h.command("formulas.defineName", p)

    l_value = l_refersto
    l_referstolocal = l_refersto

    def g_referstorange(self, ex):
        d = self._def()
        ws = self.h.app.g_activesheet(ex) if self.scope is None else self.h.sheet_object(self.h.wb().sheets[self.scope].name)
        try:
            r = parse_reference(self.h, ws, d.formula.lstrip("="))
        except VbaError:
            raise VbaError(1004, f"The name {self.name} doesn't refer to a range")
        return r

    def g_referstor1c1(self, ex):
        from .host import a1_formula_to_r1c1

        return "=" + a1_formula_to_r1c1(self._def().formula.lstrip("="), (0, 0))

    def g_visible(self, ex):
        return not self._def().hidden

    def l_visible(self, ex, v):
        self.h.require_writable()
        _set_name_hidden(self.h, self.name, self.h.wb().sheets[self.scope].name if self.scope is not None else None, not _bool(v))

    def g_comment(self, ex):
        return self._def().comment

    def g_delete(self, ex):
        self.h.require_writable()
        p: Dict[str, Any] = {"name": self.name}
        if self.scope is not None:
            p["scope"] = self.h.wb().sheets[self.scope].name
        self.h.command("formulas.deleteName", p)
        return EMPTY

    def g_parent(self, ex):
        return self.h.workbook if self.scope is None else self.h.sheet_object(self.h.wb().sheets[self.scope].name)


# ---------------------------------------------------------------- tables


class ListObjects(HostObject):
    TYPE_NAME = "ListObjects"
    DEFAULT = "item"
    PARAMS = {"add": ("SourceType", "Source", "LinkSource", "XlListObjectHasHeaders", "Destination", "TableStyleName")}

    def __init__(self, h, ws: Worksheet):
        self.h = h
        self.ws = ws

    def _tables(self):
        return list(self.ws._sheet().tables)

    def g_count(self, ex):
        return len(self._tables())

    def g_item(self, ex, index=MISSING):
        v = _plain(index)
        ts = self._tables()
        if isinstance(v, str):
            for t in ts:
                if t.name.lower() == v.lower():
                    return ListObject(self.h, self.ws, t.name)
            raise VbaError(9)
        i = V.to_long(v)
        if not 1 <= i <= len(ts):
            raise VbaError(9)
        return ListObject(self.h, self.ws, ts[i - 1].name)

    def vba_enum(self, ex):
        return iter([ListObject(self.h, self.ws, t.name) for t in self._tables()])

    def g_add(self, ex, sourcetype=MISSING, source=MISSING, link=MISSING, headers=MISSING, dest=MISSING, style=MISSING):
        self.h.require_writable()
        if _long(sourcetype, 1) != 1:
            raise NotSupported("Tables from external data")
        if _m(source):
            sel = self.h.selection()
            from ..selection import current_region

            rng = current_region(self.ws._sheet(), sel.active)
        else:
            rng = source.first if isinstance(source, Range) else make_range(self.h, self.ws, source).first
        p: Dict[str, Any] = {"range": rng.a1()}
        hv = _long(headers, 0)
        if hv == 1:
            p["header"] = True
        elif hv == 2:
            p["header"] = False
        if not _m(style):
            p["style"] = _str(style)
        before = {t.name for t in self.ws._sheet().tables}
        self.h.command("insert.table", p, self.ws.si)
        new = [t.name for t in self.ws._sheet().tables if t.name not in before]
        return ListObject(self.h, self.ws, new[0] if new else self.ws._sheet().tables[-1].name)


class ListObject(HostObject):
    TYPE_NAME = "ListObject"

    def __init__(self, h, ws: Worksheet, name: str):
        self.h = h
        self.ws = ws
        self.name = name

    def _t(self):
        for t in self.ws._sheet().tables:
            if t.name.lower() == self.name.lower():
                return t
        raise VbaError(1004, f"The table {self.name} no longer exists")

    def g_name(self, ex):
        return self._t().name

    def l_name(self, ex, v):
        self.h.require_writable()
        new = _str(v)
        self.h.command("table.rename", {"name": new, "table": self.name}, self.ws.si)
        self.name = new

    g_displayname = g_name
    l_displayname = l_name

    def g_range(self, ex):
        return Range(self.h, self.ws, [self._t().range])

    def g_databodyrange(self, ex):
        d = self._t().data_range()
        return Range(self.h, self.ws, [d]) if d is not None else NOTHING

    def g_headerrowrange(self, ex):
        t = self._t()
        if not t.header_row:
            return NOTHING
        return Range(self.h, self.ws, [RangeRef(t.range.r0, t.range.c0, t.range.r0, t.range.c1)])

    def g_totalsrowrange(self, ex):
        t = self._t()
        if not t.totals_row:
            return NOTHING
        return Range(self.h, self.ws, [RangeRef(t.range.r1, t.range.c0, t.range.r1, t.range.c1)])

    def g_listcolumns(self, ex, index=MISSING):
        c = ListColumns(self)
        return c if _m(index) else c.g_item(ex, index)

    def g_listrows(self, ex, index=MISSING):
        c = ListRows(self)
        return c if _m(index) else c.g_item(ex, index)

    def g_showtotals(self, ex):
        return self._t().totals_row

    def l_showtotals(self, ex, v):
        self.h.require_writable()
        self.h.command("table.totalRow", {"on": _bool(v), "table": self.name}, self.ws.si)

    def g_showheaders(self, ex):
        return self._t().header_row

    def l_showheaders(self, ex, v):
        self.h.require_writable()
        self.h.command("table.headerRow", {"on": _bool(v), "table": self.name}, self.ws.si)

    def g_tablestyle(self, ex):
        return self._t().style or ""

    def l_tablestyle(self, ex, v):
        self.h.require_writable()
        self.h.command("table.style", {"style": _str(v), "table": self.name}, self.ws.si)

    def g_resize(self, ex, rng=MISSING):
        self.h.require_writable()
        self.h.command("table.resize", {"range": rng.first.a1(), "table": self.name}, self.ws.si)
        return EMPTY

    def g_unlist(self, ex):
        self.h.require_writable()
        self.h.command("table.convertToRange", {"table": self.name}, self.ws.si)
        return EMPTY

    def g_delete(self, ex):
        self.h.require_writable()
        r = self._t().range
        self.h.command("table.convertToRange", {"table": self.name}, self.ws.si)
        self.h.command("edit.clearAll", {"range": r.a1()}, self.ws.si)
        return EMPTY

    def g_parent(self, ex):
        return self.ws

    def g_showautofilter(self, ex):
        return True


class ListColumns(HostObject):
    TYPE_NAME = "ListColumns"
    DEFAULT = "item"
    PARAMS = {"add": ("Position",)}

    def __init__(self, lo: ListObject):
        self.lo = lo

    def g_count(self, ex):
        return len(self.lo._t().columns)

    def g_item(self, ex, index=MISSING):
        t = self.lo._t()
        v = _plain(index)
        if isinstance(v, str):
            i = t.column_index(v)
            if i is None:
                raise VbaError(9)
        else:
            i = V.to_long(v) - 1
            if not 0 <= i < len(t.columns):
                raise VbaError(9)
        return ListColumn(self.lo, i)

    def vba_enum(self, ex):
        return iter([ListColumn(self.lo, i) for i in range(len(self.lo._t().columns))])

    def g_add(self, ex, position=MISSING):
        self.lo.h.require_writable()
        t = self.lo._t()
        col = t.range.c1 + 1 if _m(position) else t.range.c0 + _long(position) - 1
        if col > t.range.c1:
            self.lo.h.command("table.resize", {"range": RangeRef(t.range.r0, t.range.c0, t.range.r1, t.range.c1 + 1).a1(), "table": self.lo.name}, self.lo.ws.si)
        else:
            self.lo.h.command("home.insertCells", {"shift": "right", "range": RangeRef(t.range.r0, col, t.range.r1, col).a1()}, self.lo.ws.si)
        return ListColumn(self.lo, col - self.lo._t().range.c0)


class ListColumn(HostObject):
    TYPE_NAME = "ListColumn"

    def __init__(self, lo: ListObject, i: int):
        self.lo = lo
        self.i = i

    def g_name(self, ex):
        return self.lo._t().columns[self.i].name

    def l_name(self, ex, v):
        t = self.lo._t()
        if t.header_row:
            Range(self.lo.h, self.lo.ws, [RangeRef.of((t.range.r0, t.range.c0 + self.i))]).l_value(ex, v)

    def g_index(self, ex):
        return self.i + 1

    def g_range(self, ex):
        t = self.lo._t()
        c = t.range.c0 + self.i
        return Range(self.lo.h, self.lo.ws, [RangeRef(t.range.r0, c, t.range.r1, c)])

    def g_databodyrange(self, ex):
        t = self.lo._t()
        d = t.data_range()
        if d is None:
            return NOTHING
        c = t.range.c0 + self.i
        return Range(self.lo.h, self.lo.ws, [RangeRef(d.r0, c, d.r1, c)])

    def g_totalscalculation(self, ex):
        return 0

    def l_totalscalculation(self, ex, v):
        fn = {0: "none", 1: "sum", 2: "average", 3: "count", 4: "countNums", 5: "min", 6: "max", 7: "stdDev", 8: "var"}.get(_long(v), "none")
        self.lo.h.command("table.totalFunction", {"col": self.i, "function": fn, "table": self.lo.name}, self.lo.ws.si)

    def g_delete(self, ex):
        t = self.lo._t()
        c = t.range.c0 + self.i
        self.lo.h.command("home.deleteCells", {"shift": "left", "range": RangeRef(t.range.r0, c, t.range.r1, c).a1()}, self.lo.ws.si)
        return EMPTY


class ListRows(HostObject):
    TYPE_NAME = "ListRows"
    DEFAULT = "item"
    PARAMS = {"add": ("Position", "AlwaysInsert")}

    def __init__(self, lo: ListObject):
        self.lo = lo

    def g_count(self, ex):
        d = self.lo._t().data_range()
        return d.height if d is not None else 0

    def g_item(self, ex, index=MISSING):
        i = _long(index)
        if not 1 <= i <= self.g_count(ex):
            raise VbaError(9)
        return ListRow(self.lo, i - 1)

    def vba_enum(self, ex):
        return iter([ListRow(self.lo, i) for i in range(self.g_count(ex))])

    def g_add(self, ex, position=MISSING, always=MISSING):
        self.lo.h.require_writable()
        t = self.lo._t()
        d = t.data_range()
        n = d.height if d is not None else 0
        end = t.range.r1 - (1 if t.totals_row else 0)
        if _m(position) or _long(position) > n:
            if t.totals_row:
                self.lo.h.command("home.insertCells", {"shift": "down", "range": RangeRef(t.range.r1, t.range.c0, t.range.r1, t.range.c1).a1()}, self.lo.ws.si)
            else:
                self.lo.h.command("table.resize", {"range": RangeRef(t.range.r0, t.range.c0, t.range.r1 + 1, t.range.c1).a1(), "table": self.lo.name}, self.lo.ws.si)
            return ListRow(self.lo, n)
        p = _long(position) - 1
        if p < 0:
            raise VbaError(5)
        r = (d.r0 if d is not None else end) + p
        before = t.range
        self.lo.h.command("home.insertCells", {"shift": "down", "range": RangeRef(r, t.range.c0, r, t.range.c1).a1()}, self.lo.ws.si)
        t = self.lo._t()
        if t.range.r1 == before.r1:
            # The rows below moved down: the table takes them back in.
            self.lo.h.command("table.resize", {"range": RangeRef(t.range.r0, t.range.c0, t.range.r1 + 1, t.range.c1).a1(), "table": self.lo.name}, self.lo.ws.si)
        return ListRow(self.lo, p)


class ListRow(HostObject):
    TYPE_NAME = "ListRow"

    def __init__(self, lo: ListObject, i: int):
        self.lo = lo
        self.i = i

    def _rr(self) -> RangeRef:
        t = self.lo._t()
        d = t.data_range()
        if d is None:
            raise VbaError(1004)
        r = d.r0 + self.i
        return RangeRef(r, t.range.c0, r, t.range.c1)

    def g_range(self, ex):
        return Range(self.lo.h, self.lo.ws, [self._rr()])

    def g_index(self, ex):
        return self.i + 1

    def g_delete(self, ex):
        self.lo.h.require_writable()
        self.lo.h.command("home.deleteCells", {"shift": "up", "range": self._rr().a1()}, self.lo.ws.si)
        return EMPTY


# ---------------------------------------------------------------- hyperlinks and comments


class Hyperlinks(HostObject):
    TYPE_NAME = "Hyperlinks"
    DEFAULT = "item"
    PARAMS = {"add": ("Anchor", "Address", "SubAddress", "ScreenTip", "TextToDisplay")}

    def __init__(self, h, ws: Worksheet, rng: Optional[Range]):
        self.h = h
        self.ws = ws
        self.rng = rng

    def _cells(self) -> List[Tuple[int, int]]:
        links = self.ws._sheet().hyperlinks
        keys = sorted(links)
        if self.rng is not None:
            keys = [k for k in keys if any(a.contains(k) for a in self.rng.areas)]
        return keys

    def g_count(self, ex):
        return len(self._cells())

    def g_item(self, ex, index=MISSING):
        ks = self._cells()
        i = _long(index)
        if not 1 <= i <= len(ks):
            raise VbaError(9)
        return Hyperlink(self.h, self.ws, ks[i - 1])

    def vba_enum(self, ex):
        return iter([Hyperlink(self.h, self.ws, k) for k in self._cells()])

    def g_add(self, ex, anchor=MISSING, address=MISSING, subaddress=MISSING, tip=MISSING, text=MISSING):
        self.h.require_writable()
        if not isinstance(anchor, Range):
            raise NotSupported("Links on shapes")
        target = _str(address)
        if not target and not _m(subaddress):
            target = _str(subaddress)
        elif not _m(subaddress) and _str(subaddress):
            target = target + "#" + _str(subaddress)
        p: Dict[str, Any] = {"cell": a1(anchor.first.r0, anchor.first.c0), "target": target}
        if not _m(text):
            p["text"] = _str(text)
        if not _m(tip):
            p["tooltip"] = _str(tip)
        self.h.command("insert.hyperlink", p, anchor.si)
        return Hyperlink(self.h, anchor.ws, (anchor.first.r0, anchor.first.c0))

    def g_delete(self, ex):
        self.h.require_writable()
        rng = self.rng.first.a1() if self.rng is not None else RangeRef(0, 0, MAX_ROWS - 1, MAX_COLS - 1).a1()
        self.h.command("insert.removeHyperlink", {"range": rng}, self.ws.si)
        return EMPTY


class Hyperlink(HostObject):
    TYPE_NAME = "Hyperlink"

    def __init__(self, h, ws: Worksheet, rc: Tuple[int, int]):
        self.h = h
        self.ws = ws
        self.rc = rc

    def _l(self):
        l = self.ws._sheet().hyperlinks.get(self.rc)
        if l is None:
            raise VbaError(1004, "The link was removed")
        return l

    def g_address(self, ex):
        t = self._l().target
        return t if "://" in t or t.startswith("mailto:") else ""

    def g_subaddress(self, ex):
        t = self._l().target
        if "://" in t or t.startswith("mailto:"):
            return t.split("#", 1)[1] if "#" in t else ""
        return t.lstrip("#")

    def g_texttodisplay(self, ex):
        return V.text(Range(self.h, self.ws, [RangeRef.of(self.rc)]).g_value(ex))

    def g_screentip(self, ex):
        return getattr(self._l(), "tooltip", "") or ""

    def g_range(self, ex):
        return Range(self.h, self.ws, [RangeRef.of(self.rc)])

    def g_name(self, ex):
        return self.g_texttodisplay(ex)

    def g_delete(self, ex):
        self.h.require_writable()
        self.h.command("insert.removeHyperlink", {"range": a1(*self.rc)}, self.ws.si)
        return EMPTY

    def g_follow(self, ex, *args):
        raise NotSupported("Opening links from a macro")


class Comments(HostObject):
    TYPE_NAME = "Comments"
    DEFAULT = "item"

    def __init__(self, h, ws: Worksheet):
        self.h = h
        self.ws = ws

    def _keys(self):
        return sorted(self.ws._sheet().comments)

    def g_count(self, ex):
        return len(self._keys())

    def g_item(self, ex, index=MISSING):
        from .range import Comment

        ks = self._keys()
        i = _long(index)
        if not 1 <= i <= len(ks):
            raise VbaError(9)
        return Comment(self.h, self.ws, ks[i - 1])

    def vba_enum(self, ex):
        from .range import Comment

        return iter([Comment(self.h, self.ws, k) for k in self._keys()])


# ---------------------------------------------------------------- validation and formats


class Validation(HostObject):
    TYPE_NAME = "Validation"
    PARAMS = {"add": ("Type", "AlertStyle", "Operator", "Formula1", "Formula2"), "modify": ("Type", "AlertStyle", "Operator", "Formula1", "Formula2")}
    _KINDS = {0: "any", 1: "whole", 2: "decimal", 3: "list", 4: "date", 5: "time", 6: "textLength", 7: "custom"}
    _OPS = {1: "between", 2: "notBetween", 3: "equal", 4: "notEqual", 5: "greaterThan", 6: "lessThan", 7: "greaterThanOrEqual", 8: "lessThanOrEqual"}

    def __init__(self, rng: Range):
        self.rng = rng

    def _v(self):
        from ..entry import validation_at

        sh = self.rng._sheet()
        return validation_at(sh, (self.rng.first.r0, self.rng.first.c0))

    def _require(self):
        v = self._v()
        if v is None:
            raise VbaError(1004, "The cells have no data validation")
        return v

    def g_add(self, ex, type_=MISSING, alert=MISSING, operator=MISSING, f1=MISSING, f2=MISSING):
        h = self.rng.h
        h.require_writable()
        if self._v() is not None:
            raise VbaError(1004, "Add method of Validation class failed (the cells already have validation; Delete it first)")
        return self._set(type_, alert, operator, f1, f2)

    def g_modify(self, ex, type_=MISSING, alert=MISSING, operator=MISSING, f1=MISSING, f2=MISSING):
        self.rng.h.require_writable()
        return self._set(type_, alert, operator, f1, f2)

    def _set(self, type_, alert, operator, f1, f2):
        kind = self._KINDS.get(_long(type_, 0), "any")
        p: Dict[str, Any] = {"range": self.rng.address(False, False), "kind": kind}
        if not _m(operator):
            p["operator"] = self._OPS.get(_long(operator), "between")
        if not _m(f1):
            t = V.text(_plain(f1))
            p["f1"] = t
        if not _m(f2):
            p["f2"] = V.text(_plain(f2))
        if not _m(alert):
            p["errorStyle"] = {1: "stop", 2: "warning", 3: "information"}.get(_long(alert), "stop")
        for r in self.rng.areas:
            q = dict(p)
            q["range"] = r.a1()
            self.rng.h.command("data.validation", q, self.rng.si)
        return EMPTY

    def g_delete(self, ex):
        self.rng.h.require_writable()
        for r in self.rng.areas:
            self.rng.h.command("data.clearValidation", {"range": r.a1()}, self.rng.si)
        return EMPTY

    def g_type(self, ex):
        v = self._require()
        return {k: n for n, k in self._KINDS.items()}.get(v.kind, 0)

    def g_operator(self, ex):
        v = self._require()
        return {k: n for n, k in self._OPS.items()}.get(v.operator, 1)

    def g_formula1(self, ex):
        v = self._require()
        f = v.f1 or ""
        if v.kind == "list" and f.startswith('"'):
            return f.strip('"')
        return f if f.startswith("=") or v.kind == "list" else ("=" + f if f else "")

    def g_formula2(self, ex):
        v = self._require()
        return ("=" + v.f2) if v.f2 and not v.f2.startswith("=") else (v.f2 or "")

    def _update(self, **kw) -> None:
        v = self._require()
        p = {
            "range": ",".join(r.a1() for r in v.ranges),
            "kind": v.kind,
            "operator": v.operator,
            "f1": v.f1,
            "f2": v.f2,
            "allowBlank": v.allow_blank,
            "inCellDropdown": getattr(v, "in_cell_dropdown", True),
            "inputTitle": v.input_title,
            "inputMessage": v.input_message,
            "showInput": getattr(v, "show_input", True),
            "errorTitle": v.error_title,
            "errorMessage": v.error_message,
            "errorStyle": v.error_style,
            "showError": v.show_error,
        }
        p.update(kw)
        self.rng.h.require_writable()
        self.rng.h.command("data.validation", p, self.rng.si)

    def g_ignoreblank(self, ex):
        return self._require().allow_blank

    def l_ignoreblank(self, ex, v):
        self._update(allowBlank=_bool(v))

    def g_incelldropdown(self, ex):
        return getattr(self._require(), "in_cell_dropdown", True)

    def l_incelldropdown(self, ex, v):
        self._update(inCellDropdown=_bool(v))

    def g_inputtitle(self, ex):
        return self._require().input_title

    def l_inputtitle(self, ex, v):
        self._update(inputTitle=_str(v))

    def g_inputmessage(self, ex):
        return self._require().input_message

    def l_inputmessage(self, ex, v):
        self._update(inputMessage=_str(v))

    def g_errortitle(self, ex):
        return self._require().error_title

    def l_errortitle(self, ex, v):
        self._update(errorTitle=_str(v))

    def g_errormessage(self, ex):
        return self._require().error_message

    def l_errormessage(self, ex, v):
        self._update(errorMessage=_str(v))

    def g_showinput(self, ex):
        return getattr(self._require(), "show_input", True)

    def l_showinput(self, ex, v):
        self._update(showInput=_bool(v))

    def g_showerror(self, ex):
        return self._require().show_error

    def l_showerror(self, ex, v):
        self._update(showError=_bool(v))

    def g_value(self, ex):
        from ..entry import is_valid

        v = self._v()
        if v is None:
            return True
        h = self.rng.h
        rc = (self.rng.first.r0, self.rng.first.c0)
        return bool(is_valid(h.wb(), self.rng.si, rc, h.read_value(self.rng.si, rc), v))


class FormatConditions(HostObject):
    """``Range.FormatConditions``: the conditional formats over a range, top priority first."""

    TYPE_NAME = "FormatConditions"
    DEFAULT = "Item"
    PARAMS = {"add": ("Type", "Operator", "Formula1", "Formula2", "String", "TextOperator", "DateOperator", "ScopeType"), "addcolorscale": ("ColorScaleType",), "addtop10": (), "addiconsetcondition": ()}

    def __init__(self, rng: Range):
        self.rng = rng

    def _indexes(self) -> List[int]:
        sh = self.rng._sheet()
        return [i for i, cf in enumerate(sh.cond_formats) if any(r.intersects(a) for r in cf.ranges for a in self.rng.areas)]

    def g_count(self, ex):
        return len(self._indexes())

    def g_item(self, ex, index=MISSING):
        idx = self._indexes()
        k = _long(index)
        if not 1 <= k <= len(idx):
            raise VbaError(9)
        return FormatCondition(self.rng, idx[k - 1])

    def vba_enum(self, ex):
        return iter([FormatCondition(self.rng, i) for i in self._indexes()])

    def g_delete(self, ex):
        self.rng.h.require_writable()
        for i in reversed(self._indexes()):
            self.rng.h.command("home.deleteRule", {"index": i}, self.rng.si)
        return EMPTY

    def _add(self, rule: Dict[str, Any]) -> "FormatCondition":
        self.rng.h.require_writable()
        rule.setdefault("style", {})  # VBA's new rules have no formatting until it is set
        ranges = ",".join(r.a1() for r in self.rng.areas)
        self.rng.h.command("home.conditionalFormat", {"range": ranges, "rule": rule}, self.rng.si)
        return FormatCondition(self.rng, 0)  # new rules go on top

    def g_add(self, ex, type_=MISSING, operator=MISSING, f1=MISSING, f2=MISSING, string=MISSING, textop=MISSING, dateop=MISSING, scope=MISSING):
        t = _long(type_, 1)
        text = lambda v: V.text(_plain(v))
        if t == 1:  # xlCellValue
            op = _CF_OPERATORS.get(_long(operator, 1))
            if op is None:
                raise VbaError(5)
            rule: Dict[str, Any] = {"type": "cellIs", "operator": op, "value": _cf_value(f1)}
            if op in ("between", "notBetween"):
                rule["value2"] = _cf_value(f2)
            return self._add(rule)
        if t == 2:  # xlExpression
            return self._add({"type": "expression", "formula": text(f1)})
        if t == 9:  # xlTextString
            kind = {0: "containsText", 1: "notContainsText", 2: "beginsWith", 3: "endsWith"}.get(_long(textop, 0), "containsText")
            return self._add({"type": kind, "text": text(string)})
        if t in (10, 13, 16, 17):
            return self._add({"type": {10: "containsBlanks", 13: "notContainsBlanks", 16: "containsErrors", 17: "notContainsErrors"}[t]})
        if t == 11:  # xlTimePeriod
            period = {0: "today", 1: "yesterday", 2: "last7Days", 3: "thisWeek", 4: "lastWeek", 5: "lastMonth", 6: "tomorrow", 7: "nextWeek", 8: "nextMonth", 9: "thisMonth"}.get(_long(dateop, 0), "today")
            return self._add({"type": "timePeriod", "period": period})
        raise NotSupported("This kind of conditional format (FormatConditions.Add Type:=%d)" % t)

    def g_addtop10(self, ex):
        return self._add({"type": "top10", "rank": 10})

    def g_addaboveaverage(self, ex):
        return self._add({"type": "aboveAverage"})

    def g_adduniquevalues(self, ex):
        return self._add({"type": "duplicateValues"})

    def g_addcolorscale(self, ex, kind=MISSING):
        n = _long(kind, 3)
        return self._add({"type": "colorScale", "colors": ["#F8696B", "#FFEB84", "#63BE7B"] if n == 3 else ["#F8696B", "#63BE7B"], "style": None})

    def g_adddatabar(self, ex):
        return self._add({"type": "dataBar", "style": None})

    def g_addiconsetcondition(self, ex):
        return self._add({"type": "iconSet", "style": None})


_CF_OPERATORS = {1: "between", 2: "notBetween", 3: "equal", 4: "notEqual", 5: "greaterThan", 6: "lessThan", 7: "greaterThanOrEqual", 8: "lessThanOrEqual"}
_CF_TYPES = {"cellIs": 1, "expression": 2, "colorScale": 3, "dataBar": 4, "top10": 5, "iconSet": 6, "duplicateValues": 8, "uniqueValues": 8, "containsText": 9, "notContainsText": 9, "beginsWith": 9, "endsWith": 9, "containsBlanks": 10, "timePeriod": 11, "aboveAverage": 12, "notContainsBlanks": 13, "containsErrors": 16, "notContainsErrors": 17}


def _cf_value(v: Any) -> Any:
    """A rule's value as the rule command takes it (``=formula``, a number or text)."""
    pv = _plain(v)
    if pv is MISSING or pv is EMPTY:
        return ""
    if isinstance(pv, str):
        return pv
    return float(V.to_double(pv))


class _CfStyleTarget:
    """Lets Range's Font and Interior objects format a conditional-format rule."""

    def __init__(self, cond: "FormatCondition"):
        self.cond = cond
        self.h = cond.rng.h

    def _uniform(self, f: Callable) -> Any:
        from ...model.style import Style

        return f(self.cond._cf().rule.style or Style())

    def _style(self, f: Callable) -> None:
        from ...model.style import Style

        self.cond._change(lambda cf: setattr(cf.rule, "style", f(cf.rule.style or Style())))


class FormatCondition(HostObject):
    TYPE_NAME = "FormatCondition"
    PARAMS = {"modify": ("Type", "Operator", "Formula1", "Formula2")}

    def __init__(self, rng: Range, index: int):
        self.rng = rng
        self.index = index

    def _cf(self):
        sh = self.rng._sheet()
        if not 0 <= self.index < len(sh.cond_formats):
            raise VbaError(1004, "That conditional format was deleted.")
        return sh.cond_formats[self.index]

    def _change(self, f: Callable) -> None:
        self.rng.h.require_writable()
        self._cf()
        ctx = self.rng.h.edit()
        f(ctx.sheet_mut(self.rng.si).cond_formats[self.index])

    def g_type(self, ex):
        return _CF_TYPES.get(self._cf().rule.type, 1)

    def g_operator(self, ex):
        rule = self._cf().rule
        if rule.type != "cellIs":
            raise VbaError(1004, "Unable to get the Operator property of the FormatCondition class")
        return {v: k for k, v in _CF_OPERATORS.items()}.get(rule.operator, 1)

    def _formula(self, i: int) -> str:
        rule = self._cf().rule
        if rule.type not in ("cellIs", "expression") or len(rule.formulas) <= i:
            raise VbaError(1004, f"Unable to get the Formula{i + 1} property of the FormatCondition class")
        return "=" + rule.formulas[i]

    def g_formula1(self, ex):
        return self._formula(0)

    def g_formula2(self, ex):
        return self._formula(1)

    def g_text(self, ex):
        return self._cf().rule.text

    def g_interior(self, ex):
        from .range import Interior

        return Interior(_CfStyleTarget(self))  # type: ignore[arg-type]

    def g_font(self, ex):
        from .range import Font

        return Font(_CfStyleTarget(self))  # type: ignore[arg-type]

    def g_numberformat(self, ex):
        return self._uniform_style(lambda st: st.num_fmt)

    def _uniform_style(self, f: Callable) -> Any:
        return _CfStyleTarget(self)._uniform(f)

    def l_numberformat(self, ex, v):
        _CfStyleTarget(self)._style(lambda st: replace(st, num_fmt=_str(v)))

    def g_stopiftrue(self, ex):
        return self._cf().stop_if_true

    def l_stopiftrue(self, ex, v):
        on = _bool(v)
        self._change(lambda cf: setattr(cf, "stop_if_true", on))

    def g_priority(self, ex):
        return self.index + 1

    def g_appliesto(self, ex):
        return Range(self.rng.h, self.rng.ws, list(self._cf().ranges))

    def g_delete(self, ex):
        self.rng.h.require_writable()
        self.rng.h.command("home.deleteRule", {"index": self.index}, self.rng.si)
        return EMPTY

    def g_modify(self, ex, type_=MISSING, operator=MISSING, f1=MISSING, f2=MISSING):
        from ..cmd_cf import build_rule

        t = _long(type_, 1)
        if t == 1:
            op = _CF_OPERATORS.get(_long(operator, 1))
            if op is None:
                raise VbaError(5)
            spec: Dict[str, Any] = {"type": "cellIs", "operator": op, "value": _cf_value(f1), "style": {}}
            if op in ("between", "notBetween"):
                spec["value2"] = _cf_value(f2)
        elif t == 2:
            spec = {"type": "expression", "formula": V.text(_plain(f1)), "style": {}}
        else:
            raise NotSupported("FormatCondition.Modify with this type")
        try:
            new_rule = build_rule(spec)
        except EngineError as e:
            raise engine_error(e)

        def f(cf):
            new_rule.style = cf.rule.style
            cf.rule = new_rule

        self._change(f)
        return EMPTY

    def g_modifyappliestorange(self, ex, rng=MISSING):
        r = _as_range(rng)
        areas = list(r.areas)
        self._change(lambda cf: setattr(cf, "ranges", areas))
        return EMPTY


# ---------------------------------------------------------------- Sort and AutoFilter objects


class SortState(HostObject):
    """Worksheet.Sort as the macro recorder writes it: SortFields.Add …, SetRange, Header, Apply."""

    TYPE_NAME = "Sort"
    PARAMS = {"setrange": ("Rng",)}

    def __init__(self, ws: Worksheet):
        self.ws = ws
        self.fields: List[Dict[str, Any]] = []
        self.rng: Optional[Range] = None
        self.header = 0
        self.match_case = False
        self.orientation = 1

    def g_sortfields(self, ex):
        return SortFields(self)

    def g_setrange(self, ex, rng=MISSING):
        self.rng = rng if isinstance(rng, Range) else make_range(self.ws.h, self.ws, rng)
        return EMPTY

    def g_rng(self, ex):
        return self.rng if self.rng is not None else NOTHING

    def g_header(self, ex):
        return self.header

    def l_header(self, ex, v):
        self.header = _long(v)

    def g_matchcase(self, ex):
        return self.match_case

    def l_matchcase(self, ex, v):
        self.match_case = _bool(v)

    def g_orientation(self, ex):
        return self.orientation

    def l_orientation(self, ex, v):
        self.orientation = _long(v)

    def g_sortmethod(self, ex):
        return 1

    def l_sortmethod(self, ex, v):
        pass

    def g_apply(self, ex):
        h = self.ws.h
        h.require_writable()
        if self.rng is None:
            raise VbaError(1004, "The sort range isn't set (SetRange)")
        r = self.rng.first
        by_cols = self.orientation == 2
        keys = []
        for f in self.fields:
            k = f["key"].first
            col = (k.r0 - r.r0) if by_cols else (k.c0 - r.c0)
            entry: Dict[str, Any] = {"col": col, "order": "desc" if f["order"] == 2 else "asc"}
            if f["on"] == 1:
                entry["by"] = "cellColor"
                if f.get("color") is not None:
                    entry["color"] = f["color"]
            elif f["on"] == 2:
                entry["by"] = "fontColor"
                if f.get("color") is not None:
                    entry["color"] = f["color"]
            if f.get("custom"):
                entry["customList"] = f["custom"]
            keys.append(entry)
        if not keys:
            keys = [{"col": 0, "order": "asc"}]
        p: Dict[str, Any] = {"range": r.a1(), "keys": keys, "matchCase": self.match_case, "byColumns": by_cols}
        if self.header == 1:
            p["header"] = True
        elif self.header == 2:
            p["header"] = False
        h.command("data.sort", p, self.rng.si)
        return EMPTY


class SortFields(HostObject):
    TYPE_NAME = "SortFields"
    PARAMS = {"add": ("Key", "SortOn", "Order", "CustomOrder", "DataOption"), "add2": ("Key", "SortOn", "Order", "CustomOrder", "DataOption", "SubField")}

    def __init__(self, st: SortState):
        self.st = st

    def g_clear(self, ex):
        self.st.fields.clear()
        return EMPTY

    def g_count(self, ex):
        return len(self.st.fields)

    def g_add(self, ex, key=MISSING, sorton=MISSING, order=MISSING, custom=MISSING, dataoption=MISSING, subfield=MISSING):
        if _m(key):
            raise VbaError(449)
        k = key if isinstance(key, Range) else make_range(self.st.ws.h, self.st.ws, key)
        entry = {"key": k, "on": _long(sorton, 0), "order": _long(order, 1)}
        if not _m(custom):
            c = _plain(custom)
            entry["custom"] = [x.strip() for x in V.text(c).split(",")] if not isinstance(c, VArray) else [V.text(x) for x in c.data]
        self.st.fields.append(entry)
        return SortField(entry)

    g_add2 = g_add


class SortField(HostObject):
    TYPE_NAME = "SortField"

    def __init__(self, entry: Dict[str, Any]):
        self.entry = entry

    def g_sortoncolor(self, ex):
        return SortOnValue(self.entry)


class SortOnValue(HostObject):
    TYPE_NAME = "SortOnValue"

    def __init__(self, entry: Dict[str, Any]):
        self.entry = entry

    def g_color(self, ex):
        return 0.0

    def l_color(self, ex, v):
        self.entry["color"] = color_value(v)


class AutoFilter(HostObject):
    TYPE_NAME = "AutoFilter"

    def __init__(self, ws: Worksheet):
        self.ws = ws

    def g_range(self, ex):
        af = self.ws._sheet().autofilter
        return Range(self.ws.h, self.ws, [af.range]) if af is not None else NOTHING

    def g_filtermode(self, ex):
        return self.ws.g_filtermode(ex)

    def g_showalldata(self, ex):
        return self.ws.g_showalldata(ex)

    def g_filters(self, ex):
        af = self.ws._sheet().autofilter
        return EmptyCollection("Filters") if af is None else FiltersCollection(self.ws)

    def g_applyfilter(self, ex):
        self.ws.h.require_writable()
        self.ws.h.command("data.reapply", {}, self.ws.si)
        return EMPTY

    def g_sort(self, ex):
        return self.ws.h.sort_state(self.ws)


class FiltersCollection(HostObject):
    TYPE_NAME = "Filters"
    DEFAULT = "item"

    def __init__(self, ws: Worksheet):
        self.ws = ws

    def g_count(self, ex):
        af = self.ws._sheet().autofilter
        return af.range.width if af is not None else 0

    def g_item(self, ex, i=MISSING):
        return FilterItem(self.ws, _long(i) - 1)


class FilterItem(HostObject):
    TYPE_NAME = "Filter"

    def __init__(self, ws: Worksheet, i: int):
        self.ws = ws
        self.i = i

    def g_on(self, ex):
        af = self.ws._sheet().autofilter
        return af is not None and any(getattr(c, "col", None) == self.i for c in (af.criteria or []))


# ---------------------------------------------------------------- page setup, shapes


class PageSetup(HostObject):
    TYPE_NAME = "PageSetup"

    def __init__(self, ws: Worksheet):
        self.ws = ws

    def _p(self):
        return self.ws._sheet().print

    def g_orientation(self, ex):
        return 2 if self._p().landscape else 1

    def l_orientation(self, ex, v):
        self.ws.h.require_writable()
        self.ws.h.command("pageLayout.orientation", {"value": "landscape" if _long(v) == 2 else "portrait"}, self.ws.si)

    def g_printarea(self, ex):
        pa = self._p().print_area
        return pa.abs_a1() if pa is not None else ""

    def l_printarea(self, ex, v):
        self.ws.h.require_writable()
        t = _str(v)
        if not t:
            self.ws.h.command("pageLayout.clearPrintArea", {}, self.ws.si)
        else:
            self.ws.h.command("pageLayout.printArea", {"range": t.replace("$", "").split("!")[-1]}, self.ws.si)

    def g_printtitlerows(self, ex):
        tr = self._p().title_rows
        return f"${tr[0] + 1}:${tr[1] + 1}" if tr else ""

    def l_printtitlerows(self, ex, v):
        self.ws.h.require_writable()
        self.ws.h.command("pageLayout.printTitles", {"rows": _str(v).replace("$", "")}, self.ws.si)

    def g_printtitlecolumns(self, ex):
        tc = self._p().title_cols
        return f"${col_to_letters(tc[0])}:${col_to_letters(tc[1])}" if tc else ""

    def l_printtitlecolumns(self, ex, v):
        self.ws.h.require_writable()
        self.ws.h.command("pageLayout.printTitles", {"cols": _str(v).replace("$", "")}, self.ws.si)

    def g_zoom(self, ex):
        s = self._p().scale
        return s if s else False

    def l_zoom(self, ex, v):
        pv = _plain(v)
        if type(pv) is bool:
            return
        self.ws.h.require_writable()
        self.ws.h.command("pageLayout.scale", {"scale": V.to_long(pv)}, self.ws.si)

    def g_fittopageswide(self, ex):
        f = self._p().fit_to
        return f[0] if f else 1

    def l_fittopageswide(self, ex, v):
        self.ws.h.require_writable()
        pv = _plain(v)
        self.ws.h.command("pageLayout.scale", {"fitWidth": 0 if type(pv) is bool else V.to_long(pv)}, self.ws.si)

    def g_fittopagestall(self, ex):
        f = self._p().fit_to
        return f[1] if f else 1

    def l_fittopagestall(self, ex, v):
        self.ws.h.require_writable()
        pv = _plain(v)
        self.ws.h.command("pageLayout.scale", {"fitHeight": 0 if type(pv) is bool else V.to_long(pv)}, self.ws.si)

    def g_printgridlines(self, ex):
        return bool(self._p().gridlines)

    def l_printgridlines(self, ex, v):
        self.ws.h.require_writable()
        self.ws.h.command("pageLayout.printGridlines", {"on": _bool(v)}, self.ws.si)

    def g_printheadings(self, ex):
        return bool(self._p().headings)

    def l_printheadings(self, ex, v):
        self.ws.h.require_writable()
        self.ws.h.command("pageLayout.printHeadings", {"on": _bool(v)}, self.ws.si)

    def _margin(self, key: str) -> float:
        return float(getattr(self._p().margins, key, 0.75)) * 72

    def _set_margin(self, key: str, v) -> None:
        self.ws.h.require_writable()
        self.ws.h.command("pageLayout.margins", {key: V.to_double(_plain(v)) / 72.0}, self.ws.si)

    def g_leftmargin(self, ex):
        return self._margin("left")

    def l_leftmargin(self, ex, v):
        self._set_margin("left", v)

    def g_rightmargin(self, ex):
        return self._margin("right")

    def l_rightmargin(self, ex, v):
        self._set_margin("right", v)

    def g_topmargin(self, ex):
        return self._margin("top")

    def l_topmargin(self, ex, v):
        self._set_margin("top", v)

    def g_bottommargin(self, ex):
        return self._margin("bottom")

    def l_bottommargin(self, ex, v):
        self._set_margin("bottom", v)

    def g_centerheader(self, ex):
        return ""

    def l_centerheader(self, ex, v):
        self.ws.h.require_writable()
        self.ws.h.command("pageLayout.headerFooter", {"header": "&C" + _str(v)}, self.ws.si)

    def g_centerfooter(self, ex):
        return ""

    def l_centerfooter(self, ex, v):
        self.ws.h.require_writable()
        self.ws.h.command("pageLayout.headerFooter", {"footer": "&C" + _str(v)}, self.ws.si)

    def g_papersize(self, ex):
        return 1

    def l_papersize(self, ex, v):
        names = {1: "Letter", 5: "Legal", 9: "A4", 8: "A3", 11: "A5", 7: "Executive", 3: "Tabloid"}
        p = names.get(_long(v))
        if p:
            self.ws.h.require_writable()
            self.ws.h.command("pageLayout.size", {"paper": p}, self.ws.si)

    def vba_get_missing(self, ex, name, args, named):
        return EMPTY

    def vba_let(self, ex, name, args, named, value):
        f = getattr(self, "l_" + name.lower(), None)
        if f is None:
            return None  # other page settings are accepted and ignored
        return super().vba_let(ex, name, args, named, value)


class Shapes(HostObject):
    TYPE_NAME = "Shapes"
    DEFAULT = "item"
    PARAMS = {"addchart2": ("Style", "XlChartType", "Left", "Top", "Width", "Height", "NewLayout"), "addchart": ("XlChartType", "Left", "Top", "Width", "Height")}

    def __init__(self, ws: Worksheet):
        self.ws = ws

    def g_count(self, ex):
        sh = self.ws._sheet()
        return len(sh.charts) + len(sh.images)

    def g_item(self, ex, index=MISSING):
        raise NotSupported("Shapes")

    def vba_enum(self, ex):
        return iter([])

    def g_addchart2(self, ex, style=MISSING, charttype=MISSING, *rest):
        return ChartObjects(self.ws).g_add(ex)._chart_type(charttype)

    def g_addchart(self, ex, charttype=MISSING, *rest):
        return ChartObjects(self.ws).g_add(ex)._chart_type(charttype)

    def g_addshape(self, ex, *args):
        raise NotSupported("Drawing shapes")

    g_addpicture = g_addshape
    g_addtextbox = g_addshape
    g_addformcontrol = g_addshape


_CHART_KINDS = {
    51: "columnClustered", 52: "columnStacked", 53: "columnStacked100", 57: "barClustered", 58: "barStacked",
    59: "barStacked100", 4: "line", 65: "lineMarkers", 63: "lineStacked", 5: "pie", -4120: "doughnut",
    1: "area", 76: "areaStacked", -4169: "scatter", 74: "scatterLines", -4151: "radar",
}


class ChartObjects(HostObject):
    TYPE_NAME = "ChartObjects"
    DEFAULT = "item"
    PARAMS = {"add": ("Left", "Top", "Width", "Height")}

    def __init__(self, ws: Worksheet):
        self.ws = ws

    def g_count(self, ex):
        return len(self.ws._sheet().charts)

    def g_item(self, ex, index=MISSING):
        charts = self.ws._sheet().charts
        i = _long(index)
        if not 1 <= i <= len(charts):
            raise VbaError(9)
        return ChartObject(self.ws, charts[i - 1].id)

    def vba_enum(self, ex):
        return iter([ChartObject(self.ws, c.id) for c in self.ws._sheet().charts])

    def g_add(self, ex, left=MISSING, top=MISSING, width=MISSING, height=MISSING):
        h = self.ws.h
        h.require_writable()
        sel = h.selection()
        before = {c.id for c in self.ws._sheet().charts}
        h.command("insert.chart", {"kind": "columnClustered", "range": sel.current().a1()}, self.ws.si)
        new = [c.id for c in self.ws._sheet().charts if c.id not in before]
        if not new:
            raise VbaError(1004, "The chart couldn't be added")
        return ChartObject(self.ws, new[0])

    def g_delete(self, ex):
        h = self.ws.h
        h.require_writable()
        for c in list(self.ws._sheet().charts):
            h.command("object.delete", {"id": c.id}, self.ws.si)
        return EMPTY


class ChartObject(HostObject):
    TYPE_NAME = "ChartObject"

    def __init__(self, ws: Worksheet, cid: int):
        self.ws = ws
        self.cid = cid

    def _chart_type(self, kind) -> "ChartObject":
        if not _m(kind):
            Chart(self).l_charttype(None, kind)
        return self

    def g_chart(self, ex):
        return Chart(self)

    def g_delete(self, ex):
        self.ws.h.require_writable()
        self.ws.h.command("object.delete", {"id": self.cid}, self.ws.si)
        return EMPTY

    def g_name(self, ex):
        return f"Chart {self.cid}"


class Chart(HostObject):
    TYPE_NAME = "Chart"
    PARAMS = {"setsourcedata": ("Source", "PlotBy")}

    def __init__(self, co: ChartObject):
        self.co = co

    def _upd(self, **kw) -> None:
        self.co.ws.h.require_writable()
        p = {"id": self.co.cid}
        p.update(kw)
        self.co.ws.h.command("chart.update", p, self.co.ws.si)

    def g_setsourcedata(self, ex, source=MISSING, plotby=MISSING):
        from ...formula.printer import quote_sheet

        self.co.ws.h.require_writable()
        r = source if isinstance(source, Range) else make_range(self.co.ws.h, self.co.ws, source)
        p: Dict[str, Any] = {"id": self.co.cid, "range": f"{quote_sheet(r.ws.name)}!{r.first.a1()}"}
        if not _m(plotby):
            p["byRows"] = _long(plotby) == 1
        self.co.ws.h.command("chart.selectData", p, self.co.ws.si)
        return EMPTY

    def g_charttype(self, ex):
        for c in self.co.ws._sheet().charts:
            if c.id == self.co.cid:
                inv = {v: k for k, v in _CHART_KINDS.items()}
                return inv.get(c.kind, 51)
        return 51

    def l_charttype(self, ex, v):
        kind = _CHART_KINDS.get(_long(v))
        if kind is None:
            raise NotSupported("This chart type")
        self._upd(kind=kind)

    def g_hastitle(self, ex):
        for c in self.co.ws._sheet().charts:
            if c.id == self.co.cid:
                return bool(c.title)
        return False

    def l_hastitle(self, ex, v):
        if not _bool(v):
            self._upd(title=None)

    def g_charttitle(self, ex):
        return ChartTitle(self)

    def g_haslegend(self, ex):
        return True

    def l_haslegend(self, ex, v):
        self._upd(legend="right" if _bool(v) else "none")


class ChartTitle(HostObject):
    TYPE_NAME = "ChartTitle"
    DEFAULT = "text"

    def __init__(self, ch: Chart):
        self.ch = ch

    def g_text(self, ex):
        for c in self.ch.co.ws._sheet().charts:
            if c.id == self.ch.co.cid:
                return c.title or ""
        return ""

    def l_text(self, ex, v):
        self.ch._upd(title=_str(v))

    g_caption = g_text
    l_caption = l_text


def _merge_areas(areas: List[RangeRef]) -> List[RangeRef]:
    """Union's areas as Excel gives them: blocks that together form a rectangle (overlapping or
    side by side) become one (``Union(A1, A2)`` is ``$A$1:$A$2``)."""
    out = list(areas)
    merged = True
    while merged and len(out) > 1:
        merged = False
        for i in range(len(out)):
            for j in range(i + 1, len(out)):
                a, b = out[i], out[j]
                m = None
                if a.contains_range(b):
                    m = a
                elif b.contains_range(a):
                    m = b
                elif a.c0 == b.c0 and a.c1 == b.c1 and a.r0 <= b.r1 + 1 and b.r0 <= a.r1 + 1:
                    m = RangeRef(min(a.r0, b.r0), a.c0, max(a.r1, b.r1), a.c1)
                elif a.r0 == b.r0 and a.r1 == b.r1 and a.c0 <= b.c1 + 1 and b.c0 <= a.c1 + 1:
                    m = RangeRef(a.r0, min(a.c0, b.c0), a.r1, max(a.c1, b.c1))
                if m is not None:
                    out[i] = m
                    del out[j]
                    merged = True
                    break
            if merged:
                break
    return out
