"""Excel's Range object (and Font, Interior, Borders, Comment… that hang off it).

A Range is a sheet and one or more areas (``RangeRef``); ``mode`` says whether it was reached as
``Rows`` or ``Columns`` (which changes Count and For Each, as in Excel). Reading goes through the
host; changing values and formats writes into the host's open edit; everything else runs the
engine's commands with explicit ranges, so a macro does exactly what the ribbon would.
"""

from __future__ import annotations

import math
from dataclasses import replace
from typing import Any, Callable, Dict, Iterator, List, Optional, Sequence, Tuple

from ...core.addr import MAX_COLS, MAX_ROWS, RangeRef, a1, col_to_letters, letters_to_col
from ...core.value import Array, CellError
from ...model.style import AUTO, BorderLine, Color, Fill
from ...vba import values as V
from ...vba.errors import NotSupported, VbaError
from ...vba.runtime import Execution, HostObject
from ...vba.values import EMPTY, MISSING, NOTHING, NULL, VArray, is_object
from .host import MAX_CELLS_PER_OP, XlHost, a1_formula_to_r1c1, cell_to_vba

MAX_ARRAY_CELLS = 5_000_000


def _m(v: Any) -> bool:
    return v is MISSING


def _plain(v: Any) -> Any:
    if is_object(v) and v is not NOTHING:
        return V.value_of(v)
    return v


def _long(v: Any, default: Optional[int] = None) -> int:
    if v is MISSING:
        if default is None:
            raise VbaError(449)
        return default
    return V.to_long(_plain(v))


def _bool(v: Any, default: bool = False) -> bool:
    if v is MISSING:
        return default
    return V.to_boolean(_plain(v))


def _str(v: Any, default: str = "") -> str:
    if v is MISSING:
        return default
    return V.to_string(_plain(v))


# ---------------------------------------------------------------- colours


def bgr_to_hex(v: int) -> str:
    v = int(v) & 0xFFFFFF
    return f"#{v & 0xFF:02X}{(v >> 8) & 0xFF:02X}{(v >> 16) & 0xFF:02X}"


def rgb_to_bgr(rgb: Tuple[int, int, int]) -> int:
    return rgb[0] | (rgb[1] << 8) | (rgb[2] << 16)


def palette_hex(index: int) -> str:
    from ...numfmt import PALETTE56

    if not 1 <= index <= 56:
        raise VbaError(1004, "Unable to set the ColorIndex property (it takes 1 to 56)")
    c = PALETTE56[index - 1]
    return f"#{c:06X}"


def color_index_of(rgb: Optional[Tuple[int, int, int]]) -> int:
    from ...numfmt import PALETTE56

    if rgb is None:
        return -4142
    want = (rgb[0] << 16) | (rgb[1] << 8) | rgb[2]
    best, dist = 1, None
    for i, c in enumerate(PALETTE56):
        d = sum((((c >> s) & 255) - ((want >> s) & 255)) ** 2 for s in (16, 8, 0))
        if dist is None or d < dist:
            best, dist = i + 1, d
    return best


def color_value(c: Any) -> Any:
    """A VBA colour (BGR Long) as the engine's colour text."""
    n = V.to_long(_plain(c))
    if n < 0:
        raise VbaError(1004, "Unable to set the Color property")
    return bgr_to_hex(n)


# ---------------------------------------------------------------- Range


class Range(HostObject):
    TYPE_NAME = "Range"
    DEFAULT = "_default"
    PARAMS = {
        "_default": ("RowIndex", "ColumnIndex"),
        "item": ("RowIndex", "ColumnIndex"),
        "cells": ("RowIndex", "ColumnIndex"),
        "offset": ("RowOffset", "ColumnOffset"),
        "resize": ("RowSize", "ColumnSize"),
        "address": ("RowAbsolute", "ColumnAbsolute", "ReferenceStyle", "External", "RelativeTo"),
        "addresslocal": ("RowAbsolute", "ColumnAbsolute", "ReferenceStyle", "External", "RelativeTo"),
        "find": ("What", "After", "LookIn", "LookAt", "SearchOrder", "SearchDirection", "MatchCase", "MatchByte", "SearchFormat"),
        "findnext": ("After",),
        "findprevious": ("After",),
        "replace": ("What", "Replacement", "LookAt", "SearchOrder", "MatchCase", "MatchByte", "SearchFormat", "ReplaceFormat"),
        "sort": ("Key1", "Order1", "Key2", "Type", "Order2", "Key3", "Order3", "Header", "OrderCustom", "MatchCase", "Orientation", "SortMethod", "DataOption1", "DataOption2", "DataOption3"),
        "autofilter": ("Field", "Criteria1", "Operator", "Criteria2", "VisibleDropDown"),
        "removeduplicates": ("Columns", "Header"),
        "specialcells": ("Type", "Value"),
        "copy": ("Destination",),
        "cut": ("Destination",),
        "pastespecial": ("Paste", "Operation", "SkipBlanks", "Transpose"),
        "insert": ("Shift", "CopyOrigin"),
        "delete": ("Shift",),
        "autofill": ("Destination", "Type"),
        "borderaround": ("LineStyle", "Weight", "ColorIndex", "Color", "ThemeColor"),
        "addcomment": ("Text",),
        "notetext": ("Text", "Start", "Length"),
        "range": ("Cell1", "Cell2"),
        "value": ("RangeValueDataType",),
        "merge": ("Across",),
        "select": ("Replace",),
        "texttocolumns": ("Destination", "DataType", "TextQualifier", "ConsecutiveDelimiter", "Tab", "Semicolon", "Comma", "Space", "Other", "OtherChar", "FieldInfo", "DecimalSeparator", "ThousandsSeparator", "TrailingMinusNumbers"),
        "goalseek": ("Goal", "ChangingCell"),
        "characters": ("Start", "Length"),
        "end": ("Direction",),
    }

    def __init__(self, h: XlHost, ws, areas: List[RangeRef], mode: str = "cells"):
        if not areas:
            raise VbaError(1004)
        self.h = h
        self.ws = ws
        self.areas = areas
        self.mode = mode

    def __repr__(self) -> str:
        return f"<Range {self.ws.name}!{self.address()}>"

    # ------------------------------------------------------------ helpers

    @property
    def si(self) -> int:
        return self.h.index_of(self.ws.name)

    @property
    def first(self) -> RangeRef:
        return self.areas[0]

    def _new(self, areas: List[RangeRef], mode: str = "cells") -> "Range":
        return Range(self.h, self.ws, areas, mode)

    def address(self, row_abs: bool = True, col_abs: bool = True) -> str:
        return ",".join(_addr(r, row_abs, col_abs) for r in self.areas)

    def cell_list(self) -> List[Tuple[int, int]]:
        out: List[Tuple[int, int]] = []
        for r in self.areas:
            if r.count() > MAX_CELLS_PER_OP:
                raise VbaError(1004, "The range is too large for this")
            out.extend(r.cells())
        return out

    def _sheet(self):
        return self.h.wb().sheets[self.si]

    def _trim(self, r: RangeRef) -> Optional[RangeRef]:
        """Whole rows or columns cut to the sheet's used range (what reading them needs)."""
        if r.count() <= 1_000_000:
            return r
        used = self._sheet().used_range()
        if used is None:
            return None
        return r.intersection(used)

    # ------------------------------------------------------------ navigation

    def g_application(self, ex):
        return self.h.app

    def g_creator(self, ex):
        return 1480803660

    def g_parent(self, ex):
        return self.ws

    def g_worksheet(self, ex):
        return self.ws

    def g_cells(self, ex, row=MISSING, col=MISSING):
        if _m(row) and _m(col):
            return self._new(list(self.areas))
        return self.g_item(ex, row, col)

    def g_item(self, ex, row=MISSING, col=MISSING):
        f = self.first
        if _m(col):
            if _m(row):
                raise VbaError(449)
            rv = _plain(row)
            if isinstance(rv, str):
                return self._new([_parse_simple(rv, f)])
            n = V.to_long(rv)
            if self.mode == "rows":
                r = f.r0 + n - 1
                if not 0 <= r < MAX_ROWS:
                    raise VbaError(1004)
                return self._new([RangeRef(r, f.c0, r, f.c1)])
            if self.mode == "cols":
                c = f.c0 + n - 1
                if not 0 <= c < MAX_COLS:
                    raise VbaError(1004)
                return self._new([RangeRef(f.r0, c, f.r1, c)])
            w = f.width
            r = f.r0 + (n - 1) // w
            c = f.c0 + (n - 1) % w
        else:
            n = 1 if _m(row) else V.to_long(_plain(row))
            cv = _plain(col)
            if isinstance(cv, str):
                ci = letters_to_col(cv.strip().upper())
                if ci is None:
                    raise VbaError(13)
                c = f.c0 + ci
            else:
                c = f.c0 + V.to_long(cv) - 1
            r = f.r0 + n - 1
        if not (0 <= r < MAX_ROWS and 0 <= c < MAX_COLS):
            raise VbaError(1004, "Application-defined or object-defined error")
        return self._new([RangeRef(r, c, r, c)])

    def g__default(self, ex, row=MISSING, col=MISSING):
        if _m(row) and _m(col):
            return self.g_value(ex)
        return self.g_item(ex, row, col)

    def l__default(self, ex, value, row=MISSING, col=MISSING):
        if _m(row) and _m(col):
            return self.l_value(ex, value)
        return self.g_item(ex, row, col).l_value(ex, value)

    def g_range(self, ex, cell1=MISSING, cell2=MISSING):
        """``rng.Range("A1")``: relative to this range's top-left."""
        f = self.first
        sub = make_range(self.h, self.ws, cell1, cell2)
        out = [_shift(r, f.r0, f.c0) for r in sub.areas]
        for r in out:
            if r.r1 >= MAX_ROWS or r.c1 >= MAX_COLS:
                raise VbaError(1004)
        return self._new(out)

    def g_offset(self, ex, rows=MISSING, cols=MISSING):
        dr, dc = _long(rows, 0), _long(cols, 0)
        out = []
        for r in self.areas:
            if not (0 <= r.r0 + dr and r.r1 + dr < MAX_ROWS and 0 <= r.c0 + dc and r.c1 + dc < MAX_COLS):
                raise VbaError(1004, "Application-defined or object-defined error")
            out.append(_shift(r, dr, dc))
        return self._new(out, self.mode)

    def g_resize(self, ex, rows=MISSING, cols=MISSING):
        f = self.first
        h = f.height if _m(rows) else _long(rows)
        w = f.width if _m(cols) else _long(cols)
        if h < 1 or w < 1 or f.r0 + h > MAX_ROWS or f.c0 + w > MAX_COLS:
            raise VbaError(1004, "Application-defined or object-defined error")
        return self._new([RangeRef(f.r0, f.c0, f.r0 + h - 1, f.c0 + w - 1)])

    def g_rows(self, ex, index=MISSING):
        r = self._new(list(self.areas), "rows")
        return r if _m(index) else r.g_item(ex, index)

    def g_columns(self, ex, index=MISSING):
        r = self._new(list(self.areas), "cols")
        return r if _m(index) else r.g_item(ex, index)

    def g_entirerow(self, ex):
        return self._new([RangeRef(r.r0, 0, r.r1, MAX_COLS - 1) for r in self.areas], "rows")

    def g_entirecolumn(self, ex):
        return self._new([RangeRef(0, r.c0, MAX_ROWS - 1, r.c1) for r in self.areas], "cols")

    def g_row(self, ex):
        return self.first.r0 + 1

    def g_column(self, ex):
        return self.first.c0 + 1

    def g_count(self, ex):
        # Rows/Columns of several areas count the first area's, as Excel does.
        if self.mode == "rows":
            return self.areas[0].height
        if self.mode == "cols":
            return self.areas[0].width
        n = sum(r.count() for r in self.areas)
        if n > 2147483647:
            raise VbaError(6)
        return n

    def g_countlarge(self, ex):
        return V.LongLong(sum(r.count() for r in self.areas)) if self.mode == "cells" else V.LongLong(self.g_count(ex))

    def g_areas(self, ex, index=MISSING):
        coll = Areas(self)
        return coll if _m(index) else coll.g_item(ex, index)

    def g_address(self, ex, rowabs=MISSING, colabs=MISSING, style=MISSING, external=MISSING, relto=MISSING):
        ra, ca = _bool(rowabs, True), _bool(colabs, True)
        st = _long(style, 1)
        if st == -4150:
            base = (0, 0)
            if not _m(relto):
                rel = relto if isinstance(relto, Range) else None
                if rel is not None:
                    base = (rel.first.r0, rel.first.c0)
            text = ",".join(_r1c1_addr(r, ra, ca, base) for r in self.areas)
        else:
            text = self.address(ra, ca)
        if _bool(external, False):
            from ...formula.printer import quote_sheet

            text = f"[{self.h.workbook.name()}]{quote_sheet(self.ws.name)}!{text}"
        return text

    g_addresslocal = g_address

    def g_end(self, ex, direction=MISSING):
        from ..selection import jump

        d = _long(direction)
        dr, dc = {-4162: (-1, 0), -4121: (1, 0), -4159: (0, -1), -4161: (0, 1)}.get(d, (None, None))
        if dr is None:
            raise VbaError(5)
        self._flush_for_read()
        sh = self._sheet()
        r, c = jump(sh, (self.first.r0, self.first.c0), dr, dc)
        return self._new([RangeRef(r, c, r, c)])

    def g_currentregion(self, ex):
        from ..selection import current_region

        self._flush_for_read()
        return self._new([current_region(self._sheet(), (self.first.r0, self.first.c0))])

    def _flush_for_read(self) -> None:
        flush = getattr(self.h, "flush", None)
        if flush is not None and getattr(self.h, "pending", None) is not None and self.h.pending.changed:
            flush()

    def g_mergearea(self, ex):
        sh = self._sheet()
        m = sh.merge_at((self.first.r0, self.first.c0))
        return self._new([m]) if m is not None else self._new([RangeRef.of((self.first.r0, self.first.c0))])

    def g_next(self, ex):
        f = self.first
        return self._new([RangeRef.of((f.r0, min(f.c0 + 1, MAX_COLS - 1)))])

    def g_previous(self, ex):
        f = self.first
        return self._new([RangeRef.of((f.r0, max(f.c0 - 1, 0)))])

    def g_listobject(self, ex):
        from .excel import ListObject

        sh = self._sheet()
        for t in sh.tables:
            if t.range.contains((self.first.r0, self.first.c0)):
                return ListObject(self.h, self.ws, t.name)
        return NOTHING

    def vba_enum(self, ex: Execution) -> Iterator[Any]:
        def gen():
            for r in list(self.areas):
                if self.mode == "rows":
                    for i in range(r.r0, r.r1 + 1):
                        yield self._new([RangeRef(i, r.c0, i, r.c1)])
                elif self.mode == "cols":
                    for j in range(r.c0, r.c1 + 1):
                        yield self._new([RangeRef(r.r0, j, r.r1, j)])
                else:
                    t = self._trim(r) if r.count() > 1_000_000 else r
                    if t is None:
                        continue
                    for rc in t.cells():
                        yield self._new([RangeRef.of(rc)])

        return gen()

    def vba_same(self, other: Any) -> bool:
        return False

    # ------------------------------------------------------------ values

    def _values(self, ex, kind: str) -> Any:
        f = self.first
        si = self.si
        h = self.h
        if f.is_single():
            rc = (f.r0, f.c0)
            if kind in ("value", "value2"):
                return h.cell_vba_value(si, rc, kind == "value2")
            if kind == "text":
                return self._text(si, rc)
            return self._formula(si, rc, kind)
        if f.count() > MAX_ARRAY_CELLS:
            raise VbaError(7)
        if kind == "text":
            # Current Excel gives each cell's text (older versions gave Null when they differed).
            return VArray.from_rows([[self._text(si, (r, c)) for c in range(f.c0, f.c1 + 1)] for r in range(f.r0, f.r1 + 1)], 1)
        rows: List[List[Any]] = []
        if kind in ("value", "value2"):
            block = h.read_block(si, f)
            sh = h.wb().sheets[si]
            wb = h.wb()
            d1904 = h.d1904()
            value2 = kind == "value2"
            for i, row in enumerate(block):
                out = []
                for j, v in enumerate(row):
                    if v is None:
                        out.append(EMPTY)
                        continue
                    code = "" if value2 or type(v) is not float else wb.styles.get(sh.style_id((f.r0 + i, f.c0 + j))).num_fmt
                    out.append(cell_to_vba(v, code, d1904, value2))
                rows.append(out)
        else:
            for r in range(f.r0, f.r1 + 1):
                rows.append([self._formula(si, (r, c), kind) for c in range(f.c0, f.c1 + 1)])
        return VArray.from_rows(rows, 1)

    def _text(self, si: int, rc: Tuple[int, int]) -> str:
        from ..display import value_text

        self.h.read_value(si, rc)
        wb = self.h.wb()
        return value_text(wb, wb.sheets[si], rc)

    def _formula(self, si: int, rc: Tuple[int, int], kind: str) -> Any:
        wb = self.h.wb()
        sh = wb.sheets[si]
        cell = sh.cell(rc)
        if cell is None or cell.formula is None:
            arr = next((a for a in _array_ranges(sh) if a.contains(rc)), None)
            if arr is not None and arr.start != rc:
                # Every cell of a legacy array formula shows the formula.
                cell, rc = sh.cell(arr.start), arr.start
        if cell is not None and cell.formula is not None:
            if kind in ("formular1c1",):
                return "=" + a1_formula_to_r1c1(cell.formula.text, rc)
            return "=" + cell.formula.text
        if cell is None or cell.value is None:
            return ""
        return constant_text(cell.value)

    def g_value(self, ex, datatype=MISSING):
        return self._values(ex, "value")

    def g_value2(self, ex):
        return self._values(ex, "value2")

    def g_text(self, ex):
        return self._values(ex, "text")

    def g_formula(self, ex):
        return self._values(ex, "formula")

    g_formula2 = g_formula
    g_formulalocal = g_formula

    def g_formular1c1(self, ex):
        return self._values(ex, "formular1c1")

    g_formula2r1c1 = g_formular1c1
    g_formular1c1local = g_formular1c1

    def g_hasformula(self, ex):
        sh = self._sheet()
        states = set()
        arrays = _array_ranges(sh)
        for r in self.areas:
            for rc in (r.cells() if r.count() <= 100_000 else []):
                c = sh.cell(rc)
                states.add((c is not None and c.formula is not None) or any(a.contains(rc) for a in arrays))
                if len(states) > 1:
                    return NULL
        return states.pop() if states else False

    def g_hasarray(self, ex):
        sh = self._sheet()
        rc = (self.first.r0, self.first.c0)
        c = sh.cell(rc)
        if c is not None and c.formula is not None and c.formula.array is not None:
            return True
        return any(a.contains(rc) for a in _array_ranges(sh))

    def g_hasspill(self, ex):
        sh = self._sheet()
        return (self.first.r0, self.first.c0) in sh.spill_ranges or (self.first.r0, self.first.c0) in sh.spill

    def g_spillingtorange(self, ex):
        r = self._sheet().spill_ranges.get((self.first.r0, self.first.c0))
        return self._new([r]) if r is not None else NOTHING

    def g_currentarray(self, ex):
        from ..cmd_edit import array_at

        arr = array_at(self._sheet(), (self.first.r0, self.first.c0))
        if arr is None:
            raise VbaError(1004, "No cells were found.")
        return self._new([arr])

    def g_prefixcharacter(self, ex):
        """``'`` for text kept as text with an apostrophe (it would read as a number or date)."""
        from ..display import edit_text

        wb = self.h.wb()
        rc = (self.first.r0, self.first.c0)
        return "'" if edit_text(wb, wb.sheets[self.si], rc).startswith("'") else ""

    def g_formulaarray(self, ex):
        sh = self._sheet()
        from ..cmd_edit import array_at

        arr = array_at(sh, (self.first.r0, self.first.c0))
        if arr is None:
            return self.g_formula(ex)
        c = sh.cell(arr.start)
        return "=" + c.formula.text if c is not None and c.formula is not None else ""

    def l_value(self, ex, value, datatype=MISSING):
        self._write(value, "value")

    def l_value2(self, ex, value):
        self._write(value, "value")

    def l_formula(self, ex, value):
        self._write(value, "formula")

    l_formula2 = l_formula
    l_formulalocal = l_formula

    def l_formular1c1(self, ex, value):
        self._write(value, "formulaR1C1")

    l_formula2r1c1 = l_formular1c1
    l_formular1c1local = l_formular1c1

    def l_formulaarray(self, ex, value):
        """A legacy (Ctrl+Shift+Enter) array formula over the whole range."""
        self.h.require_writable()
        text = _str(value)
        if not text.startswith("="):
            text = "=" + text
        si = self.si
        top = a1(self.first.r0, self.first.c0)
        if not self.first.is_single():
            # cell.set takes the array's extent from the selection.
            saved = self.h.selection().copy()
            self.h.command("selection.set", {"range": self.first.a1(), "active": top}, si)
            try:
                self.h.command("cell.set", {"cell": top, "input": text, "array": True}, si)
            finally:
                if si == self.h.wb().active_sheet:
                    self.h.command("selection.set", {"range": saved.a1(), "active": a1(*saved.active)})
            return
        self.h.command("cell.set", {"cell": top, "input": text, "array": True}, si)

    def _write(self, value: Any, kind: str) -> None:
        self.h.require_writable()
        si = self.si
        if is_object(value) and isinstance(value, Range):
            # A one-cell range gives its value; a bigger one empties the cells, as in Excel.
            f = value.first
            value = value._values(None, "value") if f.is_single() and len(value.areas) == 1 else EMPTY
        elif is_object(value) and value is not NOTHING:
            value = V.value_of(value)
        cells: List[Tuple[Tuple[int, int], Any]] = []
        for rng in self.areas:
            if rng.count() > MAX_CELLS_PER_OP:
                raise VbaError(1004, "The range is too large to fill")
            if type(value) is VArray:
                rows = _as_rows(value)
                ah, aw = len(rows), (len(rows[0]) if rows else 0)
                for i, r in enumerate(range(rng.r0, rng.r1 + 1)):
                    ai = i if ah > 1 else 0
                    for j, c in enumerate(range(rng.c0, rng.c1 + 1)):
                        aj = j if aw > 1 else 0
                        if ai < ah and aj < aw:
                            v = rows[ai][aj]
                        else:
                            v = V.ErrorValue(2042)
                        cells.append(((r, c), v))
            elif kind in ("formula", "formulaR1C1") and isinstance(value, str) and value.startswith("=") and not rng.is_single():
                # One formula for many cells: each cell gets it relative to itself.
                if kind == "formulaR1C1":
                    for rc in rng.cells():
                        cells.append((rc, value))
                else:
                    cells.extend(_relative_formulas(value, rng))
            else:
                for rc in rng.cells():
                    cells.append((rc, value))
        if cells and len(cells) > 1:
            # A merged area takes the value in its top-left cell only.
            sh = self._sheet()
            skip = set()
            for m in sh.merges:
                if any(a.intersects(m) and a.contains(m.start) for a in self.areas):
                    skip.update(rc for rc in m.cells() if rc != m.start)
            if skip:
                cells = [x for x in cells if x[0] not in skip]
        if cells:
            self.h.write_values(si, cells, kind if kind == "formulaR1C1" else "value")
        self._fire_change()

    def _fire_change(self, target: Optional["Range"] = None) -> None:
        """Worksheet_Change for what a macro changed (as Excel runs it for macros' edits too)."""
        changed = getattr(self.h, "changed", None)
        if changed is not None:
            from ...vba.runtime import current

            changed(current(), target if target is not None else self)

    def _fire_selection(self, before) -> None:
        """Worksheet_SelectionChange when a macro moved the selection."""
        fire = getattr(self.h, "selection_changed", None)
        if fire is not None:
            from ...vba.runtime import current

            fire(current(), self.si, before)

    # ------------------------------------------------------------ formats

    def _style(self, f: Callable) -> None:
        self.h.require_writable()
        self.h.style(self.si, list(self.areas), f)

    def _style_of(self):
        wb = self.h.wb()
        sh = wb.sheets[self.si]
        return wb.styles.get(sh.style_id((self.first.r0, self.first.c0)))

    def _uniform(self, get: Callable) -> Any:
        """A format property: its value when every cell agrees, else Null."""
        wb = self.h.wb()
        sh = wb.sheets[self.si]
        seen = None
        n = 0
        for r in self.areas:
            t = self._trim(r) if r.count() > 10_000 else r
            if t is None:
                continue
            for rc in t.cells():
                v = get(wb.styles.get(sh.style_id(rc)))
                if seen is None:
                    seen = (v,)
                elif seen[0] != v:
                    return NULL
                n += 1
                if n > 10_000:
                    return seen[0]
        return seen[0] if seen is not None else get(wb.styles.get(0))

    def g_numberformat(self, ex):
        return self._uniform(lambda st: st.num_fmt)

    g_numberformatlocal = g_numberformat

    def l_numberformat(self, ex, v):
        code = _str(v)
        self._style(lambda st: replace(st, num_fmt=code or "General"))

    l_numberformatlocal = l_numberformat

    def g_font(self, ex):
        return Font(self)

    def g_interior(self, ex):
        return Interior(self)

    def g_borders(self, ex, index=MISSING):
        b = Borders(self)
        return b if _m(index) else b.g_item(ex, index)

    def g_borderaround(self, ex, linestyle=MISSING, weight=MISSING, colorindex=MISSING, color=MISSING, themecolor=MISSING):
        style = _line_style(_long(linestyle, 1) if not _m(linestyle) else 1, _long(weight, 2) if not _m(weight) else 2)
        col = None
        if not _m(color):
            col = color_value(color)
        elif not _m(colorindex) and _long(colorindex) > 0:
            col = palette_hex(_long(colorindex))
        for r in self.areas:
            self.h.require_writable()
            self.h.command("home.borders", {"preset": "outside", "style": style or "thin", "color": col or "#000000", "range": r.a1()}, self.si)
        return True

    def _align_get(self, attr: str, table: Dict[str, int]) -> Any:
        v = self._uniform(lambda st: getattr(st.align, attr))
        if v is NULL:
            return NULL
        return table.get(v, 1)

    def g_horizontalalignment(self, ex):
        return self._align_get("h", {"general": 1, "left": -4131, "center": -4108, "right": -4152, "fill": 5, "justify": -4130, "centerContinuous": 7, "distributed": -4117})

    def l_horizontalalignment(self, ex, v):
        n = _long(v)
        h = {1: "general", -4131: "left", -4108: "center", -4152: "right", 5: "fill", -4130: "justify", 7: "centerContinuous", -4117: "distributed"}.get(n)
        if h is None:
            raise VbaError(1004, "Unable to set the HorizontalAlignment property of the Range class")
        self._style(lambda st: replace(st, align=replace(st.align, h=h)))

    def g_verticalalignment(self, ex):
        return self._align_get("v", {"top": -4160, "center": -4108, "bottom": -4107, "justify": -4130, "distributed": -4117})

    def l_verticalalignment(self, ex, v):
        n = _long(v)
        val = {-4160: "top", -4108: "center", -4107: "bottom", -4130: "justify", -4117: "distributed"}.get(n)
        if val is None:
            raise VbaError(1004, "Unable to set the VerticalAlignment property of the Range class")
        self._style(lambda st: replace(st, align=replace(st.align, v=val)))

    def g_wraptext(self, ex):
        return self._uniform(lambda st: st.align.wrap)

    def l_wraptext(self, ex, v):
        on = _bool(v)
        self._style(lambda st: replace(st, align=replace(st.align, wrap=on)))

    def g_shrinktofit(self, ex):
        return self._uniform(lambda st: st.align.shrink)

    def l_shrinktofit(self, ex, v):
        on = _bool(v)
        self._style(lambda st: replace(st, align=replace(st.align, shrink=on)))

    def g_indentlevel(self, ex):
        return self._uniform(lambda st: st.align.indent)

    def l_indentlevel(self, ex, v):
        n = max(0, min(_long(v), 250))

        def f(st):
            # An indent needs a side: General (and centred) text becomes left-aligned.
            h = st.align.h if (n == 0 or st.align.h in ("left", "right", "distributed")) else "left"
            return replace(st, align=replace(st.align, indent=n, h=h))

        self._style(f)

    def g_orientation(self, ex):
        v = self._uniform(lambda st: st.align.rotation)
        if v is NULL:
            return NULL
        return {0: -4128, 255: -4166}.get(v, v)

    def l_orientation(self, ex, v):
        n = _long(v)
        rot = {-4128: 0, -4166: 255, -4171: 90, -4170: -90}.get(n, n)
        if not (-90 <= rot <= 90 or rot == 255):
            raise VbaError(1004)
        self._style(lambda st: replace(st, align=replace(st.align, rotation=rot)))

    def g_locked(self, ex):
        return self._uniform(lambda st: st.protection.locked)

    def l_locked(self, ex, v):
        on = _bool(v)
        self.h.require_writable()
        for r in self.areas:
            self.h.command("home.lockCell", {"on": on, "range": r.a1()}, self.si)

    def g_formulahidden(self, ex):
        return self._uniform(lambda st: st.protection.hidden)

    def l_formulahidden(self, ex, v):
        on = _bool(v)
        self._style(lambda st: replace(st, protection=replace(st.protection, hidden=on)))

    def g_style(self, ex):
        return "Normal"

    def l_style(self, ex, v):
        name = _str(_plain(v) if not hasattr(v, "g_name") else v.g_name(ex))
        for r in self.areas:
            self.h.command("home.cellStyle", {"name": name, "range": r.a1()}, self.si)

    def g_mergecells(self, ex):
        """True when every cell is in a merged area, False when none is, else Null."""
        sh = self._sheet()
        f = self.first
        if f.is_single():
            return sh.merge_at((f.r0, f.c0)) is not None
        covered = 0
        for m in sh.merges:
            x = m.intersection(f)
            if x is not None:
                covered += x.count()
        if covered == 0:
            return False
        return True if covered >= f.count() else NULL

    def l_mergecells(self, ex, v):
        if _bool(v):
            self.g_merge(ex)
        else:
            self.g_unmerge(ex)

    def g_merge(self, ex, across=MISSING):
        self.h.require_writable()
        cmd = "home.mergeAcross" if _bool(across) else "home.mergeCells"
        for r in self.areas:
            self.h.command(cmd, {"range": r.a1()}, self.si)
        return EMPTY

    def g_unmerge(self, ex):
        self.h.require_writable()
        for r in self.areas:
            self.h.command("home.unmergeCells", {"range": r.a1()}, self.si)
        return EMPTY

    # ------------------------------------------------------------ sizes and visibility

    def g_columnwidth(self, ex):
        sh = self._sheet()
        ws = {round(px_to_chars(sh.col_width(c)), 2) for c in range(self.first.c0, min(self.first.c1, self.first.c0 + 300) + 1)}
        return ws.pop() if len(ws) == 1 else NULL

    def l_columnwidth(self, ex, v):
        chars = V.to_double(_plain(v))
        if not 0 <= chars <= 255:
            raise VbaError(1004, "Unable to set the ColumnWidth property of the Range class")
        self.h.require_writable()
        self.h.command("home.columnWidth", {"width": chars_to_px(chars), "cols": _cols_text(self.first)}, self.si)

    def g_rowheight(self, ex):
        sh = self._sheet()
        hs = {round(sh.row_height(r) * 0.75, 2) for r in range(self.first.r0, min(self.first.r1, self.first.r0 + 300) + 1)}
        return hs.pop() if len(hs) == 1 else NULL

    def l_rowheight(self, ex, v):
        pts = V.to_double(_plain(v))
        if not 0 <= pts <= 409:
            raise VbaError(1004, "Unable to set the RowHeight property of the Range class")
        self.h.require_writable()
        self.h.command("home.rowHeight", {"height": pts / 0.75, "rows": _rows_text(self.first)}, self.si)

    def g_width(self, ex):
        sh = self._sheet()
        return sum(sh.col_width(c) for c in range(self.first.c0, min(self.first.c1, self.first.c0 + 16383) + 1)) * 0.75

    def g_height(self, ex):
        sh = self._sheet()
        return sum(sh.row_height(r) for r in range(self.first.r0, min(self.first.r1, self.first.r0 + 1_048_575) + 1)) * 0.75

    def g_left(self, ex):
        sh = self._sheet()
        return sum(sh.col_width(c) for c in range(0, self.first.c0)) * 0.75

    def g_top(self, ex):
        sh = self._sheet()
        if self.first.r0 > 100_000:
            return float(self.first.r0) * sh.default_row_height * 0.75
        return sum(sh.row_height(r) for r in range(0, self.first.r0)) * 0.75

    def g_hidden(self, ex):
        sh = self._sheet()
        f = self.first
        if self.mode == "cols" or f.is_full_cols():
            return all(sh.is_col_hidden(c) for c in range(f.c0, f.c1 + 1))
        return all(sh.is_row_hidden(r) for r in range(f.r0, min(f.r1, f.r0 + 100_000) + 1))

    def l_hidden(self, ex, v):
        on = _bool(v)
        self.h.require_writable()
        f = self.first
        if self.mode == "cols" or f.is_full_cols():
            self.h.command("home.hideColumns" if on else "home.unhideColumns", {"cols": _cols_text(f)}, self.si)
        elif self.mode == "rows" or f.is_full_rows():
            self.h.command("home.hideRows" if on else "home.unhideRows", {"rows": _rows_text(f)}, self.si)
        else:
            raise VbaError(1004, "Unable to set the Hidden property of the Range class")

    def g_autofit(self, ex):
        self.h.require_writable()
        f = self.first
        if self.mode == "rows" or f.is_full_rows():
            self.h.command("home.autofitRowHeight", {"rows": _rows_text(f)}, self.si)
        else:
            self.h.command("home.autofitColumnWidth", {"cols": _cols_text(f)}, self.si)
        return True

    # ------------------------------------------------------------ editing

    def _each(self, cid: str, extra: Optional[Dict[str, Any]] = None, fire: bool = False) -> Any:
        self.h.require_writable()
        for r in self.areas:
            p = {"range": r.a1()}
            if extra:
                p.update(extra)
            self.h.command(cid, p, self.si)
        if fire:
            self._fire_change()
        return True

    def g_clearcontents(self, ex):
        return self._each("edit.clearContents", fire=True)

    def g_clear(self, ex):
        return self._each("edit.clearAll", fire=True)

    def g_clearformats(self, ex):
        return self._each("edit.clearFormats")

    def g_clearcomments(self, ex):
        return self._each("edit.clearComments")

    g_clearnotes = g_clearcomments

    def g_clearhyperlinks(self, ex):
        return self._each("edit.clearHyperlinks")

    def g_clearoutline(self, ex):
        return True

    def g_delete(self, ex, shift=MISSING):
        self.h.require_writable()
        f = self.first
        if self.mode == "rows" or f.is_full_rows():
            self.h.command("home.deleteRows", {"rows": _rows_text(f)}, self.si)
            self._fire_change(self._new([RangeRef(f.r0, 0, f.r1, MAX_COLS - 1)], "rows"))
            return True
        if self.mode == "cols" or f.is_full_cols():
            self.h.command("home.deleteColumns", {"cols": _cols_text(f)}, self.si)
            self._fire_change(self._new([RangeRef(0, f.c0, MAX_ROWS - 1, f.c1)], "cols"))
            return True
        if _m(shift):
            # Excel decides by the range's shape.
            sh = "up" if f.width >= f.height else "left"
        else:
            sh = {-4162: "up", -4159: "left"}.get(_long(shift))
            if sh is None:
                raise VbaError(1004, "Delete method of Range class failed")
        self.h.command("home.deleteCells", {"shift": sh, "range": f.a1()}, self.si)
        self._fire_change(self._new([f]))
        return True

    def g_insert(self, ex, shift=MISSING, origin=MISSING):
        self.h.require_writable()
        f = self.first
        if self.mode == "rows" or f.is_full_rows():
            self.h.command("home.insertRows", {"rows": _rows_text(f)}, self.si)
            self._fire_change(self._new([RangeRef(f.r0, 0, f.r1, MAX_COLS - 1)], "rows"))
            return True
        if self.mode == "cols" or f.is_full_cols():
            self.h.command("home.insertColumns", {"cols": _cols_text(f)}, self.si)
            self._fire_change(self._new([RangeRef(0, f.c0, MAX_ROWS - 1, f.c1)], "cols"))
            return True
        if _m(shift):
            sh = "down" if f.width >= f.height else "right"
        else:
            sh = {-4121: "down", -4161: "right"}.get(_long(shift))
            if sh is None:
                raise VbaError(1004, "Insert method of Range class failed")
        self.h.command("home.insertCells", {"shift": sh, "range": f.a1()}, self.si)
        self._fire_change(self._new([f]))
        return True

    def g_select(self, ex, replace_=MISSING):
        self.h.require_writable()
        si = self.si
        if si != self.h.wb().active_sheet:
            raise VbaError(1004, "Select method of Range class failed (activate the sheet first)")
        before = self.h.selection().copy()
        self.h.command("selection.set", {"range": self.address(False, False).replace("$", ""), "active": a1(self.first.r0, self.first.c0)})
        self._fire_selection(before)
        return True

    def g_activate(self, ex):
        self.h.require_writable()
        si = self.si
        if si != self.h.wb().active_sheet:
            raise VbaError(1004, "Activate method of Range class failed (activate the sheet first)")
        sel = self.h.selection()
        before = sel.copy()
        rc = (self.first.r0, self.first.c0)
        if sel.contains(rc) and not sel.is_single_cell():
            self.h.command("selection.set", {"range": sel.a1(), "active": a1(*rc)})
        else:
            self.h.command("selection.set", {"range": self.first.a1(), "active": a1(*rc)})
        self._fire_selection(before)
        return True

    def g_show(self, ex):
        return True

    def g_calculate(self, ex):
        self.h.require_writable()
        self.h.command("formulas.calculateNow", {})
        return True

    def g_copy(self, ex, destination=MISSING):
        self.h.require_writable()
        s = self.h.s
        src = self.first
        if _m(destination):
            self.h.command("edit.copy", {"range": src.a1()}, self.si)
            self.h.cut_copy_mode = True
            return True
        dst = _as_range(destination)
        saved = s.clipboard
        try:
            self.h.command("edit.copy", {"range": src.a1()}, self.si)
            r = self.h.command("edit.paste", {"range": dst.first.a1()} if not dst.first.is_single() else {"at": a1(dst.first.r0, dst.first.c0)}, dst.si)
        finally:
            s.clipboard = saved
        dst._fire_change(dst._pasted(r, src))
        return True

    def _pasted(self, result: Any, src: Optional[RangeRef] = None) -> "Range":
        """The range a paste filled (from the command's answer), else this one."""
        text = result.get("range") if isinstance(result, dict) else None
        r = RangeRef.parse(text) if isinstance(text, str) else None
        if r is None and src is not None:
            f = self.first
            r = RangeRef(f.r0, f.c0, f.r0 + src.height - 1, f.c0 + src.width - 1)
        return self._new([r]) if r is not None else self

    def g_cut(self, ex, destination=MISSING):
        self.h.require_writable()
        src = self.first
        if _m(destination):
            self.h.command("edit.cut", {"range": src.a1()}, self.si)
            self.h.cut_copy_mode = True
            return True
        dst = _as_range(destination)
        if dst.si != self.si:
            raise NotSupported("Cutting cells to another sheet")
        self.h.command("edit.moveCells", {"from": src.a1(), "to": a1(dst.first.r0, dst.first.c0), "sheet": self.si}, self.si)
        moved = RangeRef(dst.first.r0, dst.first.c0, dst.first.r0 + src.height - 1, dst.first.c0 + src.width - 1)
        self._fire_change(self._new([src]))
        dst._fire_change(dst._new([moved]))
        return True

    def g_pastespecial(self, ex, paste=MISSING, operation=MISSING, skipblanks=MISSING, transpose=MISSING):
        self.h.require_writable()
        what = {-4104: "all", -4163: "values", -4122: "formats", -4123: "formulas", -4144: "comments", 6: "validation", 7: "allExceptBorders", 8: "columnWidths", 11: "formulasAndNumberFormats", 12: "valuesAndNumberFormats", 13: "all", 14: "all"}.get(_long(paste, -4104))
        if what is None:
            raise VbaError(1004, "PasteSpecial method of Range class failed")
        op = {-4142: "none", 2: "add", 3: "subtract", 4: "multiply", 5: "divide"}.get(_long(operation, -4142), "none")
        if self.h.s.clipboard is None:
            raise VbaError(1004, "PasteSpecial method of Range class failed (nothing was copied)")
        r = self.h.command("edit.pasteSpecial", {"what": what, "operation": op, "skipBlanks": _bool(skipblanks), "transpose": _bool(transpose), "at": a1(self.first.r0, self.first.c0)}, self.si)
        if what not in ("formats", "columnWidths", "comments", "validation"):
            self._fire_change(self._pasted(r))
        return True

    def g_filldown(self, ex):
        return self._each("edit.fillDown", fire=True)

    def g_fillright(self, ex):
        return self._each("edit.fillRight", fire=True)

    def g_fillup(self, ex):
        return self._each("edit.fillUp", fire=True)

    def g_fillleft(self, ex):
        return self._each("edit.fillLeft", fire=True)

    def g_autofill(self, ex, destination=MISSING, type_=MISSING):
        self.h.require_writable()
        dst = _as_range(destination)
        mode = {0: "series", 1: "copy", 2: "fillseries", 3: "formats", 4: "values"}.get(_long(type_, 0), "series")
        self.h.command("edit.autoFill", {"source": self.first.a1(), "target": dst.first.a1(), "mode": mode}, self.si)
        dst._fire_change()
        return True

    def g_sort(self, ex, key1=MISSING, order1=MISSING, key2=MISSING, type_=MISSING, order2=MISSING, key3=MISSING, order3=MISSING, header=MISSING, ordercustom=MISSING, matchcase=MISSING, orientation=MISSING, sortmethod=MISSING, *rest):
        self.h.require_writable()
        rng = self.first
        if rng.is_single():
            from ..selection import current_region

            rng = current_region(self._sheet(), (rng.r0, rng.c0))
        keys = []
        for k, o in ((key1, order1), (key2, order2), (key3, order3)):
            if _m(k):
                continue
            kr = k if isinstance(k, Range) else make_range(self.h, self.ws, k)
            col = kr.first.c0 - rng.c0 if _long(orientation, 1) != 2 else kr.first.r0 - rng.r0
            keys.append({"col": col, "order": "desc" if _long(o, 1) == 2 else "asc"})
        if not keys:
            keys = [{"col": 0, "order": "asc"}]
        hdr = _long(header, 2)
        p: Dict[str, Any] = {"range": rng.a1(), "keys": keys, "matchCase": _bool(matchcase), "byColumns": _long(orientation, 1) == 2}
        if hdr == 1:
            p["header"] = True
        elif hdr == 2:
            p["header"] = False
        self.h.command("data.sort", p, self.si)
        return True

    def g_removeduplicates(self, ex, columns=MISSING, header=MISSING):
        self.h.require_writable()
        cols: List[int] = []
        c = _plain(columns) if not _m(columns) else MISSING
        if type(c) is VArray:
            cols = [V.to_long(x) - 1 for x in c.data]
        elif not _m(columns):
            cols = [V.to_long(c) - 1]
        else:
            cols = list(range(self.first.width))
        self.h.command("data.removeDuplicates", {"range": self.first.a1(), "columns": cols, "header": _long(header, 2) == 1}, self.si)
        return True

    def g_autofilter(self, ex, field=MISSING, criteria1=MISSING, operator=MISSING, criteria2=MISSING, dropdown=MISSING):
        self.h.require_writable()
        si = self.si
        sh = self._sheet()
        rng = self.first
        if _m(field):
            on = sh.autofilter is None
            self.h.command("data.filter", {"range": rng.a1(), "on": on}, si)
            return True
        if sh.autofilter is None:
            self.h.command("data.filter", {"range": rng.a1(), "on": True}, si)
        col = _long(field) - 1
        op = _long(operator, 1)
        crit = _filter_criterion(criteria1, criteria2, op)
        self.h.command("data.filterColumn", {"col": col, "criterion": crit}, si)
        return True

    def g_texttocolumns(self, ex, destination=MISSING, datatype=MISSING, qualifier=MISSING, consecutive=MISSING, tab=MISSING, semicolon=MISSING, comma=MISSING, space=MISSING, other=MISSING, otherchar=MISSING, *rest):
        self.h.require_writable()
        delims = [ch for flag, ch in ((tab, "\t"), (semicolon, ";"), (comma, ","), (space, " ")) if _bool(flag)]
        p: Dict[str, Any] = {"range": self.first.a1(), "delimiters": delims or ["\t"], "consecutive": _bool(consecutive)}
        if _bool(other) and not _m(otherchar):
            p["other"] = _str(otherchar)[:1]
        if not _m(destination):
            p["destination"] = a1(_as_range(destination).first.r0, _as_range(destination).first.c0)
        self.h.command("data.textToColumns", p, self.si)
        return True

    def g_goalseek(self, ex, goal=MISSING, changing=MISSING):
        self.h.require_writable()
        ch = _as_range(changing)
        r = self.h.command("data.goalSeek", {"setCell": a1(self.first.r0, self.first.c0), "toValue": V.to_double(_plain(goal)), "byChangingCell": a1(ch.first.r0, ch.first.c0)}, self.si)
        return bool(r.get("found")) if isinstance(r, dict) else True

    # ------------------------------------------------------------ find and replace

    def g_find(self, ex, what=MISSING, after=MISSING, lookin=MISSING, lookat=MISSING, order=MISSING, direction=MISSING, matchcase=MISSING, matchbyte=MISSING, fmt=MISSING):
        if _m(what):
            raise VbaError(449)
        # Excel keeps LookIn, LookAt and SearchOrder from one Find (or Replace) to the next.
        saved = _find_settings(self.h)
        if not _m(lookin):
            saved["lookin"] = _long(lookin)
        if not _m(lookat):
            saved["lookat"] = _long(lookat)
        if not _m(order):
            saved["order"] = _long(order)
        opts = {
            "what": _plain(what),
            "lookin": saved["lookin"],
            "whole": saved["lookat"] == 1,
            "bycols": saved["order"] == 2,
            "prev": _long(direction, 1) == 2,
            "case": _bool(matchcase) if not _m(matchcase) else saved["case"],
            "range": self,
        }
        self.h.last_find = opts
        start = _as_range(after).first.start if not _m(after) else None
        return self._find(opts, start, opts["prev"])

    def g_findnext(self, ex, after=MISSING):
        opts = self.h.last_find
        if opts is None:
            return NOTHING
        start = _as_range(after).first.start if not _m(after) else None
        return self._find(opts, start, False)

    def g_findprevious(self, ex, after=MISSING):
        opts = self.h.last_find
        if opts is None:
            return NOTHING
        start = _as_range(after).first.start if not _m(after) else None
        return self._find(opts, start, True)

    def _find(self, o: Dict[str, Any], start: Optional[Tuple[int, int]], prev: bool) -> Any:
        self._flush_for_read()
        wb = self.h.wb()
        si = self.si
        sh = wb.sheets[si]
        area = self.first
        if area.is_single() and o.get("range") is self and len(self.areas) == 1 and False:
            area = sh.used_range() or area
        used = sh.used_range()
        if used is None:
            return NOTHING
        box = area.intersection(used)
        if box is None:
            return NOTHING
        what = o["what"]
        needle = V.text(what) if not isinstance(what, str) else what
        if not o["case"]:
            needle = needle.lower()
        match = _find_matcher(needle, o["whole"])
        from ..display import value_text

        candidates: List[Tuple[int, int]] = []
        lookin = o["lookin"]
        for rc, cell in sh.cells.iter_range(box):
            if cell.value is None and cell.formula is None:
                continue
            if lookin == -4123 and cell.formula is not None:
                hay = "=" + cell.formula.text
            elif lookin == -4144:
                cm = sh.comments.get(rc)
                if cm is None:
                    continue
                hay = cm.text
            else:
                hay = value_text(wb, sh, rc) if lookin == -4163 else (cell.input_text() if cell.formula is None else "=" + cell.formula.text)
            if match(hay if o["case"] else hay.lower()):
                candidates.append(rc)
        if not candidates:
            return NOTHING
        key = (lambda rc: (rc[1], rc[0])) if o["bycols"] else (lambda rc: rc)
        candidates.sort(key=key)
        if start is None:
            # Excel starts after (or, searching backwards, before) the range's first cell,
            # wrapping around.
            start = (area.r0, area.c0)
        sk = key(start)
        if prev:
            before = [rc for rc in candidates if key(rc) < sk]
            hit = before[-1] if before else candidates[-1]
        else:
            after_ = [rc for rc in candidates if key(rc) > sk]
            hit = after_[0] if after_ else candidates[0]
        return self._new([RangeRef.of(hit)])

    def g_replace(self, ex, what=MISSING, replacement=MISSING, lookat=MISSING, order=MISSING, matchcase=MISSING, *rest):
        self.h.require_writable()
        import re

        if _m(what) or _m(replacement):
            raise VbaError(449)
        self._flush_for_read()
        needle = V.text(_plain(what))
        rep = V.text(_plain(replacement))
        # Replace keeps LookAt, SearchOrder and MatchCase for the next Find or Replace.
        saved = _find_settings(self.h)
        if not _m(lookat):
            saved["lookat"] = _long(lookat)
        if not _m(order):
            saved["order"] = _long(order)
        if not _m(matchcase):
            saved["case"] = _bool(matchcase)
        whole = saved["lookat"] == 1
        case = saved["case"]
        wb = self.h.wb()
        si = self.si
        sh = wb.sheets[si]
        used = sh.used_range()
        if used is None or needle == "":
            return True
        rx = re.compile(re.escape(needle), 0 if case else re.IGNORECASE)
        writes = []
        for area in self.areas:
            box = area.intersection(used)
            if box is None:
                continue
            for rc, cell in sh.cells.iter_range(box):
                if cell.value is None and cell.formula is None:
                    continue
                text = "=" + cell.formula.text if cell.formula is not None else cell.input_text()
                if whole:
                    if (text == needle) if case else (text.lower() == needle.lower()):
                        writes.append((rc, rep))
                elif rx.search(text):
                    writes.append((rc, rx.sub(lambda m: rep, text)))
        if writes:
            self.h.write_values(si, writes)
        return True

    def g_specialcells(self, ex, type_=MISSING, value=MISSING):
        self._flush_for_read()
        t = _long(type_)
        wb = self.h.wb()
        si = self.si
        sh = wb.sheets[si]
        used = sh.used_range()
        if t == 11:
            if used is None:
                return self._new([RangeRef(0, 0, 0, 0)])
            return self._new([RangeRef.of((used.r1, used.c1))])
        scope = self.areas if not (len(self.areas) == 1 and self.first.is_single()) else ([used] if used is not None else [])
        mask = _long(value, 23)
        hits: List[Tuple[int, int]] = []
        for area in scope:
            if used is None:
                break
            box = area.intersection(used)
            if box is None:
                continue
            if t == 4:
                for rc in box.cells():
                    c = sh.cell(rc)
                    if (c is None or (c.value is None and c.formula is None)) and rc not in sh.spill:
                        hits.append(rc)
            elif t == 12:
                for r in range(box.r0, box.r1 + 1):
                    if sh.is_row_hidden(r):
                        continue
                    for c in range(box.c0, box.c1 + 1):
                        if not sh.is_col_hidden(c):
                            hits.append((r, c))
            elif t == -4144:
                hits.extend(rc for rc in sh.comments if box.contains(rc))
            elif t in (2, -4123):
                for rc, c in sh.cells.iter_range(box):
                    is_f = c.formula is not None
                    if (t == -4123) != is_f or (c.value is None and not is_f):
                        continue
                    v = c.value
                    kind = 16 if isinstance(v, CellError) else (4 if isinstance(v, bool) else (1 if isinstance(v, float) else 2))
                    if mask & kind:
                        hits.append(rc)
            elif t in (-4174, -4175):
                for vd in sh.validations:
                    for r in vd.ranges:
                        inter = r.intersection(box)
                        if inter is not None:
                            hits.extend(inter.cells())
            elif t in (-4172, -4173):
                for cf in sh.cond_formats:
                    for r in cf.ranges:
                        inter = r.intersection(box)
                        if inter is not None:
                            hits.extend(inter.cells())
            else:
                raise VbaError(1004, "SpecialCells method of Range class failed")
        if not hits:
            raise VbaError(1004, "No cells were found.")
        from ..ranges import cells_to_ranges

        return self._new(cells_to_ranges(sorted(set(hits))))

    # ------------------------------------------------------------ comments, links, names, validation

    def g_comment(self, ex):
        rc = (self.first.r0, self.first.c0)
        if rc in self._sheet().comments:
            return Comment(self.h, self.ws, rc)
        return NOTHING

    def g_addcomment(self, ex, text=MISSING):
        self.h.require_writable()
        rc = (self.first.r0, self.first.c0)
        if rc in self._sheet().comments:
            raise VbaError(1004, "AddComment method of Range class failed (the cell already has a note)")
        self.h.command("review.newNote", {"cell": a1(*rc), "text": _str(text), "author": self.h.app.g_username(ex)}, self.si)
        return Comment(self.h, self.ws, rc)

    def g_notetext(self, ex, text=MISSING, start=MISSING, length=MISSING):
        rc = (self.first.r0, self.first.c0)
        cm = self._sheet().comments.get(rc)
        cur = cm.text if cm is not None else ""
        if _m(text):
            return cur
        self.h.require_writable()
        new = _str(text)
        if not _m(start):
            st = _long(start)
            ln = _long(length, len(cur))
            new = cur[:st - 1] + new + cur[st - 1 + ln:]
        if cm is None:
            self.h.command("review.newNote", {"cell": a1(*rc), "text": new}, self.si)
        else:
            self.h.command("review.editNote", {"cell": a1(*rc), "text": new}, self.si)
        return new

    def g_hyperlinks(self, ex, index=MISSING):
        from .excel import Hyperlinks

        h = Hyperlinks(self.h, self.ws, self)
        return h if _m(index) else h.g_item(ex, index)

    def g_name(self, ex):
        from .excel import Name

        wb = self.h.wb()
        me = self.address()
        for n in wb.names:
            f = n.formula.lstrip("=")
            if "!" in f and f.split("!", 1)[1].replace("$", "") == me.replace("$", "") and f.split("!", 1)[0].strip("'") == self.ws.name:
                return Name(self.h, n.name, n.scope)
        raise VbaError(1004, "The range has no name")

    def l_name(self, ex, v):
        from ...formula.printer import quote_sheet

        self.h.require_writable()
        name = _str(v)
        scope = None
        if "!" in name:
            sheet, name = name.rsplit("!", 1)
            scope = sheet.strip("'")
        p = {"name": name, "refersTo": f"={quote_sheet(self.ws.name)}!{self.address()}", "replace": True}
        if scope:
            p["scope"] = scope
        self.h.command("formulas.defineName", p)

    def g_validation(self, ex):
        from .excel import Validation

        return Validation(self)

    def g_formatconditions(self, ex):
        from .excel import FormatConditions

        return FormatConditions(self)

    def g_characters(self, ex, start=MISSING, length=MISSING):
        return Characters(self, start, length)

    def g_displayformat(self, ex):
        return self

    def g_printout(self, ex, *args):
        raise NotSupported("Printing from a macro")

    g_printpreview = g_printout

    def g_dependents(self, ex):
        return self._trace("formulas.traceDependents", "dependents")

    g_directdependents = g_dependents

    def g_precedents(self, ex):
        return self._trace("formulas.tracePrecedents", "precedents")

    g_directprecedents = g_precedents

    def _trace(self, cid: str, key: str) -> "Range":
        r = self.h.command(cid, {"cell": a1(self.first.r0, self.first.c0)}, self.si) if self.h.writable else None
        items = (r or {}).get(key) or []
        areas = []
        for it in items:
            if it.get("sheet") != self.ws.name:
                continue
            rr = RangeRef.parse(it.get("range") or it.get("cell") or "")
            if rr is not None:
                areas.append(rr)
        if not areas:
            raise VbaError(1004, "No cells were found.")
        return self._new(areas)


# ---------------------------------------------------------------- range helpers


def _addr(r: RangeRef, row_abs: bool = True, col_abs: bool = True) -> str:
    def cell(row: int, col: int) -> str:
        return ("$" if col_abs else "") + col_to_letters(col) + ("$" if row_abs else "") + str(row + 1)

    if r.is_full_rows():
        return ("$" if row_abs else "") + str(r.r0 + 1) + ":" + ("$" if row_abs else "") + str(r.r1 + 1)
    if r.is_full_cols():
        return ("$" if col_abs else "") + col_to_letters(r.c0) + ":" + ("$" if col_abs else "") + col_to_letters(r.c1)
    if r.is_single():
        return cell(r.r0, r.c0)
    return cell(r.r0, r.c0) + ":" + cell(r.r1, r.c1)


def _r1c1_addr(r: RangeRef, row_abs: bool, col_abs: bool, base: Tuple[int, int]) -> str:
    def part(letter: str, v: int, absolute: bool, b: int) -> str:
        if absolute:
            return f"{letter}{v + 1}"
        d = v - b
        return letter if d == 0 else f"{letter}[{d}]"

    a = part("R", r.r0, row_abs, base[0]) + part("C", r.c0, col_abs, base[1])
    if r.is_single():
        return a
    return a + ":" + part("R", r.r1, row_abs, base[0]) + part("C", r.c1, col_abs, base[1])


def _rows_text(r: RangeRef) -> str:
    return f"{r.r0 + 1}:{r.r1 + 1}"


def _cols_text(r: RangeRef) -> str:
    return f"{col_to_letters(r.c0)}:{col_to_letters(r.c1)}"


def px_to_chars(px: float) -> float:
    if px <= 12:
        return px / 12.0
    return (px - 5) / 7.0


def chars_to_px(chars: float) -> float:
    if chars < 1:
        return round(chars * 12)
    return round(chars * 7 + 5)


def _as_rows(a: VArray) -> List[List[Any]]:
    if a.dims == 1:
        return [list(a.data)]
    if a.dims == 2:
        return a.rows()
    raise VbaError(13)


def _relative_formulas(text: str, rng: RangeRef) -> List[Tuple[Tuple[int, int], str]]:
    """A formula meant for the top-left cell, adjusted for every cell of ``rng`` (as Excel does
    when one Formula is given for many cells)."""
    from ...formula import ParseError, parse, print_formula
    from ...formula.adjust import shift_relative

    body = text[1:]
    try:
        e = parse(body)
    except ParseError:
        return [(rc, text) for rc in rng.cells()]
    out = []
    for rc in rng.cells():
        dr, dc = rc[0] - rng.r0, rc[1] - rng.c0
        out.append((rc, "=" + print_formula(shift_relative(e, dr, dc)) if (dr or dc) else text))
    return out


def _parse_simple(text: str, base: RangeRef) -> RangeRef:
    """``Rows("2:5")``/``Columns("B:D")``/``Cells("A1")``-style text."""
    t = text.strip().replace("$", "")
    r = RangeRef.parse(t)
    if r is None:
        raise VbaError(1004, f"Method 'Range' failed: {text}")
    return r


def _line_style(style: int, weight: int) -> Optional[str]:
    if style in (-4142,):
        return None
    if style == -4119:
        return "double"
    if style == -4115:
        return "dashed" if weight != -4138 else "mediumDashed"
    if style == -4118:
        return "dotted"
    if style == 4:
        return "dashDot" if weight != -4138 else "mediumDashDot"
    if style == 5:
        return "dashDotDot" if weight != -4138 else "mediumDashDotDot"
    if style == 13:
        return "slantDashDot"
    if weight == 4:
        return "thick"
    if weight == -4138:
        return "medium"
    if weight == 1:
        return "hair"
    return "thin"


def _style_line(name: Optional[str]) -> Tuple[int, int]:
    """(LineStyle, Weight) of a border line style name."""
    return {
        None: (-4142, 2), "thin": (1, 2), "medium": (1, -4138), "thick": (1, 4), "hair": (1, 1), "double": (-4119, 4),
        "dashed": (-4115, 2), "mediumDashed": (-4115, -4138), "dotted": (-4118, 2), "dashDot": (4, 2),
        "mediumDashDot": (4, -4138), "dashDotDot": (5, 2), "mediumDashDotDot": (5, -4138), "slantDashDot": (13, -4138),
    }.get(name, (1, 2))


def _filter_criterion(c1: Any, c2: Any, op: int) -> Optional[Dict[str, Any]]:
    def one(c: Any) -> Tuple[str, str]:
        t = V.text(_plain(c))
        for prefix, o in (("<>", "notEqual"), (">=", "greaterThanOrEqual"), ("<=", "lessThanOrEqual"), (">", "greaterThan"), ("<", "lessThan"), ("=", "equal")):
            if t.startswith(prefix):
                return o, t[len(prefix):]
        return "equal", t

    if _m(c1):
        return None
    v1 = c1 if type(c1) is VArray else _plain(c1)
    if op == 7 or type(v1) is VArray:
        vals = [V.text(x) for x in v1.data] if type(v1) is VArray else [V.text(v1)]
        return {"kind": "values", "values": [x.lstrip("=") for x in vals]}
    if op in (3, 4, 5, 6):
        n = V.to_long(_plain(c1))
        return {"kind": "top10", "count": n, "bottom": op in (4, 6), "percent": op in (5, 6)}
    if op in (8, 9):
        col = color_value(c1)
        return {"kind": "fillColor" if op == 8 else "fontColor", "color": col}
    conds = [dict(zip(("op", "value"), one(c1)))]
    if not _m(c2):
        conds.append(dict(zip(("op", "value"), one(c2))))
    if len(conds) == 1 and conds[0]["op"] == "equal" and not any(ch in conds[0]["value"] for ch in "*?"):
        return {"kind": "values", "values": [conds[0]["value"]]}
    return {"kind": "custom", "conditions": conds, "and": op != 2}


def _as_range(v: Any, base: Optional[Range] = None) -> Range:
    if isinstance(v, Range):
        return v
    if v is NOTHING:
        raise VbaError(91)
    if isinstance(v, str) and base is not None:
        return make_range(base.h, base.ws, v)
    raise VbaError(424)


_ARRAYS: Dict[int, Tuple[Any, List[RangeRef]]] = {}


def _array_ranges(sh) -> List[RangeRef]:
    """The legacy array formulas of a sheet snapshot (computed once per snapshot)."""
    hit = _ARRAYS.get(id(sh))
    if hit is not None and hit[0] is sh:
        return hit[1]
    out: List[RangeRef] = []
    used = sh.used_range()
    if used is not None:
        for _, cell in sh.cells.iter_range(used):
            f = cell.formula
            if f is not None and f.array is not None and not f.array.is_single():
                out.append(f.array)
    if len(_ARRAYS) > 64:
        _ARRAYS.clear()
    _ARRAYS[id(sh)] = (sh, out)
    return out


def constant_text(v: Any) -> str:
    """``Range.Formula`` of a constant: the value itself (a number as Excel writes it, up to 15
    digits, whatever its format: 0.5 for 50 %, the serial number for a date)."""
    from ...core.value import CellError

    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return number_text(v)
    if isinstance(v, CellError):
        return str(v)
    return str(v)


def number_text(v: float) -> str:
    if v == 0:
        return "0"
    s = f"{v:.15g}"
    if "e" not in s:
        return s
    mant, exp = s.split("e")
    e = int(exp)
    if -10 < e < 21:
        import decimal

        t = format(decimal.Decimal(s), "f")
        if "." in t:
            t = t.rstrip("0").rstrip(".")
        return t
    return f"{mant.upper()}E{'+' if e > 0 else '-'}{abs(e):02d}"


def _find_settings(h) -> Dict[str, Any]:
    """The Find/Replace settings Excel remembers (per person here)."""
    owner = getattr(h, "s", None) or h
    d = getattr(owner, "vba_find_settings", None)
    if d is None:
        d = {"lookin": -4123, "lookat": 2, "order": 1, "case": False}
        try:
            owner.vba_find_settings = d
        except AttributeError:
            pass
    return d


def _find_matcher(needle: str, whole: bool):
    """Find's test: ``*`` and ``?`` are wildcards and ``~`` escapes them, as in Excel."""
    if not any(ch in needle for ch in "*?~"):
        return (lambda h: h == needle) if whole else (lambda h: needle in h)
    import re

    out, i = [], 0
    while i < len(needle):
        ch = needle[i]
        if ch == "~" and i + 1 < len(needle) and needle[i + 1] in "*?~":
            out.append(re.escape(needle[i + 1]))
            i += 2
            continue
        out.append(".*" if ch == "*" else "." if ch == "?" else re.escape(ch))
        i += 1
    rx = re.compile("".join(out), re.S)
    return (lambda h: rx.fullmatch(h) is not None) if whole else (lambda h: rx.search(h) is not None)


def _shift(r: RangeRef, dr: int, dc: int) -> RangeRef:
    return RangeRef(r.r0 + dr, r.c0 + dc, r.r1 + dr, r.c1 + dc)


def make_range(h: XlHost, ws, cell1: Any, cell2: Any = MISSING) -> Range:
    """``Range(cell1[, cell2])`` on worksheet ``ws``: an address, a name, or two corners."""
    if _m(cell1):
        raise VbaError(449)
    if not _m(cell2):
        a = make_range(h, ws, cell1)
        b = make_range(h, ws, cell2)
        if b.ws.name.lower() != a.ws.name.lower():
            raise VbaError(1004, "Method 'Range' of object '_Worksheet' failed")
        box = a.first
        for r in a.areas[1:] + b.areas:
            box = box.union(r)
        return Range(h, a.ws, [box])
    if isinstance(cell1, Range):
        return Range(h, cell1.ws, list(cell1.areas))
    text = V.to_string(_plain(cell1)).strip()
    return parse_reference(h, ws, text)


def parse_reference(h: XlHost, ws, text: str) -> Range:
    if not text:
        raise VbaError(1004, "Method 'Range' of object '_Global' failed")
    plain = text.replace("$", "")
    # Plain addresses (the common case) without evaluating anything.
    areas: List[RangeRef] = []
    target = ws
    ok = True
    for part in plain.split(","):
        p = part.strip()
        sheet_name = None
        if "!" in p:
            sheet_name, p = p.rsplit("!", 1)
            sheet_name = sheet_name.strip().strip("'").replace("''", "'")
        r = RangeRef.parse(p)
        if r is None:
            ok = False
            break
        if sheet_name is not None:
            if h.wb().sheet_index(sheet_name) is None:
                raise VbaError(1004, f"Method 'Range' failed: no sheet named {sheet_name}")
            target = h.sheet_object(h.wb().sheets[h.wb().sheet_index(sheet_name)].name)
        areas.append(r)
    if ok and areas:
        return Range(h, target, areas)
    # Names, tables, structured references, intersections: the formula engine resolves them.
    from ...calc.eval import Evaluator, RefVal
    from ...calc.recalc import Calc, _PassHost
    from ...core.date import now_serial
    from ...formula import ParseError, parse

    try:
        e = parse(text)
    except ParseError:
        raise VbaError(1004, f"Method 'Range' failed: {text}")
    wb = h.wb()
    si = h.index_of(ws.name)
    host = _PassHost(wb, {}, Calc(), now_serial())
    ref = Evaluator(host, si, (0, 0)).eval(e)
    if not isinstance(ref, RefVal) or not ref.areas:
        raise VbaError(1004, f"Method 'Range' failed: {text}")
    s0 = ref.areas[0].sheet
    if any(a.sheet != s0 for a in ref.areas):
        raise VbaError(1004, f"Method 'Range' failed: {text}")
    return Range(h, h.sheet_object(wb.sheets[s0].name), [a.rng for a in ref.areas])


class Areas(HostObject):
    TYPE_NAME = "Areas"
    DEFAULT = "item"

    def __init__(self, rng: Range):
        self.rng = rng

    def g_count(self, ex):
        return len(self.rng.areas)

    def g_item(self, ex, i=MISSING):
        n = _long(i)
        if not 1 <= n <= len(self.rng.areas):
            raise VbaError(9)
        return self.rng._new([self.rng.areas[n - 1]])

    def vba_enum(self, ex):
        return iter([self.rng._new([a]) for a in self.rng.areas])


# ---------------------------------------------------------------- Font, Interior, Borders


class Font(HostObject):
    TYPE_NAME = "Font"

    def __init__(self, rng: Range):
        self.rng = rng

    def _get(self, f: Callable) -> Any:
        return self.rng._uniform(lambda st: f(st.font))

    def _set(self, **kw: Any) -> None:
        self.rng._style(lambda st: replace(st, font=replace(st.font, **kw)))

    def g_bold(self, ex):
        return self._get(lambda f: f.bold)

    def l_bold(self, ex, v):
        self._set(bold=_bool(v))

    def g_italic(self, ex):
        return self._get(lambda f: f.italic)

    def l_italic(self, ex, v):
        self._set(italic=_bool(v))

    def g_strikethrough(self, ex):
        return self._get(lambda f: f.strike)

    def l_strikethrough(self, ex, v):
        self._set(strike=_bool(v))

    def g_underline(self, ex):
        v = self._get(lambda f: f.underline)
        if v is NULL:
            return NULL
        return {"none": -4142, "single": 2, "double": -4119, "singleAccounting": 4, "doubleAccounting": 5}.get(v, 2)

    def l_underline(self, ex, v):
        pv = _plain(v)
        if type(pv) is bool:
            u = "single" if pv else "none"
        else:
            u = {-4142: "none", 2: "single", -4119: "double", 4: "singleAccounting", 5: "doubleAccounting", 0: "none", 1: "single", -1: "single"}.get(V.to_long(pv), "single")
        self._set(underline=u)

    def g_size(self, ex):
        return self._get(lambda f: f.size)

    def l_size(self, ex, v):
        n = V.to_double(_plain(v))
        if not 1 <= n <= 409:
            raise VbaError(1004, "Unable to set the Size property of the Font class")
        self._set(size=float(n))

    def g_name(self, ex):
        return self._get(lambda f: f.name)

    def l_name(self, ex, v):
        self._set(name=_str(v))

    def g_color(self, ex):
        wb = self.rng.h.wb()
        v = self._get(lambda f: f.color)
        if v is NULL:
            return NULL
        rgb = v.resolve(wb.theme) if not v.is_auto else None
        return float(rgb_to_bgr(rgb)) if rgb is not None else 0.0

    def l_color(self, ex, v):
        from ..cmd_format import parse_color

        self._set(color=parse_color(color_value(v)))

    def g_colorindex(self, ex):
        wb = self.rng.h.wb()
        v = self._get(lambda f: f.color)
        if v is NULL:
            return NULL
        if v.is_auto:
            return -4105
        return color_index_of(v.resolve(wb.theme))

    def l_colorindex(self, ex, v):
        from ..cmd_format import parse_color

        n = _long(v)
        self._set(color=AUTO if n in (-4105, 0, -4142) else parse_color(palette_hex(n)))

    def g_themecolor(self, ex):
        v = self._get(lambda f: f.color)
        if v is NULL:
            return NULL
        return v.value + 1 if not v.is_auto and v.kind == "theme" else 0

    def l_themecolor(self, ex, v):
        n = _long(v)
        self._set(color=Color.theme(max(n - 1, 0)))

    def g_tintandshade(self, ex):
        v = self._get(lambda f: f.color)
        return 0.0 if v is NULL or v.is_auto else v.tint / 1000.0

    def l_tintandshade(self, ex, v):
        t = int(round(V.to_double(_plain(v)) * 1000))
        self.rng._style(lambda st: replace(st, font=replace(st.font, color=replace(st.font.color, tint=t) if not st.font.color.is_auto else st.font.color)))

    def g_superscript(self, ex):
        return self._get(lambda f: f.vert == "superscript")

    def l_superscript(self, ex, v):
        self._set(vert="superscript" if _bool(v) else "baseline")

    def g_subscript(self, ex):
        return self._get(lambda f: f.vert == "subscript")

    def l_subscript(self, ex, v):
        self._set(vert="subscript" if _bool(v) else "baseline")

    def g_fontstyle(self, ex):
        b = self._get(lambda f: (f.bold, f.italic))
        if b is NULL:
            return NULL
        return {(False, False): "Regular", (True, False): "Bold", (False, True): "Italic", (True, True): "Bold Italic"}[b]

    def l_fontstyle(self, ex, v):
        t = _str(v).lower()
        self._set(bold="bold" in t, italic="italic" in t)

    def g_parent(self, ex):
        return self.rng


class Interior(HostObject):
    TYPE_NAME = "Interior"

    def __init__(self, rng: Range):
        self.rng = rng

    def _fill(self) -> Any:
        return self.rng._uniform(lambda st: st.fill)

    def g_color(self, ex):
        f = self._fill()
        if f is NULL:
            return NULL
        if f.pattern == "none":
            return 16777215.0
        rgb = f.fg.resolve(self.rng.h.wb().theme)
        return float(rgb_to_bgr(rgb)) if rgb is not None else 16777215.0

    def l_color(self, ex, v):
        from ..cmd_format import parse_color

        c = parse_color(color_value(v))
        self.rng._style(lambda st: replace(st, fill=Fill.solid(c)))

    def g_colorindex(self, ex):
        f = self._fill()
        if f is NULL:
            return NULL
        if f.pattern == "none":
            return -4142
        return color_index_of(f.fg.resolve(self.rng.h.wb().theme))

    def l_colorindex(self, ex, v):
        from ..cmd_format import parse_color

        n = _long(v)
        if n in (-4142, 0):
            self.rng._style(lambda st: replace(st, fill=Fill()))
        elif n == -4105:
            self.rng._style(lambda st: replace(st, fill=Fill()))
        else:
            c = parse_color(palette_hex(n))
            self.rng._style(lambda st: replace(st, fill=Fill.solid(c)))

    def g_pattern(self, ex):
        f = self._fill()
        if f is NULL:
            return NULL
        return {"none": -4142, "solid": 1}.get(f.pattern, 1)

    def l_pattern(self, ex, v):
        n = _long(v)
        if n in (-4142, 0):
            self.rng._style(lambda st: replace(st, fill=Fill()))
        elif n == 1:
            self.rng._style(lambda st: replace(st, fill=Fill.solid(st.fill.fg) if st.fill.pattern != "none" else Fill.solid(Color.rgb(0xFFFFFF))))

    def g_themecolor(self, ex):
        f = self._fill()
        if f is NULL:
            return NULL
        return f.fg.value + 1 if f.pattern != "none" and f.fg.kind == "theme" else 0

    def l_themecolor(self, ex, v):
        n = _long(v)
        self.rng._style(lambda st: replace(st, fill=Fill.solid(Color.theme(max(n - 1, 0)))))

    def g_tintandshade(self, ex):
        f = self._fill()
        return 0.0 if f is NULL or f.pattern == "none" else f.fg.tint / 1000.0

    def l_tintandshade(self, ex, v):
        t = int(round(V.to_double(_plain(v)) * 1000))
        self.rng._style(lambda st: replace(st, fill=Fill.solid(replace(st.fill.fg, tint=t))) if st.fill.pattern != "none" else st)

    def g_patterncolor(self, ex):
        return 0.0

    def l_patterncolor(self, ex, v):
        pass

    def g_patterncolorindex(self, ex):
        return -4105

    def l_patterncolorindex(self, ex, v):
        pass

    def g_parent(self, ex):
        return self.rng


_EDGES = {5: "diagDown", 6: "diagUp", 7: "left", 8: "top", 9: "bottom", 10: "right", 11: "insideV", 12: "insideH"}


class Borders(HostObject):
    TYPE_NAME = "Borders"
    DEFAULT = "item"

    def __init__(self, rng: Range):
        self.rng = rng

    def g_item(self, ex, index=MISSING):
        n = _long(index)
        if n not in _EDGES:
            raise VbaError(1004, "Unable to get the Item property of the Borders class")
        return Border(self.rng, [n])

    def g_count(self, ex):
        return 6

    def vba_enum(self, ex):
        return iter([Border(self.rng, [n]) for n in (7, 8, 9, 10, 5, 6)])

    def _all(self) -> "Border":
        return Border(self.rng, [7, 8, 9, 10, 11, 12])

    def g_linestyle(self, ex):
        return self._all().g_linestyle(ex)

    def l_linestyle(self, ex, v):
        self._all().l_linestyle(ex, v)

    def g_weight(self, ex):
        return self._all().g_weight(ex)

    def l_weight(self, ex, v):
        self._all().l_weight(ex, v)

    def g_color(self, ex):
        return self._all().g_color(ex)

    def l_color(self, ex, v):
        self._all().l_color(ex, v)

    def g_colorindex(self, ex):
        return self._all().g_colorindex(ex)

    def l_colorindex(self, ex, v):
        self._all().l_colorindex(ex, v)

    def g_value(self, ex):
        return self.g_linestyle(ex)

    def l_value(self, ex, v):
        self.l_linestyle(ex, v)


class Border(HostObject):
    TYPE_NAME = "Border"

    def __init__(self, rng: Range, edges: List[int]):
        self.rng = rng
        self.edges = edges

    def _lines(self) -> Any:
        f = self.rng.first
        sh = self.rng._sheet()
        wb = self.rng.h.wb()
        e = self.edges[0]
        name = _EDGES[e]
        if name == "insideV":
            # The lines between the range's columns: the first cell's right edge.
            if f.width < 2:
                return BorderLine()
            name = "right"
            f = RangeRef(f.r0, f.c0, f.r0, f.c0)
        elif name == "insideH":
            if f.height < 2:
                return BorderLine()
            name = "bottom"
            f = RangeRef(f.r0, f.c0, f.r0, f.c0)
        r, c = (f.r0, f.c0)
        if name == "bottom":
            r = f.r1
        if name == "right":
            c = f.c1
        st = wb.styles.get(sh.style_id((r, c)))
        return getattr(st.borders, {"diagDown": "diag_down", "diagUp": "diag_up"}.get(name, name))

    def g_linestyle(self, ex):
        ln = self._lines()
        if not ln or not getattr(ln, "style", None) or ln.style == "none":
            return -4142
        return _style_line(ln.style)[0]

    def g_weight(self, ex):
        ln = self._lines()
        if not ln or not getattr(ln, "style", None) or ln.style == "none":
            return 2
        return _style_line(ln.style)[1]

    def g_color(self, ex):
        ln = self._lines()
        if not ln or not getattr(ln, "style", None) or ln.style == "none":
            return 0.0
        rgb = ln.color.resolve(self.rng.h.wb().theme) if not ln.color.is_auto else None
        return float(rgb_to_bgr(rgb)) if rgb else 0.0

    def g_colorindex(self, ex):
        ln = self._lines()
        if not ln or ln.color.is_auto:
            return -4105
        return color_index_of(ln.color.resolve(self.rng.h.wb().theme))

    def _apply(self, style: Optional[str] = MISSING, color: Any = MISSING) -> None:  # type: ignore[assignment]
        """Sets the edges (outer edges of the range, and the lines between its cells for the
        inside edges) to ``style`` and/or ``color``."""
        from ..cmd_format import parse_color

        rng = self.rng
        f = rng.first
        new_color = parse_color(color) if color is not MISSING and color is not None else None

        def line(cur: BorderLine) -> BorderLine:
            st = cur.style if style is MISSING else style
            if not st:
                return BorderLine()
            col = new_color if new_color is not None else (cur.color if cur.style else AUTO)
            return BorderLine(st, col)

        for e in self.edges:
            name = _EDGES[e]
            if name == "left":
                rng.h.style(rng.si, [RangeRef(f.r0, f.c0, f.r1, f.c0)], lambda s: replace(s, borders=replace(s.borders, left=line(s.borders.left))))
            elif name == "right":
                rng.h.style(rng.si, [RangeRef(f.r0, f.c1, f.r1, f.c1)], lambda s: replace(s, borders=replace(s.borders, right=line(s.borders.right))))
            elif name == "top":
                rng.h.style(rng.si, [RangeRef(f.r0, f.c0, f.r0, f.c1)], lambda s: replace(s, borders=replace(s.borders, top=line(s.borders.top))))
            elif name == "bottom":
                rng.h.style(rng.si, [RangeRef(f.r1, f.c0, f.r1, f.c1)], lambda s: replace(s, borders=replace(s.borders, bottom=line(s.borders.bottom))))
            elif name == "insideV" and f.width > 1:
                rng.h.style(rng.si, [RangeRef(f.r0, f.c0, f.r1, f.c1 - 1)], lambda s: replace(s, borders=replace(s.borders, right=line(s.borders.right))))
                rng.h.style(rng.si, [RangeRef(f.r0, f.c0 + 1, f.r1, f.c1)], lambda s: replace(s, borders=replace(s.borders, left=line(s.borders.left))))
            elif name == "insideH" and f.height > 1:
                rng.h.style(rng.si, [RangeRef(f.r0, f.c0, f.r1 - 1, f.c1)], lambda s: replace(s, borders=replace(s.borders, bottom=line(s.borders.bottom))))
                rng.h.style(rng.si, [RangeRef(f.r0 + 1, f.c0, f.r1, f.c1)], lambda s: replace(s, borders=replace(s.borders, top=line(s.borders.top))))
            elif name == "diagDown":
                rng.h.style(rng.si, [f], lambda s: replace(s, borders=replace(s.borders, diag_down=line(s.borders.diag_down))))
            elif name == "diagUp":
                rng.h.style(rng.si, [f], lambda s: replace(s, borders=replace(s.borders, diag_up=line(s.borders.diag_up))))

    def l_linestyle(self, ex, v):
        rng = self.rng
        rng.h.require_writable()
        n = _long(v)
        cur_w = 2
        style = _line_style(n, cur_w)
        self._apply(style=style)

    def l_weight(self, ex, v):
        self.rng.h.require_writable()
        w = _long(v)
        ln = self._lines()
        cur = _style_line(ln.style)[0] if ln and getattr(ln, "style", None) else 1
        self._apply(style=_line_style(cur if cur != -4142 else 1, w))

    def l_color(self, ex, v):
        self.rng.h.require_writable()
        ln = self._lines()
        style = ln.style if ln and getattr(ln, "style", None) else "thin"
        self._apply(style=style, color=color_value(v))

    def l_colorindex(self, ex, v):
        self.rng.h.require_writable()
        n = _long(v)
        ln = self._lines()
        style = ln.style if ln and getattr(ln, "style", None) else "thin"
        self._apply(style=style, color=None if n in (-4105, -4142) else palette_hex(n))

    def g_themecolor(self, ex):
        return 0

    def l_themecolor(self, ex, v):
        pass

    def g_tintandshade(self, ex):
        return 0.0

    def l_tintandshade(self, ex, v):
        pass


class Comment(HostObject):
    """A note (Excel's legacy comment) on a cell."""

    TYPE_NAME = "Comment"
    PARAMS = {"text": ("Text", "Start", "Overwrite")}

    def __init__(self, h: XlHost, ws, rc: Tuple[int, int]):
        self.h = h
        self.ws = ws
        self.rc = rc

    def _note(self):
        sh = self.h.wb().sheets[self.h.index_of(self.ws.name)]
        cm = sh.comments.get(self.rc)
        if cm is None:
            raise VbaError(1004, "The note was deleted")
        return cm

    def g_text(self, ex, text=MISSING, start=MISSING, overwrite=MISSING):
        cur = self._note().text
        if _m(text):
            return cur
        self.h.require_writable()
        new = _str(text)
        if not _m(start):
            st = _long(start)
            if _bool(overwrite, True):
                new = cur[:st - 1] + new + cur[st - 1 + len(new):]
            else:
                new = cur[:st - 1] + new + cur[st - 1:]
        self.h.command("review.editNote", {"cell": a1(*self.rc), "text": new}, self.h.index_of(self.ws.name))
        return new

    def g_author(self, ex):
        return self._note().author or ""

    def g_visible(self, ex):
        return bool(getattr(self._note(), "visible", False))

    def l_visible(self, ex, v):
        self.h.command("review.showNote", {"cell": a1(*self.rc), "on": _bool(v)}, self.h.index_of(self.ws.name))

    def g_delete(self, ex):
        self.h.require_writable()
        self.h.command("review.deleteNote", {"cell": a1(*self.rc)}, self.h.index_of(self.ws.name))
        return EMPTY

    def g_parent(self, ex):
        return Range(self.h, self.ws, [RangeRef.of(self.rc)])

    def g_shape(self, ex):
        raise NotSupported("Shapes of notes")


class Characters(HostObject):
    """Part of a cell's text. ReflexGrid formats whole cells, so formatting applies to the cell."""

    TYPE_NAME = "Characters"

    def __init__(self, rng: Range, start: Any, length: Any):
        self.rng = rng
        self.start = start
        self.length = length

    def g_text(self, ex):
        t = V.text(self.rng.g_value(ex))
        st = 1 if _m(self.start) else _long(self.start)
        if _m(self.length):
            return t[st - 1:]
        return t[st - 1:st - 1 + _long(self.length)]

    def g_font(self, ex):
        return Font(self.rng)

    def g_count(self, ex):
        return len(self.g_text(ex))
