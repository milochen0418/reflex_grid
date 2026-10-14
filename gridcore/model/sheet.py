"""Worksheets."""

from __future__ import annotations

import bisect
import copy
from typing import Dict, Iterator, List, Optional, Tuple

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef
from ..core.value import Value
from .cell import Cell
from .features import (
    LINE_DEFAULT,
    AutoFilter,
    Chart,
    Comment,
    CondFormat,
    Hyperlink,
    Image,
    LineInfo,
    PrintSettings,
    SheetProtection,
    Sparkline,
    Table,
    Validation,
)
from .store import CellStore

# Default column width and row height in px (Calibri 11 at 96 dpi: 8.43 characters, 15 pt).
DEFAULT_COL_WIDTH = 64.0
DEFAULT_ROW_HEIGHT = 20.0


class Sheet:
    def __init__(self, name: str):
        self.name = name
        self.cells = CellStore()
        # Values spilled by dynamic-array formulas (maintained by calc), keyed by cell.
        self.spill: Dict[Tuple[int, int], Value] = {}
        # Anchor -> spill range.
        self.spill_ranges: Dict[Tuple[int, int], RangeRef] = {}
        self.cols: Dict[int, LineInfo] = {}
        self.rows: Dict[int, LineInfo] = {}
        self.default_col_width = DEFAULT_COL_WIDTH
        self.default_row_height = DEFAULT_ROW_HEIGHT
        self.merges: List[RangeRef] = []
        # Frozen rows and columns (top/left).
        self.freeze: Optional[Tuple[int, int]] = None
        self.tab_color = None
        self.visibility = "visible"  # visible hidden veryHidden
        self.show_gridlines = True
        self.show_headings = True
        self.show_formulas = False
        self.show_zeros = True
        self.right_to_left = False
        self.zoom = 100
        self.cond_formats: List[CondFormat] = []
        self.validations: List[Validation] = []
        self.tables: List[Table] = []
        self.autofilter: Optional[AutoFilter] = None
        self.comments: Dict[Tuple[int, int], Comment] = {}
        self.hyperlinks: Dict[Tuple[int, int], Hyperlink] = {}
        self.charts: List[Chart] = []
        self.images: List[Image] = []
        self.sparklines: List[Sparkline] = []
        self.protection: Optional[SheetProtection] = None
        self.print = PrintSettings()
        # Last selection and scroll position, saved with the file.
        self.view_active: Tuple[int, int] = (0, 0)
        self.view_top_left: Tuple[int, int] = (0, 0)
        # Unmodelled worksheet XML kept from a file (passed through on save).
        self.kept_xml: Dict[str, str] = {}
        # The sheet's name in VBA (its document module, e.g. ``Sheet1``); "" until it has one.
        self.code_name = ""
        self._row_index: Optional[_LineIndex] = None
        self._col_index: Optional[_LineIndex] = None

    # ------------------------------------------------------------ copy-on-write

    def clone(self) -> "Sheet":
        s = Sheet.__new__(Sheet)
        d = self.__dict__
        s.__dict__.update(d)
        s.cells = self.cells.clone()
        s.spill = dict(self.spill)
        s.spill_ranges = dict(self.spill_ranges)
        s.cols = dict(self.cols)
        s.rows = dict(self.rows)
        s.merges = list(self.merges)
        s.comments = dict(self.comments)
        s.hyperlinks = dict(self.hyperlinks)
        for k in ("cond_formats", "validations", "tables", "autofilter", "charts", "images", "sparklines", "protection", "print", "kept_xml"):
            setattr(s, k, copy.deepcopy(d[k]))
        s._row_index = None
        s._col_index = None
        return s

    # ------------------------------------------------------------ cells

    def cell(self, rc: Tuple[int, int]) -> Optional[Cell]:
        return self.cells.get(rc)

    def value(self, rc: Tuple[int, int]) -> Value:
        """Displayed value of a cell: its own value, or a spilled value."""
        cell = self.cells.get(rc)
        if cell is not None and (cell.value is not None or cell.formula is not None):
            return cell.value
        return self.spill.get(rc)

    def style_id(self, rc: Tuple[int, int]) -> int:
        cell = self.cells.get(rc)
        if cell is not None:
            return cell.style
        ri = self.rows.get(rc[0])
        if ri is not None and ri.style is not None:
            return ri.style
        ci = self.cols.get(rc[1])
        if ci is not None and ci.style is not None:
            return ci.style
        return 0

    def set_cell(self, rc: Tuple[int, int], cell: Cell) -> None:
        self.cells.set(rc, cell)

    def set_value(self, rc: Tuple[int, int], v: Value) -> None:
        self.cells.set(rc, Cell(v, None, self.style_id(rc)))

    def set_style(self, rc: Tuple[int, int], sid: int) -> None:
        cell = self.cells.get(rc)
        if cell is None:
            self.cells.set(rc, Cell(None, None, sid))
        elif cell.style != sid:
            self.cells.set(rc, cell.with_style(sid))

    def remove_cell(self, rc: Tuple[int, int]) -> Optional[Cell]:
        return self.cells.remove(rc)

    def input_text(self, rc: Tuple[int, int]) -> str:
        cell = self.cells.get(rc)
        return cell.input_text() if cell is not None else ""

    def used_range(self) -> Optional[RangeRef]:
        r = self.cells.used_range()
        for sr in self.spill_ranges.values():
            r = sr if r is None else r.union(sr)
        return r

    # ------------------------------------------------------------ geometry

    def _invalidate_geometry(self) -> None:
        self._row_index = None
        self._col_index = None

    def set_row(self, r: int, info: LineInfo) -> None:
        if info.is_default():
            self.rows.pop(r, None)
        else:
            self.rows[r] = info
        self._row_index = None

    def set_col(self, c: int, info: LineInfo) -> None:
        if info.is_default():
            self.cols.pop(c, None)
        else:
            self.cols[c] = info
        self._col_index = None

    def row_info(self, r: int) -> LineInfo:
        return self.rows.get(r, LINE_DEFAULT)

    def col_info(self, c: int) -> LineInfo:
        return self.cols.get(c, LINE_DEFAULT)

    def col_width(self, col: int) -> float:
        i = self.cols.get(col)
        if i is None:
            return self.default_col_width
        if i.hidden:
            return 0.0
        return i.size if i.size is not None else self.default_col_width

    def row_height(self, row: int) -> float:
        i = self.rows.get(row)
        if i is None:
            return self.default_row_height
        if i.hidden:
            return 0.0
        return i.size if i.size is not None else self.default_row_height

    def is_row_hidden(self, row: int) -> bool:
        i = self.rows.get(row)
        return i is not None and i.hidden

    def is_col_hidden(self, col: int) -> bool:
        i = self.cols.get(col)
        return i is not None and i.hidden

    def is_row_filtered(self, row: int) -> bool:
        """Whether ``row`` is hidden by an applied filter (as opposed to hidden by hand)."""
        if not self.is_row_hidden(row):
            return False
        af = self.autofilter
        if af is not None and af.criteria and af.range.r0 < row <= af.range.r1:
            return True
        for t in self.tables:
            if t.filter:
                d = t.data_range()
                if d is not None and d.r0 <= row <= d.r1:
                    return True
        return False

    def row_index(self) -> "_LineIndex":
        if self._row_index is None:
            self._row_index = _LineIndex(self.rows, self.default_row_height, MAX_ROWS)
        return self._row_index

    def col_index(self) -> "_LineIndex":
        if self._col_index is None:
            self._col_index = _LineIndex(self.cols, self.default_col_width, MAX_COLS)
        return self._col_index

    def row_top(self, row: int) -> float:
        return self.row_index().start(row)

    def col_left(self, col: int) -> float:
        return self.col_index().start(col)

    def merge_at(self, rc: Tuple[int, int]) -> Optional[RangeRef]:
        for m in self.merges:
            if m.contains(rc):
                return m
        return None

    def table_at(self, rc: Tuple[int, int]) -> Optional[Table]:
        for t in self.tables:
            if t.range.contains(rc):
                return t
        return None

    def is_protected(self) -> bool:
        return self.protection is not None

    def has_filter_button(self, rc: Tuple[int, int]) -> bool:
        af = self.autofilter
        if af is not None and af.range.r0 == rc[0] and af.range.contains(rc):
            return True
        return any(t.filter_button and t.header_row and t.range.r0 == rc[0] and t.range.contains(rc) for t in self.tables)


class _LineIndex:
    """Prefix sums over the custom-sized lines of an axis: O(log n) position <-> line lookups."""

    def __init__(self, custom: Dict[int, LineInfo], default: float, mx: int):
        self.default = default
        self.max = mx
        self.lines: List[int] = []
        self.pos: List[float] = []
        self.size: List[float] = []
        pos = 0.0
        prev = 0
        for line in sorted(custom):
            info = custom[line]
            pos += (line - prev) * default
            size = 0.0 if info.hidden else (info.size if info.size is not None else default)
            self.lines.append(line)
            self.pos.append(pos)
            self.size.append(size)
            pos += size
            prev = line + 1

    def start(self, line: int) -> float:
        i = bisect.bisect_left(self.lines, line)
        if i < len(self.lines) and self.lines[i] == line:
            return self.pos[i]
        if i == 0:
            return line * self.default
        l, p, s = self.lines[i - 1], self.pos[i - 1], self.size[i - 1]
        return p + s + (line - l - 1) * self.default

    def size_of(self, line: int) -> float:
        i = bisect.bisect_left(self.lines, line)
        if i < len(self.lines) and self.lines[i] == line:
            return self.size[i]
        return self.default

    def at(self, pos: float) -> int:
        """The line containing ``pos``."""
        if pos <= 0:
            return 0
        i = bisect.bisect_right(self.pos, pos)
        if i > 0:
            l, p, s = self.lines[i - 1], self.pos[i - 1], self.size[i - 1]
            if pos < p + s:
                return min(l, self.max - 1)
            base_line, base_pos = l + 1, p + s
        else:
            base_line, base_pos = 0, 0.0
        if self.default <= 0:
            return min(base_line, self.max - 1)
        r = base_line + int((pos - base_pos) // self.default)
        return min(r, self.max - 1)
