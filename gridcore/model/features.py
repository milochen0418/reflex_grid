"""Sheet features beyond cells: row/column info, tables, conditional formatting, data
validation, comments, hyperlinks, AutoFilter, charts, pictures, sparklines, print settings,
protection and defined names."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from ..core.addr import RangeRef
from .style import AUTO, Color, Style


@dataclass(frozen=True)
class LineInfo:
    """A row or column: size in px (``None`` = default), visibility, outline and style."""

    size: Optional[float] = None
    hidden: bool = False
    outline: int = 0
    collapsed: bool = False
    style: Optional[int] = None
    # The size was set by the user (rows): automatic heights follow the content otherwise.
    custom: bool = False

    def is_default(self) -> bool:
        return self == LINE_DEFAULT


LINE_DEFAULT = LineInfo()


# ---------------------------------------------------------------- tables

TOTALS_SUBTOTAL = {"average": 101, "count": 103, "countNums": 102, "max": 104, "min": 105, "sum": 109, "stdDev": 107, "var": 110}


@dataclass
class TableColumn:
    name: str
    totals: str = "none"  # none average count countNums max min sum stdDev var custom
    # Totals row label (first column) or custom formula.
    totals_label: Optional[str] = None
    # Calculated column formula.
    formula: Optional[str] = None


@dataclass
class Table:
    id: int
    name: str
    # Whole table including header and totals rows.
    range: RangeRef
    header_row: bool = True
    totals_row: bool = False
    columns: List[TableColumn] = field(default_factory=list)
    style: str = "TableStyleMedium2"
    banded_rows: bool = True
    banded_cols: bool = False
    first_col: bool = False
    last_col: bool = False
    filter_button: bool = True
    # Column offset within the table -> criterion.
    filter: List[Tuple[int, "FilterCriterion"]] = field(default_factory=list)

    def data_range(self) -> Optional[RangeRef]:
        r0 = self.range.r0 + (1 if self.header_row else 0)
        r1 = self.range.r1 - (1 if self.totals_row else 0)
        if r1 < r0:
            return None
        return RangeRef(r0, self.range.c0, r1, self.range.c1)

    def header_range(self) -> Optional[RangeRef]:
        if not self.header_row:
            return None
        return RangeRef(self.range.r0, self.range.c0, self.range.r0, self.range.c1)

    def totals_range(self) -> Optional[RangeRef]:
        if not self.totals_row:
            return None
        return RangeRef(self.range.r1, self.range.c0, self.range.r1, self.range.c1)

    def column_index(self, name: str) -> Optional[int]:
        n = name.lower()
        for i, c in enumerate(self.columns):
            if c.name.lower() == n:
                return i
        return None


# ---------------------------------------------------------------- conditional formatting


@dataclass
class CfValue:
    """A threshold of a colour scale, data bar or icon set: ``min``, ``max``, ``num``,
    ``percent``, ``percentile`` or ``formula``."""

    kind: str = "min"
    value: Any = None


@dataclass
class CfRule:
    """One conditional-format rule.

    ``type``: cellIs expression containsText notContainsText beginsWith endsWith containsBlanks
    notContainsBlanks containsErrors notContainsErrors duplicateValues uniqueValues top10
    aboveAverage timePeriod colorScale dataBar iconSet."""

    type: str
    operator: str = "between"  # cellIs: between notBetween equal notEqual greaterThan lessThan …
    formulas: List[str] = field(default_factory=list)
    text: str = ""
    style: Optional[Style] = None
    rank: int = 10
    bottom: bool = False
    percent: bool = False
    below: bool = False
    equal_average: bool = False
    std_dev: int = 0
    period: str = "today"
    # colorScale: (threshold, colour) stops
    stops: List[Tuple[CfValue, Color]] = field(default_factory=list)
    # dataBar
    bar_min: CfValue = field(default_factory=lambda: CfValue("min"))
    bar_max: CfValue = field(default_factory=lambda: CfValue("max"))
    bar_color: Color = AUTO
    show_value: bool = True
    # iconSet
    icon_set: str = "3TrafficLights1"
    thresholds: List[CfValue] = field(default_factory=list)
    reverse: bool = False


@dataclass
class CondFormat:
    ranges: List[RangeRef]
    rule: CfRule
    priority: int = 1
    stop_if_true: bool = False


def icon_count(icon_set: str) -> int:
    c = icon_set[:1]
    n = int(c) if c.isdigit() else 3
    return min(max(n, 3), 5)


def default_icon_thresholds(icon_set: str) -> List[CfValue]:
    k = icon_count(icon_set)
    return [CfValue("percent", float(round(i * 100.0 / k))) for i in range(k)]


# ---------------------------------------------------------------- data validation


@dataclass
class Validation:
    ranges: List[RangeRef]
    kind: str = "any"  # any whole decimal list date time textLength custom
    operator: str = "between"
    # Formula 1 (list source: `"a,b,c"` literal list or `=$A$1:$A$5`).
    f1: str = ""
    f2: Optional[str] = None
    allow_blank: bool = True
    in_cell_dropdown: bool = True
    input_title: str = ""
    input_message: str = ""
    show_input: bool = True
    error_title: str = ""
    error_message: str = ""
    error_style: str = "stop"  # stop warning information
    show_error: bool = True


# ---------------------------------------------------------------- comments, links


@dataclass(frozen=True)
class Comment:
    """A note (classic comment) or a threaded comment."""

    author: str
    text: str
    replies: Tuple[Tuple[str, str], ...] = ()
    threaded: bool = False
    resolved: bool = False
    visible: bool = False
    # Where the note's box sits and how big it is, from the note's VML shape: the <x:Anchor>
    # corners relative to the note's cell (columns, pixels, rows, pixels; top-left then bottom-right,
    # so the box moves with the note when rows or columns move) and the shape's width and height in
    # points. None: a note made here, or a file without the shape, gets DEFAULT_NOTE_BOX.
    box: Optional[Tuple[float, ...]] = None


# The box a note made here gets: two columns right of its cell, four rows tall, 108 x 59.25 pt.
DEFAULT_NOTE_BOX: Tuple[float, ...] = (1, 15, 0, 2, 3, 15, 4, 16, 108.0, 59.25)


@dataclass(frozen=True)
class Hyperlink:
    # URL, `mailto:`, another file (`Source.xlsx`), or an in-workbook location like `Sheet2!A1`.
    target: str
    tooltip: Optional[str] = None
    # Outside this workbook (a web page or another file): files keep it as a relationship.
    external: bool = False


# ---------------------------------------------------------------- filter


@dataclass
class FilterCriterion:
    """``values`` (shown display texts + blanks), ``custom`` (one or two comparisons),
    ``top10``, ``aboveAverage``/``belowAverage``, ``fillColor``/``fontColor``."""

    kind: str
    values: List[str] = field(default_factory=list)
    blanks: bool = False
    # custom: [(op, value)], joined with and/or
    conditions: List[Tuple[str, str]] = field(default_factory=list)
    and_: bool = False
    bottom: bool = False
    percent: bool = False
    count: int = 10
    color: Color = AUTO


@dataclass
class AutoFilter:
    range: RangeRef
    # Column offset within `range` -> criterion.
    criteria: List[Tuple[int, FilterCriterion]] = field(default_factory=list)


# ---------------------------------------------------------------- drawing objects


@dataclass
class Anchor:
    """Position of a floating object: top-left cell plus offset (px), size in px."""

    row: int = 0
    col: int = 0
    dx: float = 0.0
    dy: float = 0.0
    width: float = 480.0
    height: float = 288.0
    mode: str = "twoCell"  # twoCell oneCell absolute


@dataclass
class Series:
    # Formula text for the series name (`Sheet1!$B$1`) or a literal.
    name: Optional[str] = None
    categories: Optional[str] = None
    values: str = ""
    color: Optional[Color] = None


CHART_KINDS = [
    "columnClustered", "columnStacked", "columnStacked100", "barClustered", "barStacked", "barStacked100",
    "line", "lineMarkers", "lineStacked", "pie", "doughnut", "area", "areaStacked", "areaStacked100", "scatter", "scatterLines",
    "radar",
]


@dataclass
class Chart:
    id: int
    kind: str = "columnClustered"
    anchor: Anchor = field(default_factory=Anchor)
    title: Optional[str] = None
    series: List[Series] = field(default_factory=list)
    legend: str = "bottom"  # none bottom top left right
    data_labels: bool = False
    gridlines: bool = True
    x_title: Optional[str] = None
    y_title: Optional[str] = None
    # Source range when created from a selection (for Switch Row/Column).
    source: Optional[str] = None
    by_rows: bool = False


@dataclass
class Image:
    id: int
    anchor: Anchor
    data: bytes
    mime: str = "image/png"
    alt: str = ""


@dataclass
class Sparkline:
    row: int
    col: int
    source: str
    kind: str = "line"  # line column stacked(win/loss)
    color: Color = AUTO
    markers: bool = False


# ---------------------------------------------------------------- print & protection


@dataclass
class PrintSettings:
    orientation: str = "portrait"
    paper: str = "Letter"
    # Margins in inches: left, right, top, bottom, header, footer.
    margins: Tuple[float, float, float, float, float, float] = (0.7, 0.7, 0.75, 0.75, 0.3, 0.3)
    print_area: Optional[RangeRef] = None
    title_rows: Optional[Tuple[int, int]] = None
    title_cols: Optional[Tuple[int, int]] = None
    scale: int = 100
    fit_to: Optional[Tuple[int, int]] = None
    gridlines: bool = False
    headings: bool = False
    center_h: bool = False
    center_v: bool = False
    header: str = ""
    footer: str = ""
    row_breaks: List[int] = field(default_factory=list)
    col_breaks: List[int] = field(default_factory=list)


@dataclass
class SheetProtection:
    # Legacy 16-bit password hash (hex), or the modern hash attributes as read from a file.
    password_hash: Optional[str] = None
    hash_attrs: Dict[str, str] = field(default_factory=dict)
    select_locked: bool = True
    select_unlocked: bool = True
    format_cells: bool = False
    format_columns: bool = False
    format_rows: bool = False
    insert_columns: bool = False
    insert_rows: bool = False
    delete_columns: bool = False
    delete_rows: bool = False
    sort: bool = False
    autofilter: bool = False


def legacy_password_hash(password: str) -> str:
    """The 16-bit hash Excel stores for sheet/workbook protection passwords (ECMA-376 legacy)."""
    h = 0
    for ch in reversed(password[:15]):
        h = ((h >> 14) & 0x01) | ((h << 1) & 0x7FFF)
        h ^= ord(ch) & 0xFF
    h = ((h >> 14) & 0x01) | ((h << 1) & 0x7FFF)
    h ^= len(password[:15])
    h ^= 0xCE4B
    return f"{h:04X}"


# ---------------------------------------------------------------- names


@dataclass
class DefinedName:
    name: str
    # Formula text without `=` (`Sheet1!$A$1:$A$10`, `0.07`, `LAMBDA(x,x*2)`).
    formula: str
    # None = workbook scope; otherwise the sheet index.
    scope: Optional[int] = None
    comment: str = ""
    hidden: bool = False
