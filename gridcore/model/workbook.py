"""Workbooks."""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

from ..core.date import DateSystem
from ..formula import Call, Expr, ParseError, parse, print_formula
from .features import DefinedName
from .sheet import Sheet
from .style import StyleTable, Theme


@dataclass
class CalcSettings:
    mode: str = "auto"  # auto autoNoTable manual
    iterative: bool = False
    max_iterations: int = 100
    max_change: float = 0.001
    precision_as_displayed: bool = False


@dataclass
class DocProps:
    title: str = ""
    subject: str = ""
    author: str = ""
    company: str = ""
    keywords: str = ""
    description: str = ""


INVALID_SHEET_CHARS = set(':\\/?*[]')


class Workbook:
    def __init__(self, empty: bool = False):
        self.sheets: List[Sheet] = [] if empty else [Sheet("Sheet1")]
        self.names: List[DefinedName] = []
        self.styles = StyleTable()
        self.theme = Theme()
        self.date_system = DateSystem.D1900
        self.calc = CalcSettings()
        self.active_sheet = 0
        self.props = DocProps()
        self.protected_structure = False
        self.structure_password: Optional[str] = None
        # Named cell styles (Cell Styles gallery): name -> style.
        self.cell_styles: List[Tuple[str, object]] = []
        # Field separator of the CSV this workbook was read from.
        self.csv_delimiter: Optional[str] = None
        # The VBA project (macros) and the workbook's name in it (``ThisWorkbook``); a project is
        # immutable, so snapshots share it (see :mod:`gridcore.model.vba`).
        self.vba = None
        self.code_name = ""
        # Sheets this workbook object may mutate in place (copy-on-write; ids of sheet objects).
        self._owned: set = set(id(s) for s in self.sheets)

    # ------------------------------------------------------------ copy-on-write

    def clone(self) -> "Workbook":
        """A cheap copy: sheets are shared until :meth:`sheet_mut` copies them."""
        w = Workbook.__new__(Workbook)
        w.__dict__.update(self.__dict__)
        w.sheets = list(self.sheets)
        w.names = copy.deepcopy(self.names)
        w.calc = copy.deepcopy(self.calc)
        w.props = copy.deepcopy(self.props)
        w.cell_styles = list(self.cell_styles)
        w._owned = set()
        self._owned = set()
        return w

    def sheet(self, i: int) -> Optional[Sheet]:
        if 0 <= i < len(self.sheets):
            return self.sheets[i]
        return None

    def sheet_mut(self, i: int) -> Optional[Sheet]:
        """The sheet at ``i``, copied first if it is shared with another workbook snapshot."""
        if not 0 <= i < len(self.sheets):
            return None
        s = self.sheets[i]
        if id(s) not in self._owned:
            s = s.clone()
            self.sheets[i] = s
            self._owned.add(id(s))
        return s

    def add_sheet(self, sheet: Sheet, index: Optional[int] = None) -> int:
        if index is None:
            index = len(self.sheets)
        self.sheets.insert(index, sheet)
        self._owned.add(id(sheet))
        return index

    # ------------------------------------------------------------ lookups

    def active(self) -> Optional[Sheet]:
        return self.sheet(self.active_sheet)

    def sheet_index(self, name: str) -> Optional[int]:
        n = name.lower()
        for i, s in enumerate(self.sheets):
            if s.name.lower() == n:
                return i
        return None

    def next_sheet_name(self) -> str:
        n = 1
        while self.sheet_index(f"Sheet{n}") is not None:
            n += 1
        return f"Sheet{n}"

    def check_sheet_name(self, name: str, except_: Optional[int] = None) -> Optional[str]:
        """Validates a sheet name like Excel; returns an error message or ``None``."""
        if not name:
            return "A sheet name can't be blank."
        if len(name) > 31:
            return "A sheet name can't be longer than 31 characters."
        if any(c in INVALID_SHEET_CHARS for c in name):
            return "A sheet name can't contain any of these characters: : \\ / ? * [ ]"
        if name.startswith("'") or name.endswith("'"):
            return "A sheet name can't begin or end with an apostrophe."
        if name.lower() == "history":
            return '"History" is a reserved name.'
        i = self.sheet_index(name)
        if i is not None and i != except_:
            return "That name is already taken. Try a different one."
        return None

    def name(self, name: str, sheet: int) -> Optional[DefinedName]:
        """A defined name visible from ``sheet``: sheet scope first, then workbook scope."""
        n = name.lower()
        for d in self.names:
            if d.scope == sheet and d.name.lower() == n:
                return d
        for d in self.names:
            if d.scope is None and d.name.lower() == n:
                return d
        return None

    def table(self, name: str) -> Optional[Tuple[int, int]]:
        """Table by name across sheets: (sheet index, table index)."""
        n = name.lower()
        for si, s in enumerate(self.sheets):
            for ti, t in enumerate(s.tables):
                if t.name.lower() == n:
                    return si, ti
        return None

    def next_object_id(self) -> int:
        m = 0
        for s in self.sheets:
            for o in s.charts:
                m = max(m, o.id)
            for o in s.images:
                m = max(m, o.id)
            for o in s.tables:
                m = max(m, o.id)
        return m + 1

    def lambda_names(self) -> List[str]:
        out = []
        for d in self.names:
            f = d.formula.lstrip("=").lstrip()
            if f[:7].upper() == "LAMBDA(":
                out.append(d.name)
        return out
