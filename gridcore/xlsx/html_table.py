"""Web pages that hold a table. A web app's "Export to Excel" is often an HTML table saved as
``.xls`` (GridCraft #3166); Excel opens its table as a sheet, after warning that the format and the
extension don't match. Here too: the cells read as typed input does (through the CSV reader, so
numbers, dates and text read exactly as they do there), header cells (``<th>``) are bold and
centred, and the sheet takes the name the page gives it (``<x:Name>``) or the file's."""

from __future__ import annotations

import re
from dataclasses import replace
from html.parser import HTMLParser
from typing import Dict, List, Optional, Tuple

from ..core.addr import MAX_COLS, MAX_ROWS
from ..model.workbook import Workbook
from .csv_io import read_csv

_STARTS = ("<!doctype html", "<html", "<table", "<head", "<body", "<meta")
_CHARSET = re.compile(rb"""charset\s*=\s*["']?\s*([A-Za-z0-9_.:\-]+)""", re.I)
_SHEET_NAME = re.compile(r"<x:Name>(.*?)</x:Name>", re.I | re.S)
_SPACES = re.compile(r"[ \t\r\n\f\v]+")


def decode_page(data: bytes) -> str:
    """A page's text: BOMs, else UTF-8, else the charset its ``<meta>`` names (``big5`` from a
    bank's export), else Windows-1252, as a CSV is read."""
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8", errors="replace")
    if data.startswith(b"\xff\xfe"):
        return data[2:].decode("utf-16-le", errors="replace")
    if data.startswith(b"\xfe\xff"):
        return data[2:].decode("utf-16-be", errors="replace")
    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        pass
    m = _CHARSET.search(data[:4096])
    if m is not None:
        try:
            return data.decode(m.group(1).decode("ascii"), errors="replace")
        except (LookupError, ValueError):
            pass
    return data.decode("cp1252", errors="replace")


def looks_like_page(text: str) -> bool:
    """Whether a text file is a web page: it starts with one of a page's tags, after a BOM, spaces
    and comments an exporter writes before the page."""
    rest = text.lstrip("﻿").lstrip()
    while rest.startswith("<!--"):
        end = rest.find("-->")
        if end < 0:
            break
        rest = rest[end + 3 :].lstrip()
    head = rest[:16].lower()
    return any(head.startswith(t) for t in _STARTS)


class _Cell:
    __slots__ = ("parts", "num", "header", "colspan", "rowspan")

    def __init__(self, num: Optional[float], header: bool, colspan: int, rowspan: int):
        self.parts: List[str] = []
        self.num = num
        self.header = header
        self.colspan = colspan
        self.rowspan = rowspan

    def text(self) -> str:
        # Runs of spaces and line breaks show as one space, as a browser shows them; <br> breaks
        # the line. Non-breaking spaces stay, as in Excel.
        lines = "".join(self.parts).split("\x00")
        return "\n".join(_SPACES.sub(" ", line).strip(" ") for line in lines).strip("\n")


class _Tables(HTMLParser):
    """The cells of a page's tables, row by row, spans laid out: ``rows[r][c]`` is a ``_Cell`` or
    ``None``. A table inside a cell adds its text to that cell."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.rows: List[Dict[int, _Cell]] = []
        # A cell spanning rows keeps its columns of the rows below: column -> last row it holds.
        self.held: Dict[int, int] = {}
        self.depth = 0
        self.skip = 0
        self.cell: Optional[_Cell] = None
        self.row: Optional[int] = None
        self.col = 0

    def _num(self, attrs) -> Optional[float]:
        for k, v in attrs:
            if k == "x:num" and v:
                try:
                    n = float(v)
                except ValueError:
                    return None
                return n if n == n and abs(n) != float("inf") else None
        return None

    @staticmethod
    def _span(attrs, key: str, limit: int) -> int:
        for k, v in attrs:
            if k == key and v:
                try:
                    return min(max(int(v.strip()), 1), limit)
                except ValueError:
                    return 1
        return 1

    def _end_cell(self) -> None:
        c = self.cell
        self.cell = None
        if c is None or self.row is None or self.row >= MAX_ROWS:
            return
        col = self.col
        while self.held.get(col, -1) >= self.row:
            col += 1
        if col >= MAX_COLS:
            return
        self.rows[self.row][col] = c
        if c.rowspan > 1:
            for dc in range(min(c.colspan, MAX_COLS - col)):
                self.held[col + dc] = max(self.held.get(col + dc, -1), self.row + c.rowspan - 1)
        self.col = col + c.colspan

    def _start_row(self) -> None:
        self._end_cell()
        self.row = len(self.rows)
        self.rows.append({})
        self.col = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self.skip += 1
            return
        if tag == "table":
            self.depth += 1
            return
        if self.depth == 1:
            if tag == "tr":
                self._start_row()
            elif tag in ("td", "th"):
                self._end_cell()
                if self.row is None:
                    self._start_row()
                self.cell = _Cell(self._num(attrs), tag == "th", self._span(attrs, "colspan", 1000), self._span(attrs, "rowspan", MAX_ROWS))
                return
        if self.cell is not None and tag in ("br", "p", "div", "li", "tr"):
            self.cell.parts.append("\x00")

    def handle_startendtag(self, tag, attrs):
        if self.cell is not None and tag == "br":
            self.cell.parts.append("\x00")
        elif tag not in ("br",):
            self.handle_starttag(tag, attrs)

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self.skip = max(self.skip - 1, 0)
            return
        if tag == "table":
            if self.depth == 1:
                self._end_cell()
                self.row = None
            self.depth = max(self.depth - 1, 0)
            return
        if self.depth == 1:
            if tag in ("td", "th"):
                self._end_cell()
            elif tag == "tr":
                self._end_cell()

    def handle_data(self, data):
        if self.cell is not None and not self.skip:
            self.cell.parts.append(data)


def read_page_table(data: bytes, name: str) -> Optional[Tuple[Workbook, List[str]]]:
    """The tables of a web page as a sheet, with a warning when its name says it is something else;
    ``None`` when the bytes aren't such a page (or it has no table cells)."""
    text = decode_page(data)
    if not looks_like_page(text):
        return None
    p = _Tables()
    try:
        p.feed(text)
        p.close()
    except Exception:  # never crash on a broken page: it is read as text, as before
        return None
    rows = p.rows[:MAX_ROWS]
    if not any(rows):
        return None
    lines = []
    budget = 5_000_000  # fields: spans can't blow a small page up into a huge sheet
    for r in rows:
        fields = []
        width = max(r) + 1 if r else 0
        budget -= width
        if budget < 0:
            break
        for c in range(width):
            cell = r.get(c)
            if cell is None:
                fields.append("")
                continue
            # The number the page recorded (`x:num`, Excel's own export) wins over the text shown.
            if cell.num is not None:
                field = repr(cell.num) if cell.num != int(cell.num) or abs(cell.num) >= 1e15 else str(int(cell.num))
            else:
                field = cell.text()
            if any(ch in field for ch in '\t\n\r"'):
                field = '"' + field.replace('"', '""') + '"'
            fields.append(field)
        lines.append("\t".join(fields))
    wb = read_csv(("\n".join(lines) + "\n").encode("utf-8"), "\t")
    wb.csv_delimiter = None
    sh = wb.sheets[0]
    bold: Dict[int, int] = {}
    for ri, r in enumerate(rows):
        for ci, cell in r.items():
            if cell.header and cell.text().strip():
                base = sh.style_id((ri, ci))
                sid = bold.get(base)
                if sid is None:
                    sid = wb.styles.derive(base, lambda st: replace(st, font=replace(st.font, bold=True), align=replace(st.align, h="center")))
                    bold[base] = sid
                existing = sh.cell((ri, ci))
                if existing is not None:
                    sh.cells.set((ri, ci), existing.with_style(sid))
    m = _SHEET_NAME.search(text)
    title = _SPACES.sub(" ", m.group(1)).strip() if m else ""
    if not title:
        title = name.rsplit("/", 1)[-1].rsplit(".", 1)[0]
    if title and wb.check_sheet_name(title[:31]) is None:
        sh.name = title[:31]
    warnings = []
    if not name.lower().endswith((".htm", ".html")):
        warnings.append(f"`{name}` is a web page (HTML), not an Excel workbook: its table was opened as a sheet.")
    return wb, warnings
