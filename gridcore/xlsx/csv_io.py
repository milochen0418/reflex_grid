"""CSV/TSV import and export (RFC 4180 with the usual real-world leniency)."""

from __future__ import annotations

from typing import Callable, Dict, List, Optional, Set

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef
from ..core.date import DateSystem
from ..core.parse import parse_input
from ..model.cell import Cell
from ..model.sheet import Sheet
from ..model.workbook import Workbook


def decode_text(data: bytes) -> str:
    """CSV bytes: UTF-8/UTF-16 BOMs, else UTF-8, else Windows-1252."""
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8", errors="replace")
    if data.startswith(b"\xff\xfe"):
        return data[2:].decode("utf-16-le", errors="replace")
    if data.startswith(b"\xfe\xff"):
        return data[2:].decode("utf-16-be", errors="replace")
    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace")


def sniff_delimiter(text: str, quote: str = '"') -> str:
    best = (",", 0, 0)
    for d in (",", ";", "\t"):
        counts: List[int] = []
        n = 0
        in_q = False
        blank = True
        for ch in text[: 64 * 1024]:
            if not in_q and ch == "\n":
                if not blank:
                    counts.append(n)
                n = 0
                blank = True
                if len(counts) >= 20:
                    break
                continue
            if ch != "\r":
                blank = False
            if ch == quote:
                in_q = not in_q
            elif not in_q and ch == d:
                n += 1
        if not blank:
            counts.append(n)
        mn = min(counts) if counts else 0
        total = sum(counts)
        if mn > 0 and (mn, total) > (best[1], best[2]):
            best = (d, mn, total)
    return best[0]


def parse_records(text: str, delim: str, quote: str = '"'):
    field: List[str] = []
    record: List[str] = []
    in_q = False
    started_quoted = False
    i = 0
    n = len(text)
    while i < n:
        c = text[i]
        if in_q:
            if c == quote:
                if i + 1 < n and text[i + 1] == quote:
                    field.append(quote)
                    i += 2
                    continue
                in_q = False
            else:
                field.append(c)
            i += 1
            continue
        if c == quote and not field and not started_quoted:
            in_q = True
            started_quoted = True
        elif c == delim:
            record.append("".join(field))
            field = []
            started_quoted = False
        elif c in "\r\n":
            if c == "\r" and i + 1 < n and text[i + 1] == "\n":
                i += 1
            record.append("".join(field))
            field = []
            started_quoted = False
            yield record
            record = []
        else:
            field.append(c)
        i += 1
    if field or record or started_quoted:
        record.append("".join(field))
        yield record


def read_csv(data: bytes, delimiter: Optional[str] = None, parse_values: bool = True, sys: DateSystem = DateSystem.D1900) -> Workbook:
    text = decode_text(data)
    delim = delimiter or sniff_delimiter(text)
    wb = Workbook(empty=True)
    wb.date_system = sys
    wb.csv_delimiter = delim
    sheet = Sheet("Sheet1")
    fmt_cache = {}
    for row, rec in enumerate(parse_records(text, delim)):
        if row >= MAX_ROWS:
            break
        for col, f in enumerate(rec):
            if col >= MAX_COLS:
                break
            if f == "":
                continue
            if parse_values:
                p = parse_input(f, sys)
                style = 0
                if p.format is not None:
                    style = fmt_cache.get(p.format)
                    if style is None:
                        from dataclasses import replace

                        style = wb.styles.derive(0, lambda s, code=p.format: replace(s, num_fmt=code))
                        fmt_cache[p.format] = style
                sheet.cells.set((row, col), Cell(p.value, None, style))
            else:
                sheet.cells.set((row, col), Cell(f))
    wb.add_sheet(sheet)
    return wb


# A CSV is capped: two cells far apart (A1 and XFD1048576) would make gigabytes of commas.
MAX_CSV_BYTES = 512 * 1024 * 1024


class CsvTooLarge(ValueError):
    """The CSV would be bigger than :data:`MAX_CSV_BYTES`."""


def write_csv(sheet: Sheet, cell_text: Callable[[tuple], str], delimiter: str = ",", limit: int = MAX_CSV_BYTES) -> bytes:
    """The sheet from A1 to its last used cell as CSV, as Excel saves it (CRLF line ends, UTF-8
    with BOM so Excel reads Unicode). Only cells that hold something are formatted, so a sparse
    sheet is quick; raises :class:`CsvTooLarge` past ``limit`` (before anything is built when the
    delimiters alone would pass it)."""
    u = sheet.used_range()
    if u is None:
        return b""
    rows, cols = u.r1 + 1, u.c1 + 1

    def too_large(size: int) -> CsvTooLarge:
        return CsvTooLarge(
            f"the CSV would be at least {size / 2 ** 20:,.0f} MiB ({rows:,} rows by {cols:,} columns), over the limit of "
            f"{limit // 2 ** 20:,} MiB; clear the cells far from the data, or copy a smaller range"
        )

    least = rows * (cols + 1)  # every row has its delimiters and a CRLF
    if least > limit:
        raise too_large(least)
    filled: Dict[int, Set[int]] = {}
    for (r, c), cell in sheet.cells.iter_range(RangeRef(0, 0, u.r1, u.c1)):
        if cell.value is not None or cell.formula is not None:
            filled.setdefault(r, set()).add(c)
    for r, c in sheet.spill:
        filled.setdefault(r, set()).add(c)
    blank = delimiter * (cols - 1)
    lines = []
    size = 0
    for r in range(rows):
        cs = filled.get(r)
        if cs:
            fields = [""] * cols
            for c in cs:
                t = cell_text((r, c))
                if delimiter in t or '"' in t or "\n" in t or "\r" in t:
                    t = '"' + t.replace('"', '""') + '"'
                fields[c] = t
            line = delimiter.join(fields)
        else:
            line = blank
        size += len(line) + 2
        if size > limit:
            raise too_large(size)
        lines.append(line)
    return ("﻿" + "\r\n".join(lines) + "\r\n").encode("utf-8")
