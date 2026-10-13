"""How cells look as text: number formats applied, for the grid, clipboard, CSV and agents."""

from __future__ import annotations

import math
from typing import Dict, List, Optional, Tuple

from ..core.addr import RangeRef
from ..core.date import datetime_from_serial
from ..core.parse import parse_input
from ..core.value import Array, CellError, Value
from ..numfmt import Formatted, format_value, parse as parse_fmt


def fmt(v: Value, code: str, wb, width_chars: int = 11) -> Formatted:
    return format_value(v, parse_fmt(code), wb.date_system, width_chars)


def edit_text(wb, sheet, rc: Tuple[int, int]) -> str:
    """What the cell editor and the formula bar start from: the formula, or the typed value; a
    constant date/time shows as a date that reads back as the same serial."""
    cell = sheet.cell(rc)
    if cell is None:
        return ""
    raw = cell.input_text()
    v = cell.value
    if cell.formula is not None or not isinstance(v, float):
        if cell.formula is None and isinstance(v, str):
            p = parse_input(v, wb.date_system)
            if not isinstance(p.value, str) or v.startswith("'"):
                # Text that would read as something else keeps it as text with an apostrophe.
                return "'" + v
        return raw
    code = wb.styles.get(cell.style).num_fmt
    nf = parse_fmt(code)
    if nf.is_percent() and math.isfinite(v):
        from ..core.value import number_to_text

        return number_to_text(v * 100.0) + "%"
    if not nf.is_date():
        return raw

    def reads_back(s: str) -> bool:
        p = parse_input(s, wb.date_system).value
        return isinstance(p, float) and abs(p - v) < 1e-9

    shown = fmt(v, code, wb).text
    if reads_back(shown):
        return shown
    dt = datetime_from_serial(wb.date_system, v)
    if dt is None:
        return raw
    t = f"{dt.hour:02d}:{dt.minute:02d}:{dt.second:02d}"
    d = f"{dt.year:04d}-{dt.month:02d}-{dt.day:02d}"
    iso = t if v < 1 else (d if v == int(v) else f"{d} {t}")
    return iso if reads_back(iso) else raw


def value_text(wb, sheet, rc: Tuple[int, int]) -> str:
    """A cell's value with its number format applied, independent of column width."""
    st = wb.styles.get(sheet.style_id(rc))
    f = fmt(sheet.value(rc), st.num_fmt, wb, 11)
    if f.fill is not None:
        ch, pos = f.fill
        return f.text[:pos] + " " + f.text[pos:]
    return f.text


def cell_text(wb, sheet, rc: Tuple[int, int]) -> str:
    if sheet.show_formulas:
        cell = sheet.cell(rc)
        if cell is not None and cell.formula is not None:
            return "=" + cell.formula.text
    return value_text(wb, sheet, rc)


def range_text(wb, sheet, rng: RangeRef) -> str:
    """Tab-separated text of a range, as the system clipboard gets it."""
    if rng.count() > 2_000_000:
        rng = RangeRef(rng.r0, rng.c0, min(rng.r1, rng.r0 + 1999), min(rng.c1, rng.c0 + 999))
    lines = []
    for r in range(rng.r0, rng.r1 + 1):
        if sheet.is_row_hidden(r):
            continue
        parts = []
        for c in range(rng.c0, rng.c1 + 1):
            t = cell_text(wb, sheet, (r, c))
            if "\t" in t or "\n" in t or '"' in t:
                t = '"' + t.replace('"', '""') + '"'
            parts.append(t)
        lines.append("\t".join(parts))
    return "\n".join(lines) + ("\n" if lines else "")


def stats(sheet, ranges: List[RangeRef]) -> Dict[str, Optional[float]]:
    """Status-bar statistics for a selection."""
    count = numeric = 0
    total = 0.0
    mn, mx = math.inf, -math.inf
    seen = set()
    used = sheet.used_range()
    for r in ranges:
        if used is None:
            break
        inter = r.intersection(used)
        if inter is None:
            continue

        def visit(rc, v):
            nonlocal count, numeric, total, mn, mx
            if rc in seen or sheet.is_row_hidden(rc[0]) or v is None:
                return
            seen.add(rc)
            count += 1
            if isinstance(v, float):
                numeric += 1
                total += v
                mn = min(mn, v)
                mx = max(mx, v)

        for rc, cell in sheet.cells.iter_range(inter):
            visit(rc, cell.value)
        if sheet.spill:
            for rc, v in sheet.spill.items():
                if inter.contains(rc):
                    visit(rc, v)
    return {
        "count": float(count),
        "numeric": float(numeric),
        "sum": total if numeric else 0.0,
        "min": mn if numeric else 0.0,
        "max": mx if numeric else 0.0,
        "average": total / numeric if numeric else None,
    }
