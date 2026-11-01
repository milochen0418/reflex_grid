"""What a workbook offers to documents that show part of it (CoDoc in MD, say): its UserForms, and
cells, which they only read.

Cells are named by a *reference* that goes on meaning the same thing while the workbook grows:

    Orders          an Excel table (it grows by itself), or a defined name
    Sheet1!A1:D     from A1 down to the last row with something in columns A to D
    Sheet1!A1:D20   these cells and no others

:func:`resolve` says which cells a reference stands for now, :func:`suggest` picks the reference
for a selection, the way a person would, and :func:`table_view` draws the cells."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

from ..core.addr import RangeRef, a1, col_to_letters, letters_to_col
from ..formula.printer import quote_sheet
from ..model.vba import FORM
from .cmd_edit import resolve_reference
from .display import value_text
from .html_export import table_markup
from .selection import current_region
from .session import EngineError, Session, command

# `A1:D`: a start cell and a last column, without a last row.
_OPEN_END = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*):\$?([A-Za-z]{1,3})$")
# What a table on a page shows: more of it stays in the workbook.
TABLE_ROWS = 1000
# What a preview to choose cells from shows.
PREVIEW_ROWS, PREVIEW_COLS = 200, 40


def _sheet_of(ref: str) -> Tuple[Optional[str], str]:
    """``'Q1 Sales'!A1:D`` → ``("Q1 Sales", "A1:D")``."""
    head, bang, tail = ref.rpartition("!")
    if not bang:
        return None, ref
    if len(head) >= 2 and head[0] == head[-1] == "'":
        head = head[1:-1].replace("''", "'")
    return head, tail


def _first_shown(wb) -> int:
    return next((i for i, s in enumerate(wb.sheets) if s.visibility == "visible"), 0)


def _filled(sh, rc: Tuple[int, int]) -> bool:
    """Whether a cell shows something: formulas waiting for data, which give "", don't."""
    v = sh.value(rc)
    return v is not None and v != ""


def last_row(sh, r0: int, c0: int, c1: int) -> int:
    """The last row from ``r0`` on with something to show in columns ``c0``..``c1``; ``r0 - 1``
    when there is none."""
    used = sh.used_range()
    if used is None or used.r1 < r0:
        return r0 - 1
    last = r0 - 1
    for r in reversed(sh.cells.rows_in(r0, used.r1)):
        if any(c0 <= c <= c1 and _filled(sh, (r, c)) for c in sh.cells.row(r)):
            last = r
            break
    # Dynamic arrays: their spilled cells aren't stored cells.
    for sr in sh.spill_ranges.values():
        if sr.c0 <= c1 and c0 <= sr.c1 and sr.r1 > last:
            last = sr.r1
    return last


def _first_row(sh, c0: int, c1: int) -> int:
    used = sh.used_range()
    if used is None:
        return 0
    for r in sh.cells.rows_in(used.r0, used.r1):
        if any(c0 <= c <= c1 and _filled(sh, (r, c)) for c in sh.cells.row(r)):
            return r
    return used.r0


def resolve(wb, ref: str) -> Tuple[int, RangeRef]:
    """The sheet and the cells ``ref`` stands for now."""
    text = (ref or "").strip().lstrip("=")
    if not text:
        raise EngineError("Reference isn't valid.")
    sheet, rest = _sheet_of(text)
    si = _first_shown(wb)
    if sheet is not None:
        si = wb.sheet_index(sheet)
        if si is None:
            raise EngineError(f"There is no sheet named {sheet}.")
    m = _OPEN_END.match(rest)
    if m:
        c0, c1 = sorted((letters_to_col(m.group(1).upper()), letters_to_col(m.group(3).upper())))
        r0 = int(m.group(2)) - 1
        rng = RangeRef(r0, c0, max(r0, last_row(wb.sheets[si], r0, c0, c1)), c1)
    else:
        si, rng = resolve_reference(wb, si, text)
        if rng.is_full_cols():
            sh = wb.sheets[si]
            r0 = _first_row(sh, rng.c0, rng.c1)
            rng = RangeRef(r0, rng.c0, max(r0, last_row(sh, r0, rng.c0, rng.c1)), rng.c1)
    if wb.sheets[si].visibility != "visible":
        raise EngineError("Reference isn't valid.")
    return si, rng


def suggest(wb, si: int, sel: RangeRef) -> Dict[str, Any]:
    """How to name the cells selected on sheet ``si``. ``ref`` is the best fit: the table when
    the selection is one, the cells down to wherever the data ends when the selection reaches
    that end, else the very cells. ``fixed`` and ``growing`` are the two ways to name cells of
    no table (``growing`` is empty for a single row: nothing says it is a list)."""
    sh = wb.sheets[si]
    prefix = quote_sheet(sh.name) + "!"
    rng = sel
    if rng.is_single():
        rng = current_region(sh, rng.start)
    elif rng.is_full_cols():
        r0 = _first_row(sh, rng.c0, rng.c1)
        rng = RangeRef(r0, rng.c0, max(r0, last_row(sh, r0, rng.c0, rng.c1)), rng.c1)
    elif rng.is_full_rows():
        used = sh.used_range() or RangeRef(0, 0, 0, 0)
        rng = RangeRef(rng.r0, used.c0, min(rng.r1, used.r1), used.c1)
    t = sh.table_at(rng.start)
    if t is not None and (sel.is_single() or rng == t.range or rng == t.data_range()):
        return {"ref": t.name, "kind": "table", "grows": True, "fixed": "", "growing": "", "sheet": sh.name, "area": t.range.a1(), "rows": t.range.height}
    fixed = prefix + rng.a1()
    growing = prefix + f"{a1(rng.r0, rng.c0)}:{col_to_letters(rng.c1)}" if rng.height > 1 else ""
    grows = bool(growing) and rng.r1 >= last_row(sh, rng.r0, rng.c0, rng.c1)
    return {"ref": growing if grows else fixed, "kind": "cells", "grows": grows, "fixed": fixed, "growing": growing, "sheet": sh.name, "area": rng.a1(), "rows": rng.height}


def main_block(sh) -> RangeRef:
    """Where a sheet's data is, to start from: its first table, else the biggest block of cells
    (a title above a list is a block of its own)."""
    if sh.tables:
        return sh.tables[0].range
    used = sh.used_range()
    if used is None:
        return RangeRef(0, 0, 0, 0)
    best, r = RangeRef(used.r0, used.c0, used.r0, used.c0), used.r0
    while r <= used.r1 and r < used.r0 + PREVIEW_ROWS:
        c = next((c for c in sorted(sh.cells.row(r)) if _filled(sh, (r, c))), None)
        if c is None:
            r += 1
            continue
        block = current_region(sh, (r, c))
        if block.count() > best.count():
            best = block
        r = block.r1 + 1
    return best


def forms(wb) -> List[Dict[str, str]]:
    """The workbook's UserForms."""
    from .macro.project import form_design

    proj = wb.vba
    if proj is None or proj.unreadable:
        return []
    return [{"name": m.name, "caption": form_design(proj, m).caption or m.name} for m in proj.modules if m.kind == FORM]


def tables(wb) -> List[Dict[str, Any]]:
    """The tables on the sheets people see."""
    return [{"name": t.name, "sheet": s.name, "area": t.range.a1(), "rows": t.range.height} for s in wb.sheets if s.visibility == "visible" for t in s.tables]


def names(wb) -> List[Dict[str, Any]]:
    """The defined names that stand for cells on the sheets people see."""
    out = []
    for d in wb.names:
        if d.hidden or d.name.startswith("_xlnm."):
            continue
        try:
            si, rng = resolve(wb, quote_sheet(wb.sheets[d.scope].name) + "!" + d.name if d.scope is not None else d.name)
        except (EngineError, IndexError):
            continue
        out.append({"name": d.name, "sheet": wb.sheets[si].name, "area": rng.a1(), "rows": rng.height})
    return out


def preview(wb, si: int) -> Dict[str, Any]:
    """A sheet's cells as text, to choose from: ``r0`` and ``c0`` are where ``rows`` starts."""
    sh = wb.sheets[si]
    used = sh.used_range()
    if used is None:
        return {"r0": 0, "c0": 0, "rows": [[""] * 8 for _ in range(12)], "more": False}
    r1, c1 = min(used.r1, used.r0 + PREVIEW_ROWS - 1), min(used.c1, used.c0 + PREVIEW_COLS - 1)
    rows = [[value_text(wb, sh, (r, c)) if sh.value((r, c)) is not None else "" for c in range(used.c0, c1 + 1)] for r in range(used.r0, r1 + 1)]
    return {"r0": used.r0, "c0": used.c0, "rows": rows, "more": r1 < used.r1 or c1 < used.c1}


def table_view(wb, ref: str, limit: int = TABLE_ROWS) -> Dict[str, Any]:
    """The cells ``ref`` stands for as an HTML table: ``html``, the ``rows`` they have, and how
    many are ``shown``. With more than ``limit`` rows it shows the first (the headings, usually)
    and the latest."""
    si, rng = resolve(wb, ref)
    sh = wb.sheets[si]
    rows = [r for r in range(rng.r0, rng.r1 + 1) if not sh.is_row_hidden(r)]
    total = len(rows)
    if total > limit:
        rows = rows[:1] + rows[-(limit - 1):]
    return {"html": table_markup(wb, si, rng, rows=rows), "rows": total, "shown": len(rows)}


@command("embed.info", "Embed in a Document", menu=("Insert", "Links"), params="{} → {selection: how to name the cells selected, forms, tables, names}", undoable=False, journal=False)
def embed_info(s: Session, p: Dict[str, Any]):
    d = s.doc()
    wb = d.wb
    return {"selection": suggest(wb, wb.active_sheet, d.selection.current()), "forms": forms(wb), "tables": tables(wb), "names": names(wb)}
