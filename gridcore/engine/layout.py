"""Geometry helpers: approximate text widths, automatic row heights and AutoFit widths."""

from __future__ import annotations

import math
from dataclasses import replace
from typing import Iterable, Tuple

from ..model.features import LineInfo
from ..model.style import DEFAULT_FONT_SIZE
from .display import cell_text

# Advance widths (em) close to Calibri's, the default font; Excel's column widths are measured in
# Calibri digits (7 px at 11 pt).
_EM = {}
for _chars, _w in (
    ("0123456789$", 0.477),
    ("iljI'|", 0.23),
    (".,:;!", 0.25),
    ("ftr()[]{}-/\\ ", 0.33),
    ("mwMW%@", 0.8),
    ("ABCDEFGHJKLNOPQRSTUVXYZ&#", 0.58),
    ("abcdeghknopqsuvxyz*+=<>?^_~", 0.47),
):
    for _c in _chars:
        _EM[_c] = _w


def approx_text_width(text: str, size: float, bold: bool) -> float:
    """Width in px of ``text`` at ``size`` points with Calibri-like metrics (the server has no font
    rasteriser; the browser measures exactly when it draws)."""
    w = 0.0
    for c in text:
        e = _EM.get(c)
        if e is None:
            e = 0.5 if c.isascii() else 1.0  # CJK and other wide scripts take a full em
        w += e
    return w * size * (1.05 if bold else 1.0) * 96.0 / 72.0


def auto_row_height(wb, si: int, row: int) -> None:
    """Recomputes an automatic row height from its content: the largest font, and wrapped or
    multi-line text. Rows with a user-set height keep it."""
    sh = wb.sheet(si)
    if sh is None:
        return
    info = sh.rows.get(row)
    if info is not None and info.custom:
        return
    default = sh.default_row_height
    cells = sh.cells.row(row)
    plain = True
    for c, cell in cells.items():
        st = wb.styles.get(cell.style)
        if st.font.size > DEFAULT_FONT_SIZE or st.align.wrap or (isinstance(cell.value, str) and "\n" in cell.value):
            plain = False
            break
    if plain:
        if info is not None and info.size is not None:
            sm = wb.sheet_mut(si)
            sm.set_row(row, replace(info, size=None))
        return
    need = default
    for c, cell in cells.items():
        st = wb.styles.get(cell.style)
        size = st.font.size
        line = max(round(size * 1.8), default * size / DEFAULT_FONT_SIZE)
        if cell.value is not None or cell.formula is not None:
            text = cell_text(wb, sh, (row, c))
            lines_txt = text.split("\n") or [""]
            if st.align.wrap:
                w = max(sh.col_width(c) - 6.0, 1.0)
                lines = max(sum(max(math.ceil(approx_text_width(l, size, st.font.bold) / w), 1) for l in lines_txt), 1)
            else:
                lines = max(len(lines_txt), 1)
        elif size <= DEFAULT_FONT_SIZE:
            continue
        else:
            lines = 1
        h = math.ceil(size * 96.0 / 72.0 * 1.22 * lines + 5.0) if lines > 1 else line
        need = max(need, min(h, 546.0))
    sm = wb.sheet_mut(si)
    cur = sm.rows.get(row, LineInfo())
    size = None if abs(need - default) < 0.5 else float(need)
    if cur.size != size:
        sm.set_row(row, replace(cur, size=size))


def autofit_width(wb, sh, col: int) -> float:
    used = sh.used_range()
    if used is None:
        return 0.0
    from ..core.addr import RangeRef

    mx = 0.0
    for rc, _cell in sh.cells.iter_range(RangeRef(used.r0, col, used.r1, col)):
        m = sh.merge_at(rc)
        if m is not None and not m.is_single():
            continue
        st = wb.styles.get(sh.style_id(rc))
        text = cell_text(wb, sh, rc)
        extra = 16.0 if sh.has_filter_button(rc) else 0.0
        for line in text.split("\n"):
            mx = max(mx, approx_text_width(line, st.font.size, st.font.bold) + 6.0 + st.align.indent * 9.0 + extra)
    return min(max(mx, 8.0), 1785.0) if mx > 0 else 0.0
