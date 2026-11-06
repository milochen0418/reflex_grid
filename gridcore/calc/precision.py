""""Set precision as displayed" (File › Options › Advanced in Excel; ``fullPrecision="0"`` in
``.xlsx``): with it on, a number is stored as its cell's number format shows it, and formulas read
those rounded values (GridCraft PR 1366). The rounding itself is the number format's
(:meth:`gridcore.numfmt.NumberFormat.round_as_displayed`); this looks each cell's format up,
parsing every style's code once."""

from __future__ import annotations

from typing import Dict, Optional, Tuple

from ..core.value import Value
from ..numfmt import NumberFormat, parse as parse_fmt

_UNSEEN = object()


class DisplayedPrecision:
    """Rounds values to their cell's displayed precision, remembering each style's parsed format.
    Styles don't change while one is used (one recalculation pass, one sweep of a workbook)."""

    def __init__(self) -> None:
        # Parsed format per style; None for General, which keeps the full value.
        self.formats: Dict[int, Optional[NumberFormat]] = {}

    def round(self, wb, sheet: int, rc: Tuple[int, int], n: float) -> float:
        """``n`` as the format of ``rc`` on ``sheet`` displays it (the cell's own style, else its
        row's, else its column's, as the grid shows it)."""
        sh = wb.sheet(sheet)
        if sh is None:
            return n
        sid = sh.style_id(rc)
        f = self.formats.get(sid, _UNSEEN)
        if f is _UNSEEN:
            nf = parse_fmt(wb.styles.get(sid).num_fmt)
            f = None if nf.is_general() else nf
            self.formats[sid] = f
        return n if f is None else f.round_as_displayed(n)

    def round_value(self, wb, sheet: int, rc: Tuple[int, int], v: Value) -> Value:
        """Like :meth:`round` for a value: only numbers change."""
        if type(v) is float:
            r = self.round(wb, sheet, rc, v)
            return v if r == v else r
        return v
