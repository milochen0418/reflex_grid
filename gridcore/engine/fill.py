"""Fill: copy, AutoFill series (numbers, dates, weekdays, months, text with numbers, custom
lists) and the Series command."""

from __future__ import annotations

import math
import re
from typing import List, Optional, Tuple

from ..core.addr import RangeRef
from ..core.date import MONTHS, WEEKDAYS, DateSystem, datetime_from_serial, days_in_month, max_serial, serial_from_ymd
from ..formula.adjust import shift_relative
from ..model.cell import Cell, Formula
from ..numfmt import parse as parse_fmt


def _lists() -> List[List[str]]:
    return [
        [d[:3] for d in WEEKDAYS],
        list(WEEKDAYS),
        [m[:3] for m in MONTHS],
        list(MONTHS),
        ["Q1", "Q2", "Q3", "Q4"],
    ]


def _round15(x: float) -> float:
    if x == 0 or not math.isfinite(x):
        return x
    mag = math.floor(math.log10(abs(x)))
    p = 10.0 ** (14 - mag)
    if not math.isfinite(p) or p == 0:
        return x
    y = x * p
    return (math.floor(y + 0.5) if y >= 0 else -math.floor(-y + 0.5)) / p


def _digit(ch: str) -> bool:
    return ch.isdigit() and ch.isascii()


def _split_num(t: str) -> Optional[Tuple[str, int, str, int]]:
    """The number AutoFill counts on: the one the text starts with (``1st Period``, ``1 to 5``),
    else the one it ends with (``Item 007`` → ("Item ", 7, "", 3)). A number inside the text
    (``a1b``, ``Item 5th``) is copied as it is (Excel)."""
    if t and _digit(t[0]):
        start = end = 0
        while end + 1 < len(t) and _digit(t[end + 1]):
            end += 1
    elif t and _digit(t[-1]):
        start = end = len(t) - 1
        while start > 0 and _digit(t[start - 1]):
            start -= 1
    else:
        return None
    digits = t[start:end + 1]
    if len(digits) > 15:
        return None
    width = len(digits) if digits.startswith("0") and len(digits) > 1 else 1
    return t[:start], int(digits), t[end + 1:], width


# Ordinal numbers continue as ordinals: 1st Period, 2nd Period, 3rd Period.
_ORDINAL = re.compile(r"^(st|nd|rd|th)(?![a-z])", re.IGNORECASE)
# Quarters wrap after 4 (Q3, Q4, Q1), however they are spelled.
_QUARTER = ("q", "qtr", "quarter")


def _ordinal(n: int) -> str:
    if 10 <= n % 100 <= 20:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")


def _add_months(sys: DateSystem, serial: float, months: int) -> Optional[float]:
    d = datetime_from_serial(sys, serial)
    if d is None:
        return None
    total = d.year * 12 + d.month - 1 + months
    y, m = total // 12, total % 12 + 1
    if not 1900 <= y <= 9999:
        return None
    return serial_from_ymd(sys, y, m, min(d.day, days_in_month(y, m)))


class Pattern:
    kind = "copy"

    def __init__(self, kind: str = "copy", **kw):
        self.kind = kind
        self.__dict__.update(kw)

    @staticmethod
    def detect(cells: List[Optional[Cell]], custom: List[List[str]], sys: DateSystem, fmt: str) -> "Pattern":
        if any(c is not None and c.formula is not None for c in cells):
            return Pattern()
        vals = [c.value if c is not None else None for c in cells]
        if vals and all(isinstance(v, float) for v in vals):
            nums = vals
            first, last = nums[0], nums[-1]
            is_date = parse_fmt(fmt).is_date() and not parse_fmt(fmt).is_time_only()
            if len(nums) == 1:
                if parse_fmt(fmt).is_time_only():
                    # A time steps by an hour (9:00, 10:00, 11:00).
                    return Pattern("linear", last=last, first=first, step=1 / 24)
                return Pattern("linear", last=last, first=first, step=1.0) if is_date else Pattern()
            if is_date:
                a, b = datetime_from_serial(sys, first), datetime_from_serial(sys, nums[1])
                if a is not None and b is not None and a.day == b.day and (a.month != b.month or a.year != b.year):
                    months = (b.year - a.year) * 12 + b.month - a.month
                    return Pattern("months", last=last, first=first, months=months)
            n = float(len(nums))
            mx = (n - 1) / 2
            my = sum(nums) / n
            num = sum((i - mx) * (y - my) for i, y in enumerate(nums))
            den = sum((i - mx) ** 2 for i in range(len(nums)))
            step = 0.0 if den == 0 else num / den
            return Pattern("linear", last=my + step * (n - 1 - mx), first=my - step * mx, step=step)
        if vals and all(isinstance(v, str) for v in vals):
            texts = vals
            for lst in list(custom) + _lists():
                low = [x.lower() for x in lst]
                idx = [low.index(t.lower()) if t.lower() in low else None for t in texts]
                if all(i is not None for i in idx):
                    step = idx[1] - idx[0] if len(idx) >= 2 else 1
                    t0 = texts[0]
                    upper = t0.upper() == t0 and any(ch.isalpha() for ch in t0) and len(t0) > 1
                    lower = t0.lower() == t0
                    return Pattern("list", lst=lst, last=idx[-1], first=idx[0], step=step, upper=upper, lower=lower)
            parts = [_split_num(t) for t in texts]
            if all(p is not None for p in parts):
                ordinal = all(_ORDINAL.match(p[2]) for p in parts)
                # 1st Period, 2nd Period: the same text once the ordinal ending is taken off.
                key = [(p[0], p[2][2:] if ordinal else p[2]) for p in parts]
                if all(k == key[0] for k in key):
                    step = parts[1][1] - parts[0][1] if len(parts) >= 2 else 1
                    prefix, suffix = key[0]
                    wrap = prefix.strip().lower() in _QUARTER and suffix == "" and all(1 <= p[1] <= 4 for p in parts)
                    return Pattern(
                        "textnum", prefix=prefix, suffix=suffix, last=parts[-1][1], first=parts[0][1], step=step,
                        width=parts[0][3], ordinal=ordinal, wrap=4 if wrap else 0,
                    )
        return Pattern()

    def alternate(self, cells: List[Optional[Cell]]) -> "Pattern":
        """Ctrl+drag swaps copying and extending."""
        if self.kind == "copy" and len(cells) == 1 and cells[0] is not None and cells[0].formula is None and isinstance(cells[0].value, float):
            n = cells[0].value
            return Pattern("linear", last=n, first=n, step=1.0)
        return Pattern()

    def value(self, k: int, forward: bool, sys: DateSystem):
        sign = 1 if forward else -1
        if self.kind == "linear":
            base = self.last if forward else self.first
            return _round15(base + self.step * k * sign)
        if self.kind == "textnum":
            base = self.last if forward else self.first
            n = base + self.step * k * sign
            n = (n - 1) % self.wrap + 1 if self.wrap else abs(n)
            end = _ordinal(n) + self.suffix if self.ordinal else self.suffix
            return f"{self.prefix}{str(n).zfill(self.width)}{end}"
        if self.kind == "list":
            ln = len(self.lst)
            base = self.last if forward else self.first
            s = self.lst[(base + self.step * k * sign) % ln]
            if self.upper:
                return s.upper()
            if self.lower and not s[:1].isupper():
                return s.lower()
            return s
        if self.kind == "months":
            base = self.last if forward else self.first
            return _add_months(sys, base, self.months * k * sign)
        return None


def fill(ctx, sheet: int, src: RangeRef, target: RangeRef, mode: str = "series") -> None:
    """Fills ``target`` (which contains ``src`` at one edge) from the pattern in ``src``.
    ``mode``: series, alternate (Ctrl+drag), fillseries (always a series: a single number counts
    up, as VBA's ``xlFillSeries``), copy, formats, values."""
    wb = ctx.wb
    sh = wb.sheet(sheet)
    vertical = target.height != src.height or target.width == src.width
    if vertical:
        forward = target.r1 > src.r1 or target.r0 == src.r0
    else:
        forward = target.c1 > src.c1 or target.c0 == src.c0
    lanes = range(src.c0, src.c1 + 1) if vertical else range(src.r0, src.r1 + 1)
    n = src.height if vertical else src.width
    writes: List[Tuple[Tuple[int, int], Optional[Cell]]] = []
    for lane in lanes:
        src_cells = []
        for i in range(n):
            c = (src.r0 + i, lane) if vertical else (lane, src.c0 + i)
            src_cells.append((c, sh.cell(c)))
        if vertical:
            idxs = range(src.r1 + 1, target.r1 + 1) if forward else range(src.r0 - 1, target.r0 - 1, -1)
            targets = [(r, lane) for r in idxs]
        else:
            idxs = range(src.c1 + 1, target.c1 + 1) if forward else range(src.c0 - 1, target.c0 - 1, -1)
            targets = [(lane, c) for c in idxs]
        if mode in ("series", "alternate", "fillseries"):
            first = src_cells[0][1]
            fmt = wb.styles.get(first.style).num_fmt if first is not None else "General"
            pattern = Pattern.detect([c for _, c in src_cells], [], wb.date_system, fmt)
            if mode == "alternate" or (mode == "fillseries" and pattern.kind == "copy"):
                pattern = pattern.alternate([c for _, c in src_cells])
        else:
            pattern = Pattern()
        # Cells without a common pattern continue in runs of neighbours of one kind, as Excel's
        # AutoFill does: 1, 2, x, 5 → 3, 4, x, 6; Product 1, On backorder → Product 2, On backorder.
        runs: Optional[List[Tuple[Pattern, int, int]]] = None
        if mode == "series" and pattern.kind == "copy" and n > 1:
            runs = _runs([c for _, c in src_cells], wb)
        for k, t in enumerate(targets):
            step = k + 1
            idx = k % len(src_cells)
            sc, scell = src_cells[idx if forward else len(src_cells) - 1 - idx]
            if mode == "formats":
                ex = sh.cell(t)
                style = scell.style if scell is not None else 0
                writes.append((t, ex.with_style(style) if ex is not None else Cell(None, None, style)))
                continue
            style = (sh.cell(t).style if sh.cell(t) is not None else 0) if mode == "values" else (scell.style if scell is not None else 0)
            if scell is None:
                writes.append((t, Cell(None, None, style) if style else None))
                continue
            if scell.formula is not None and mode != "values":
                e = scell.formula.get_expr()
                if e is not None:
                    f = Formula.from_expr(shift_relative(e, t[0] - sc[0], t[1] - sc[1]))
                    writes.append((t, Cell(None, f, style)))
                else:
                    writes.append((t, Cell(None, Formula(scell.formula.text), style)))
                continue
            if runs is not None:
                run, j, m = runs[idx if forward else n - 1 - idx]
                r = k // n + 1  # the how-many-th copy of the source this target is
                v = run.value(m * (r - 1) + j + 1 if forward else m * r - j, forward, wb.date_system)
            else:
                v = pattern.value(step, forward, wb.date_system)
            if v is None:
                v = scell.value
            writes.append((t, Cell(v, None, style)))
    sm = ctx.sheet_mut(sheet)
    for c, cell in writes:
        if cell is None or cell.is_blank():
            sm.remove_cell(c)
        else:
            sm.set_cell(c, cell)
        ctx.touch(sheet, c)


def _run_kind(c: Optional[Cell]) -> Optional[tuple]:
    """What makes neighbouring source cells one run: numbers; names of one list (Mon, Tue); texts
    around a number with the same words (Item 1, Item 3)."""
    if c is None or c.formula is not None:
        return None
    v = c.value
    if isinstance(v, float):
        return ("num",)
    if isinstance(v, str):
        low = v.lower()
        for i, lst in enumerate(_lists()):
            if low in (x.lower() for x in lst):
                return ("list", i)
        p = _split_num(v)
        if p is not None:
            ordinal = bool(_ORDINAL.match(p[2]))
            return ("textnum", p[0], p[2][2:] if ordinal else p[2], ordinal)
    return None


def _runs(cells: List[Optional[Cell]], wb) -> List[Tuple[Pattern, int, int]]:
    """For each source cell of a source without one pattern: the pattern of its run, its place in
    the run and the run's length. A number alone counts on by 1."""
    out: List[Tuple[Pattern, int, int]] = []
    i = 0
    while i < len(cells):
        kind = _run_kind(cells[i])
        j = i + 1
        while kind is not None and j < len(cells) and _run_kind(cells[j]) == kind:
            j += 1
        run = cells[i:j]
        if kind is None:
            p = Pattern()
        else:
            fmt = wb.styles.get(run[0].style).num_fmt
            p = Pattern.detect(run, [], wb.date_system, fmt)
            if p.kind == "copy" and kind == ("num",) and len(run) == 1:
                p = Pattern("linear", last=run[0].value, first=run[0].value, step=1.0)
        out.extend((p, k, len(run)) for k in range(len(run)))
        i = j
    return out


def series(ctx, sheet: int, rng: RangeRef, rows: bool, kind: str, step: float, stop: Optional[float], unit: str = "day") -> None:
    """Home › Fill › Series: linear, growth or date series from the first cell of each lane."""
    wb = ctx.wb
    sh = wb.sheet(sheet)
    sm = ctx.sheet_mut(sheet)
    lanes = range(rng.r0, rng.r1 + 1) if rows else range(rng.c0, rng.c1 + 1)
    for lane in lanes:
        cells = [(lane, c) for c in range(rng.c0, rng.c1 + 1)] if rows else [(r, lane) for r in range(rng.r0, rng.r1 + 1)]
        first = sh.cell(cells[0])
        if first is None or not isinstance(first.value, float):
            continue
        v = first.value
        for i, c in enumerate(cells[1:], 1):
            prev = v
            if kind == "growth":
                v = v * step
            elif kind == "date":
                # Months and years count from the first date: Jan 31, Feb 28, Mar 31.
                if unit == "month":
                    v = _add_months(wb.date_system, first.value, int(step) * i) or v
                elif unit == "year":
                    v = _add_months(wb.date_system, first.value, int(step) * 12 * i) or v
                elif unit == "weekday":
                    v = _add_weekdays(wb.date_system, v, int(step))
                else:
                    v = v + step
            else:
                v = _round15(v + step)
            # The stop value bounds the series in the direction it runs (a falling growth too).
            if stop is not None and ((v >= prev and v > stop) or (v < prev and v < stop)):
                break
            ex = sh.cell(c)
            sm.set_cell(c, Cell(v, None, ex.style if ex is not None else first.style))
            ctx.touch(sheet, c)


def _add_weekdays(sys: DateSystem, v: float, k: int) -> float:
    """``v`` moved by ``k`` weekdays (Saturdays and Sundays skipped), in whole weeks where it can, so
    a huge step can't keep the server busy."""
    from .session import EngineError

    off = 6 if sys is DateSystem.D1900 else 5

    def weekend(x: float) -> bool:
        return (int(x) + off) % 7 in (0, 6)

    d = 1 if k > 0 else -1
    steps = abs(k)
    if steps:
        # The first step lands on a weekday; from there five weekdays are a week.
        v += d
        while weekend(v):
            v += d
        steps -= 1
    weeks, rest = divmod(steps, 5)
    v += d * 7 * weeks
    for _ in range(rest):
        v += d
        while weekend(v):
            v += d
    if not 0 <= v <= max_serial(sys):
        raise EngineError("The weekday series runs past the supported date range.")
    return v
