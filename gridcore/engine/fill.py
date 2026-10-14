"""Fill: copy, AutoFill series (numbers, dates, weekdays, months, text with numbers, custom
lists) and the Series command."""

from __future__ import annotations

import math
from typing import List, Optional, Tuple

from ..core.addr import RangeRef
from ..core.date import MONTHS, WEEKDAYS, DateSystem, datetime_from_serial, days_in_month, serial_from_ymd
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


def _split_num(t: str) -> Optional[Tuple[str, int, str, int]]:
    """``Item 007`` → ("Item ", 7, "", 3)."""
    end = None
    for i in range(len(t) - 1, -1, -1):
        if t[i].isdigit() and t[i].isascii():
            end = i
            break
    if end is None:
        return None
    start = end
    while start > 0 and t[start - 1].isdigit() and t[start - 1].isascii():
        start -= 1
    digits = t[start:end + 1]
    if len(digits) > 15:
        return None
    width = len(digits) if digits.startswith("0") and len(digits) > 1 else 1
    return t[:start], int(digits), t[end + 1:], width


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
            if all(p is not None for p in parts) and all(p[0] == parts[0][0] and p[2] == parts[0][2] for p in parts):
                step = parts[1][1] - parts[0][1] if len(parts) >= 2 else 1
                return Pattern("textnum", prefix=parts[0][0], suffix=parts[0][2], last=parts[-1][1], first=parts[0][1], step=step, width=parts[0][3])
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
            n = abs(base + self.step * k * sign)
            return f"{self.prefix}{str(n).zfill(self.width)}{self.suffix}"
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
        for c in cells[1:]:
            if kind == "growth":
                v = v * step
            elif kind == "date":
                if unit == "month":
                    v = _add_months(wb.date_system, v, int(step)) or v
                elif unit == "year":
                    v = _add_months(wb.date_system, v, int(step) * 12) or v
                elif unit == "weekday":
                    k = int(step)
                    d = 1 if k > 0 else -1
                    for _ in range(abs(k)):
                        v += d
                        while (int(v) + 6) % 7 in (0, 6):
                            v += d
                else:
                    v = v + step
            else:
                v = _round15(v + step)
            if stop is not None and ((step >= 0 and v > stop) or (step < 0 and v < stop)):
                break
            ex = sh.cell(c)
            sm.set_cell(c, Cell(v, None, ex.style if ex is not None else first.style))
            ctx.touch(sheet, c)
