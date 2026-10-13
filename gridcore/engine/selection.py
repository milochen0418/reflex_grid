"""Cell selection: one or more ranges with an active cell, Ctrl+arrow jumps and the current
region."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional, Tuple

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef, a1


@dataclass
class Selection:
    active: Tuple[int, int] = (0, 0)
    anchor: Tuple[int, int] = (0, 0)
    ranges: List[RangeRef] = field(default_factory=lambda: [RangeRef(0, 0, 0, 0)])

    @staticmethod
    def at(c: Tuple[int, int]) -> "Selection":
        return Selection(c, c, [RangeRef(c[0], c[1], c[0], c[1])])

    @staticmethod
    def of_range(r: RangeRef, active: Optional[Tuple[int, int]] = None) -> "Selection":
        a = active if active is not None and r.contains(active) else r.start
        return Selection(a, a, [r])

    def current(self) -> RangeRef:
        return self.ranges[-1] if self.ranges else RangeRef(self.active[0], self.active[1], self.active[0], self.active[1])

    def bounds(self) -> RangeRef:
        r = self.ranges[0] if self.ranges else self.current()
        for x in self.ranges[1:]:
            r = r.union(x)
        return r

    def contains(self, c: Tuple[int, int]) -> bool:
        return any(r.contains(c) for r in self.ranges)

    def is_single_cell(self) -> bool:
        return len(self.ranges) == 1 and self.current().is_single()

    def a1(self) -> str:
        return ",".join(r.a1() for r in self.ranges)

    def copy(self) -> "Selection":
        return Selection(self.active, self.anchor, list(self.ranges))

    def to_json(self) -> dict:
        return {
            "active": list(self.active),
            "anchor": list(self.anchor),
            "ranges": [[r.r0, r.c0, r.r1, r.c1] for r in self.ranges],
            "a1": self.a1(),
        }

    def expand_merges(self, sheet) -> None:
        out = []
        for r in self.ranges:
            cur = r
            while True:
                grown = cur
                for m in sheet.merges:
                    if m.intersects(grown):
                        grown = grown.union(m)
                if grown == cur:
                    break
                cur = grown
            out.append(cur)
        self.ranges = out

    def cycle(self, forward: bool, by_row: bool) -> None:
        """Moves the active cell within the selection (Enter/Tab inside a multi-cell selection)."""
        r = self.current()
        if r.is_single() and len(self.ranges) == 1:
            return
        row, col = self.active
        step = 1 if forward else -1
        if by_row:
            col += step
            if col > r.c1:
                col, row = r.c0, row + 1
            elif col < r.c0:
                col, row = r.c1, row - 1
            if row > r.r1:
                row = r.r0
            elif row < r.r0:
                row = r.r1
        else:
            row += step
            if row > r.r1:
                row, col = r.r0, col + 1
            elif row < r.r0:
                row, col = r.r1, col - 1
            if col > r.c1:
                col = r.c0
            elif col < r.c0:
                col = r.c1
        self.active = (min(max(row, 0), MAX_ROWS - 1), min(max(col, 0), MAX_COLS - 1))


def _filled(sheet, c: Tuple[int, int]) -> bool:
    v = sheet.value(c)
    return v is not None


def jump(sheet, frm: Tuple[int, int], dr: int, dc: int) -> Tuple[int, int]:
    """Ctrl+arrow: jump to the edge of the current data region (Excel's End mode)."""
    if dr == 0 and dc == 0:
        return frm
    nxt = (frm[0] + dr, frm[1] + dc)
    if not (0 <= nxt[0] < MAX_ROWS and 0 <= nxt[1] < MAX_COLS):
        return frm
    cur = nxt
    if _filled(sheet, frm) and _filled(sheet, cur):
        while True:
            n = (cur[0] + dr, cur[1] + dc)
            if not (0 <= n[0] < MAX_ROWS and 0 <= n[1] < MAX_COLS) or not _filled(sheet, n):
                return cur
            cur = n
    used = sheet.used_range()
    while True:
        if _filled(sheet, cur):
            return cur
        beyond = (
            used is None
            or (dr > 0 and cur[0] > used.r1)
            or (dc > 0 and cur[1] > used.c1)
            or (dr < 0 and cur[0] < used.r0)
            or (dc < 0 and cur[1] < used.c0)
        )
        if beyond:
            r = MAX_ROWS - 1 if dr > 0 else (0 if dr < 0 else cur[0])
            c = MAX_COLS - 1 if dc > 0 else (0 if dc < 0 else cur[1])
            return (r, c)
        n = (cur[0] + dr, cur[1] + dc)
        if not (0 <= n[0] < MAX_ROWS and 0 <= n[1] < MAX_COLS):
            return cur
        cur = n


def current_region(sheet, c: Tuple[int, int]) -> RangeRef:
    """The contiguous data region around ``c`` (Ctrl+A / Ctrl+Shift+8, AutoSum, sort, tables)."""

    def filled(rc) -> bool:
        cell = sheet.cell(rc)
        return (cell is not None and (cell.value is not None or cell.formula is not None)) or rc in sheet.spill

    def row_has(row: int, a: int, b: int) -> bool:
        return any(filled((row, col)) for col in range(a, b + 1))

    def col_has(col: int, a: int, b: int) -> bool:
        return any(filled((row, col)) for row in range(a, b + 1))

    r0 = r1 = c[0]
    c0 = c1 = c[1]
    while True:
        before = (r0, r1, c0, c1)
        lo_c, hi_c = max(c0 - 1, 0), min(c1 + 1, MAX_COLS - 1)
        while r0 > 0 and row_has(r0 - 1, lo_c, hi_c):
            r0 -= 1
        while r1 + 1 < MAX_ROWS and row_has(r1 + 1, lo_c, hi_c):
            r1 += 1
        lo_r, hi_r = max(r0 - 1, 0), min(r1 + 1, MAX_ROWS - 1)
        while c0 > 0 and col_has(c0 - 1, lo_r, hi_r):
            c0 -= 1
        while c1 + 1 < MAX_COLS and col_has(c1 + 1, lo_r, hi_r):
            c1 += 1
        if (r0, r1, c0, c1) == before:
            break
    return RangeRef(r0, c0, r1, c1)
