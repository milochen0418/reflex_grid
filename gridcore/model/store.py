"""Copy-on-write sparse cell storage.

Cells live in bands of 64 rows: ``band -> {row -> {col -> Cell}}``. Cloning a store (an undo
snapshot) copies only the band map; editing a cell copies just its band (two dict levels).
Cell objects themselves are immutable and shared.
"""

from __future__ import annotations

from typing import Dict, Iterator, List, Optional, Set, Tuple

from ..core.addr import RangeRef
from .cell import Cell

BAND = 64

Row = Dict[int, Cell]
Band = Dict[int, Row]


class CellStore:
    __slots__ = ("_bands", "_owned", "_len", "_used", "_used_valid")

    def __init__(self) -> None:
        self._bands: Dict[int, Band] = {}
        # Bands this store may mutate in place (created or copied since the last clone).
        self._owned: Set[int] = set()
        self._len = 0
        self._used: Optional[RangeRef] = None
        self._used_valid = True

    def clone(self) -> "CellStore":
        n = CellStore.__new__(CellStore)
        n._bands = dict(self._bands)
        n._owned = set()
        n._len = self._len
        n._used = self._used
        n._used_valid = self._used_valid
        # Both stores now share the band dicts: neither may mutate them in place.
        self._owned = set()
        return n

    def __len__(self) -> int:
        return self._len

    def get(self, rc: Tuple[int, int]) -> Optional[Cell]:
        band = self._bands.get(rc[0] >> 6)
        if band is None:
            return None
        row = band.get(rc[0])
        if row is None:
            return None
        return row.get(rc[1])

    def has(self, rc: Tuple[int, int]) -> bool:
        return self.get(rc) is not None

    def _band_mut(self, b: int, create: bool) -> Optional[Band]:
        band = self._bands.get(b)
        if band is None:
            if not create:
                return None
            band = {}
            self._bands[b] = band
            self._owned.add(b)
            return band
        if b not in self._owned:
            band = {r: dict(cols) for r, cols in band.items()}
            self._bands[b] = band
            self._owned.add(b)
        return band

    def set(self, rc: Tuple[int, int], cell: Cell) -> None:
        """Stores a cell; a blank (empty, unstyled) cell is removed instead."""
        if cell.is_blank():
            self.remove(rc)
            return
        r, c = rc
        band = self._band_mut(r >> 6, True)
        row = band.get(r)
        if row is None:
            row = {}
            band[r] = row
        if c not in row:
            self._len += 1
            self._extend_used(r, c)
        row[c] = cell

    def remove(self, rc: Tuple[int, int]) -> Optional[Cell]:
        r, c = rc
        b = r >> 6
        band = self._bands.get(b)
        if band is None or c not in band.get(r, ()):
            return None
        band = self._band_mut(b, False)
        row = band[r]
        old = row.pop(c)
        if not row:
            del band[r]
        if not band:
            del self._bands[b]
            self._owned.discard(b)
        self._len -= 1
        self._used_valid = False
        return old

    def _extend_used(self, r: int, c: int) -> None:
        if not self._used_valid:
            return
        u = self._used
        if u is None:
            self._used = RangeRef(r, c, r, c)
        elif not u.contains((r, c)):
            self._used = RangeRef(min(u.r0, r), min(u.c0, c), max(u.r1, r), max(u.c1, c))

    def used_range(self) -> Optional[RangeRef]:
        """Smallest range containing every stored cell, ``None`` when empty."""
        if self._used_valid:
            return self._used
        r0 = c0 = 1 << 30
        r1 = c1 = -1
        for band in self._bands.values():
            for r, cols in band.items():
                if not cols:
                    continue
                if r < r0:
                    r0 = r
                if r > r1:
                    r1 = r
                lo, hi = min(cols), max(cols)
                if lo < c0:
                    c0 = lo
                if hi > c1:
                    c1 = hi
        self._used = None if r1 < 0 else RangeRef(r0, c0, r1, c1)
        self._used_valid = True
        return self._used

    def iter_range(self, rg: RangeRef) -> Iterator[Tuple[Tuple[int, int], Cell]]:
        """Cells in ``rg`` (unordered within a band; callers sort when order matters)."""
        r0, c0, r1, c1 = rg
        b0, b1 = r0 >> 6, r1 >> 6
        bands = self._bands
        if b1 - b0 > len(bands):
            keys = [b for b in bands if b0 <= b <= b1]
        else:
            keys = [b for b in range(b0, b1 + 1) if b in bands]
        whole_cols = c0 == 0 and c1 >= 16383
        for b in keys:
            band = bands[b]
            for r, cols in band.items():
                if r0 <= r <= r1:
                    if whole_cols:
                        for c, cell in cols.items():
                            yield (r, c), cell
                    elif len(cols) > (c1 - c0 + 1) * 2:
                        for c in range(c0, c1 + 1):
                            cell = cols.get(c)
                            if cell is not None:
                                yield (r, c), cell
                    else:
                        for c, cell in cols.items():
                            if c0 <= c <= c1:
                                yield (r, c), cell

    def iter_sorted(self, rg: Optional[RangeRef] = None) -> List[Tuple[Tuple[int, int], Cell]]:
        items = list(self.iter_range(rg) if rg is not None else self.items())
        items.sort(key=lambda kv: kv[0])
        return items

    def items(self) -> Iterator[Tuple[Tuple[int, int], Cell]]:
        for band in self._bands.values():
            for r, cols in band.items():
                for c, cell in cols.items():
                    yield (r, c), cell

    def row(self, r: int) -> Dict[int, Cell]:
        """The cells of one row (col -> cell). Do not mutate the returned dict."""
        band = self._bands.get(r >> 6)
        if band is None:
            return {}
        return band.get(r, {})

    def rows_in(self, r0: int, r1: int) -> List[int]:
        out = []
        for b in range(r0 >> 6, (r1 >> 6) + 1):
            band = self._bands.get(b)
            if band:
                out.extend(r for r in band if r0 <= r <= r1)
        out.sort()
        return out

    def take_range(self, rg: RangeRef) -> List[Tuple[Tuple[int, int], Cell]]:
        """Removes and returns the cells of a range."""
        got = list(self.iter_range(rg))
        for rc, _ in got:
            self.remove(rc)
        return got

    def diff_keys(self, other: "CellStore") -> List[Tuple[int, int]]:
        """Keys whose cell object differs between two stores (bands compared by identity first)."""
        out = []
        for b in set(self._bands) | set(other._bands):
            x, y = self._bands.get(b), other._bands.get(b)
            if x is y:
                continue
            x = x or {}
            y = y or {}
            for r in set(x) | set(y):
                cx, cy = x.get(r, {}), y.get(r, {})
                if cx is cy:
                    continue
                for c in set(cx) | set(cy):
                    if cx.get(c) is not cy.get(c):
                        out.append((r, c))
        return out
