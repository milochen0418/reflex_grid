"""Cell and range addresses.

Rows and columns are 0-based (``A1`` is ``(0, 0)``). A cell is a plain ``(row, col)`` tuple,
which keeps row-major ordering and hashing cheap. The sheet is 1,048,576 rows by 16,384
columns (``XFD1048576``), like Excel.
"""

from __future__ import annotations

from typing import Iterator, NamedTuple, Optional, Tuple

MAX_ROWS = 1_048_576
MAX_COLS = 16_384

Cell = Tuple[int, int]

_COL_CACHE: dict[int, str] = {}


def col_to_letters(col: int) -> str:
    """0 -> ``A``, 25 -> ``Z``, 26 -> ``AA``, 16383 -> ``XFD``."""
    s = _COL_CACHE.get(col)
    if s is not None:
        return s
    n = col + 1
    out = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out.append(chr(65 + rem))
    s = "".join(reversed(out))
    if len(_COL_CACHE) < 20000:
        _COL_CACHE[col] = s
    return s


def letters_to_col(s: str) -> Optional[int]:
    """``A`` -> 0, ``xfd`` -> 16383. ``None`` for empty, non-letters or beyond ``XFD``."""
    if not s or len(s) > 3:
        return None
    n = 0
    for ch in s:
        o = ord(ch)
        if 97 <= o <= 122:
            o -= 32
        if not 65 <= o <= 90:
            return None
        n = n * 26 + (o - 64)
    if n == 0 or n > MAX_COLS:
        return None
    return n - 1


def a1(row: int, col: int) -> str:
    """``(6, 1)`` -> ``B7``."""
    return f"{col_to_letters(col)}{row + 1}"


def parse_a1_prefix(s: str) -> Optional[tuple[int, int, bool, bool, str]]:
    """Parses an A1 reference at the start of ``s``: ``(col, row, col_abs, row_abs, rest)``."""
    i = 0
    n = len(s)
    col_abs = n > 0 and s[0] == "$"
    if col_abs:
        i += 1
    cs = i
    while i < n and s[i].isascii() and s[i].isalpha():
        i += 1
    if i == cs or i - cs > 3:
        return None
    col = letters_to_col(s[cs:i])
    if col is None:
        return None
    row_abs = i < n and s[i] == "$"
    if row_abs:
        i += 1
    rs = i
    while i < n and "0" <= s[i] <= "9":
        i += 1
    if i == rs or i - rs > 7:
        return None
    row = int(s[rs:i])
    if row == 0 or row > MAX_ROWS:
        return None
    return col, row - 1, col_abs, row_abs, s[i:]


def parse_cell(s: str) -> Optional[Cell]:
    """Parses ``A1``, ``$A$1``, ``xfd1048576`` (case-insensitive)."""
    p = parse_a1_prefix(s.strip())
    if p is None or p[4]:
        return None
    return (p[1], p[0])


def offset_cell(c: Cell, dr: int, dc: int) -> Optional[Cell]:
    r, k = c[0] + dr, c[1] + dc
    if 0 <= r < MAX_ROWS and 0 <= k < MAX_COLS:
        return (r, k)
    return None


def clamp_cell(r: int, c: int) -> Cell:
    return (min(max(r, 0), MAX_ROWS - 1), min(max(c, 0), MAX_COLS - 1))


class RangeRef(NamedTuple):
    """A rectangular block of cells, inclusive on both ends. Always normalised."""

    r0: int
    c0: int
    r1: int
    c1: int

    @staticmethod
    def of(a: Cell, b: Optional[Cell] = None) -> "RangeRef":
        if b is None:
            b = a
        return RangeRef(min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1]))

    @staticmethod
    def make(r0: int, c0: int, r1: int, c1: int) -> "RangeRef":
        return RangeRef(min(r0, r1), min(c0, c1), max(r0, r1), max(c0, c1))

    @staticmethod
    def rows(r0: int, r1: int) -> "RangeRef":
        return RangeRef.make(r0, 0, r1, MAX_COLS - 1)

    @staticmethod
    def cols(c0: int, c1: int) -> "RangeRef":
        return RangeRef.make(0, c0, MAX_ROWS - 1, c1)

    @staticmethod
    def all() -> "RangeRef":
        return RangeRef(0, 0, MAX_ROWS - 1, MAX_COLS - 1)

    @property
    def start(self) -> Cell:
        return (self.r0, self.c0)

    @property
    def end(self) -> Cell:
        return (self.r1, self.c1)

    @property
    def height(self) -> int:
        return self.r1 - self.r0 + 1

    @property
    def width(self) -> int:
        return self.c1 - self.c0 + 1

    def count(self) -> int:
        return self.height * self.width

    def is_single(self) -> bool:
        return self.r0 == self.r1 and self.c0 == self.c1

    def is_full_rows(self) -> bool:
        return self.c0 == 0 and self.c1 == MAX_COLS - 1

    def is_full_cols(self) -> bool:
        return self.r0 == 0 and self.r1 == MAX_ROWS - 1

    def contains(self, c: Cell) -> bool:
        return self.r0 <= c[0] <= self.r1 and self.c0 <= c[1] <= self.c1

    def contains_range(self, o: "RangeRef") -> bool:
        return self.r0 <= o.r0 and o.r1 <= self.r1 and self.c0 <= o.c0 and o.c1 <= self.c1

    def intersects(self, o: "RangeRef") -> bool:
        return self.r0 <= o.r1 and o.r0 <= self.r1 and self.c0 <= o.c1 and o.c0 <= self.c1

    def intersection(self, o: "RangeRef") -> Optional["RangeRef"]:
        if not self.intersects(o):
            return None
        return RangeRef(max(self.r0, o.r0), max(self.c0, o.c0), min(self.r1, o.r1), min(self.c1, o.c1))

    def union(self, o: "RangeRef") -> "RangeRef":
        return RangeRef(min(self.r0, o.r0), min(self.c0, o.c0), max(self.r1, o.r1), max(self.c1, o.c1))

    def cells(self) -> Iterator[Cell]:
        """Row-major iterator. Callers must bound huge ranges themselves."""
        for r in range(self.r0, self.r1 + 1):
            for c in range(self.c0, self.c1 + 1):
                yield (r, c)

    def a1(self) -> str:
        """``A1:B2``, ``A1`` for single cells, ``1:3`` for full rows, ``A:C`` for full columns."""
        if self.is_full_rows() and not self.is_full_cols():
            return f"{self.r0 + 1}:{self.r1 + 1}"
        if self.is_full_cols() and not self.is_full_rows():
            return f"{col_to_letters(self.c0)}:{col_to_letters(self.c1)}"
        if self.is_single():
            return a1(self.r0, self.c0)
        return f"{a1(self.r0, self.c0)}:{a1(self.r1, self.c1)}"

    def abs_a1(self) -> str:
        """``$A$1:$D$20`` (or ``$1:$3`` / ``$A:$C``)."""
        if self.is_full_rows() and not self.is_full_cols():
            return f"${self.r0 + 1}:${self.r1 + 1}"
        if self.is_full_cols() and not self.is_full_rows():
            return f"${col_to_letters(self.c0)}:${col_to_letters(self.c1)}"
        if self.is_single():
            return f"${col_to_letters(self.c0)}${self.r0 + 1}"
        return f"${col_to_letters(self.c0)}${self.r0 + 1}:${col_to_letters(self.c1)}${self.r1 + 1}"

    @staticmethod
    def parse(s: str) -> Optional["RangeRef"]:
        """Parses ``A1``, ``A1:B2``, ``$A$1:$B$2``, ``A:C``, ``3:5`` (no sheet prefix)."""
        s = s.strip()
        if ":" in s:
            a, b = s.split(":", 1)
            x, y = parse_cell(a), parse_cell(b)
            if x is not None and y is not None:
                return RangeRef.of(x, y)
            sa, sb = a.strip().lstrip("$"), b.strip().lstrip("$")
            ca, cb = letters_to_col(sa), letters_to_col(sb)
            if ca is not None and cb is not None:
                return RangeRef.cols(min(ca, cb), max(ca, cb))
            if sa.isdigit() and sb.isdigit():
                pa, pb = int(sa), int(sb)
                if pa == 0 or pb == 0 or pa > MAX_ROWS or pb > MAX_ROWS:
                    return None
                return RangeRef.rows(min(pa, pb) - 1, max(pa, pb) - 1)
            return None
        c = parse_cell(s)
        return RangeRef.of(c) if c is not None else None

    def __str__(self) -> str:  # pragma: no cover - debugging aid
        return self.a1()


def parse_ranges(s: str) -> list[RangeRef]:
    """A space- or comma-separated list of ranges (``sqref``)."""
    out = []
    for part in s.replace(",", " ").split():
        r = RangeRef.parse(part)
        if r is not None:
            out.append(r)
    return out
