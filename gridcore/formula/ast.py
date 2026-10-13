"""Formula syntax tree.

Nodes are frozen dataclasses (hashable, comparable). A sheet selector is ``None`` for the
formula's own sheet, a ``str`` for a named sheet, or a ``(first, last)`` tuple for a 3-D span.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Callable, Iterator, List, Optional, Tuple, Union

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef
from ..core.value import CellError

SheetSel = Union[None, str, Tuple[str, str]]

# Binary operators and their binding power (higher binds tighter).
PRECEDENCE = {
    ":": 90,
    " ": 85,
    ",": 80,
    "^": 50,
    "*": 40,
    "/": 40,
    "+": 30,
    "-": 30,
    "&": 20,
    "=": 10,
    "<>": 10,
    "<": 10,
    ">": 10,
    "<=": 10,
    ">=": 10,
}

TRIM_MODES = {".:": 1, ":.": 2, ".:.": 3}  # symbol -> TRIMRANGE code


class Expr:
    __slots__ = ()

    def walk(self) -> Iterator["Expr"]:
        """Every node depth-first (pre-order)."""
        stack: List[Expr] = [self]
        while stack:
            e = stack.pop()
            yield e
            kids = e.children()
            if kids:
                stack.extend(reversed(kids))

    def children(self) -> List["Expr"]:
        return []

    def map(self, f: Callable[["Expr"], "Expr"]) -> "Expr":
        """Rewrites every node bottom-up."""
        return f(self._map_children(f))

    def _map_children(self, f: Callable[["Expr"], "Expr"]) -> "Expr":
        return self

    def references(self) -> List["Ref"]:
        return [e for e in self.walk() if isinstance(e, Ref)]

    def functions(self) -> List[str]:
        return [e.name for e in self.walk() if isinstance(e, Call)]


@dataclass(frozen=True)
class Num(Expr):
    value: float


@dataclass(frozen=True)
class Str(Expr):
    value: str


@dataclass(frozen=True)
class Bool(Expr):
    value: bool


@dataclass(frozen=True)
class Err(Expr):
    error: CellError


@dataclass(frozen=True)
class Missing(Expr):
    """An omitted argument ``F(1,,3)``."""


@dataclass(frozen=True)
class ArrayLit(Expr):
    """Array constant ``{1,2;3,4}`` (rows of literal expressions)."""

    rows: Tuple[Tuple[Expr, ...], ...]

    def children(self) -> List[Expr]:
        return [e for r in self.rows for e in r]

    def _map_children(self, f):
        return ArrayLit(tuple(tuple(e.map(f) for e in r) for r in self.rows))


@dataclass(frozen=True)
class Ref(Expr):
    """A reference: ``kind`` is ``cell``, ``range``, ``rows`` (``3:5``) or ``cols`` (``A:C``).

    For ``cell`` only ``r0``/``c0`` matter; for ``rows`` only ``r0``/``r1``; for ``cols`` only
    ``c0``/``c1``. Range ends are kept as written (``range()`` normalises)."""

    sheet: SheetSel
    kind: str
    r0: int = 0
    c0: int = 0
    r1: int = 0
    c1: int = 0
    r0_abs: bool = False
    c0_abs: bool = False
    r1_abs: bool = False
    c1_abs: bool = False

    @staticmethod
    def cell(row: int, col: int, row_abs: bool = False, col_abs: bool = False, sheet: SheetSel = None) -> "Ref":
        return Ref(sheet, "cell", row, col, row, col, row_abs, col_abs, row_abs, col_abs)

    @staticmethod
    def area(r0: int, c0: int, r1: int, c1: int, abs_: bool = False, sheet: SheetSel = None) -> "Ref":
        return Ref(sheet, "range", r0, c0, r1, c1, abs_, abs_, abs_, abs_)

    def range(self) -> RangeRef:
        k = self.kind
        if k == "cell":
            return RangeRef(self.r0, self.c0, self.r0, self.c0)
        if k == "range":
            return RangeRef.make(self.r0, self.c0, self.r1, self.c1)
        if k == "rows":
            return RangeRef.rows(min(self.r0, self.r1), max(self.r0, self.r1))
        return RangeRef.cols(min(self.c0, self.c1), max(self.c0, self.c1))

    def sheet_name(self) -> Optional[str]:
        return self.sheet if isinstance(self.sheet, str) else None

    def with_sheet(self, sheet: SheetSel) -> "Ref":
        return replace(self, sheet=sheet)


@dataclass(frozen=True)
class Name(Expr):
    """Defined name, table name or LET/LAMBDA parameter (``Sheet1!Name`` for sheet-scoped)."""

    name: str


@dataclass(frozen=True)
class StructRef(Expr):
    """Structured reference ``Table1[[#Headers],[Col1]:[Col3]]``, ``[@Col]``."""

    table: str
    specifiers: Tuple[str, ...] = ()  # "#All" "#Data" "#Headers" "#Totals" "#This Row"
    col_start: Optional[str] = None
    col_end: Optional[str] = None


@dataclass(frozen=True)
class Unary(Expr):
    """``-`` ``+`` ``%`` (postfix) ``@`` (implicit intersection) ``#`` (spill range) and the
    trim-reference modes ``.:`` ``:.`` ``.:.``."""

    op: str
    expr: Expr

    def children(self) -> List[Expr]:
        return [self.expr]

    def _map_children(self, f):
        return Unary(self.op, self.expr.map(f))


@dataclass(frozen=True)
class Binary(Expr):
    op: str
    left: Expr
    right: Expr

    def children(self) -> List[Expr]:
        return [self.left, self.right]

    def _map_children(self, f):
        return Binary(self.op, self.left.map(f), self.right.map(f))


@dataclass(frozen=True)
class Call(Expr):
    """Function call; the name is upper-case without ``_xlfn.``."""

    name: str
    args: Tuple[Expr, ...] = ()

    def children(self) -> List[Expr]:
        return list(self.args)

    def _map_children(self, f):
        return Call(self.name, tuple(a.map(f) for a in self.args))


@dataclass(frozen=True)
class Invoke(Expr):
    """Calling the result of an expression, e.g. ``LAMBDA(x,x+1)(2)``."""

    callee: Expr
    args: Tuple[Expr, ...] = ()

    def children(self) -> List[Expr]:
        return [self.callee, *self.args]

    def _map_children(self, f):
        return Invoke(self.callee.map(f), tuple(a.map(f) for a in self.args))


@dataclass(frozen=True)
class Paren(Expr):
    expr: Expr

    def children(self) -> List[Expr]:
        return [self.expr]

    def _map_children(self, f):
        return Paren(self.expr.map(f))


def clamp_anchor(row: int, col: int) -> Optional[Tuple[int, int]]:
    if 0 <= row < MAX_ROWS and 0 <= col < MAX_COLS:
        return row, col
    return None
