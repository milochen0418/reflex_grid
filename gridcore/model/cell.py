"""Cells and formulas.

Cells are treated as immutable: changing a value or style stores a new :class:`Cell`. That
keeps undo snapshots (which share cell objects through the copy-on-write store) consistent.
"""

from __future__ import annotations

from typing import Optional

from ..core.addr import RangeRef
from ..core.value import CellError, Value, display
from ..formula import Expr, ParseError, parse, print_formula


class Formula:
    """A parsed formula. ``text`` is the canonical formula text without the leading ``=``."""

    __slots__ = ("text", "expr", "array")

    def __init__(self, text: str, expr: Optional[Expr] = None, array: Optional[RangeRef] = None):
        self.text = text
        self.expr = expr
        # Legacy array formula (Ctrl+Shift+Enter) covering this range; the formula lives in the
        # top-left cell.
        self.array = array

    @staticmethod
    def new(text: str) -> "Formula":
        """Parses ``text`` (with or without ``=``). Unparseable formulas keep their text and
        evaluate to ``#NAME?``."""
        body = text[1:] if text.startswith("=") else text
        try:
            e = parse(body)
        except ParseError:
            return Formula(body, None)
        return Formula(print_formula(e), e)

    @staticmethod
    def from_expr(e: Expr, array: Optional[RangeRef] = None) -> "Formula":
        return Formula(print_formula(e), e, array)

    def get_expr(self) -> Optional[Expr]:
        if self.expr is not None:
            return self.expr
        try:
            return parse(self.text)
        except ParseError:
            return None

    def __eq__(self, other: object) -> bool:
        return isinstance(other, Formula) and self.text == other.text and self.array == other.array

    def __hash__(self) -> int:
        return hash((self.text, self.array))

    def __repr__(self) -> str:
        return f"Formula({self.text!r}{', array=' + self.array.a1() if self.array else ''})"


class Cell:
    __slots__ = ("value", "formula", "style")

    def __init__(self, value: Value = None, formula: Optional[Formula] = None, style: int = 0):
        # Constant value, or the cached result of the formula.
        self.value = value
        self.formula = formula
        self.style = style

    def is_blank(self) -> bool:
        """No value, no formula and the default style: storing it is pointless."""
        return self.value is None and self.formula is None and self.style == 0

    def with_value(self, v: Value) -> "Cell":
        return Cell(v, self.formula, self.style)

    def with_style(self, s: int) -> "Cell":
        return Cell(self.value, self.formula, s)

    def input_text(self) -> str:
        """What the formula bar shows."""
        if self.formula is not None:
            return "=" + self.formula.text
        v = self.value
        if isinstance(v, str):
            return v
        return display(v)

    def bar_text(self) -> str:
        """Formula bar text when not editing: a legacy array formula in braces."""
        if self.formula is not None and self.formula.array is not None:
            return "{=" + self.formula.text + "}"
        return self.input_text()

    def __eq__(self, other: object) -> bool:
        return (
            isinstance(other, Cell)
            and self.style == other.style
            and self.formula == other.formula
            and _same_value(self.value, other.value)
        )

    def __repr__(self) -> str:
        return f"Cell({self.value!r}, {self.formula!r}, s={self.style})"


def _same_value(a: Value, b: Value) -> bool:
    if type(a) is not type(b):
        return False
    return a == b


EMPTY_CELL = Cell()
