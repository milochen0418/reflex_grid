"""Cell values and Excel's coercion and comparison rules.

Values are plain Python objects, which keeps the evaluator fast:

* ``None`` – empty
* ``float`` – number (ints are normalised to float at the boundaries)
* ``str`` – text
* ``bool`` – logical
* :class:`CellError` – an error value such as ``#DIV/0!``
* :class:`Array` – a 2-D array (formula results, references read as values)
"""

from __future__ import annotations

import enum
import math
from typing import Any, Iterable, List, Optional, Union


class CellError(enum.Enum):
    """Excel's error values."""

    NULL = "#NULL!"
    DIV0 = "#DIV/0!"
    VALUE = "#VALUE!"
    REF = "#REF!"
    NAME = "#NAME?"
    NUM = "#NUM!"
    NA = "#N/A"
    GETTING_DATA = "#GETTING_DATA"
    SPILL = "#SPILL!"
    CALC = "#CALC!"
    # Circular reference that could not be resolved (Excel shows 0; we keep it explicit).
    CIRC = "#CIRC!"

    def __str__(self) -> str:
        return self.value

    def __repr__(self) -> str:
        return f"CellError({self.value})"

    @property
    def code(self) -> int:
        """``ERROR.TYPE`` number."""
        return _ERROR_CODES[self]

    @staticmethod
    def parse(s: str) -> Optional["CellError"]:
        return _ERROR_BY_TEXT.get(s.strip().upper())


_ERROR_CODES = {
    CellError.NULL: 1,
    CellError.DIV0: 2,
    CellError.VALUE: 3,
    CellError.REF: 4,
    CellError.NAME: 5,
    CellError.NUM: 6,
    CellError.NA: 7,
    CellError.GETTING_DATA: 8,
    CellError.SPILL: 9,
    CellError.CALC: 14,
    CellError.CIRC: 15,
}
_ERROR_BY_TEXT = {e.value: e for e in CellError}

# Errors written as literals in formulas (the lexer's list).
ERROR_LITERALS = [e for e in CellError if e is not CellError.CIRC]


class XlError(Exception):
    """Raised inside the engine to abort with an Excel error value (like Rust's ``?``)."""

    __slots__ = ("error",)

    def __init__(self, error: CellError):
        super().__init__(error.value)
        self.error = error


class Array:
    """A 2-D array of values (row-major). Always at least 1×1."""

    __slots__ = ("rows", "cols", "data")

    def __init__(self, rows: int, cols: int, data: List[Any]):
        self.rows = rows
        self.cols = cols
        self.data = data

    @staticmethod
    def new(rows: int, cols: int, data: List[Any]) -> Optional["Array"]:
        if rows <= 0 or cols <= 0 or rows * cols != len(data):
            return None
        return Array(rows, cols, data)

    @staticmethod
    def filled(rows: int, cols: int, v: Any) -> "Array":
        rows, cols = max(rows, 1), max(cols, 1)
        return Array(rows, cols, [v] * (rows * cols))

    @staticmethod
    def scalar(v: Any) -> "Array":
        return Array(1, 1, [v])

    @staticmethod
    def column(values: List[Any]) -> "Array":
        if not values:
            return Array.scalar(CellError.CALC)
        return Array(len(values), 1, list(values))

    @staticmethod
    def row(values: List[Any]) -> "Array":
        if not values:
            return Array.scalar(CellError.CALC)
        return Array(1, len(values), list(values))

    @staticmethod
    def from_rows(rows: List[List[Any]]) -> "Array":
        if not rows or not rows[0]:
            return Array.scalar(CellError.CALC)
        w = len(rows[0])
        data: List[Any] = []
        for r in rows:
            if len(r) != w:
                raise ValueError("ragged rows")
            data.extend(r)
        return Array(len(rows), w, data)

    def get(self, r: int, c: int, default: Any = None) -> Any:
        if 0 <= r < self.rows and 0 <= c < self.cols:
            return self.data[r * self.cols + c]
        return default

    def get_broadcast(self, r: int, c: int) -> Any:
        """Excel's broadcasting read: a 1-row or 1-column array repeats along that axis; outside
        the array is ``#N/A``."""
        rr = 0 if self.rows == 1 else r
        cc = 0 if self.cols == 1 else c
        if rr < self.rows and cc < self.cols:
            return self.data[rr * self.cols + cc]
        return CellError.NA

    def transpose(self) -> "Array":
        data = []
        for c in range(self.cols):
            for r in range(self.rows):
                data.append(self.data[r * self.cols + c])
        return Array(self.cols, self.rows, data)

    def row_list(self, r: int) -> List[Any]:
        return self.data[r * self.cols:(r + 1) * self.cols]

    def col_list(self, c: int) -> List[Any]:
        return self.data[c::self.cols]

    def to_rows(self) -> List[List[Any]]:
        return [self.data[r * self.cols:(r + 1) * self.cols] for r in range(self.rows)]

    def __iter__(self):
        return iter(self.data)

    def __eq__(self, other: object) -> bool:
        return isinstance(other, Array) and self.rows == other.rows and self.cols == other.cols and self.data == other.data

    def __hash__(self) -> int:  # pragma: no cover - arrays are rarely hashed
        return hash((self.rows, self.cols, tuple(map(repr, self.data))))

    def __repr__(self) -> str:
        return f"Array({self.rows}x{self.cols}, {self.data!r})"


Value = Union[None, float, str, bool, CellError, Array]


def is_number(v: Any) -> bool:
    return isinstance(v, float) or (isinstance(v, int) and not isinstance(v, bool))


def is_error(v: Any) -> bool:
    return isinstance(v, CellError)


def number(v: float) -> Value:
    """A number, or ``#NUM!`` when it is NaN or infinite (Excel never stores those)."""
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    v = float(v)
    if math.isfinite(v):
        return 0.0 if v == 0.0 else v
    return CellError.NUM


def normalize(v: Any) -> Value:
    """Maps Python values to engine values (ints → floats, NaN → #NUM!)."""
    if v is None or isinstance(v, (str, bool, CellError, Array)):
        return v
    if isinstance(v, (int, float)):
        return number(v)
    return str(v)


def scalar(v: Value) -> Value:
    """Top-left element of an array, the value itself otherwise."""
    if isinstance(v, Array):
        return v.data[0] if v.data else None
    return v


def to_number(v: Value) -> float:
    """Coercion to a number as arithmetic does it: empty = 0, TRUE = 1, numeric text parses
    (dates, times, percentages and currency too), other text = ``#VALUE!``."""
    if v is None:
        return 0.0
    t = type(v)
    if t is float:
        return v  # type: ignore[return-value]
    if t is bool:
        return 1.0 if v else 0.0
    if t is str:
        from .parse import parse_number_text

        n = parse_number_text(v)  # type: ignore[arg-type]
        if n is None:
            raise XlError(CellError.VALUE)
        return n
    if t is CellError:
        raise XlError(v)  # type: ignore[arg-type]
    if t is int:
        return float(v)  # type: ignore[arg-type]
    if isinstance(v, Array):
        return to_number(v.data[0]) if v.data else 0.0
    raise XlError(CellError.VALUE)


def to_number_in(v: Value, sys) -> float:
    """:func:`to_number` with date text read in the date system ``sys`` (in a 1904 workbook
    ``"1/1/1904"+0`` is 0)."""
    if type(v) is str and sys is not None and sys.value == "1904":
        from .parse import parse_number_text_in

        n = parse_number_text_in(v, sys)  # type: ignore[arg-type]
        if n is None:
            raise XlError(CellError.VALUE)
        return n
    return to_number(v)


def to_text(v: Value) -> str:
    """Coercion to text as ``&`` does it."""
    if v is None:
        return ""
    t = type(v)
    if t is str:
        return v  # type: ignore[return-value]
    if t is float or t is int:
        return number_to_text(float(v))  # type: ignore[arg-type]
    if t is bool:
        return "TRUE" if v else "FALSE"
    if t is CellError:
        raise XlError(v)  # type: ignore[arg-type]
    if isinstance(v, Array):
        return to_text(v.data[0]) if v.data else ""
    return str(v)


def to_bool(v: Value) -> bool:
    """Coercion to a boolean as ``IF`` does it."""
    if v is None:
        return False
    t = type(v)
    if t is bool:
        return v  # type: ignore[return-value]
    if t is float or t is int:
        return v != 0
    if t is str:
        u = v.upper()  # type: ignore[union-attr]
        if u == "TRUE":
            return True
        if u == "FALSE":
            return False
        raise XlError(CellError.VALUE)
    if t is CellError:
        raise XlError(v)  # type: ignore[arg-type]
    if isinstance(v, Array):
        return to_bool(v.data[0]) if v.data else False
    raise XlError(CellError.VALUE)


def display(v: Value) -> str:
    """Display text with General formatting (for logs, CSV and tests)."""
    if isinstance(v, CellError):
        return v.value
    try:
        return to_text(v)
    except XlError as e:
        return e.error.value


def _rank(v: Value) -> int:
    """Type rank for comparison and sorting: numbers < text < booleans < errors; empty last."""
    if v is None:
        return 4
    t = type(v)
    if t is float or t is int:
        return 0
    if t is str:
        return 1
    if t is bool:
        return 2
    if t is CellError:
        return 3
    return 5


def _round15(x: float) -> float:
    if not math.isfinite(x) or x == 0.0:
        return x
    return float(f"{x:.14e}")


def compare_numbers(a: float, b: float) -> int:
    """Compares numbers the way Excel does: both rounded to 15 significant digits first, so
    binary noise such as ``0.1+0.2`` versus ``0.3`` counts as equal."""
    if a == b:
        return 0
    if abs(a - b) > max(abs(a), abs(b)) * 1e-13:
        return -1 if a < b else 1
    ra, rb = _round15(a), _round15(b)
    if ra == rb:
        return 0
    return -1 if ra < rb else 1


def _fold_char(c: str) -> str:
    f = c.casefold()
    if len(f) == 1:
        return f
    low = c.lower()  # ẞ folds to ß; ß and İ, whose folding is two letters, stay as they are
    return low if len(low) == 1 else c


def fold(s: str) -> str:
    """Excel's case folding (text comparison, SEARCH, criteria): Unicode simple case folding, one
    character for each character, so a position in the folded text is one in the text. ς, σ and Σ
    are one letter, µ and μ too; ß, İ and ı are letters of their own (Excel 16 for Mac)."""
    f = s.casefold()
    if len(f) == len(s):
        return f
    return "".join(_fold_char(c) for c in s)


def compare_text(a: str, b: str) -> int:
    """Case-insensitive text order."""
    la, lb = fold(a), fold(b)
    if la == lb:
        return 0
    return -1 if la < lb else 1


def compare(a: Value, b: Value) -> int:
    """Excel's comparison for ``=``, ``<``… Text compares case-insensitively; empty acts as 0,
    "" or FALSE depending on the other side. Errors are handled by the caller."""
    ta, tb = type(a), type(b)
    if a is None and b is None:
        return 0
    if a is None:
        if tb is float or tb is int:
            return compare_numbers(0.0, float(b))  # type: ignore[arg-type]
        if tb is str:
            return 0 if b == "" else -1
        if tb is bool:
            return 0 if b is False else -1
        return -1 if _rank(None) < _rank(b) else 1
    if b is None:
        return -compare(b, a)
    if (ta is float or ta is int) and (tb is float or tb is int):
        return compare_numbers(float(a), float(b))  # type: ignore[arg-type]
    if ta is str and tb is str:
        return compare_text(a, b)  # type: ignore[arg-type]
    if ta is bool and tb is bool:
        return (a > b) - (a < b)  # type: ignore[operator]
    if ta is CellError and tb is CellError:
        ca, cb = a.code, b.code  # type: ignore[union-attr]
        return (ca > cb) - (ca < cb)
    ra, rb = _rank(a), _rank(b)
    return (ra > rb) - (ra < rb)


def sort_compare(a: Value, b: Value) -> int:
    """Sorting order used by Sort (ascending): numbers, text, booleans, errors, then blanks."""
    if a is None and b is None:
        return 0
    if a is None:
        return 1
    if b is None:
        return -1
    if isinstance(a, CellError) and isinstance(b, CellError):
        return 0
    return compare(a, b)


def trim_decimal(s: str) -> str:
    """Removes trailing zeros (and a trailing point) from a decimal string."""
    if "." in s:
        t = s.rstrip("0").rstrip(".")
        return "0" if t == "-0" else t
    return s


def sci_text(n: float, sig: int) -> str:
    """Scientific text like ``1.23456789012346E+15`` with up to ``sig`` significant digits."""
    s = f"{n:.{max(sig - 1, 0)}e}"
    mant, _, exp = s.partition("e")
    mant = trim_decimal(mant)
    e = int(exp or 0)
    return f"{mant}E{'-' if e < 0 else '+'}{abs(e):02d}"


def number_to_text(n: float) -> str:
    """Number → text the way Excel converts in ``&`` (Excel 16 for Mac): 15 significant digits,
    written out in full while that takes at most 20 characters, the sign not counted —
    ``1E-18`` is ``0.000000000000000001``, ``2^63`` is ``9223372036854780000``,
    ``1E+15+0.5`` is ``1000000000000000`` — and in scientific notation past that: ``1E-19``,
    ``1E+20``, ``3.33333333333333E-06`` (which would need 22 characters)."""
    if n == 0.0 or not math.isfinite(n):
        return "0"
    mant, _, exp = f"{abs(n):.14e}".partition("e")
    digits = mant.replace(".", "")
    e = int(exp)
    if e >= 14:
        plain = digits + "0" * (e - 14)
    elif e >= 0:
        plain = trim_decimal(digits[: e + 1] + "." + digits[e + 1:])
    else:
        plain = trim_decimal("0." + "0" * (-e - 1) + digits)
    if len(plain) > 20:
        return sci_text(n, 15)
    return ("-" if n < 0 else "") + plain


def flatten_values(vals: Iterable[Value]) -> List[Value]:
    out: List[Value] = []
    for v in vals:
        if isinstance(v, Array):
            out.extend(v.data)
        else:
            out.append(v)
    return out
