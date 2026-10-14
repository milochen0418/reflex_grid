"""VBA values and their rules: Variant subtypes, conversions (Let-coercion), arithmetic,
comparison and string forms, after [MS-VBAL] and observed Excel VBA behaviour.

Values are plain Python objects tagged by type, so the interpreter stays fast:

========  ==================  ==========================================
VarType   VBA type            Python
========  ==================  ==========================================
0         Empty               :data:`EMPTY`
1         Null                :data:`NULL`
2         Integer             :class:`Integer` (an ``int``)
3         Long                ``int``
4         Single              :class:`Single` (a ``float``)
5         Double              ``float``
6         Currency            :class:`Currency` (exact, units of 1/10000)
7         Date                :class:`Date` (a ``float``: OLE date serial)
8         String              ``str``
9         Object              any object; :data:`NOTHING` is Nothing
10        Error               :class:`ErrorValue`; :data:`MISSING`
11        Boolean             ``bool``
14        Decimal             ``decimal.Decimal``
17        Byte                :class:`Byte` (an ``int``)
20        LongLong            :class:`LongLong` (an ``int``)
36        user-defined type   :class:`Record`
8192+     array               :class:`VArray`
========  ==================  ==========================================

Operators take a ``variant`` flag: VBA promotes a result that overflows (Integer → Long →
Double) only when an operand is a Variant; with typed operands (``Dim i As Integer``,
literals) the overflow is error 6, as in ``Debug.Print 200 * 200``.
"""

from __future__ import annotations

import datetime as _dt
import decimal
import math
import re
from typing import Any, Callable, Iterator, List, Optional, Sequence, Tuple

from .errors import VbaError

_DEC = decimal.Context(prec=29, rounding=decimal.ROUND_HALF_EVEN)


class _Special:
    __slots__ = ("name",)

    def __init__(self, name: str):
        self.name = name

    def __repr__(self) -> str:
        return self.name

    def __reduce__(self):
        return (_special, (self.name,))


def _special(name: str) -> "_Special":
    return {"Empty": EMPTY, "Null": NULL, "Nothing": NOTHING, "Missing": MISSING}[name]


EMPTY = _Special("Empty")
NULL = _Special("Null")
NOTHING = _Special("Nothing")
MISSING = _Special("Missing")


class Integer(int):
    __slots__ = ()


class Byte(int):
    __slots__ = ()


class LongLong(int):
    __slots__ = ()


class Single(float):
    __slots__ = ()


class Date(float):
    __slots__ = ()


class Currency:
    """A Currency: a 64-bit integer count of 1/10000ths."""

    __slots__ = ("units",)

    def __init__(self, units: int):
        if not -(1 << 63) <= units < (1 << 63):
            raise VbaError(6)
        self.units = units

    @staticmethod
    def of(x) -> "Currency":
        if isinstance(x, Currency):
            return x
        if isinstance(x, decimal.Decimal):
            q = (x * 10000).to_integral_value(rounding=decimal.ROUND_HALF_EVEN)
            return Currency(int(q))
        if isinstance(x, float):
            if not math.isfinite(x):
                raise VbaError(6)
            d = decimal.Decimal(repr(x)) * 10000
            return Currency(int(d.to_integral_value(rounding=decimal.ROUND_HALF_EVEN)))
        return Currency(int(x) * 10000)

    def value(self) -> decimal.Decimal:
        return decimal.Decimal(self.units) / 10000

    def __float__(self) -> float:
        return self.units / 10000.0

    def __eq__(self, other) -> bool:
        return isinstance(other, Currency) and other.units == self.units

    def __hash__(self) -> int:
        return hash(("cur", self.units))

    def __repr__(self) -> str:
        return f"Currency({currency_text(self)})"


class ErrorValue:
    """A Variant of subtype Error (``CVErr``); Excel's cell errors are 2000-2050."""

    __slots__ = ("code",)

    def __init__(self, code: int):
        self.code = int(code)

    def __eq__(self, other) -> bool:
        return isinstance(other, ErrorValue) and other.code == self.code

    def __hash__(self) -> int:
        return hash(("err", self.code))

    def __repr__(self) -> str:
        return f"Error {self.code}"


# ---------------------------------------------------------------- arrays and records


class VArray:
    """A VBA array: bounds per dimension and the elements in column-major order (the first
    index varies fastest, as VBA stores them — For Each and ReDim Preserve depend on it).
    ``coerce`` converts what is stored into the element type (None: Variant)."""

    __slots__ = ("bounds", "data", "coerce", "etype", "fixed", "locked")

    def __init__(self, bounds: List[Tuple[int, int]], data: Optional[list] = None, coerce: Optional[Callable] = None, etype: str = "Variant", fixed: bool = False, default: Any = None):
        self.bounds = [(int(a), int(b)) for a, b in bounds]
        n = 1
        for lb, ub in self.bounds:
            if ub < lb - 1:
                raise VbaError(9)
            n *= max(ub - lb + 1, 0)
        if n > MAX_ARRAY:
            raise VbaError(7)
        if data is None:
            fill = EMPTY if default is None else default
            if callable(fill):
                data = [fill() for _ in range(n)]
            else:
                data = [fill] * n
        self.data = data
        self.coerce = coerce
        self.etype = etype
        self.fixed = fixed
        self.locked = 0

    @staticmethod
    def from_list(items: Sequence, lb: int = 0) -> "VArray":
        return VArray([(lb, lb + len(items) - 1)], list(items))

    @staticmethod
    def from_rows(rows: Sequence[Sequence], lb: int = 1) -> "VArray":
        """A 2-D array (rows × columns) from row lists, stored column-major."""
        h = len(rows)
        w = len(rows[0]) if h else 0
        data = [None] * (h * w)
        for r, row in enumerate(rows):
            for c, v in enumerate(row):
                data[c * h + r] = v
        return VArray([(lb, lb + h - 1), (lb, lb + w - 1)], data)

    @property
    def dims(self) -> int:
        return len(self.bounds)

    def offset(self, idx: Sequence[int]) -> int:
        b = self.bounds
        if len(idx) != len(b):
            raise VbaError(9)
        off = 0
        stride = 1
        for i, (lb, ub) in zip(idx, b):
            if not lb <= i <= ub:
                raise VbaError(9)
            off += (i - lb) * stride
            stride *= ub - lb + 1
        return off

    def get(self, idx: Sequence[int]) -> Any:
        return self.data[self.offset(idx)]

    def set(self, idx: Sequence[int], v: Any) -> None:
        c = self.coerce
        self.data[self.offset(idx)] = c(v) if c is not None else v

    def rows(self) -> List[list]:
        """A 1-D or 2-D array as row lists (for writing into cells)."""
        if self.dims == 1:
            return [list(self.data)]
        if self.dims != 2:
            raise VbaError(13)
        (r0, r1), (c0, c1) = self.bounds
        h, w = r1 - r0 + 1, c1 - c0 + 1
        return [[self.data[c * h + r] for c in range(w)] for r in range(h)]

    def copy(self) -> "VArray":
        a = VArray.__new__(VArray)
        a.bounds = list(self.bounds)
        a.data = [copy_value(x) for x in self.data]
        a.coerce = self.coerce
        a.etype = self.etype
        a.fixed = False
        a.locked = 0
        return a

    def __iter__(self) -> Iterator[Any]:
        return iter(self.data)

    def __repr__(self) -> str:
        return f"VArray({self.bounds}, {self.data[:20]!r})"


MAX_ARRAY = 20_000_000


class Record:
    """An instance of a user-defined type (``Type … End Type``): copied on assignment."""

    __slots__ = ("typ", "fields")

    def __init__(self, typ, fields: list):
        self.typ = typ
        self.fields = fields

    def copy(self) -> "Record":
        return Record(self.typ, [copy_value(x) for x in self.fields])

    def __repr__(self) -> str:
        return f"Record({self.typ.name})"


def copy_value(v: Any) -> Any:
    """Arrays and user-defined types are values in VBA: assignment copies them."""
    t = type(v)
    if t is VArray or t is Record:
        return v.copy()
    return v


# ---------------------------------------------------------------- type information

INTEGRAL = (Byte, Integer, int, LongLong)
_TYPE_NAMES = {
    Integer: "Integer",
    int: "Long",
    Single: "Single",
    float: "Double",
    Currency: "Currency",
    Date: "Date",
    str: "String",
    bool: "Boolean",
    decimal.Decimal: "Decimal",
    Byte: "Byte",
    LongLong: "LongLong",
    ErrorValue: "Error",
}
_VARTYPES = {
    Integer: 2,
    int: 3,
    Single: 4,
    float: 5,
    Currency: 6,
    Date: 7,
    str: 8,
    ErrorValue: 10,
    bool: 11,
    decimal.Decimal: 14,
    Byte: 17,
    LongLong: 20,
    Record: 36,
}
ETYPE_VARTYPE = {
    "Integer": 2, "Long": 3, "Single": 4, "Double": 5, "Currency": 6, "Date": 7, "String": 8, "Object": 9,
    "Boolean": 11, "Variant": 12, "Decimal": 14, "Byte": 17, "LongLong": 20, "LongPtr": 20,
}


def is_object(v: Any) -> bool:
    return v is NOTHING or not (v is EMPTY or v is NULL or v is MISSING or type(v) in _SCALAR_TYPES or type(v) is VArray or type(v) is Record)


_SCALAR_TYPES = frozenset({Integer, int, Single, float, Currency, Date, str, bool, decimal.Decimal, Byte, LongLong, ErrorValue})


def is_numeric_type(v: Any) -> bool:
    return type(v) in _NUMERIC_TYPES


_NUMERIC_TYPES = frozenset({Integer, int, Single, float, Currency, Date, bool, decimal.Decimal, Byte, LongLong})


def vartype(v: Any) -> int:
    """VarType (a Long, as VBA 7 returns it)."""
    if v is EMPTY:
        return 0
    if v is NULL:
        return 1
    if v is MISSING:
        return 10
    t = type(v)
    vt = _VARTYPES.get(t)
    if vt is not None:
        return vt
    if t is VArray:
        return 8192 + ETYPE_VARTYPE.get(v.etype, 12)
    return 9


def type_name(v: Any) -> str:
    if v is EMPTY:
        return "Empty"
    if v is NULL:
        return "Null"
    if v is NOTHING:
        return "Nothing"
    if v is MISSING:
        return "Error"
    t = type(v)
    n = _TYPE_NAMES.get(t)
    if n is not None:
        return n
    if t is VArray:
        return v.etype + "()"
    if t is Record:
        return v.typ.name
    tn = getattr(v, "vba_type_name", None)
    if tn is not None:
        return tn() if callable(tn) else str(tn)
    return "Object"


# ---------------------------------------------------------------- default members

# Set by the runtime: evaluates an object's default member (its value), raising 438 when it has none.
_default_value: Callable[[Any], Any] = None  # type: ignore[assignment]


def set_default_value_hook(fn: Callable[[Any], Any]) -> None:
    global _default_value
    _default_value = fn


def value_of(v: Any) -> Any:
    """A value for Let-coercion: an object gives its default member's value."""
    if is_object(v):
        if v is NOTHING:
            raise VbaError(91)
        return _default_value(v)
    return v


# ---------------------------------------------------------------- string → number

_NUM_TEXT = re.compile(r"^\s*([+-]?)\s*\$?\s*(\d[\d,]*\.?\d*|\.\d+)(?:[eEdD]([+-]?\d+))?\s*$")
_HEX_TEXT = re.compile(r"^\s*&([HhOo])([0-9A-Fa-f]+)&?\s*$")


def parse_number(s: str) -> Optional[Any]:
    """The number a string holds as VBA reads it (``" 1,234.5 "``, ``"1E3"``, ``"&HFF"``,
    ``"$12"``), as an ``int`` or ``decimal.Decimal``; ``None`` if it isn't one."""
    m = _HEX_TEXT.match(s)
    if m:
        base = 16 if m.group(1) in "Hh" else 8
        try:
            v = int(m.group(2), base)
        except ValueError:
            return None
        if v <= 0xFFFF and not s.strip().endswith("&"):
            return v - 0x10000 if v >= 0x8000 else v
        if v <= 0xFFFFFFFF:
            return v - 0x100000000 if v >= 0x80000000 else v
        return v
    m = _NUM_TEXT.match(s)
    if not m:
        t = s.strip()
        if t[:1] == "(" and t[-1:] == ")":
            inner = parse_number(t[1:-1])
            if inner is not None:
                return -inner
        return None
    sign, body, exp = m.group(1), m.group(2), m.group(3)
    if "," in body:
        intpart = body.split(".")[0]
        groups = intpart.split(",")
        if any(g == "" for g in groups):
            return None
    body = body.replace(",", "")
    try:
        d = decimal.Decimal(body)
    except decimal.InvalidOperation:
        return None
    if exp is not None:
        d = d.scaleb(int(exp))
    if sign == "-":
        d = -d
    if d == d.to_integral_value() and exp is None and "." not in body:
        return int(d)
    return d


def _str_num(s: str) -> Any:
    v = parse_number(s)
    if v is None:
        raise VbaError(13)
    return v


# ---------------------------------------------------------------- conversions (Let-coercion)


def _round_half_even(x: float) -> int:
    if not math.isfinite(x):
        raise VbaError(6)
    return int(round(x))


def _integral(v: Any) -> int:
    """The integer a value rounds to (banker's rounding, as CInt and CLng)."""
    t = type(v)
    if t is int or t is Integer or t is Byte or t is LongLong:
        return int(v)
    if t is bool:
        return -1 if v else 0
    if t is float or t is Single or t is Date:
        return _round_half_even(v)
    if t is Currency:
        q, r = divmod(v.units, 10000)
        if r > 5000 or (r == 5000 and q % 2):
            q += 1
        return q
    if t is decimal.Decimal:
        return int(v.to_integral_value(rounding=decimal.ROUND_HALF_EVEN))
    if t is str:
        x = _str_num(v)
        return _integral(x if not isinstance(x, decimal.Decimal) else x)
    if v is EMPTY:
        return 0
    if v is NULL:
        raise VbaError(94)
    if t is ErrorValue or v is MISSING:
        raise VbaError(13)
    if t is VArray or t is Record:
        raise VbaError(13)
    return _integral(value_of(v))


def to_long(v: Any) -> int:
    t = type(v)
    if t is int:
        return v
    n = _integral(v)
    if not -2147483648 <= n <= 2147483647:
        raise VbaError(6)
    return n


def to_integer(v: Any) -> Integer:
    if type(v) is Integer:
        return v
    n = _integral(v)
    if not -32768 <= n <= 32767:
        raise VbaError(6)
    return Integer(n)


def to_byte(v: Any) -> Byte:
    if type(v) is Byte:
        return v
    n = _integral(v)
    if not 0 <= n <= 255:
        raise VbaError(6)
    return Byte(n)


def to_longlong(v: Any) -> LongLong:
    if type(v) is LongLong:
        return v
    n = _integral(v)
    if not -(1 << 63) <= n < (1 << 63):
        raise VbaError(6)
    return LongLong(n)


def to_double(v: Any) -> float:
    t = type(v)
    if t is float:
        return v
    if t is int or t is Integer or t is Byte or t is LongLong or t is Single or t is Date:
        return float(v)
    if t is bool:
        return -1.0 if v else 0.0
    if t is Currency:
        return v.units / 10000.0
    if t is decimal.Decimal:
        return float(v)
    if t is str:
        x = _str_num(v)
        return float(x)
    if v is EMPTY:
        return 0.0
    if v is NULL:
        raise VbaError(94)
    if t is ErrorValue or v is MISSING or t is VArray or t is Record:
        raise VbaError(13)
    return to_double(value_of(v))


def to_single(v: Any) -> Single:
    if type(v) is Single:
        return v
    x = to_double(v)
    if x != 0 and not math.isinf(x) and abs(x) > 3.4028234663852886e38:
        raise VbaError(6)
    return Single(_f32(x))


def _f32(x: float) -> float:
    import struct

    try:
        return struct.unpack("f", struct.pack("f", x))[0]
    except OverflowError:
        raise VbaError(6)


def to_currency(v: Any) -> Currency:
    t = type(v)
    if t is Currency:
        return v
    if t is str:
        x = _str_num(v)
        return Currency.of(decimal.Decimal(x) if not isinstance(x, (float, Date)) else float(x))
    if t is bool:
        return Currency(-10000 if v else 0)
    if t is float or t is Single or t is Date:
        return Currency.of(float(v))
    if t is decimal.Decimal:
        return Currency.of(v)
    if t in (int, Integer, Byte, LongLong):
        return Currency(int(v) * 10000)
    if v is EMPTY:
        return Currency(0)
    if v is NULL:
        raise VbaError(94)
    if t is ErrorValue or v is MISSING or t is VArray or t is Record:
        raise VbaError(13)
    return to_currency(value_of(v))


def to_decimal(v: Any) -> decimal.Decimal:
    t = type(v)
    if t is decimal.Decimal:
        return v
    if t is Currency:
        return v.value()
    if t is str:
        x = _str_num(v)
        return decimal.Decimal(float(x)) if isinstance(x, Date) else _DEC.create_decimal(x)
    if t is bool:
        return decimal.Decimal(-1 if v else 0)
    if t in (int, Integer, Byte, LongLong):
        return decimal.Decimal(int(v))
    if t is float or t is Single or t is Date:
        if not math.isfinite(v):
            raise VbaError(6)
        return _DEC.create_decimal(repr(float(v)))
    if v is EMPTY:
        return decimal.Decimal(0)
    if v is NULL:
        raise VbaError(94)
    if not is_object(v):
        raise VbaError(13)
    return to_decimal(value_of(v))


def to_boolean(v: Any) -> bool:
    t = type(v)
    if t is bool:
        return v
    if t in (int, Integer, Byte, LongLong, float, Single, Date):
        return v != 0
    if t is Currency:
        return v.units != 0
    if t is decimal.Decimal:
        return v != 0
    if t is str:
        s = v.strip().lower()
        if s == "true" or s == "#true#":
            return True
        if s == "false" or s == "#false#":
            return False
        x = parse_number(v)
        if x is None:
            raise VbaError(13)
        return x != 0
    if v is EMPTY:
        return False
    if v is NULL:
        raise VbaError(94)
    if t is ErrorValue or v is MISSING or t is VArray or t is Record:
        raise VbaError(13)
    return to_boolean(value_of(v))


def to_date(v: Any) -> Date:
    t = type(v)
    if t is Date:
        return v
    if t is str:
        d = parse_date_text(v)
        if d is None:
            x = parse_number(v)
            if x is None:
                raise VbaError(13)
            d = float(x)
        return _check_date(d)
    if t in (int, Integer, Byte, LongLong, float, Single):
        return _check_date(float(v))
    if t is bool:
        return Date(-1.0 if v else 0.0)
    if t is Currency or t is decimal.Decimal:
        return _check_date(float(v))
    if v is EMPTY:
        return Date(0.0)
    if v is NULL:
        raise VbaError(94)
    if t is ErrorValue or v is MISSING or t is VArray or t is Record:
        raise VbaError(13)
    return to_date(value_of(v))


DATE_MIN = -657434.0  # 1/1/100
DATE_MAX = 2958465.9999999  # 12/31/9999 23:59:59


def _check_date(x: float) -> Date:
    if not (DATE_MIN <= x < 2958466.0):
        raise VbaError(6)
    return Date(x)


def to_string(v: Any) -> str:
    t = type(v)
    if t is str:
        return v
    if v is NULL:
        raise VbaError(94)
    return text(v)


def text(v: Any) -> str:
    """A value as CStr shows it (Null and Empty as ``""``; used by ``&`` and Print)."""
    t = type(v)
    if t is str:
        return v
    if t is bool:
        return "True" if v else "False"
    if t is int or t is Integer or t is Byte or t is LongLong:
        return str(int(v))
    if t is float:
        return double_text(v)
    if t is Single:
        return single_text(v)
    if t is Date:
        return date_text(v)
    if t is Currency:
        return currency_text(v)
    if t is decimal.Decimal:
        return decimal_text(v)
    if v is EMPTY or v is NULL:
        return ""
    if t is ErrorValue:
        return f"Error {v.code}"
    if v is MISSING:
        return "Error 448"
    if t is VArray or t is Record:
        raise VbaError(13)
    return text(value_of(v))


def to_variant(v: Any) -> Any:
    return v


def to_object(v: Any) -> Any:
    if v is NOTHING or is_object(v):
        return v
    raise VbaError(424)


def fixed_string(n: int) -> Callable[[Any], str]:
    """``String * n``: padded with spaces or cut to ``n`` characters."""

    def co(v: Any) -> str:
        s = to_string(v)
        return s[:n] if len(s) >= n else s + " " * (n - len(s))

    return co


COERCE = {
    "Integer": to_integer,
    "Long": to_long,
    "LongLong": to_longlong,
    "LongPtr": to_longlong,
    "Byte": to_byte,
    "Single": to_single,
    "Double": to_double,
    "Currency": to_currency,
    "Decimal": to_decimal,
    "Date": to_date,
    "String": to_string,
    "Boolean": to_boolean,
    "Variant": None,
    "Any": None,
}
DEFAULTS = {
    "Integer": Integer(0),
    "Long": 0,
    "LongLong": LongLong(0),
    "LongPtr": LongLong(0),
    "Byte": Byte(0),
    "Single": Single(0.0),
    "Double": 0.0,
    "Currency": Currency(0),
    "Decimal": decimal.Decimal(0),
    "Date": Date(0.0),
    "String": "",
    "Boolean": False,
    "Variant": EMPTY,
    "Object": NOTHING,
}


# ---------------------------------------------------------------- number text


def double_text(x: float) -> str:
    """A Double as VBA prints it: up to 15 significant digits, scientific (``1E+15``,
    ``1E-05``) outside 1E-4 … 1E15. Negative zero prints as ``-0``, as in VBA."""
    if x == 0:
        return "-0" if math.copysign(1.0, x) < 0 else "0"
    if not math.isfinite(x):
        return "1.#INF" if x > 0 else ("-1.#INF" if x < 0 else "-1.#IND")
    return _sig_text(x, 15)


def single_text(x: float) -> str:
    if x == 0:
        return "0"
    return _sig_text(float(x), 7)


def _sig_text(x: float, digits: int) -> str:
    s = f"{abs(x):.{digits - 1}e}"
    mant, exp = s.split("e")
    e = int(exp)
    mant = mant.replace(".", "")
    mant = mant.rstrip("0") or "0"
    neg = "-" if x < 0 else ""
    if e >= digits or e < -4:
        m = mant[0] + ("." + mant[1:] if len(mant) > 1 else "")
        return f"{neg}{m}E{'+' if e >= 0 else '-'}{abs(e):02d}"
    if e >= 0:
        if len(mant) <= e + 1:
            return neg + mant + "0" * (e + 1 - len(mant))
        return neg + mant[:e + 1] + "." + mant[e + 1:]
    return neg + "0." + "0" * (-e - 1) + mant


def currency_text(c: Currency) -> str:
    neg = c.units < 0
    q, r = divmod(abs(c.units), 10000)
    s = str(q)
    if r:
        s += "." + f"{r:04d}".rstrip("0")
    return ("-" if neg else "") + s


def decimal_text(d: decimal.Decimal) -> str:
    if d == 0:
        return "0"
    s = format(d.normalize(), "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s


# ---------------------------------------------------------------- dates

_EPOCH = _dt.datetime(1899, 12, 30)
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]


def serial_to_datetime(x: float) -> _dt.datetime:
    """An OLE date serial as a datetime. Negative serials count days back from 12/30/1899 and
    keep the time of day positive (-1.25 is 12/29/1899 6:00 AM)."""
    if not (DATE_MIN <= x < 2958466.0):
        raise VbaError(6)
    days = math.floor(x) if x >= 0 else math.ceil(x)
    frac = abs(x - days)
    secs = round(frac * 86400.0)
    if secs >= 86400:
        secs = 86399 if days == int(DATE_MAX) else secs
        if secs >= 86400:
            days += 1 if x >= 0 else -1
            secs -= 86400
    try:
        return _EPOCH + _dt.timedelta(days=days, seconds=secs)
    except OverflowError:
        raise VbaError(6)


def datetime_to_serial(d: _dt.datetime) -> float:
    delta = d - _EPOCH
    days = delta.days
    secs = delta.seconds + delta.microseconds / 1e6
    if days < 0 and secs:
        # Before the epoch the time of day is added "away from zero".
        return float(days + 1) - secs / 86400.0
    return days + secs / 86400.0


def serial_from_parts(y: int, m: int, d: int, hh: int = 0, mm: int = 0, ss: float = 0.0) -> float:
    if not 100 <= y <= 9999:
        raise VbaError(5)
    base = _dt.datetime(y, m, d)
    days = (base - _EPOCH).days
    t = (hh * 3600 + mm * 60 + ss) / 86400.0
    return days + t if days >= 0 else days - t


def date_serial(y: int, m: int, d: int) -> float:
    """DateSerial: months and days roll over (DateSerial(2020, 13, 1) is 1/1/2021)."""
    if 0 <= y <= 29:
        y += 2000
    elif 30 <= y <= 99:
        y += 1900
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    if not 100 <= y <= 9999:
        raise VbaError(5)
    base = _dt.date(y, m, 1)
    days = (base - _EPOCH.date()).days + (d - 1)
    if not DATE_MIN <= days <= DATE_MAX:
        raise VbaError(5)
    return float(days)


def time_serial(h: int, m: int, s: int) -> float:
    total = h * 3600 + m * 60 + s
    v = total / 86400.0
    if not -657434 <= v <= 2958466:
        raise VbaError(5)
    return v


def date_text(x: float) -> str:
    """A Date as CStr shows it in the US locale: ``1/2/2020``, ``10:30:00 AM`` or both. A
    date part of zero (12/30/1899) shows the time alone, midnight included."""
    d = serial_to_datetime(x)
    date_part = int(x) != 0
    has_time = x != int(x)
    out = []
    if date_part:
        out.append(f"{d.month}/{d.day}/{d.year}")
    if has_time or not date_part:
        out.append(time_text(d))
    return " ".join(out)


def time_text(d: _dt.datetime) -> str:
    h = d.hour % 12 or 12
    return f"{h}:{d.minute:02d}:{d.second:02d} {'AM' if d.hour < 12 else 'PM'}"


_MONTH_ABBR = {m[:3].lower(): i + 1 for i, m in enumerate(MONTHS)}
_MONTH_FULL = {m.lower(): i + 1 for i, m in enumerate(MONTHS)}
_TIME_RE = re.compile(r"^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}(?:\.\d+)?))?\s*([AaPp][Mm]?)?$")


def _year(y: int, two_digit: bool) -> int:
    if two_digit:
        return 2000 + y if y < 30 else 1900 + y
    return y


def parse_time_text(s: str) -> Optional[float]:
    m = _TIME_RE.match(s.strip())
    if not m:
        return None
    if m.group(2) is None and m.group(4) is None:
        return None
    h = int(m.group(1))
    mi = int(m.group(2) or 0)
    sec = float(m.group(3) or 0)
    ap = (m.group(4) or "").upper()
    if ap:
        if not 1 <= h <= 12:
            return None
        if ap.startswith("P") and h != 12:
            h += 12
        if ap.startswith("A") and h == 12:
            h = 0
    if h > 23 or mi > 59 or sec >= 60:
        return None
    return (h * 3600 + mi * 60 + sec) / 86400.0


def parse_date_text(s: str) -> Optional[float]:
    """A date/time string as VBA's CDate reads it in the US locale; ``None`` if it isn't one."""
    t = " ".join(s.strip().replace(",", " ").split())
    if not t:
        return None
    # A trailing time.
    time_part = 0.0
    has_time = False
    m = re.match(r"^(.*?)\s*(\d{1,2}:\d{1,2}(?::\d{1,2}(?:\.\d+)?)?\s*(?:[AaPp][Mm]?)?|\d{1,2}\s*[AaPp][Mm])$", t)
    if m and (":" in m.group(2) or re.search(r"[AaPp]", m.group(2))):
        tv = parse_time_text(m.group(2))
        if tv is not None:
            time_part, has_time = tv, True
            t = m.group(1).strip()
            if not t:
                return time_part
    d = _parse_date_only(t)
    if d is None:
        return None
    return d + time_part if d >= 0 else d - time_part


def _parse_date_only(t: str) -> Optional[float]:
    parts = re.split(r"[/\-. ]+", t)
    try:
        if len(parts) == 3 and all(p.isdigit() for p in parts):
            a, b, c = (int(p) for p in parts)
            if len(parts[0]) == 4:  # yyyy-mm-dd
                return _ymd(a, b, c)
            # m/d/y; when that isn't a date, d/m/y (as VBA does)
            y = _year(c, len(parts[2]) <= 2)
            v = _ymd(y, a, b)
            if v is None:
                v = _ymd(y, b, a)
            return v
        if len(parts) == 2 and all(p.isdigit() for p in parts):
            a, b = int(parts[0]), int(parts[1])
            today = _dt.date.today()
            if len(parts[1]) == 4 and 1 <= a <= 12:  # m/yyyy
                return _ymd(b, a, 1)
            v = _ymd(today.year, a, b)
            if v is None and 1 <= a <= 12:
                return _ymd(_year(b, len(parts[1]) <= 2), a, 1)
            return v
        words = [p for p in parts if p]
        if 2 <= len(words) <= 4:
            month = None
            nums: List[str] = []
            for w in words:
                lw = w.lower()
                if lw in _MONTH_FULL or lw[:3] in _MONTH_ABBR and lw.isalpha():
                    if month is not None:
                        return None
                    month = _MONTH_FULL.get(lw) or _MONTH_ABBR.get(lw[:3])
                elif lw in [d.lower() for d in DAYS] or lw[:3] in [d[:3].lower() for d in DAYS]:
                    continue
                elif w.isdigit():
                    nums.append(w)
                else:
                    return None
            if month is None or not nums:
                return None
            if len(nums) == 1:
                n0 = int(nums[0])
                if len(nums[0]) == 4:
                    return _ymd(n0, month, 1)
                return _ymd(_dt.date.today().year, month, n0)
            if len(nums) == 2:
                a, b = nums
                if len(a) == 4:
                    return _ymd(int(a), month, int(b))
                return _ymd(_year(int(b), len(b) <= 2), month, int(a))
    except (ValueError, OverflowError):
        return None
    return None


def _ymd(y: int, m: int, d: int) -> Optional[float]:
    try:
        if not 100 <= y <= 9999:
            return None
        return float((_dt.date(y, m, d) - _EPOCH.date()).days)
    except ValueError:
        return None


def parse_date_literal(s: str) -> Optional[Date]:
    v = parse_date_text(s)
    return Date(v) if v is not None else None


# ---------------------------------------------------------------- arithmetic

# Ranks for result types: the larger operand type wins.
_R = {Byte: 1, bool: 2, Integer: 2, int: 3, LongLong: 4, Single: 5, float: 6, Currency: 7, decimal.Decimal: 8}
_RANK_TYPE = {1: Byte, 2: Integer, 3: int, 4: LongLong, 5: Single, 6: float, 7: Currency, 8: decimal.Decimal}
_LIMITS = {1: (0, 255), 2: (-32768, 32767), 3: (-2147483648, 2147483647), 4: (-(1 << 63), (1 << 63) - 1)}
_PROMOTE = {1: 2, 2: 3, 3: 6, 4: 6, 5: 6}


def _num_operand(v: Any) -> Tuple[Any, int]:
    """(value, rank) of an arithmetic operand: strings read as Double, Empty as Integer 0,
    Boolean as Integer; Dates keep rank 6 with a date marker handled by the callers."""
    t = type(v)
    r = _R.get(t)
    if r is not None:
        if t is bool:
            return Integer(-1 if v else 0), 2
        return v, r
    if t is Date:
        return float(v), 6
    if t is str:
        x = _str_num(v)
        if isinstance(x, Date):
            return float(x), 6
        return float(x), 6
    if v is EMPTY:
        return Integer(0), 2
    if t is ErrorValue or v is MISSING or t is VArray or t is Record:
        raise VbaError(13)
    if v is NULL:
        return NULL, 0
    return _num_operand(value_of(v))


def _result_rank(ra: int, rb: int) -> int:
    r = max(ra, rb)
    if r == 5 and min(ra, rb) in (3, 4):
        return 6
    return r


def _fit(x: Any, rank: int, variant: bool) -> Any:
    """``x`` (a Python number) as the result type ``rank``, promoting on overflow for Variants."""
    while True:
        if rank in _LIMITS:
            lo, hi = _LIMITS[rank]
            if lo <= x <= hi:
                return _RANK_TYPE[rank](x) if rank != 3 else int(x)
            if not variant:
                raise VbaError(6)
            rank = _PROMOTE[rank]
            if rank == 6:
                return float(x)
            continue
        if rank == 5:
            f = float(x)
            if abs(f) > 3.4028234663852886e38:
                if not variant:
                    raise VbaError(6)
                return _check_double(f)
            return Single(_f32(f))
        if rank == 6:
            return _check_double(float(x))
        if rank == 7:
            return Currency.of(x if isinstance(x, decimal.Decimal) else (decimal.Decimal(x) if isinstance(x, int) else float(x)))
        if rank == 8:
            d = x if isinstance(x, decimal.Decimal) else decimal.Decimal(x)
            if abs(d) >= decimal.Decimal("79228162514264337593543950336"):
                raise VbaError(6)
            return _DEC.plus(d)
        return x


def _check_double(f: float) -> float:
    if math.isinf(f) or math.isnan(f):
        raise VbaError(6)
    return f


def _as_py(v: Any, rank: int) -> Any:
    if rank <= 4:
        return int(v)
    if rank in (5, 6):
        return float(v)
    if rank == 7:
        return to_currency(v).value()
    return to_decimal(v)


def add(a: Any, b: Any, variant: bool = True) -> Any:
    ta, tb = type(a), type(b)
    if ta is int and tb is int:
        r = a + b
        if -2147483648 <= r <= 2147483647:
            return r
        if variant:
            return float(r)
        raise VbaError(6)
    if ta is float and tb is float:
        return _check_double(a + b)
    if a is NULL or b is NULL:
        return NULL
    if ta is str and tb is str:
        return a + b
    if ta is str and b is EMPTY:
        return a
    if a is EMPTY and tb is str:
        return b
    if is_object(a):
        return add(value_of(a), b, True)
    if is_object(b):
        return add(a, value_of(b), True)
    if ta is Date or tb is Date:
        x = to_double(a) + to_double(b)
        return _check_date(x)
    x, ra = _num_operand(a)
    y, rb = _num_operand(b)
    if x is NULL or y is NULL:
        return NULL
    r = _result_rank(ra, rb)
    return _fit(_as_py(x, r) + _as_py(y, r), r, variant)


def sub(a: Any, b: Any, variant: bool = True) -> Any:
    ta, tb = type(a), type(b)
    if ta is int and tb is int:
        r = a - b
        if -2147483648 <= r <= 2147483647:
            return r
        if variant:
            return float(r)
        raise VbaError(6)
    if ta is float and tb is float:
        return _check_double(a - b)
    if a is NULL or b is NULL:
        return NULL
    if is_object(a):
        return sub(value_of(a), b, True)
    if is_object(b):
        return sub(a, value_of(b), True)
    if ta is Date and tb is Date:
        return float(a) - float(b)
    if ta is Date or tb is Date:
        return _check_date(to_double(a) - to_double(b))
    x, ra = _num_operand(a)
    y, rb = _num_operand(b)
    if x is NULL or y is NULL:
        return NULL
    r = _result_rank(ra, rb)
    return _fit(_as_py(x, r) - _as_py(y, r), r, variant)


def mul(a: Any, b: Any, variant: bool = True) -> Any:
    ta, tb = type(a), type(b)
    if ta is int and tb is int:
        r = a * b
        if -2147483648 <= r <= 2147483647:
            return r
        if variant:
            return float(r)
        raise VbaError(6)
    if ta is float and tb is float:
        return _check_double(a * b)
    if a is NULL or b is NULL:
        return NULL
    if is_object(a):
        return mul(value_of(a), b, True)
    if is_object(b):
        return mul(a, value_of(b), True)
    x, ra = _num_operand(a)
    y, rb = _num_operand(b)
    if x is NULL or y is NULL:
        return NULL
    r = _result_rank(ra, rb)
    if r == 7:
        if min(ra, rb) in (5, 6) and ra != rb:
            # Currency times a Single or Double is a Double.
            return _check_double(float(x) * float(y))
        p = _as_py(x, 7) * _as_py(y, 7)
        return _fit(p, 7, variant)
    return _fit(_as_py(x, r) * _as_py(y, r), r, variant)


def div(a: Any, b: Any, variant: bool = True) -> Any:
    if a is NULL or b is NULL:
        return NULL
    if is_object(a):
        return div(value_of(a), b)
    if is_object(b):
        return div(a, value_of(b))
    x, ra = _num_operand(a)
    y, rb = _num_operand(b)
    if x is NULL or y is NULL:
        return NULL
    if ra == 8 or rb == 8:
        dx, dy = to_decimal(x), to_decimal(y)
        if dy == 0:
            raise VbaError(6 if dx == 0 else 11)
        return _fit(_DEC.divide(dx, dy), 8, False)
    fx = float(x) if not isinstance(x, Currency) else x.units / 10000.0
    fy = float(y) if not isinstance(y, Currency) else y.units / 10000.0
    if fy == 0:
        raise VbaError(6 if fx == 0 else 11)
    q = fx / fy
    if ra <= 5 and rb <= 5 and (ra == 5 or rb == 5) and max(ra, rb) == 5 and min(ra, rb) <= 2:
        return _fit(q, 5, True)
    if ra == 5 and rb == 5:
        return _fit(q, 5, True)
    return _check_double(q)


def _int_operand(v: Any) -> Tuple[int, int]:
    t = type(v)
    if t is Byte:
        return int(v), 1
    if t is Integer:
        return int(v), 2
    if t is int:
        return v, 3
    if t is LongLong:
        return int(v), 4
    if t is bool:
        return (-1 if v else 0), 2
    if v is EMPTY:
        return 0, 2
    if is_object(v):
        return _int_operand(value_of(v))
    n = _integral(v)
    if not -2147483648 <= n <= 2147483647:
        raise VbaError(6)
    return n, 3


def intdiv(a: Any, b: Any, variant: bool = True) -> Any:
    if a is NULL or b is NULL:
        return NULL
    x, ra = _int_operand(a)
    y, rb = _int_operand(b)
    if y == 0:
        raise VbaError(11)
    q = abs(x) // abs(y)
    if (x < 0) != (y < 0):
        q = -q
    return _fit(q, max(ra, rb), variant)


def mod(a: Any, b: Any, variant: bool = True) -> Any:
    if a is NULL or b is NULL:
        return NULL
    x, ra = _int_operand(a)
    y, rb = _int_operand(b)
    if y == 0:
        raise VbaError(11)
    r = abs(x) % abs(y)
    if x < 0:
        r = -r
    return _fit(r, max(ra, rb), variant)


def power(a: Any, b: Any, variant: bool = True) -> Any:
    if a is NULL or b is NULL:
        return NULL
    x, y = to_double(value_of(a) if is_object(a) else a), to_double(value_of(b) if is_object(b) else b)
    if x == 0 and y < 0:
        raise VbaError(5)
    if x < 0 and y != int(y):
        raise VbaError(5)
    try:
        r = math.pow(x, y)
    except OverflowError:
        raise VbaError(6)
    except ValueError:
        raise VbaError(5)
    return _check_double(r)


def neg(a: Any, variant: bool = True) -> Any:
    t = type(a)
    if t is int:
        if a == -2147483648:
            if variant:
                return 2147483648.0
            raise VbaError(6)
        return -a
    if t is float:
        return -a
    if a is NULL:
        return NULL
    if a is EMPTY:
        return Integer(0)
    if t is Date:
        return _check_date(-float(a))
    if t is Currency:
        return Currency(-a.units)
    if t is decimal.Decimal:
        return -a
    if is_object(a):
        return neg(value_of(a), True)
    if t is Byte:
        return Integer(-int(a))
    x, r = _num_operand(a)
    return _fit(-_as_py(x, r), r, variant)


def concat(a: Any, b: Any) -> Any:
    if a is NULL and b is NULL:
        return NULL
    ta, tb = type(a), type(b)
    sa = a if ta is str else ("" if a is NULL else text(a))
    sb = b if tb is str else ("" if b is NULL else text(b))
    return sa + sb


# ---------------------------------------------------------------- logical and bitwise


def _logic_operand(v: Any) -> Tuple[Any, int]:
    t = type(v)
    if t is bool:
        return v, 0
    if v is NULL:
        return NULL, -1
    if t is str:
        s = v.strip().lower()
        if s in ("true", "false"):
            return s == "true", 0
    n, r = _int_operand(v)
    return n, r


def _bool_or_int(x: int, ra: int, rb: int, variant: bool) -> Any:
    if ra == 0 and rb == 0:
        return bool(x)
    r = max(max(ra, 2), max(rb, 2)) if (ra == 0 or rb == 0) else max(ra, rb)
    return _fit(x, r, True)


def _as_int(x: Any) -> int:
    return (-1 if x else 0) if type(x) is bool else int(x)


def op_and(a: Any, b: Any, variant: bool = True) -> Any:
    x, ra = _logic_operand(a)
    y, rb = _logic_operand(b)
    if ra < 0 or rb < 0:
        other = y if ra < 0 else x
        if other is NULL:
            return NULL
        if _as_int(other) == 0:
            return False if type(other) is bool else _bool_or_int(0, max(ra, rb, 2), max(ra, rb, 2), variant)
        return NULL
    return _bool_or_int(_as_int(x) & _as_int(y), ra, rb, variant)


def op_or(a: Any, b: Any, variant: bool = True) -> Any:
    x, ra = _logic_operand(a)
    y, rb = _logic_operand(b)
    if ra < 0 or rb < 0:
        other = y if ra < 0 else x
        if other is NULL:
            return NULL
        if type(other) is bool and other:
            return True
        if type(other) is not bool and _as_int(other) == -1:
            return _bool_or_int(-1, max(ra, rb), max(ra, rb), variant)
        return NULL
    return _bool_or_int(_as_int(x) | _as_int(y), ra, rb, variant)


def op_xor(a: Any, b: Any, variant: bool = True) -> Any:
    x, ra = _logic_operand(a)
    y, rb = _logic_operand(b)
    if ra < 0 or rb < 0:
        return NULL
    return _bool_or_int(_as_int(x) ^ _as_int(y), ra, rb, variant)


def op_eqv(a: Any, b: Any, variant: bool = True) -> Any:
    x, ra = _logic_operand(a)
    y, rb = _logic_operand(b)
    if ra < 0 or rb < 0:
        return NULL
    return _bool_or_int(~(_as_int(x) ^ _as_int(y)), ra, rb, variant)


def op_imp(a: Any, b: Any, variant: bool = True) -> Any:
    x, ra = _logic_operand(a)
    y, rb = _logic_operand(b)
    if ra < 0 and rb < 0:
        return NULL
    if ra < 0:
        return True if (type(y) is bool and y) or (type(y) is not bool and _as_int(y) == -1) else NULL
    if rb < 0:
        return True if _as_int(x) == 0 else NULL
    return _bool_or_int((~_as_int(x)) | _as_int(y), ra, rb, variant)


def op_not(a: Any, variant: bool = True) -> Any:
    x, r = _logic_operand(a)
    if r < 0:
        return NULL
    if r == 0:
        return not x
    return _fit(~int(x), r, True)


# ---------------------------------------------------------------- comparison


def _num_value(v: Any) -> Any:
    t = type(v)
    if t is bool:
        return -1 if v else 0
    if t is Currency:
        return v.value()
    if t is decimal.Decimal:
        return v
    if t in (int, Integer, Byte, LongLong):
        return int(v)
    return float(v)


def _cmp_nums(a: Any, b: Any) -> int:
    x, y = _num_value(a), _num_value(b)
    if isinstance(x, decimal.Decimal) or isinstance(y, decimal.Decimal):
        x = x if isinstance(x, decimal.Decimal) else decimal.Decimal(repr(x)) if isinstance(x, float) else decimal.Decimal(x)
        y = y if isinstance(y, decimal.Decimal) else decimal.Decimal(repr(y)) if isinstance(y, float) else decimal.Decimal(y)
    return (x > y) - (x < y)


def _cmp_text(a: str, b: str, text_mode: bool) -> int:
    if text_mode:
        a, b = a.lower(), b.lower()
    return (a > b) - (a < b)


def compare(a: Any, b: Any, text_mode: bool = False, va: bool = True, vb: bool = True) -> Any:
    """-1, 0 or 1, or :data:`NULL`. ``va`` and ``vb``: whether each side is a Variant, which
    decides what happens when a number meets a string (see VBA's comparison rules)."""
    if a is NULL or b is NULL:
        return NULL
    if is_object(a):
        a = value_of(a)
        va = True
    if is_object(b):
        b = value_of(b)
        vb = True
    if a is NULL or b is NULL:
        return NULL
    ta, tb = type(a), type(b)
    na, nb = ta in _NUMERIC_TYPES, tb in _NUMERIC_TYPES
    if na and nb:
        return _cmp_nums(a, b)
    sa, sb = ta is str, tb is str
    if sa and sb:
        return _cmp_text(a, b, text_mode)
    if a is EMPTY and b is EMPTY:
        return 0
    if a is EMPTY:
        return _cmp_text("", b, text_mode) if sb else (_cmp_nums(0, b) if nb else _cmp_err())
    if b is EMPTY:
        return _cmp_text(a, "", text_mode) if sa else (_cmp_nums(a, 0) if na else _cmp_err())
    if (na and sb) or (sa and nb):
        num, s, num_variant, s_variant, flip = (a, b, va, vb, False) if na else (b, a, vb, va, True)
        if num_variant and s_variant:
            # Both Variants: numbers sort before strings.
            r = -1
        elif not s_variant and not num_variant:
            x = _str_num(s)
            r = _cmp_nums(num, x)
        elif s_variant:
            # A typed number and a Variant string: numerically when the string is a number,
            # else the number sorts first (as for two Variants).
            x = parse_number(s)
            if x is None:
                d = parse_date_text(s)
                x = d
            r = _cmp_nums(num, x) if x is not None else -1
        else:
            # A typed string and a Variant number: compared as strings.
            r = _cmp_text(text(num), s, text_mode)
        return -r if flip else r
    if ta is ErrorValue and tb is ErrorValue:
        return (a.code > b.code) - (a.code < b.code)
    raise VbaError(13)


def _cmp_err() -> int:
    raise VbaError(13)


def eq(a, b, t=False, va=True, vb=True):
    r = compare(a, b, t, va, vb)
    return NULL if r is NULL else r == 0


def ne(a, b, t=False, va=True, vb=True):
    r = compare(a, b, t, va, vb)
    return NULL if r is NULL else r != 0


def lt(a, b, t=False, va=True, vb=True):
    r = compare(a, b, t, va, vb)
    return NULL if r is NULL else r < 0


def le(a, b, t=False, va=True, vb=True):
    r = compare(a, b, t, va, vb)
    return NULL if r is NULL else r <= 0


def gt(a, b, t=False, va=True, vb=True):
    r = compare(a, b, t, va, vb)
    return NULL if r is NULL else r > 0


def ge(a, b, t=False, va=True, vb=True):
    r = compare(a, b, t, va, vb)
    return NULL if r is NULL else r >= 0


_LIKE_CACHE: dict = {}


def like(s: Any, pattern: Any, text_mode: bool = False) -> Any:
    """The ``Like`` operator: ``?`` ``*`` ``#`` ``[a-z]`` ``[!a-z]``."""
    if s is NULL or pattern is NULL:
        return NULL
    s = to_string(value_of(s) if is_object(s) else s)
    pattern = to_string(value_of(pattern) if is_object(pattern) else pattern)
    key = (pattern, text_mode)
    rx = _LIKE_CACHE.get(key)
    if rx is None:
        rx = _like_regex(pattern, text_mode)
        if len(_LIKE_CACHE) > 2000:
            _LIKE_CACHE.clear()
        _LIKE_CACHE[key] = rx
    return rx.fullmatch(s) is not None


def _like_regex(p: str, text_mode: bool):
    out = []
    i = 0
    n = len(p)
    while i < n:
        c = p[i]
        if c == "?":
            out.append(".")
        elif c == "*":
            out.append(".*")
        elif c == "#":
            out.append("[0-9]")
        elif c == "[":
            j = p.find("]", i + 1)
            if j < 0:
                raise VbaError(93)
            body = p[i + 1:j]
            neg_ = body.startswith("!")
            if neg_:
                body = body[1:]
            if body == "" and not neg_:
                out.append("(?:)")
            else:
                parts = []
                k = 0
                while k < len(body):
                    if k + 2 < len(body) and body[k + 1] == "-":
                        lo, hi = body[k], body[k + 2]
                        if lo > hi:
                            raise VbaError(93)
                        parts.append(re.escape(lo) + "-" + re.escape(hi))
                        k += 3
                    else:
                        parts.append(re.escape(body[k]))
                        k += 1
                out.append("[" + ("^" if neg_ else "") + "".join(parts) + "]")
            i = j
        else:
            out.append(re.escape(c))
        i += 1
    flags = re.DOTALL | (re.IGNORECASE if text_mode else 0)
    return re.compile("".join(out), flags)


def is_true(v: Any) -> bool:
    """A condition (If, While, Do…): Null counts as False, as in VBA."""
    t = type(v)
    if t is bool:
        return v
    if v is NULL:
        return False
    return to_boolean(v)
