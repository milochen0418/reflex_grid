"""The VBA library: string, math, conversion, date and information functions, Format, MsgBox
and InputBox (through the host), Debug.Print, Rnd (VBA's own generator, so a sequence without
``Randomize`` matches Excel's) and the ``vb…`` constants.

Functions get evaluated arguments (MISSING for omitted optional ones) and a call context with
the calling module's ``Option Base`` and ``Option Compare``. The sandbox has no files, shell,
registry or network: those functions raise a clear "not available" error.
"""

from __future__ import annotations

import datetime as _dt
import decimal
import math
import re
import struct
import time
from typing import Any, Callable, Dict, List, Optional, Tuple

from . import values as V
from .errors import NotSupported, VbaError
from .values import EMPTY, MISSING, NOTHING, NULL, Byte, Currency, Date, ErrorValue, Integer, LongLong, Single, VArray, is_object


class Builtin:
    __slots__ = ("name", "fn", "ret", "params", "min_args", "max_args", "refs", "const_ok")

    def __init__(self, name: str, fn: Callable, ret: str, params: tuple, min_args: int, max_args: int, const_ok: bool):
        self.name = name
        self.fn = fn
        self.ret = ret
        self.params = params
        self.min_args = min_args
        self.max_args = max_args
        self.refs = False
        self.const_ok = const_ok


BUILTINS: Dict[str, Builtin] = {}


def bi(name: str, *params: str, ret: str = "Variant", const_ok: bool = False, names: tuple = ()):
    """Registers a library function. Parameters starting with ``?`` are optional; ``*`` takes
    any number of arguments."""

    def deco(fn: Callable) -> Callable:
        mn = sum(1 for p in params if not p.startswith("?") and not p.startswith("*"))
        mx = -1 if any(p.startswith("*") for p in params) else len(params)
        clean = tuple(p.lstrip("?*") for p in params)
        for n in (name,) + tuple(names):
            BUILTINS[n.lower()] = Builtin(n, fn, ret if not n.endswith("$") else "String", clean, mn, mx, const_ok)
        return fn

    return deco


def lookup(name: str) -> Optional[Builtin]:
    return BUILTINS.get(name.lower())


def _arg(args: List[Any], i: int, default: Any = MISSING) -> Any:
    if i < len(args):
        v = args[i]
        return default if v is MISSING else v
    return default


def _val(v: Any) -> Any:
    return V.value_of(v) if is_object(v) else v


def _s(v: Any) -> str:
    return V.to_string(_val(v))


def _l(v: Any) -> int:
    return V.to_long(_val(v))


def _d(v: Any) -> float:
    return V.to_double(_val(v))


def _compare_mode(args: List[Any], i: int, ctx) -> bool:
    c = _arg(args, i)
    if c is MISSING:
        return bool(ctx.compare_text) if ctx is not None else False
    n = _l(c)
    if n not in (-1, 0, 1, 2):
        raise VbaError(5)
    if n == -1:
        return bool(ctx.compare_text) if ctx is not None else False
    return n == 1


# ---------------------------------------------------------------- strings


@bi("Len", "Expression", ret="Long")
def f_len(ex, args, ctx):
    v = _val(args[0])
    if v is NULL:
        return NULL
    if type(v) is V.Record:
        return sum(len(V.text(x)) if isinstance(x, str) else 4 for x in v.fields)
    return len(V.text(v))


@bi("LenB", "Expression", ret="Long")
def f_lenb(ex, args, ctx):
    v = _val(args[0])
    if v is NULL:
        return NULL
    return 2 * len(V.text(v))


def _null_or(fn: Callable[[Any], Any], dollar: bool = False) -> Callable:
    def f(ex, args, ctx):
        v = _val(args[0])
        if v is NULL:
            if dollar:
                raise VbaError(94)
            return NULL
        return fn(args)

    return f


def _left(args):
    n = _l(args[1])
    if n < 0:
        raise VbaError(5)
    return V.text(_val(args[0]))[:n]


def _right(args):
    n = _l(args[1])
    if n < 0:
        raise VbaError(5)
    s = V.text(_val(args[0]))
    return s[len(s) - n:] if n < len(s) else s


def _mid(args):
    s = V.text(_val(args[0]))
    start = _l(args[1])
    if start < 1:
        raise VbaError(5)
    ln = _arg(args, 2)
    if ln is MISSING:
        return s[start - 1:]
    n = _l(ln)
    if n < 0:
        raise VbaError(5)
    return s[start - 1:start - 1 + n]


bi("Left", "String", "Length")(_null_or(_left))
bi("Left$", "String", "Length")(_null_or(_left, True))
bi("Right", "String", "Length")(_null_or(_right))
bi("Right$", "String", "Length")(_null_or(_right, True))
bi("Mid", "String", "Start", "?Length")(_null_or(_mid))
bi("Mid$", "String", "Start", "?Length")(_null_or(_mid, True))
bi("LeftB", "String", "Length")(_null_or(lambda a: V.text(_val(a[0])).encode("utf-16-le")[:_l(a[1])].decode("utf-16-le", "ignore")))
bi("RightB", "String", "Length")(_null_or(lambda a: (lambda b, n: b[len(b) - n:] if n < len(b) else b)(V.text(_val(a[0])).encode("utf-16-le"), _l(a[1])).decode("utf-16-le", "ignore")))
bi("MidB", "String", "Start", "?Length")(_null_or(lambda a: (lambda b, st, n: b[st - 1:(st - 1 + n) if n is not None else None])(V.text(_val(a[0])).encode("utf-16-le"), _l(a[1]), _l(a[2]) if _arg(a, 2) is not MISSING else None).decode("utf-16-le", "ignore")))
bi("UCase", "String")(_null_or(lambda a: V.text(_val(a[0])).upper()))
bi("UCase$", "String")(_null_or(lambda a: V.text(_val(a[0])).upper(), True))
bi("LCase", "String")(_null_or(lambda a: V.text(_val(a[0])).lower()))
bi("LCase$", "String")(_null_or(lambda a: V.text(_val(a[0])).lower(), True))
bi("Trim", "String")(_null_or(lambda a: V.text(_val(a[0])).strip(" ")))
bi("Trim$", "String")(_null_or(lambda a: V.text(_val(a[0])).strip(" "), True))
bi("LTrim", "String")(_null_or(lambda a: V.text(_val(a[0])).lstrip(" ")))
bi("LTrim$", "String")(_null_or(lambda a: V.text(_val(a[0])).lstrip(" "), True))
bi("RTrim", "String")(_null_or(lambda a: V.text(_val(a[0])).rstrip(" ")))
bi("RTrim$", "String")(_null_or(lambda a: V.text(_val(a[0])).rstrip(" "), True))
bi("StrReverse", "Expression", ret="String")(_null_or(lambda a: V.text(_val(a[0]))[::-1], True))


@bi("Space", "Number", names=("Space$",))
def f_space(ex, args, ctx):
    n = _l(args[0])
    if n < 0:
        raise VbaError(5)
    _check_len(n)
    return " " * n


def _check_len(n: int) -> None:
    if n > MAX_STRING:
        raise VbaError(14)


MAX_STRING = 64 * 1024 * 1024


@bi("String", "Number", "Character", names=("String$",))
def f_string(ex, args, ctx):
    n = _l(args[0])
    c = _val(args[1])
    if c is NULL:
        return NULL
    if n < 0:
        raise VbaError(5)
    _check_len(n)
    if isinstance(c, str):
        if not c:
            raise VbaError(5)
        ch = c[0]
    else:
        ch = _chr(_l(c) % 256)
    return ch * n


def _chr(n: int) -> str:
    if not -32768 <= n <= 65535:
        raise VbaError(5)
    if n < 0:
        n += 65536
    if n < 256:
        return bytes([n]).decode("cp1252", "replace") if 128 <= n < 160 else chr(n)
    return chr(n)


@bi("Chr", "CharCode", const_ok=True, names=("Chr$",))
def f_chr(ex, args, ctx):
    return _chr(_l(args[0]))


@bi("ChrW", "CharCode", const_ok=True, names=("ChrW$",))
def f_chrw(ex, args, ctx):
    n = _l(args[0])
    if not -32768 <= n <= 65535:
        raise VbaError(5)
    return chr(n & 0xFFFF)


@bi("ChrB", "CharCode", names=("ChrB$",))
def f_chrb(ex, args, ctx):
    return chr(_l(args[0]) & 0xFF)


@bi("Asc", "String", ret="Integer", const_ok=True)
def f_asc(ex, args, ctx):
    s = _s(args[0])
    if not s:
        raise VbaError(5)
    c = s[0]
    if ord(c) < 128:
        return Integer(ord(c))
    try:
        b = c.encode("cp1252")
        return Integer(b[0])
    except UnicodeEncodeError:
        return Integer(63)


@bi("AscW", "String", ret="Integer", const_ok=True)
def f_ascw(ex, args, ctx):
    s = _s(args[0])
    if not s:
        raise VbaError(5)
    n = ord(s[0])
    if n > 0xFFFF:
        n = 0xD800 + ((n - 0x10000) >> 10)
    return Integer(n - 65536 if n > 32767 else n)


@bi("AscB", "String", ret="Integer")
def f_ascb(ex, args, ctx):
    s = _s(args[0])
    if not s:
        raise VbaError(5)
    return Integer(s.encode("utf-16-le")[0])


@bi("InStr", "Start", "String1", "?String2", "?Compare")
def f_instr(ex, args, ctx):
    if len(args) == 2 or (len(args) >= 2 and _arg(args, 2) is MISSING and not _is_num(_val(args[0]))):
        start, a, b, cmp_i = 1, _val(args[0]), _val(args[1]), 3
    else:
        start = _l(args[0])
        a, b = _val(args[1]), _val(args[2])
        cmp_i = 3
    if start < 1:
        raise VbaError(5)
    text_mode = _compare_mode(args, cmp_i, ctx) if len(args) > 3 else (bool(ctx.compare_text) if ctx else False)
    if a is NULL or b is NULL:
        return NULL
    a, b = V.text(a), V.text(b)
    if start > len(a):
        return 0
    if b == "":
        return start
    if text_mode:
        a, b = a.lower(), b.lower()
    return a.find(b, start - 1) + 1


def _is_num(v: Any) -> bool:
    return type(v) in (int, Integer, Byte, LongLong, float, Single, Currency, decimal.Decimal, bool)


@bi("InStrRev", "StringCheck", "StringMatch", "?Start", "?Compare")
def f_instrrev(ex, args, ctx):
    a, b = _val(args[0]), _val(args[1])
    if a is NULL or b is NULL:
        return NULL
    a, b = V.text(a), V.text(b)
    start = _l(_arg(args, 2, -1))
    if start == 0 or start < -1:
        raise VbaError(5)
    c = _arg(args, 3)
    text_mode = c is not MISSING and _l(c) == 1
    if start == -1:
        start = len(a)
    if b == "":
        return start if start <= len(a) else 0
    if start > len(a):
        return 0
    if text_mode:
        a, b = a.lower(), b.lower()
    return a.rfind(b, 0, start) + 1


@bi("StrComp", "String1", "String2", "?Compare")
def f_strcomp(ex, args, ctx):
    a, b = _val(args[0]), _val(args[1])
    if a is NULL or b is NULL:
        return NULL
    a, b = V.text(a), V.text(b)
    if _compare_mode(args, 2, ctx):
        a, b = a.lower(), b.lower()
    return Integer((a > b) - (a < b))


@bi("Replace", "Expression", "Find", "Replace", "?Start", "?Count", "?Compare", ret="String")
def f_replace(ex, args, ctx):
    s = _val(args[0])
    if s is NULL:
        raise VbaError(94)
    s, find, rep = V.text(s), _s(args[1]), _s(args[2])
    start = _l(_arg(args, 3, 1))
    count = _l(_arg(args, 4, -1))
    c = _arg(args, 5)
    text_mode = c is not MISSING and _l(c) == 1
    if start < 1 or count < -1:
        raise VbaError(5)
    s = s[start - 1:]
    if find == "" or count == 0:
        return s
    if not text_mode:
        res = s.replace(find, rep) if count == -1 else s.replace(find, rep, count)
    else:
        res = re.sub(re.escape(find), lambda m: rep, s, count=0 if count == -1 else count, flags=re.IGNORECASE)
    _check_len(len(res))
    return res


@bi("Split", "Expression", "?Delimiter", "?Limit", "?Compare")
def f_split(ex, args, ctx):
    s = _val(args[0])
    if s is NULL:
        raise VbaError(94)
    s = V.text(s)
    d = _s(_arg(args, 1, " "))
    limit = _l(_arg(args, 2, -1))
    c = _arg(args, 3)
    text_mode = c is not MISSING and _l(c) == 1
    if s == "":
        return VArray([(0, -1)], [], V.to_string, "String")
    if d == "":
        parts = [s]
    elif text_mode:
        parts = re.split(re.escape(d), s, maxsplit=0 if limit < 0 else max(limit - 1, 0), flags=re.IGNORECASE)
    else:
        parts = s.split(d) if limit < 0 else s.split(d, max(limit - 1, 0))
    if limit == 0:
        parts = []
    return VArray([(0, len(parts) - 1)], parts, V.to_string, "String")


@bi("Join", "SourceArray", "?Delimiter", ret="String")
def f_join(ex, args, ctx):
    a = _val(args[0])
    if type(a) is not VArray:
        raise VbaError(13)
    if a.dims > 1:
        raise VbaError(5)
    d = _s(_arg(args, 1, " "))
    return d.join(V.text(x) if x is not NULL else "" for x in a.data)


@bi("Filter", "SourceArray", "Match", "?Include", "?Compare")
def f_filter(ex, args, ctx):
    a = _val(args[0])
    if type(a) is not VArray or a.dims != 1:
        raise VbaError(13)
    m = _s(args[1])
    inc = V.to_boolean(_arg(args, 2, True))
    c = _arg(args, 3)
    text_mode = c is not MISSING and _l(c) == 1
    out = []
    for x in a.data:
        t = V.text(x)
        hit = (m.lower() in t.lower()) if text_mode else (m in t)
        if hit == inc:
            out.append(t)
    return VArray([(0, len(out) - 1)], out, V.to_string, "String")


@bi("StrConv", "String", "Conversion", "?LocaleID")
def f_strconv(ex, args, ctx):
    s = _val(args[0])
    if s is NULL:
        return NULL
    conv = _l(args[1])
    if type(s) is VArray:
        if conv & 64:
            return "".join(chr(int(b)) for b in s.data)
        raise VbaError(13)
    s = V.text(s)
    if conv & 128:
        b = s.encode("cp1252", "replace")
        return VArray([(0, len(b) - 1)], [Byte(x) for x in b], V.to_byte, "Byte")
    if conv & 64:
        return s
    if conv & 3 == 1:
        s = s.upper()
    elif conv & 3 == 2:
        s = s.lower()
    elif conv & 3 == 3:
        s = re.sub(r"[^\W_]+", lambda m: m.group(0)[0].upper() + m.group(0)[1:].lower(), s)
    if conv & 4:
        s = "".join(chr(ord(c) + 0xFEE0) if 0x21 <= ord(c) <= 0x7E else ("　" if c == " " else c) for c in s)
    if conv & 8:
        s = "".join(chr(ord(c) - 0xFEE0) if 0xFF01 <= ord(c) <= 0xFF5E else (" " if c == "　" else c) for c in s)
    return s


@bi("Val", "String", ret="Double")
def f_val(ex, args, ctx):
    v = _val(args[0])
    if v is NULL:
        raise VbaError(94)
    s = V.text(v) if not isinstance(v, str) else v
    return val_text(s)


def val_text(s: str) -> Any:
    t = "".join(ch for ch in s if ch not in " \t\n\r")
    m = re.match(r"&([Hh])([0-9A-Fa-f]+)|&([Oo]?)([0-7]+)", t)
    if m:
        if m.group(1):
            n = int(m.group(2)[:16], 16)
        else:
            n = int(m.group(4)[:22], 8)
        # Val always returns a Double (&HFFFF is -1, as an Integer would be).
        if n <= 0xFFFF:
            return float(n - 0x10000 if n >= 0x8000 else n)
        if n <= 0xFFFFFFFF:
            return float(n - 0x100000000 if n >= 0x80000000 else n)
        return float(n)
    m = re.match(r"[+-]?(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?", t)
    if not m:
        return 0.0
    txt = m.group(0).replace("d", "e").replace("D", "e")
    try:
        f = float(txt)
    except (ValueError, OverflowError):
        raise VbaError(6)
    if math.isinf(f):
        raise VbaError(6)
    return f


@bi("Str", "Number", names=("Str$",))
def f_str(ex, args, ctx):
    v = _val(args[0])
    if v is NULL:
        return NULL
    if isinstance(v, str):
        v = V.to_double(v)
    t = V.text(v)
    if _is_num(v) and type(v) is not bool:
        # Str drops the zero before the decimal point (" .5") and keeps a space for the sign.
        if t.startswith("0.") and len(t) > 2:
            t = t[1:]
        elif t.startswith("-0.") and len(t) > 3:
            t = "-" + t[2:]
        if not t.startswith("-"):
            return " " + t
    return t


@bi("Hex", "Number", names=("Hex$",))
def f_hex(ex, args, ctx):
    return _radix(args[0], 16)


@bi("Oct", "Number", names=("Oct$",))
def f_oct(ex, args, ctx):
    return _radix(args[0], 8)


def _radix(v: Any, base: int) -> Any:
    v = _val(v)
    if v is NULL:
        return NULL
    if v is EMPTY:
        return "0"
    t = type(v)
    if t in (Integer, Byte, bool):
        n = V.to_integer(v)
        bits = 16
    elif t is LongLong:
        n = int(v)
        bits = 64
    else:
        n = V.to_long(v) if t is not str else V.to_long(v)
        bits = 32
    if n < 0:
        n += 1 << bits
    return format(n, "X" if base == 16 else "o")


# ---------------------------------------------------------------- conversions


def _conv(fn: Callable[[Any], Any]) -> Callable:
    def f(ex, args, ctx):
        v = _val(args[0])
        return fn(v)

    return f


bi("CBool", "Expression", ret="Boolean", const_ok=True)(_conv(V.to_boolean))
bi("CByte", "Expression", ret="Byte", const_ok=True)(_conv(V.to_byte))
bi("CCur", "Expression", ret="Currency", const_ok=True)(_conv(V.to_currency))
bi("CDate", "Expression", ret="Date", const_ok=True, names=("CVDate",))(_conv(V.to_date))
bi("CDbl", "Expression", ret="Double", const_ok=True)(_conv(V.to_double))
bi("CDec", "Expression", ret="Variant", const_ok=True)(_conv(V.to_decimal))
bi("CInt", "Expression", ret="Integer", const_ok=True)(_conv(V.to_integer))
bi("CLng", "Expression", ret="Long", const_ok=True)(_conv(V.to_long))
bi("CLngLng", "Expression", ret="LongLong", const_ok=True)(_conv(V.to_longlong))
bi("CLngPtr", "Expression", ret="LongPtr", const_ok=True)(_conv(V.to_longlong))
bi("CSng", "Expression", ret="Single", const_ok=True)(_conv(V.to_single))
bi("CStr", "Expression", ret="String", const_ok=True)(_conv(V.to_string))


@bi("CVar", "Expression", const_ok=True)
def f_cvar(ex, args, ctx):
    return _val(args[0])


@bi("CVErr", "Expression")
def f_cverr(ex, args, ctx):
    return ErrorValue(_l(args[0]))


def _int_like(v: Any, trunc: bool) -> Any:
    v = _val(v)
    if v is NULL:
        return NULL
    t = type(v)
    if t in (int, Integer, Byte, LongLong):
        return v
    if t is bool:
        return Integer(-1 if v else 0)
    if v is EMPTY:
        return Integer(0)
    if t is Currency:
        q = abs(v.units) // 10000 if trunc else None
        if trunc:
            return Currency((q if v.units >= 0 else -q) * 10000)
        return Currency((v.units // 10000) * 10000)
    if t is decimal.Decimal:
        return v.to_integral_value(rounding=decimal.ROUND_DOWN if trunc else decimal.ROUND_FLOOR)
    if t is str:
        v = V.to_double(v)
        t = float
    f = float(v)
    r = float(math.trunc(f)) if trunc else float(math.floor(f))
    if r == 0 and f < 0:
        r = -0.0
    if t is Single:
        return Single(r)
    if t is Date:
        return Date(r)
    return r


@bi("Int", "Number", const_ok=True)
def f_int(ex, args, ctx):
    return _int_like(args[0], False)


@bi("Fix", "Number", const_ok=True)
def f_fix(ex, args, ctx):
    return _int_like(args[0], True)


# ---------------------------------------------------------------- math


def _math(fn: Callable[[float], float], domain: Callable[[float], bool] = lambda x: True) -> Callable:
    def f(ex, args, ctx):
        v = _val(args[0])
        if v is NULL:
            return NULL
        x = V.to_double(v)
        if not domain(x):
            raise VbaError(5)
        try:
            r = fn(x)
        except (OverflowError, ValueError):
            raise VbaError(6)
        if math.isinf(r) or math.isnan(r):
            raise VbaError(6)
        return r

    return f


bi("Atn", "Number", ret="Double", const_ok=True)(_math(math.atan))
bi("Cos", "Number", ret="Double", const_ok=True)(_math(math.cos))
bi("Sin", "Number", ret="Double", const_ok=True)(_math(math.sin))
bi("Tan", "Number", ret="Double", const_ok=True)(_math(math.tan))
bi("Exp", "Number", ret="Double", const_ok=True)(_math(math.exp))
bi("Log", "Number", ret="Double", const_ok=True)(_math(math.log, lambda x: x > 0))
bi("Sqr", "Number", ret="Double", const_ok=True)(_math(math.sqrt, lambda x: x >= 0))


@bi("Abs", "Number", const_ok=True)
def f_abs(ex, args, ctx):
    v = _val(args[0])
    if v is NULL:
        return NULL
    if v is EMPTY:
        return Integer(0)
    t = type(v)
    if t is str:
        return abs(V.to_double(v))
    if t is bool:
        return Integer(1 if v else 0)
    if t is Currency:
        return Currency(abs(v.units))
    if t in (Integer, int, LongLong):
        r = abs(int(v))
        return V._fit(r, {Integer: 2, int: 3, LongLong: 4}[t], True) if t is not int else (r if r <= 2147483647 else float(r))
    if t is Date:
        return Date(abs(float(v)))
    if t is Single:
        return Single(abs(float(v)))
    if t is Byte:
        return v
    if t is decimal.Decimal:
        return abs(v)
    return abs(float(v))


@bi("Sgn", "Number", ret="Integer", const_ok=True)
def f_sgn(ex, args, ctx):
    v = _val(args[0])
    if v is NULL:
        return NULL
    x = V.to_double(v) if not isinstance(v, (Currency, decimal.Decimal)) else float(v)
    return Integer((x > 0) - (x < 0))


@bi("Round", "Number", "?NumDigitsAfterDecimal")
def f_round(ex, args, ctx):
    v = _val(args[0])
    if v is NULL:
        return NULL
    n = _l(_arg(args, 1, 0))
    if n < 0:
        raise VbaError(5)
    t = type(v)
    if t is Currency:
        d = v.value().quantize(decimal.Decimal(1).scaleb(-min(n, 4)), rounding=decimal.ROUND_HALF_EVEN)
        return Currency.of(d)
    if t is decimal.Decimal:
        return v.quantize(decimal.Decimal(1).scaleb(-n), rounding=decimal.ROUND_HALF_EVEN)
    if t in (int, Integer, Byte, LongLong):
        return v
    if t is bool:
        return Integer(-1 if v else 0)
    x = V.to_double(v)
    r = round_half_even(x, n)
    if t is Single:
        return Single(V._f32(r))
    if t is Date:
        return Date(r)
    return r


def round_half_even(x: float, n: int) -> float:
    """VBA's Round: banker's rounding of the value as the decimal VBA prints (15 digits)."""
    if not math.isfinite(x):
        raise VbaError(6)
    d = decimal.Decimal(repr(x))
    d15 = decimal.Context(prec=15).create_decimal(d)
    q = d15.quantize(decimal.Decimal(1).scaleb(-n), rounding=decimal.ROUND_HALF_EVEN) if n <= 28 else d15
    return float(q)


# Rnd: VBA's own linear congruential generator.


def rnd(rt, x: Any = MISSING) -> float:
    if x is not MISSING:
        f = V.to_single(_val(x))
        if f == 0:
            return rt.rnd_last
        if f < 0:
            bits = struct.unpack("<I", struct.pack("<f", float(f)))[0]
            rt.rnd_state = (bits + (bits >> 24)) & 0xFFFFFF
    rt.rnd_state = (rt.rnd_state * 0x43FD43FD + 0xC39EC3) & 0xFFFFFF
    rt.rnd_last = Single(rt.rnd_state / 16777216.0)
    return rt.rnd_last


def randomize(rt, seed: Any = None) -> None:
    if seed is None or seed is MISSING:
        seed = timer_value()
    d = V.to_double(_val(seed))
    bits = struct.unpack("<Q", struct.pack("<d", d))[0]
    lo = bits & 0xFFFFFFFF
    # The seed's low dword, its two halves XORed, replaces the middle 16 bits of the state
    # (worked out from Excel's own sequences).
    mid = (lo ^ (lo >> 16)) & 0xFFFF
    rt.rnd_state = (rt.rnd_state & 0xFF) | (mid << 8)


@bi("Rnd", "?Number", ret="Single")
def f_rnd(ex, args, ctx):
    return rnd(ex.rt, _arg(args, 0))


# ---------------------------------------------------------------- dates and times


def now_serial() -> float:
    n = _dt.datetime.now().replace(microsecond=0)
    return V.datetime_to_serial(n)


def timer_value() -> float:
    n = _dt.datetime.now()
    return n.hour * 3600 + n.minute * 60 + n.second + n.microsecond / 1e6


@bi("Now", ret="Date")
def f_now(ex, args, ctx):
    return Date(now_serial())


@bi("Date", ret="Date")
def f_date(ex, args, ctx):
    return Date(float(int(now_serial())))


@bi("Date$", ret="String")
def f_date_s(ex, args, ctx):
    n = _dt.datetime.now()
    return f"{n.month:02d}-{n.day:02d}-{n.year:04d}"


@bi("Time", ret="Date")
def f_time(ex, args, ctx):
    s = now_serial()
    return Date(s - int(s))


@bi("Time$", ret="String")
def f_time_s(ex, args, ctx):
    n = _dt.datetime.now()
    return f"{n.hour:02d}:{n.minute:02d}:{n.second:02d}"


@bi("Timer", ret="Single")
def f_timer(ex, args, ctx):
    return Single(round(timer_value(), 2))


def _dateparts(v: Any) -> _dt.datetime:
    v = _val(v)
    return V.serial_to_datetime(float(V.to_date(v)))


def _date_null(fn: Callable[[_dt.datetime], Any]) -> Callable:
    def f(ex, args, ctx):
        v = _val(args[0])
        if v is NULL:
            return NULL
        return fn(_dateparts(v))

    return f


bi("Year", "Date", ret="Integer")(_date_null(lambda d: Integer(d.year)))
bi("Month", "Date", ret="Integer")(_date_null(lambda d: Integer(d.month)))
bi("Day", "Date", ret="Integer")(_date_null(lambda d: Integer(d.day)))
bi("Hour", "Time", ret="Integer")(_date_null(lambda d: Integer(d.hour)))
bi("Minute", "Time", ret="Integer")(_date_null(lambda d: Integer(d.minute)))
bi("Second", "Time", ret="Integer")(_date_null(lambda d: Integer(d.second)))


def _weekday(d: _dt.datetime, first: int = 1) -> int:
    # Sunday = 1 when the week starts on Sunday.
    wd = (d.isoweekday() % 7) + 1  # Sunday 1 … Saturday 7
    if first == 0:
        first = 1
    return (wd - first) % 7 + 1


@bi("Weekday", "Date", "?FirstDayOfWeek", ret="Integer")
def f_weekday(ex, args, ctx):
    v = _val(args[0])
    if v is NULL:
        return NULL
    first = _l(_arg(args, 1, 1))
    if not 0 <= first <= 7:
        raise VbaError(5)
    return Integer(_weekday(_dateparts(v), first))


@bi("DateSerial", "Year", "Month", "Day", ret="Date")
def f_dateserial(ex, args, ctx):
    return Date(V.date_serial(_l(args[0]), _l(args[1]), _l(args[2])))


@bi("TimeSerial", "Hour", "Minute", "Second", ret="Date")
def f_timeserial(ex, args, ctx):
    t = V.time_serial(_l(args[0]), _l(args[1]), _l(args[2]))
    if t < 0:
        days = math.floor(t)
        t = days + (t - days) if days == 0 else t
    return Date(t)


@bi("DateValue", "Date", ret="Date")
def f_datevalue(ex, args, ctx):
    v = _val(args[0])
    if v is NULL:
        return NULL
    d = V.to_date(v) if not isinstance(v, str) else V.to_date(v)
    return Date(float(int(d)))


@bi("TimeValue", "Time", ret="Date")
def f_timevalue(ex, args, ctx):
    v = _val(args[0])
    if v is NULL:
        return NULL
    d = float(V.to_date(v))
    return Date(abs(d - int(d)))


_INTERVALS = {"yyyy", "q", "m", "y", "d", "w", "ww", "h", "n", "s"}


def _interval(v: Any) -> str:
    s = _s(v).lower()
    if s not in _INTERVALS:
        raise VbaError(5)
    return s


def _add_months(d: _dt.datetime, months: int) -> _dt.datetime:
    y, m = divmod(d.month - 1 + months, 12)
    y += d.year
    m += 1
    if not 100 <= y <= 9999:
        raise VbaError(5)
    import calendar

    day = min(d.day, calendar.monthrange(y, m)[1])
    return d.replace(year=y, month=m, day=day)


@bi("DateAdd", "Interval", "Number", "Date", ret="Date")
def f_dateadd(ex, args, ctx):
    iv = _interval(args[0])
    n = _val(args[1])
    if _val(args[2]) is NULL:
        return NULL
    num = V.to_double(n)
    d = _dateparts(args[2])
    serial = float(V.to_date(_val(args[2])))
    try:
        if iv == "yyyy":
            r = _add_months(d, int(_fix(num)) * 12)
        elif iv == "q":
            r = _add_months(d, int(_fix(num)) * 3)
        elif iv == "m":
            r = _add_months(d, int(_fix(num)))
        elif iv in ("y", "d", "w"):
            r = d + _dt.timedelta(days=int(_fix(num)))
        elif iv == "ww":
            r = d + _dt.timedelta(weeks=int(_fix(num)))
        elif iv == "h":
            return _check_date_serial(serial + _sign_away(serial, _fix(num) / 24.0))
        elif iv == "n":
            return _check_date_serial(serial + _sign_away(serial, _fix(num) / 1440.0))
        else:
            return _check_date_serial(serial + _sign_away(serial, _fix(num) / 86400.0))
    except (OverflowError, ValueError):
        raise VbaError(5)
    return Date(V.datetime_to_serial(r))


def _fix(x: float) -> float:
    return float(math.trunc(x)) if abs(x - round(x)) > 1e-9 else float(round(x))


def _sign_away(base: float, delta: float) -> float:
    return delta


def _check_date_serial(x: float) -> Date:
    if not V.DATE_MIN <= x < 2958466.0:
        raise VbaError(5)
    # Whole seconds, as VBA keeps them.
    return Date(round(x * 86400) / 86400.0)


def _week_start(d: _dt.date, first: int) -> _dt.date:
    wd = _weekday(_dt.datetime(d.year, d.month, d.day), first)
    return d - _dt.timedelta(days=wd - 1)


@bi("DateDiff", "Interval", "Date1", "Date2", "?FirstDayOfWeek", "?FirstWeekOfYear", ret="Long")
def f_datediff(ex, args, ctx):
    iv = _interval(args[0])
    a, b = _val(args[1]), _val(args[2])
    if a is NULL or b is NULL:
        return NULL
    first = _l(_arg(args, 3, 1)) or 1
    s1, s2 = float(V.to_date(a)), float(V.to_date(b))
    d1, d2 = V.serial_to_datetime(s1), V.serial_to_datetime(s2)
    if iv == "yyyy":
        return d2.year - d1.year
    if iv == "q":
        return (d2.year * 4 + (d2.month - 1) // 3) - (d1.year * 4 + (d1.month - 1) // 3)
    if iv == "m":
        return (d2.year * 12 + d2.month) - (d1.year * 12 + d1.month)
    if iv in ("y", "d"):
        return (d2.date() - d1.date()).days
    if iv == "w":
        days = (d2.date() - d1.date()).days
        return int(days / 7)
    if iv == "ww":
        return ((_week_start(d2.date(), first) - _week_start(d1.date(), first)).days) // 7
    secs = (d2 - d1).total_seconds()
    if iv in ("h", "n"):
        # Whole minutes of each time (seconds dropped), then the difference; hours are that \ 60.
        m1 = math.floor(round(s1 * 86400) / 60)
        m2 = math.floor(round(s2 * 86400) / 60)
        n = m2 - m1
        return n if iv == "n" else int(n / 60)
    return int(round(secs))


@bi("DatePart", "Interval", "Date", "?FirstDayOfWeek", "?FirstWeekOfYear", ret="Integer")
def f_datepart(ex, args, ctx):
    iv = _interval(args[0])
    v = _val(args[1])
    if v is NULL:
        return NULL
    first = _l(_arg(args, 2, 1)) or 1
    firstweek = _l(_arg(args, 3, 1)) or 1
    d = _dateparts(v)
    if iv == "yyyy":
        return Integer(d.year)
    if iv == "q":
        return Integer((d.month - 1) // 3 + 1)
    if iv == "m":
        return Integer(d.month)
    if iv == "y":
        return Integer(d.timetuple().tm_yday)
    if iv == "d":
        return Integer(d.day)
    if iv == "w":
        return Integer(_weekday(d, first))
    if iv == "ww":
        return Integer(week_of_year(d.date(), first, firstweek))
    if iv == "h":
        return Integer(d.hour)
    if iv == "n":
        return Integer(d.minute)
    return Integer(d.second)


def week_of_year(d: _dt.date, first: int, firstweek: int) -> int:
    jan1 = _dt.date(d.year, 1, 1)
    start = _week_start(jan1, first)
    if firstweek == 2:  # the first week with four days in the year
        if (jan1 - start).days >= 4:
            start += _dt.timedelta(days=7)
    elif firstweek == 3:  # the first full week
        if start < jan1:
            start += _dt.timedelta(days=7)
    if d < start:
        return week_of_year(_dt.date(d.year - 1, 12, 31), first, firstweek)
    return (d - start).days // 7 + 1


@bi("MonthName", "Month", "?Abbreviate", ret="String")
def f_monthname(ex, args, ctx):
    m = _l(args[0])
    if not 1 <= m <= 12:
        raise VbaError(5)
    ab = V.to_boolean(_arg(args, 1, False))
    n = V.MONTHS[m - 1]
    return n[:3] if ab else n


@bi("WeekdayName", "Weekday", "?Abbreviate", "?FirstDayOfWeek", ret="String")
def f_weekdayname(ex, args, ctx):
    w = _l(args[0])
    if not 1 <= w <= 7:
        raise VbaError(5)
    ab = V.to_boolean(_arg(args, 1, False))
    first = _l(_arg(args, 2, 1)) or 1
    idx = (w - 1 + first - 1) % 7
    n = V.DAYS[idx]
    return n[:3] if ab else n


# ---------------------------------------------------------------- information


@bi("IsArray", "VarName", ret="Boolean")
def f_isarray(ex, args, ctx):
    return type(args[0]) is VArray


@bi("IsDate", "Expression", ret="Boolean")
def f_isdate(ex, args, ctx):
    v = _val(args[0]) if not (args[0] is NOTHING) else args[0]
    if type(v) is Date:
        return True
    if isinstance(v, str):
        return V.parse_date_text(v) is not None
    return False


@bi("IsEmpty", "Expression", ret="Boolean")
def f_isempty(ex, args, ctx):
    v = args[0]
    if is_object(v) and v is not NOTHING:
        try:
            v = V.value_of(v)
        except VbaError:
            return False
    return v is EMPTY


@bi("IsError", "Expression", ret="Boolean")
def f_iserror(ex, args, ctx):
    v = args[0]
    if is_object(v) and v is not NOTHING:
        try:
            v = V.value_of(v)
        except VbaError:
            return False
    return type(v) is ErrorValue or v is MISSING


@bi("IsMissing", "ArgName", ret="Boolean")
def f_ismissing(ex, args, ctx):
    return args[0] is MISSING


@bi("IsNull", "Expression", ret="Boolean")
def f_isnull(ex, args, ctx):
    v = args[0]
    if is_object(v) and v is not NOTHING:
        try:
            v = V.value_of(v)
        except VbaError:
            return False
    return v is NULL


@bi("IsNumeric", "Expression", ret="Boolean")
def f_isnumeric(ex, args, ctx):
    v = args[0]
    if is_object(v) and v is not NOTHING:
        try:
            v = V.value_of(v)
        except VbaError:
            return False
    t = type(v)
    if v is EMPTY:
        return True
    if t in (int, Integer, Byte, LongLong, float, Single, Currency, decimal.Decimal, bool):
        return True
    if t is str:
        return V.parse_number(v) is not None
    return False


@bi("IsObject", "Expression", ret="Boolean")
def f_isobject(ex, args, ctx):
    return args[0] is NOTHING or is_object(args[0])


@bi("TypeName", "VarName", ret="String")
def f_typename(ex, args, ctx):
    return V.type_name(args[0])


@bi("VarType", "VarName", ret="Long")
def f_vartype(ex, args, ctx):
    v = args[0]
    if is_object(v) and v is not NOTHING:
        try:
            return V.vartype(V.value_of(v))
        except VbaError:
            return 9
    return V.vartype(v)


@bi("Array", "*ArgList")
def f_array(ex, args, ctx):
    base = ctx.option_base if ctx is not None else 0
    items = [x if x is not MISSING else EMPTY for x in args]
    return VArray([(base, base + len(items) - 1)], items)


def _bound(args, upper: bool) -> int:
    a = args[0]
    if type(a) is not VArray:
        raise VbaError(13)
    dim = _l(_arg(args, 1, 1))
    if not a.bounds or not 1 <= dim <= len(a.bounds):
        raise VbaError(9)
    return a.bounds[dim - 1][1 if upper else 0]


@bi("LBound", "ArrayName", "?Dimension", ret="Long")
def f_lbound(ex, args, ctx):
    return _bound(args, False)


@bi("UBound", "ArrayName", "?Dimension", ret="Long")
def f_ubound(ex, args, ctx):
    return _bound(args, True)


@bi("IIf", "Expression", "TruePart", "FalsePart")
def f_iif(ex, args, ctx):
    return args[1] if V.is_true(_val(args[0])) else args[2]


@bi("Choose", "Index", "*Choice")
def f_choose(ex, args, ctx):
    v = _val(args[0])
    if v is NULL:
        return NULL
    i = int(math.trunc(V.to_double(v)))
    if 1 <= i < len(args):
        return args[i]
    return NULL


@bi("Switch", "*VarExpr")
def f_switch(ex, args, ctx):
    if len(args) % 2:
        raise VbaError(5)
    for i in range(0, len(args), 2):
        if V.is_true(_val(args[i])):
            return args[i + 1]
    return NULL


@bi("RGB", "Red", "Green", "Blue", ret="Long", const_ok=True)
def f_rgb(ex, args, ctx):
    r, g, b = (max(0, min(255, _l(x))) for x in args[:3])
    return r + g * 256 + b * 65536


_QB = [0x000000, 0x800000, 0x008000, 0x808000, 0x000080, 0x800080, 0x008080, 0xC0C0C0, 0x808080, 0xFF0000, 0x00FF00, 0xFFFF00, 0x0000FF, 0xFF00FF, 0x00FFFF, 0xFFFFFF]


@bi("QBColor", "Color", ret="Long")
def f_qbcolor(ex, args, ctx):
    i = _l(args[0])
    if not 0 <= i <= 15:
        raise VbaError(5)
    return _QB[i]


@bi("Error", "?ErrorNumber", names=("Error$",))
def f_error(ex, args, ctx):
    n = _arg(args, 0)
    if n is MISSING:
        return ex.err.description
    from .errors import message

    k = _l(n)
    if k == 0:
        return ""
    return message(k)


@bi("Erl", ret="Long")
def f_erl(ex, args, ctx):
    return ex.err.erl


@bi("Environ", "Expression", names=("Environ$",))
def f_environ(ex, args, ctx):
    """The sandbox shows a few harmless values, never the server's environment."""
    v = _val(args[0])
    safe = {"username": getattr(ex.host, "user_name", lambda: "")() if callable(getattr(ex.host, "user_name", None)) else "", "os": "ReflexGrid", "computername": "ReflexGrid", "temp": "", "tmp": "", "userprofile": "", "homepath": ""}
    if isinstance(v, str):
        return safe.get(v.lower(), "")
    return ""


@bi("Command", names=("Command$",))
def f_command(ex, args, ctx):
    return ""


@bi("DoEvents", ret="Integer")
def f_doevents(ex, args, ctx):
    ex.check()
    return Integer(0)


@bi("UserForms", "?index")
def f_userforms(ex, args, ctx):
    from .msforms import LoadedForms

    forms = LoadedForms(ex.host.form_registry(ex))
    i = _arg(args, 0)
    return forms if i is MISSING else forms.g_item(ex, i)


@bi("CallByName", "Object", "ProcName", "CallType", "*Args")
def f_callbyname(ex, args, ctx):
    from .runtime import get_member, let_member, set_member

    obj = args[0]
    name = _s(args[1])
    kind = _l(args[2])
    rest = list(args[3:])
    if kind in (1, 2):
        return get_member(ex, obj, name, rest, None)
    if kind == 4:
        if not rest:
            raise VbaError(5)
        let_member(ex, obj, name, rest[:-1], None, rest[-1])
        return EMPTY
    if kind == 8:
        if not rest:
            raise VbaError(5)
        set_member(ex, obj, name, rest[:-1], None, rest[-1])
        return EMPTY
    raise VbaError(5)


@bi("Partition", "Number", "Start", "Stop", "Interval", ret="String")
def f_partition(ex, args, ctx):
    n, start, stop, interval = (_l(x) for x in args[:4])
    if start < 0 or stop <= start or interval < 1:
        raise VbaError(5)
    w = max(len(str(stop + 1)), len(str(start - 1)))
    if n < start:
        return f"{'':>{w}}:{start - 1:>{w}}"
    if n > stop:
        return f"{stop + 1:>{w}}:{'':>{w}}"
    lo = start + ((n - start) // interval) * interval
    hi = min(lo + interval - 1, stop)
    return f"{lo:>{w}}:{hi:>{w}}"


@bi("ObjPtr", "Object", ret="LongPtr")
def f_objptr(ex, args, ctx):
    return LongLong(id(args[0]) & 0x7FFFFFFFFFFF)


@bi("VarPtr", "Var", ret="LongPtr")
def f_varptr(ex, args, ctx):
    return LongLong(id(args[0]) & 0x7FFFFFFFFFFF)


@bi("StrPtr", "Str", ret="LongPtr")
def f_strptr(ex, args, ctx):
    return LongLong(id(args[0]) & 0x7FFFFFFFFFFF)


# ---------------------------------------------------------------- financial


def _rate_args(args, *defaults):
    out = []
    for i, d in enumerate(defaults):
        v = _arg(args, i, d)
        if v is MISSING:
            raise VbaError(449)
        out.append(_d(v))
    return out


@bi("Pmt", "Rate", "NPer", "PV", "?FV", "?Type", ret="Double")
def f_pmt(ex, args, ctx):
    rate, nper, pv, fv, typ = _rate_args(args, MISSING, MISSING, MISSING, 0.0, 0.0)
    if nper == 0:
        raise VbaError(5)
    if rate == 0:
        return -(pv + fv) / nper
    f = (1 + rate) ** nper
    return -(rate * (fv + pv * f)) / ((1 + rate * (1 if typ else 0)) * (f - 1))


@bi("FV", "Rate", "NPer", "Pmt", "?PV", "?Type", ret="Double")
def f_fv(ex, args, ctx):
    rate, nper, pmt, pv, typ = _rate_args(args, MISSING, MISSING, MISSING, 0.0, 0.0)
    if rate == 0:
        return -(pv + pmt * nper)
    f = (1 + rate) ** nper
    return -(pv * f + pmt * (1 + rate * (1 if typ else 0)) * (f - 1) / rate)


@bi("PV", "Rate", "NPer", "Pmt", "?FV", "?Type", ret="Double")
def f_pv(ex, args, ctx):
    rate, nper, pmt, fv, typ = _rate_args(args, MISSING, MISSING, MISSING, 0.0, 0.0)
    if rate == 0:
        return -(fv + pmt * nper)
    f = (1 + rate) ** nper
    return -(fv + pmt * (1 + rate * (1 if typ else 0)) * (f - 1) / rate) / f


@bi("NPer", "Rate", "Pmt", "PV", "?FV", "?Type", ret="Double")
def f_nper(ex, args, ctx):
    rate, pmt, pv, fv, typ = _rate_args(args, MISSING, MISSING, MISSING, 0.0, 0.0)
    if rate == 0:
        if pmt == 0:
            raise VbaError(5)
        return -(pv + fv) / pmt
    t = 1 if typ else 0
    try:
        return math.log((pmt * (1 + rate * t) - fv * rate) / (pmt * (1 + rate * t) + pv * rate)) / math.log(1 + rate)
    except (ValueError, ZeroDivisionError):
        raise VbaError(5)


@bi("IPmt", "Rate", "Per", "NPer", "PV", "?FV", "?Type", ret="Double")
def f_ipmt(ex, args, ctx):
    rate, per, nper, pv, fv, typ = _rate_args(args, MISSING, MISSING, MISSING, MISSING, 0.0, 0.0)
    if not 1 <= per <= nper:
        raise VbaError(5)
    pmt = f_pmt(ex, [rate, nper, pv, fv, typ], ctx)
    if typ and per == 1:
        return 0.0
    bal = f_fv(ex, [rate, per - 1, pmt, pv, typ], ctx)
    ip = bal * rate
    if typ:
        ip /= 1 + rate
    return ip


@bi("PPmt", "Rate", "Per", "NPer", "PV", "?FV", "?Type", ret="Double")
def f_ppmt(ex, args, ctx):
    rate, per, nper, pv, fv, typ = _rate_args(args, MISSING, MISSING, MISSING, MISSING, 0.0, 0.0)
    return f_pmt(ex, [rate, nper, pv, fv, typ], ctx) - f_ipmt(ex, [rate, per, nper, pv, fv, typ], ctx)


@bi("SLN", "Cost", "Salvage", "Life", ret="Double")
def f_sln(ex, args, ctx):
    c, s, life = (_d(x) for x in args[:3])
    if life == 0:
        raise VbaError(5)
    return (c - s) / life


@bi("SYD", "Cost", "Salvage", "Life", "Period", ret="Double")
def f_syd(ex, args, ctx):
    c, s, life, per = (_d(x) for x in args[:4])
    if life <= 0 or not 0 < per <= life:
        raise VbaError(5)
    return (c - s) * (life - per + 1) * 2 / (life * (life + 1))


@bi("DDB", "Cost", "Salvage", "Life", "Period", "?Factor", ret="Double")
def f_ddb(ex, args, ctx):
    c, s, life, per = (_d(x) for x in args[:4])
    factor = _d(_arg(args, 4, 2.0))
    if life <= 0 or per <= 0 or per > life or factor <= 0:
        raise VbaError(5)
    total = 0.0
    dep = 0.0
    for p in range(1, int(math.ceil(per)) + 1):
        dep = min((c - total) * factor / life, max(c - s - total, 0.0))
        total += dep
    return dep


@bi("NPV", "Rate", "ValueArray", ret="Double")
def f_npv(ex, args, ctx):
    rate = _d(args[0])
    arr = _val(args[1])
    if type(arr) is not VArray:
        raise VbaError(5)
    if rate == -1:
        raise VbaError(5)
    vals = [_d(x) for x in arr.data]
    return sum(v / (1 + rate) ** (i + 1) for i, v in enumerate(vals))


@bi("IRR", "ValueArray", "?Guess", ret="Double")
def f_irr(ex, args, ctx):
    arr = _val(args[0])
    if type(arr) is not VArray:
        raise VbaError(5)
    vals = [_d(x) for x in arr.data]
    r = _d(_arg(args, 1, 0.1))
    for _ in range(100):
        f = sum(v / (1 + r) ** i for i, v in enumerate(vals))
        df = sum(-i * v / (1 + r) ** (i + 1) for i, v in enumerate(vals))
        if df == 0:
            break
        nr = r - f / df
        if abs(nr - r) < 1e-10:
            return nr
        r = nr
    raise VbaError(5)


@bi("MIRR", "ValueArray", "FinanceRate", "ReinvestRate", ret="Double")
def f_mirr(ex, args, ctx):
    arr = _val(args[0])
    if type(arr) is not VArray:
        raise VbaError(5)
    vals = [_d(x) for x in arr.data]
    fr, rr = _d(args[1]), _d(args[2])
    n = len(vals)
    pos = sum(v * (1 + rr) ** (n - 1 - i) for i, v in enumerate(vals) if v > 0)
    neg = sum(v / (1 + fr) ** i for i, v in enumerate(vals) if v < 0)
    if neg == 0 or pos == 0:
        raise VbaError(5)
    return (pos / -neg) ** (1.0 / (n - 1)) - 1


@bi("Rate", "NPer", "Pmt", "PV", "?FV", "?Type", "?Guess", ret="Double")
def f_rate(ex, args, ctx):
    nper, pmt, pv, fv, typ, guess = _rate_args(args, MISSING, MISSING, MISSING, 0.0, 0.0, 0.1)
    r = guess
    for _ in range(100):
        if r <= -1:
            raise VbaError(5)
        f = (1 + r) ** nper
        y = pv * f + pmt * (1 + r * typ) * (f - 1) / r + fv if r != 0 else pv + pmt * nper + fv
        # Numeric derivative.
        h = 1e-7
        r2 = r + h
        f2 = (1 + r2) ** nper
        y2 = pv * f2 + pmt * (1 + r2 * typ) * (f2 - 1) / r2 + fv
        d = (y2 - y) / h
        if d == 0:
            break
        nr = r - y / d
        if abs(nr - r) < 1e-10:
            return nr
        r = nr
    raise VbaError(5)


# ---------------------------------------------------------------- interaction


def msgbox_default(buttons: int) -> int:
    """The answer when nobody can be asked: the default button of the button set."""
    kind = buttons & 7
    default = (buttons >> 8) & 3
    sets = {0: [1], 1: [1, 2], 2: [3, 4, 5], 3: [6, 7, 2], 4: [6, 7], 5: [4, 2]}
    opts = sets.get(kind, [1])
    return opts[min(default, len(opts) - 1)]


@bi("MsgBox", "Prompt", "?Buttons", "?Title", "?HelpFile", "?Context", ret="Long")
def f_msgbox(ex, args, ctx):
    prompt = _val(args[0])
    text = "" if prompt is NULL else V.text(prompt)
    buttons = _l(_arg(args, 1, 0))
    t = _arg(args, 2)
    title = V.text(_val(t)) if t is not MISSING and t is not NULL else "Microsoft Excel"
    r = ex.host.prompt(ex, "msgbox", text, title, buttons)
    return int(r)  # a VbMsgBoxResult (Long)


@bi("InputBox", "Prompt", "?Title", "?Default", "?XPos", "?YPos", "?HelpFile", "?Context", ret="Variant", names=("InputBox$",))
def f_inputbox(ex, args, ctx):
    prompt = V.text(_val(args[0]))
    t = _arg(args, 1)
    title = V.text(_val(t)) if t is not MISSING else "Microsoft Excel"
    d = _arg(args, 2)
    default = V.text(_val(d)) if d is not MISSING else ""
    r = ex.host.prompt(ex, "inputbox", prompt, title, 0, default)
    return "" if r is None else V.text(r)


@bi("CreateObject", "Class", "?ServerName")
def f_createobject(ex, args, ctx):
    return ex.host.create_object(ex, _s(args[0]))


@bi("GetObject", "?PathName", "?Class")
def f_getobject(ex, args, ctx):
    cls = _arg(args, 1)
    if cls is not MISSING and _s(cls).lower() == "excel.application":
        return ex.host.global_get(ex, "Application", [], None)
    raise NotSupported("GetObject")


def _unsupported(what: str) -> Callable:
    def f(ex, args, ctx):
        raise NotSupported(what)

    return f


for _name, _what in (
    ("Shell", "Shell (running programs)"), ("Dir", "Dir (the file system)"), ("Dir$", "Dir (the file system)"),
    ("FileLen", "FileLen (the file system)"), ("FileDateTime", "FileDateTime (the file system)"),
    ("GetAttr", "GetAttr (the file system)"), ("CurDir", "CurDir (the file system)"), ("CurDir$", "CurDir (the file system)"),
    ("FreeFile", "File I/O (FreeFile)"), ("EOF", "File I/O (EOF)"), ("LOF", "File I/O (LOF)"), ("Loc", "File I/O (Loc)"),
    ("Seek", "File I/O (Seek)"), ("Input", "File I/O (Input)"), ("Input$", "File I/O (Input)"), ("InputB", "File I/O (InputB)"),
    ("FileAttr", "File I/O (FileAttr)"), ("SendKeys", "SendKeys"), ("AppActivate", "AppActivate"), ("MacScript", "MacScript"),
    ("AppleScriptTask", "AppleScriptTask"), ("GetAllSettings", "the registry (GetAllSettings)"),
    ("LoadPicture", "LoadPicture"), ("ImeStatus", "ImeStatus"),
):
    bi(_name, "*Args")(_unsupported(_what))


@bi("GetSetting", "AppName", "Section", "Key", "?Default")
def f_getsetting(ex, args, ctx):
    store = ex.rt.data.setdefault("settings", {})
    key = (_s(args[0]).lower(), _s(args[1]).lower(), _s(args[2]).lower())
    if key in store:
        return store[key]
    d = _arg(args, 3)
    return V.text(_val(d)) if d is not MISSING else ""


def declared_function(ex, lib: str, name: str, args: List[Any]) -> Any:
    """A Windows API function a ``Declare`` names: a few harmless ones are emulated."""
    n = name.lower()
    if n in ("gettickcount", "gettickcount64", "timegettime"):
        ms = int((time.monotonic()) * 1000) & 0x7FFFFFFF
        return ms
    if n == "sleep":
        ms = V.to_long(_val(args[0])) if args else 0
        if ms > 0:
            ex.check()
            time.sleep(min(ms, 2000) / 1000.0)
        return EMPTY
    if n in ("getcursorpos", "setcursorpos", "messagebeep", "beep"):
        return 0
    raise NotSupported(f"Windows API calls ({name} in {lib})")


# ---------------------------------------------------------------- Debug.Print


def print_text(v: Any) -> str:
    """A value as Print writes it: numbers with a space for the sign and one after."""
    if v is NULL:
        return "Null"
    if v is EMPTY:
        return ""
    t = type(v)
    if t is bool:
        return "True" if v else "False"
    if t is str:
        return v
    if t is Date:
        return V.date_text(v)
    if t is ErrorValue:
        return f"Error {v.code}"
    if t in (int, Integer, Byte, LongLong, float, Single, Currency, decimal.Decimal):
        s = V.text(v)
        return (s if s.startswith("-") else " " + s) + " "
    if t is VArray:
        raise VbaError(13)
    if is_object(v):
        return print_text(V.value_of(v))
    return V.text(v)


def debug_print(ex, items) -> None:
    """One Print statement: ``;`` keeps on the same spot, ``,`` moves to the next 14-column
    zone; without a trailing separator the line is finished."""
    pending = getattr(ex, "pending_print", "")
    col = len(pending)
    out = []
    newline = True
    for v, sep, arg in items:
        if v == "spc":
            n = max(V.to_long(arg), 0) if arg is not None else 0
            out.append(" " * n)
            col += n
        elif v == "tab":
            target = V.to_long(arg) - 1 if arg is not None else ((col // 14) + 1) * 14
            pad = target - col if target > col else 0
            out.append(" " * pad)
            col += pad
        elif v is not None:
            s = print_text(v)
            out.append(s)
            col += len(s)
        if sep == ",":
            nxt = (col // 14 + 1) * 14
            out.append(" " * (nxt - col))
            col = nxt
            newline = False
        elif sep == ";":
            newline = False
        else:
            newline = True
    line = pending + "".join(out)
    if newline or not items:
        ex.pending_print = ""
        ex.host.output(ex, line)
    else:
        ex.pending_print = line


def flush_print(ex) -> None:
    pending = getattr(ex, "pending_print", "")
    if pending:
        ex.pending_print = ""
        ex.host.output(ex, pending)


# ---------------------------------------------------------------- Format


@bi("Format", "Expression", "?Format", "?FirstDayOfWeek", "?FirstWeekOfYear", names=("Format$",))
def f_format(ex, args, ctx):
    v = _val(args[0])
    fmt = _arg(args, 1)
    fmt_s = "" if fmt is MISSING or fmt is NULL else _s(fmt)
    first = _l(_arg(args, 2, 1)) or 1
    firstweek = _l(_arg(args, 3, 1)) or 1
    from .format import vba_format

    return vba_format(v, fmt_s, first, firstweek)


@bi("FormatNumber", "Expression", "?NumDigitsAfterDecimal", "?IncludeLeadingDigit", "?UseParensForNegativeNumbers", "?GroupDigits", ret="String")
def f_formatnumber(ex, args, ctx):
    from .format import format_number

    return format_number(_val(args[0]), args[1:], currency=False, percent=False)


@bi("FormatCurrency", "Expression", "?NumDigitsAfterDecimal", "?IncludeLeadingDigit", "?UseParensForNegativeNumbers", "?GroupDigits", ret="String")
def f_formatcurrency(ex, args, ctx):
    from .format import format_number

    return format_number(_val(args[0]), args[1:], currency=True, percent=False)


@bi("FormatPercent", "Expression", "?NumDigitsAfterDecimal", "?IncludeLeadingDigit", "?UseParensForNegativeNumbers", "?GroupDigits", ret="String")
def f_formatpercent(ex, args, ctx):
    from .format import format_number

    return format_number(_val(args[0]), args[1:], currency=False, percent=True)


@bi("FormatDateTime", "Expression", "?NamedFormat", ret="String")
def f_formatdatetime(ex, args, ctx):
    from .format import vba_format

    v = _val(args[0])
    k = _l(_arg(args, 1, 0))
    names = {0: "General Date", 1: "Long Date", 2: "Short Date", 3: "Long Time", 4: "Short Time"}
    if k not in names:
        raise VbaError(5)
    return vba_format(V.to_date(v), names[k], 1, 1)


# ---------------------------------------------------------------- constants

CONSTANTS: Dict[str, Any] = {
    "vbcr": "\r", "vblf": "\n", "vbcrlf": "\r\n", "vbnewline": "\r\n", "vbnullchar": "\0", "vbnullstring": "",
    "vbtab": "\t", "vbback": "\b", "vbformfeed": "\f", "vbverticaltab": "\v", "vbobjecterror": -2147221504,
    "vbokonly": 0, "vbokcancel": 1, "vbabortretryignore": 2, "vbyesnocancel": 3,
    "vbyesno": 4, "vbretrycancel": 5, "vbcritical": 16, "vbquestion": 32,
    "vbexclamation": 48, "vbinformation": 64, "vbdefaultbutton1": 0,
    "vbdefaultbutton2": 256, "vbdefaultbutton3": 512, "vbdefaultbutton4": 768,
    "vbapplicationmodal": 0, "vbsystemmodal": 4096, "vbmsgboxhelpbutton": 16384,
    "vbmsgboxsetforeground": 65536, "vbmsgboxright": 524288, "vbmsgboxrtlreading": 1048576,
    "vbok": 1, "vbcancel": 2, "vbabort": 3, "vbretry": 4, "vbignore": 5,
    "vbyes": 6, "vbno": 7,
    "vbempty": 0, "vbnull": 1, "vbinteger": 2, "vblong": 3, "vbsingle": 4,
    "vbdouble": 5, "vbcurrency": 6, "vbdate": 7, "vbstring": 8, "vbobject": 9,
    "vberror": 10, "vbboolean": 11, "vbvariant": 12, "vbdataobject": 13,
    "vbdecimal": 14, "vbbyte": 17, "vblonglong": 20, "vbuserdefinedtype": 36,
    "vbarray": 8192,
    "vbbinarycompare": 0, "vbtextcompare": 1, "vbdatabasecompare": 2,
    "vbusesystemdayofweek": 0, "vbsunday": 1, "vbmonday": 2, "vbtuesday": 3, "vbwednesday": 4, "vbthursday": 5,
    "vbfriday": 6, "vbsaturday": 7, "vbusesystem": 0, "vbfirstjan1": 1, "vbfirstfourdays": 2, "vbfirstfullweek": 3,
    "vbgeneraldate": 0, "vblongdate": 1, "vbshortdate": 2, "vblongtime": 3, "vbshorttime": 4,
    "vbtrue": -1, "vbfalse": 0, "vbusedefault": -2,
    "vbuppercase": 1, "vblowercase": 2, "vbpropercase": 3, "vbwide": 4, "vbnarrow": 8, "vbkatakana": 16,
    "vbhiragana": 32, "vbunicode": 64, "vbfromunicode": 128,
    "vbblack": 0, "vbred": 255, "vbgreen": 65280, "vbyellow": 65535, "vbblue": 16711680, "vbmagenta": 16711935,
    "vbcyan": 16776960, "vbwhite": 16777215,
    "vbmethod": 1, "vbget": 2, "vblet": 4, "vbset": 8,
    "vbnormal": 0, "vbreadonly": 1, "vbhidden": 2, "vbsystem": 4, "vbvolume": 8, "vbdirectory": 16, "vbarchive": 32, "vbalias": 64,
    "vbhide": 0, "vbnormalfocus": 1, "vbminimizedfocus": 2, "vbmaximizedfocus": 3, "vbnormalnofocus": 4, "vbminimizednofocus": 6,
    "vbimenoop": 0, "vbimemodenocontrol": 0, "vbimeon": 1, "vbimemodeon": 1, "vbimeoff": 2, "vbimemodeoff": 2,
    "vbimedisable": 3, "vbimemodedisable": 3, "vbimehiragana": 4, "vbimemodehiragana": 4, "vbimekatakanadbl": 5,
    "vbimemodekatakana": 5, "vbimekatakanasng": 6, "vbimemodekatakanahalf": 6, "vbimealphadbl": 7,
    "vbimemodealphafull": 7, "vbimealphasng": 8, "vbimemodealpha": 8, "vbimemodehangulfull": 9, "vbimemodehangul": 10,
    "vbcalgreg": 0, "vbcalhijri": 1,
    "vbkeylbutton": 1, "vbkeyrbutton": 2, "vbkeycancel": 3, "vbkeymbutton": 4, "vbkeyback": 8, "vbkeytab": 9,
    "vbkeyclear": 12, "vbkeyreturn": 13, "vbkeyshift": 16, "vbkeycontrol": 17, "vbkeymenu": 18, "vbkeypause": 19,
    "vbkeycapital": 20, "vbkeyescape": 27, "vbkeyspace": 32, "vbkeypageup": 33, "vbkeypagedown": 34, "vbkeyend": 35,
    "vbkeyhome": 36, "vbkeyleft": 37, "vbkeyup": 38, "vbkeyright": 39, "vbkeydown": 40, "vbkeyselect": 41,
    "vbkeyprint": 42, "vbkeyexecute": 43, "vbkeysnapshot": 44, "vbkeyinsert": 45, "vbkeydelete": 46, "vbkeyhelp": 47,
    "vbkeynumlock": 144, "vbkeyscrolllock": 145, "vbkeymultiply": 106, "vbkeyadd": 107, "vbkeyseparator": 108,
    "vbkeysubtract": 109, "vbkeydecimal": 110, "vbkeydivide": 111,
    **{f"vbkey{chr(c).lower()}": c for c in range(65, 91)},
    **{f"vbkey{d}": 48 + d for d in range(10)},
    **{f"vbkeynumpad{d}": 96 + d for d in range(10)},
    **{f"vbkeyf{n}": 111 + n for n in range(1, 17)},
}

# VBA's enumeration types (``Dim r As VbMsgBoxResult`` declares a Long) and their members.
ENUMS: Dict[str, Tuple[str, ...]] = {
    "vbvartype": ("vbempty", "vbnull", "vbinteger", "vblong", "vbsingle", "vbdouble", "vbcurrency", "vbdate", "vbstring", "vbobject", "vberror", "vbboolean", "vbvariant", "vbdataobject", "vbdecimal", "vbbyte", "vblonglong", "vbuserdefinedtype", "vbarray"),
    "vbmsgboxstyle": ("vbokonly", "vbokcancel", "vbabortretryignore", "vbyesnocancel", "vbyesno", "vbretrycancel", "vbcritical", "vbquestion", "vbexclamation", "vbinformation", "vbdefaultbutton1", "vbdefaultbutton2", "vbdefaultbutton3", "vbdefaultbutton4", "vbapplicationmodal", "vbsystemmodal", "vbmsgboxhelpbutton", "vbmsgboxright", "vbmsgboxrtlreading", "vbmsgboxsetforeground"),
    "vbmsgboxresult": ("vbok", "vbcancel", "vbabort", "vbretry", "vbignore", "vbyes", "vbno"),
    "vbfileattribute": ("vbnormal", "vbreadonly", "vbhidden", "vbsystem", "vbvolume", "vbdirectory", "vbarchive", "vbalias"),
    "vbstrconv": ("vbuppercase", "vblowercase", "vbpropercase", "vbwide", "vbnarrow", "vbkatakana", "vbhiragana", "vbunicode", "vbfromunicode"),
    "vbdayofweek": ("vbusesystemdayofweek", "vbsunday", "vbmonday", "vbtuesday", "vbwednesday", "vbthursday", "vbfriday", "vbsaturday"),
    "vbfirstweekofyear": ("vbusesystem", "vbfirstjan1", "vbfirstfourdays", "vbfirstfullweek"),
    "vbimestatus": ("vbimenoop", "vbimeon", "vbimeoff", "vbimedisable", "vbimehiragana", "vbimekatakanadbl", "vbimekatakanasng", "vbimealphadbl", "vbimealphasng"),
    "vbappwinstyle": ("vbhide", "vbnormalfocus", "vbminimizedfocus", "vbmaximizedfocus", "vbnormalnofocus", "vbminimizednofocus"),
    "vbcomparemethod": ("vbbinarycompare", "vbtextcompare", "vbdatabasecompare"),
    "vbcalendar": ("vbcalgreg", "vbcalhijri"),
    "vbdatetimeformat": ("vbgeneraldate", "vblongdate", "vbshortdate", "vblongtime", "vbshorttime"),
    "vbtristate": ("vbusedefault", "vbtrue", "vbfalse"),
    "vbcalltype": ("vbmethod", "vbget", "vblet", "vbset"),
    "colorconstants": ("vbblack", "vbred", "vbgreen", "vbyellow", "vbblue", "vbmagenta", "vbcyan", "vbwhite"),
}
