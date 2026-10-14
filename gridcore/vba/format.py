"""VBA's ``Format`` function (US English): named formats, date and time codes (``yyyy-mm-dd
hh:nn:ss``, ``ddd``, ``q``, ``ww``…), number codes (shared with Excel's number formats) and
string codes (``@ & < > !``), plus FormatNumber/FormatCurrency/FormatPercent."""

from __future__ import annotations

import datetime as _dt
import decimal
import math
import re
from typing import Any, List, Optional, Sequence

from . import values as V
from .errors import VbaError
from .values import EMPTY, MISSING, NULL, Currency, Date, VArray, is_object

_NAMED = {
    "general number": None,
    "currency": "$#,##0.00;($#,##0.00)",
    "fixed": "0.00",
    "standard": "#,##0.00",
    "percent": "0.00%",
    "scientific": "0.00E+00",
    "yes/no": ("Yes", "No"),
    "true/false": ("True", "False"),
    "on/off": ("On", "Off"),
}
_DATE_NAMED = {
    "general date": "c",
    "long date": "dddddd",
    "medium date": "dd-mmm-yy",
    "short date": "ddddd",
    "long time": "ttttt",
    "medium time": "hh:nn AM/PM",
    "short time": "hh:nn",
}


def _num_value(v: Any) -> Optional[Any]:
    t = type(v)
    if t is str:
        x = V.parse_number(v)
        if x is None:
            d = V.parse_date_text(v)
            return None if d is None else Date(d)
        return float(x) if not isinstance(x, int) else x
    if t is bool:
        return -1 if v else 0
    if v is EMPTY:
        return 0
    if t in (int, V.Integer, V.Byte, V.LongLong, float, V.Single, Currency, decimal.Decimal, Date):
        return v
    return None


def vba_format(v: Any, fmt: str, first: int = 1, firstweek: int = 1) -> Any:
    if is_object(v):
        v = V.value_of(v)
    if type(v) is VArray:
        raise VbaError(13)
    low = fmt.strip().lower()
    if v is NULL:
        secs = _split_sections(fmt)
        return _literal_only(secs[3]) if len(secs) >= 4 else ""
    if not fmt:
        if type(v) is Date:
            return V.date_text(v)
        if type(v) is bool:
            return "True" if v else "False"
        if v is EMPTY:
            return ""
        return V.text(v)
    if low in _DATE_NAMED:
        d = _as_date(v)
        if d is None:
            return V.text(v)
        if low == "general date":
            return V.date_text(d)
        return _format_date(d, _DATE_NAMED[low], first, firstweek)
    if low in _NAMED:
        n = _num_value(v)
        if n is None:
            return V.text(v)
        named = _NAMED[low]
        if named is None:
            return V.text(n if type(n) is not Date else float(n))
        if isinstance(named, tuple):
            return named[0] if V.to_double(n) != 0 else named[1]
        return _format_number(float(n) if type(n) is not Currency else float(n), named)
    kind = _kind(fmt)
    if kind == "date":
        d = _as_date(v)
        if d is None:
            return V.text(v)
        return _format_date(d, fmt, first, firstweek)
    if kind == "number":
        n = _num_value(v)
        if n is None:
            return V.text(v)
        return _format_number(float(n) if not isinstance(n, decimal.Decimal) else float(n), fmt)
    if kind == "string":
        return _format_string(V.text(v), fmt)
    # Only literal text: VBA shows the format text itself.
    n = _num_value(v)
    if n is not None and type(v) is not str:
        return _format_number(float(n), fmt)
    return _literal_only(fmt)


def _as_date(v: Any) -> Optional[float]:
    if type(v) is Date:
        return float(v)
    if type(v) is str:
        d = V.parse_date_text(v)
        if d is not None:
            return d
        x = V.parse_number(v)
        return float(x) if x is not None else None
    n = _num_value(v)
    if n is None:
        return None
    try:
        return float(V.to_date(n))
    except VbaError:
        return None


def _split_sections(fmt: str) -> List[str]:
    out = []
    cur = []
    i = 0
    while i < len(fmt):
        c = fmt[i]
        if c == '"':
            j = fmt.find('"', i + 1)
            j = len(fmt) if j < 0 else j
            cur.append(fmt[i:j + 1])
            i = j + 1
            continue
        if c == "\\" and i + 1 < len(fmt):
            cur.append(fmt[i:i + 2])
            i += 2
            continue
        if c == ";":
            out.append("".join(cur))
            cur = []
            i += 1
            continue
        cur.append(c)
        i += 1
    out.append("".join(cur))
    return out


def _strip_literals(fmt: str) -> str:
    out = []
    i = 0
    while i < len(fmt):
        c = fmt[i]
        if c == '"':
            j = fmt.find('"', i + 1)
            i = len(fmt) if j < 0 else j + 1
            continue
        if c == "\\":
            i += 2
            continue
        out.append(c)
        i += 1
    return "".join(out)


def _kind(fmt: str) -> str:
    bare = _strip_literals(_split_sections(fmt)[0]).lower()
    if re.search(r"am/pm|a/p|ampm", bare):
        return "date"
    if re.search(r"[0#]", bare) or ("%" in bare and re.search(r"[0#.]", bare)):
        return "number"
    if re.search(r"[dmyhnswqc]|ttttt", bare):
        return "date"
    if re.search(r"[@&<>!]", bare):
        return "string"
    if "e+" in bare or "e-" in bare or "." in bare or "," in bare or "%" in bare:
        return "number"
    return "literal"


def _literal_only(fmt: str) -> str:
    out = []
    i = 0
    while i < len(fmt):
        c = fmt[i]
        if c == '"':
            j = fmt.find('"', i + 1)
            j = len(fmt) if j < 0 else j
            out.append(fmt[i + 1:j])
            i = j + 1
            continue
        if c == "\\" and i + 1 < len(fmt):
            out.append(fmt[i + 1])
            i += 2
            continue
        out.append(c)
        i += 1
    return "".join(out)


def _format_number(x: float, fmt: str) -> str:
    from ..core.date import DateSystem
    from ..numfmt import format_value, parse

    secs = _split_sections(fmt)
    code = ";".join(secs[:3])
    if x == 0 and len(secs) >= 3 and secs[2] == "":
        code = ";".join(secs[:2])
    try:
        r = format_value(float(x), parse(code), DateSystem.D1900, 255)
    except Exception:
        return V.text(x)
    t = r.text
    if r.fill is not None:
        ch, pos = r.fill
        t = t[:pos] + t[pos:]
    return t


def _format_string(s: str, fmt: str) -> str:
    secs = _split_sections(fmt)
    f = secs[1] if (s == "" and len(secs) > 1) else secs[0]
    case = None
    left_to_right = False
    toks: List[tuple] = []
    i = 0
    while i < len(f):
        c = f[i]
        if c == '"':
            j = f.find('"', i + 1)
            j = len(f) if j < 0 else j
            toks.append(("lit", f[i + 1:j]))
            i = j + 1
            continue
        if c == "\\" and i + 1 < len(f):
            toks.append(("lit", f[i + 1]))
            i += 2
            continue
        if c in "@&":
            toks.append((c, ""))
        elif c == "<":
            case = "lower"
        elif c == ">":
            case = "upper"
        elif c == "!":
            left_to_right = True
        else:
            toks.append(("lit", c))
        i += 1
    if case == "lower":
        s = s.lower()
    elif case == "upper":
        s = s.upper()
    n_ph = sum(1 for t in toks if t[0] in "@&")
    if n_ph == 0:
        return "".join(t[1] for t in toks) + s
    chars = list(s)
    out: List[str] = []
    if left_to_right:
        k = 0
        for t, lit in toks:
            if t == "lit":
                out.append(lit)
            else:
                if k < len(chars):
                    out.append(chars[k])
                elif t == "@":
                    out.append(" ")
                k += 1
        if k < len(chars):
            out.append("".join(chars[k:]))
        return "".join(out)
    # Placeholders fill from the right; extra characters go before the first placeholder.
    extra = max(len(chars) - n_ph, 0)
    k = len(chars) - 1
    rev: List[str] = []
    for t, lit in reversed(toks):
        if t == "lit":
            rev.append(lit)
        else:
            if k >= extra and k >= 0:
                rev.append(chars[k])
            elif t == "@":
                rev.append(" ")
            k -= 1
    res = "".join(reversed(rev))
    if extra:
        first = next(i for i, t in enumerate(toks) if t[0] in "@&")
        prefix = "".join(t[1] for t in toks[:first])
        res = prefix + "".join(chars[:extra]) + res[len(prefix):]
    return res


_DATE_TOKEN = re.compile(r"ttttt|AM/PM|am/pm|A/P|a/p|AMPM|ampm|dddddd|ddddd|dddd|ddd|dd|d|mmmm|mmm|mm|m|yyyy|yyy|yy|y|hh|h|nn|n|ss|s|ww|w|q|c|oooo|aaaa", re.I)


def _format_date(serial: float, fmt: str, first: int = 1, firstweek: int = 1) -> str:
    from .stdlib import _weekday, week_of_year

    d = V.serial_to_datetime(serial)
    sec = _split_sections(fmt)[0]
    toks: List[tuple] = []
    i = 0
    while i < len(sec):
        c = sec[i]
        if c == '"':
            j = sec.find('"', i + 1)
            j = len(sec) if j < 0 else j
            toks.append(("lit", sec[i + 1:j]))
            i = j + 1
            continue
        if c == "\\" and i + 1 < len(sec):
            toks.append(("lit", sec[i + 1]))
            i += 2
            continue
        m = _DATE_TOKEN.match(sec, i)
        if m:
            toks.append(("tok", m.group(0)))
            i = m.end()
            continue
        toks.append(("lit", c))
        i += 1
    has_ampm = any(t == "tok" and v.lower() in ("am/pm", "a/p", "ampm") for t, v in toks)
    out: List[str] = []
    for k, (t, v) in enumerate(toks):
        if t == "lit":
            out.append(v)
            continue
        lv = v.lower()
        if lv in ("m", "mm"):
            # Minutes when it follows an hour or precedes seconds.
            prev = next((tv.lower() for tt, tv in reversed(toks[:k]) if tt == "tok"), "")
            nxt = next((tv.lower() for tt, tv in toks[k + 1:] if tt == "tok"), "")
            if prev in ("h", "hh") or nxt in ("s", "ss"):
                out.append(str(d.minute) if lv == "m" else f"{d.minute:02d}")
            else:
                out.append(str(d.month) if lv == "m" else f"{d.month:02d}")
        elif lv == "mmm":
            out.append(V.MONTHS[d.month - 1][:3])
        elif lv == "mmmm":
            out.append(V.MONTHS[d.month - 1])
        elif lv == "d":
            out.append(str(d.day))
        elif lv == "dd":
            out.append(f"{d.day:02d}")
        elif lv == "ddd":
            out.append(V.DAYS[_weekday(d, 1) - 1][:3])
        elif lv == "dddd":
            out.append(V.DAYS[_weekday(d, 1) - 1])
        elif lv == "ddddd":
            out.append(f"{d.month}/{d.day}/{d.year}")
        elif lv == "dddddd":
            out.append(f"{V.DAYS[_weekday(d, 1) - 1]}, {V.MONTHS[d.month - 1]} {d.day}, {d.year}")
        elif lv == "aaaa":
            out.append(V.DAYS[_weekday(d, 1) - 1])
        elif lv == "oooo":
            out.append(V.MONTHS[d.month - 1])
        elif lv == "w":
            out.append(str(_weekday(d, first)))
        elif lv == "ww":
            out.append(str(week_of_year(d.date(), first, firstweek)))
        elif lv == "q":
            out.append(str((d.month - 1) // 3 + 1))
        elif lv == "y":
            out.append(str(d.timetuple().tm_yday))
        elif lv == "yy":
            out.append(f"{d.year % 100:02d}")
        elif lv == "yyy":
            out.append(f"{d.year % 100:02d}{d.timetuple().tm_yday}")
        elif lv == "yyyy":
            out.append(f"{d.year:04d}" if d.year >= 1000 else str(d.year))
        elif lv in ("h", "hh"):
            h = d.hour
            if has_ampm:
                h = h % 12 or 12
            out.append(str(h) if lv == "h" else f"{h:02d}")
        elif lv == "n":
            out.append(str(d.minute))
        elif lv == "nn":
            out.append(f"{d.minute:02d}")
        elif lv == "s":
            out.append(str(d.second))
        elif lv == "ss":
            out.append(f"{d.second:02d}")
        elif lv == "ttttt":
            h = d.hour % 12 or 12
            out.append(f"{h}:{d.minute:02d}:{d.second:02d} {'AM' if d.hour < 12 else 'PM'}")
        elif lv == "c":
            out.append(V.date_text(serial))
        elif lv == "am/pm":
            am = d.hour < 12
            out.append(("AM" if am else "PM") if v[0].isupper() else ("am" if am else "pm"))
        elif lv == "a/p":
            am = d.hour < 12
            out.append(("A" if am else "P") if v[0].isupper() else ("a" if am else "p"))
        elif lv == "ampm":
            out.append("AM" if d.hour < 12 else "PM")
        else:
            out.append(v)
    return "".join(out)


def _tristate(v: Any, default: bool) -> bool:
    if v is MISSING:
        return default
    n = V.to_long(V.value_of(v) if is_object(v) else v)
    if n == -2:
        return default
    return n != 0


def format_number(v: Any, args: Sequence[Any], currency: bool, percent: bool) -> str:
    if v is NULL:
        raise VbaError(94)
    x = V.to_double(v)
    digits = V.to_long(args[0]) if len(args) > 0 and args[0] is not MISSING else -1
    if digits == -1:
        digits = 2
    if digits < -1:
        raise VbaError(5)
    lead = _tristate(args[1] if len(args) > 1 else MISSING, True)
    parens = _tristate(args[2] if len(args) > 2 else MISSING, currency)
    group = _tristate(args[3] if len(args) > 3 else MISSING, True)
    if percent:
        x *= 100
    neg = x < 0
    q = decimal.Decimal(repr(abs(x))).quantize(decimal.Decimal(1).scaleb(-digits), rounding=decimal.ROUND_HALF_UP)
    if q == 0:
        neg = False
    body = f"{q:,.{digits}f}" if group else f"{q:.{digits}f}"
    if not lead and body.startswith("0.") and digits > 0:
        body = body[1:]
    if currency:
        body = "$" + body
    if percent:
        body += "%"
    if neg:
        return f"({body})" if parens else "-" + body
    return body
