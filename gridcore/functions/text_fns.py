"""Text functions."""

from __future__ import annotations

import math
import re
import unicodedata
import urllib.parse
from typing import List, Optional

from ..core.parse import parse_number_text_in
from ..core.value import Array, CellError, Value, XlError, number_to_text, scalar, to_text
from .criteria import fold, wildcard_find
from .registry import MAX_TEXT, A, Arg, Ctx, S, VAR, fn
from .util import array_val, as_array, boolean, has, int_, num, opt_int, opt_num, round_half_away, text

T = "Text"


def _limit(s: str) -> str:
    if len(s) > MAX_TEXT:
        raise XlError(CellError.VALUE)
    return s


def _flat_text(args: List[Arg]) -> List[str]:
    out = []
    for a in args:
        v = a.value
        items = v.data if isinstance(v, Array) else [v]
        for x in items:
            if isinstance(x, CellError):
                raise XlError(x)
            out.append(to_text(x))
    return out


@fn("CONCATENATE", 1, VAR, T, S, "CONCATENATE(text1, [text2], …)", "Joins several text items into one text item.")
def _concatenate(args, ctx):
    return _limit("".join(text(args, i) for i in range(len(args))))


@fn("CONCAT", 1, VAR, T, A, "CONCAT(text1, [text2], …)", "Combines the text from multiple ranges and/or strings.")
def _concat(args, ctx):
    return _limit("".join(_flat_text(args)))


@fn("TEXTJOIN", 3, VAR, T, (True, True, False), "TEXTJOIN(delimiter, ignore_empty, text1, [text2], …)", "Combines the text from multiple ranges and/or strings, with a delimiter.")
def _textjoin(args, ctx):
    delims = as_array(args[0].value)
    ds = [to_text(d) for d in delims.data] or [""]
    ignore = boolean(args, 1)
    parts = _flat_text(args[2:])
    if ignore:
        parts = [p for p in parts if p != ""]
    out = []
    for i, p in enumerate(parts):
        if i:
            out.append(ds[(i - 1) % len(ds)])
        out.append(p)
    return _limit("".join(out))


def _count(args, i: int, default: int) -> int:
    n = opt_num(args, i, float(default)) if i < len(args) else float(default)
    if n < 0:
        raise XlError(CellError.VALUE)
    return int(n)


@fn("LEFT", 1, 2, T, S, "LEFT(text, [num_chars])", "Returns the leftmost characters from a text value.")
def _left(args, ctx):
    return text(args, 0)[: _count(args, 1, 1)]


@fn("RIGHT", 1, 2, T, S, "RIGHT(text, [num_chars])", "Returns the rightmost characters from a text value.")
def _right(args, ctx):
    t = text(args, 0)
    n = _count(args, 1, 1)
    return t[len(t) - n:] if n else ""


@fn("MID", 3, 3, T, S, "MID(text, start_num, num_chars)", "Returns a specific number of characters from a text string starting at the position you specify.")
def _mid(args, ctx):
    t = text(args, 0)
    start, n = num(args, 1), num(args, 2)
    if start < 1 or n < 0:
        raise XlError(CellError.VALUE)
    s = int(start) - 1
    return t[s:s + int(n)]


@fn("LEN", 1, 1, T, S, "LEN(text)", "Returns the number of characters in a text string.")
def _len(args, ctx):
    return float(len(text(args, 0)))


def _byte_len(s: str) -> int:
    return sum(2 if unicodedata.east_asian_width(c) in ("W", "F") else 1 for c in s)


@fn("LENB", 1, 1, T, S, "LENB(text)", "Returns the number of bytes used to represent the characters in a text string.")
def _lenb(args, ctx):
    return float(_byte_len(text(args, 0)))


for _n, _t in (("LEFTB", "LEFT"), ("RIGHTB", "RIGHT"), ("MIDB", "MID")):
    from .registry import alias as _alias

    _alias(_n, _t)


# Excel changes the case of each letter into one letter (Excel 16 for Mac): UPPER("straße") is
# STRAßE, LOWER("İ") is i, and a letter whose upper case is two letters (ŉ, ﬀ) stays as it is.
def _upper1(c: str) -> str:
    u = c.upper()
    return u if len(u) == 1 else c


def _lower1(c: str) -> str:
    return c.lower()[0]


@fn("LOWER", 1, 1, T, S, "LOWER(text)", "Converts text to lowercase.")
def _lower(args, ctx):
    t = text(args, 0)
    low = t.lower()
    return low if len(low) == len(t) else "".join(_lower1(c) for c in t)


@fn("UPPER", 1, 1, T, S, "UPPER(text)", "Converts text to uppercase.")
def _upper(args, ctx):
    t = text(args, 0)
    u = t.upper()
    return u if len(u) == len(t) else "".join(_upper1(c) for c in t)


@fn("PROPER", 1, 1, T, S, "PROPER(text)", "Capitalizes the first letter in each word of a text value.")
def _proper(args, ctx):
    t = text(args, 0)
    out = []
    prev_letter = False
    for c in t:
        if c.isalpha():
            out.append(_lower1(c) if prev_letter else _upper1(c))
            prev_letter = True
        else:
            out.append(c)
            prev_letter = False
    return "".join(out)


@fn("TRIM", 1, 1, T, S, "TRIM(text)", "Removes spaces from text.")
def _trim(args, ctx):
    return " ".join(p for p in text(args, 0).split(" ") if p)


@fn("CLEAN", 1, 1, T, S, "CLEAN(text)", "Removes all nonprintable characters from text.")
def _clean(args, ctx):
    return "".join(c for c in text(args, 0) if ord(c) >= 32)


@fn("REPLACE", 4, 4, T, S, "REPLACE(old_text, start_num, num_chars, new_text)", "Replaces characters within text.")
def _replace(args, ctx):
    t = text(args, 0)
    start, n = num(args, 1), num(args, 2)
    if start < 1 or n < 0:
        raise XlError(CellError.VALUE)
    s = int(start) - 1
    return _limit(t[:s] + text(args, 3) + t[s + int(n):])


from .registry import alias as _alias2  # noqa: E402

_alias2("REPLACEB", "REPLACE")


@fn("SUBSTITUTE", 3, 4, T, S, "SUBSTITUTE(text, old_text, new_text, [instance_num])", "Substitutes new text for old text in a text string.")
def _substitute(args, ctx):
    t, old, new = text(args, 0), text(args, 1), text(args, 2)
    if old == "":
        return t
    if has(args, 3):
        k = int(num(args, 3))
        if k < 1:
            raise XlError(CellError.VALUE)
        idx = -1
        for _ in range(k):
            idx = t.find(old, idx + 1)
            if idx < 0:
                return t
        return _limit(t[:idx] + new + t[idx + len(old):])
    return _limit(t.replace(old, new))


@fn("REPT", 2, 2, T, S, "REPT(text, number_times)", "Repeats text a given number of times.")
def _rept(args, ctx):
    t = text(args, 0)
    n = num(args, 1)
    if n < 0:
        raise XlError(CellError.VALUE)
    if len(t) * int(n) > MAX_TEXT:
        raise XlError(CellError.VALUE)
    return t * int(n)


@fn("FIND", 2, 3, T, S, "FIND(find_text, within_text, [start_num])", "Finds one text value within another (case-sensitive).")
def _find(args, ctx):
    needle, hay = text(args, 0), text(args, 1)
    start = int(opt_num(args, 2, 1.0)) if has(args, 2) else 1
    # FIND may start just past the end (FIND("","abc",4) is 4); SEARCH may not (Excel 16 for Mac).
    if start < 1 or start > len(hay) + 1:
        raise XlError(CellError.VALUE)
    i = hay.find(needle, start - 1)
    if i < 0:
        raise XlError(CellError.VALUE)
    return float(i + 1)


@fn("SEARCH", 2, 3, T, S, "SEARCH(find_text, within_text, [start_num])", "Finds one text value within another (not case-sensitive, wildcards allowed).")
def _search(args, ctx):
    needle, hay = text(args, 0), text(args, 1)
    start = int(opt_num(args, 2, 1.0)) if has(args, 2) else 1
    # Not past the end, so nothing is found in empty text: SEARCH("","") is #VALUE!.
    if start < 1 or start > len(hay):
        raise XlError(CellError.VALUE)
    i = wildcard_find(needle, hay, start - 1)
    if i is None:
        raise XlError(CellError.VALUE)
    return float(i + 1)


_alias2("FINDB", "FIND")
_alias2("SEARCHB", "SEARCH")


@fn("EXACT", 2, 2, T, S, "EXACT(text1, text2)", "Checks to see if two text values are identical.")
def _exact(args, ctx):
    return text(args, 0) == text(args, 1)


@fn("CHAR", 1, 1, T, S, "CHAR(number)", "Returns the character specified by the code number.")
def _char(args, ctx):
    n = int(num(args, 0))
    if not 1 <= n <= 255:
        raise XlError(CellError.VALUE)
    return bytes([n]).decode("cp1252", errors="replace") if 128 <= n <= 159 else chr(n)


@fn("CODE", 1, 1, T, S, "CODE(text)", "Returns a numeric code for the first character in a text string.")
def _code(args, ctx):
    t = text(args, 0)
    if not t:
        raise XlError(CellError.VALUE)
    c = t[0]
    try:
        b = c.encode("cp1252")
        return float(b[0])
    except UnicodeEncodeError:
        return 63.0


@fn("UNICHAR", 1, 1, T, S, "UNICHAR(number)", "Returns the Unicode character that is referenced by the given numeric value.")
def _unichar(args, ctx):
    n = int(num(args, 0))
    if not 1 <= n <= 0x10FFFF or 0xD800 <= n <= 0xDFFF:
        raise XlError(CellError.VALUE)
    return chr(n)


@fn("UNICODE", 1, 1, T, S, "UNICODE(text)", "Returns the number (code point) that corresponds to the first character of the text.")
def _unicode(args, ctx):
    t = text(args, 0)
    if not t:
        raise XlError(CellError.VALUE)
    return float(ord(t[0]))


@fn("VALUE", 1, 1, T, S, "VALUE(text)", "Converts a text argument to a number.")
def _value(args, ctx):
    v = scalar(args[0].value)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, bool):
        raise XlError(CellError.VALUE)
    if v is None:
        return 0.0
    n = parse_number_text_in(to_text(v), ctx.date_system)
    if n is None:
        raise XlError(CellError.VALUE)
    return n


@fn("NUMBERVALUE", 1, 3, T, S, "NUMBERVALUE(text, [decimal_separator], [group_separator])", "Converts text to number in a locale-independent manner.")
def _numbervalue(args, ctx):
    t = text(args, 0)
    dec = text(args, 1)[:1] if has(args, 1) else "."
    grp = text(args, 2)[:1] if has(args, 2) else ","
    if not dec or dec == grp:
        raise XlError(CellError.VALUE)
    s = "".join(t.split())
    if s == "":
        return 0.0
    pct = 0
    while s.endswith("%"):
        pct += 1
        s = s[:-1]
    if dec in s:
        ip, fp = s.split(dec, 1)
        if grp in fp:
            raise XlError(CellError.VALUE)
    else:
        ip, fp = s, ""
    ip = ip.replace(grp, "")
    s2 = ip + ("." + fp if fp else "")
    try:
        n = float(s2)
    except ValueError:
        raise XlError(CellError.VALUE) from None
    return n / (100.0 ** pct)


def _fixed_text(n: float, decimals: int, commas: bool) -> str:
    from ..numfmt import format_value, parse as parse_fmt

    r = round_half_away(n, decimals)
    d = max(decimals, 0)
    code = ("#,##0" if commas else "0") + ("." + "0" * d if d else "")
    return format_value(r, parse_fmt(code)).text


@fn("FIXED", 1, 3, T, S, "FIXED(number, [decimals], [no_commas])", "Formats a number as text with a fixed number of decimals.")
def _fixed(args, ctx):
    n = num(args, 0)
    d = int(opt_num(args, 1, 2.0)) if has(args, 1) else 2
    if d > 127:
        raise XlError(CellError.VALUE)
    no_commas = boolean(args, 2) if has(args, 2) else False
    return _fixed_text(n, d, not no_commas)


@fn("DOLLAR", 1, 2, T, S, "DOLLAR(number, [decimals])", "Converts a number to text, using the $ (dollar) currency format.")
def _dollar(args, ctx):
    n = num(args, 0)
    d = int(opt_num(args, 1, 2.0)) if has(args, 1) else 2
    if d > 127:
        raise XlError(CellError.VALUE)
    body = _fixed_text(abs(n), d, True)
    return f"(${body})" if round_half_away(n, d) < 0 else f"${body}"


# An older name of DOLLAR that Excel still calculates.
_alias2("USDOLLAR", "DOLLAR", None, "Compatibility")


@fn("T", 1, 1, T, S, "T(value)", "Converts its arguments to text.")
def _t(args, ctx):
    v = scalar(args[0].value)
    if isinstance(v, CellError):
        raise XlError(v)
    return v if isinstance(v, str) else ""


def _delims(v: Value) -> List[str]:
    a = as_array(v)
    out = []
    for x in a.data:
        if isinstance(x, CellError):
            raise XlError(x)
        out.append(to_text(x))
    return out


def _find_delim(t: str, delims: List[str], start: int, nocase: bool) -> Optional[tuple]:
    best = None
    hay = fold(t) if nocase else t
    for d in delims:
        if d == "":
            continue
        i = hay.find(fold(d) if nocase else d, start)
        if i >= 0 and (best is None or i < best[0]):
            best = (i, len(d))
    return best


def _split_points_from_end(t: str, delims: List[str], nocase: bool) -> List[tuple]:
    """The delimiters of ``t`` found searching from its end, as a negative instance_num of
    TEXTBEFORE/TEXTAFTER does (``aa`` in ``aaa`` is the last two letters), in text order."""
    hay = fold(t) if nocase else t
    ds = sorted({fold(d) if nocase else d for d in delims if d != ""}, key=len, reverse=True)
    pts = []
    end = len(hay)
    while end > 0:
        d = next((d for d in ds if len(d) <= end and hay.startswith(d, end - len(d))), None)
        if d is None:
            end -= 1
            continue
        pts.append((end - len(d), len(d)))
        end -= len(d)
    pts.reverse()
    return pts


def _text_split_points(t: str, delims: List[str], nocase: bool) -> List[tuple]:
    pts = []
    pos = 0
    while True:
        f = _find_delim(t, delims, pos, nocase)
        if f is None:
            break
        pts.append(f)
        pos = f[0] + max(f[1], 1)
    return pts


def _before_after(args, after: bool):
    t = text(args, 0)
    delims = _delims(args[1].value)
    n = int(opt_num(args, 2, 1.0)) if has(args, 2) else 1
    nocase = bool(opt_num(args, 3, 0.0)) if has(args, 3) else False
    match_end = bool(opt_num(args, 4, 0.0)) if has(args, 4) else False
    if n == 0 or abs(n) > max(len(t), 1) + 1:
        raise XlError(CellError.VALUE)
    if all(d == "" for d in delims):
        return t if (n > 0) != after else ""
    pts = _text_split_points(t, delims, nocase) if n > 0 else _split_points_from_end(t, delims, nocase)
    if match_end:
        if n > 0:
            pts = pts + [(len(t), 0)]
        else:
            pts = [(0, 0)] + pts
    if abs(n) > len(pts):
        if len(args) > 5:
            return args[5].value if not isinstance(args[5].value, Array) else scalar(args[5].value)
        raise XlError(CellError.NA)
    p = pts[n - 1] if n > 0 else pts[n]
    return t[p[0] + p[1]:] if after else t[: p[0]]


@fn("TEXTBEFORE", 2, 6, T, (True, False, True, True, True, True), "TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])", "Returns text that occurs before a given character or string.")
def _textbefore(args, ctx):
    return _before_after(args, False)


@fn("TEXTAFTER", 2, 6, T, (True, False, True, True, True, True), "TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])", "Returns text that occurs after a given character or string.")
def _textafter(args, ctx):
    return _before_after(args, True)


@fn("TEXTSPLIT", 2, 6, T, (True, False, False, True, True, True), "TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])", "Splits text strings by using column and row delimiters.")
def _textsplit(args, ctx):
    t = text(args, 0)
    cols_d = _delims(args[1].value) if has(args, 1) else []
    rows_d = _delims(args[2].value) if has(args, 2) else []
    ignore = boolean(args, 3) if has(args, 3) else False
    nocase = bool(opt_num(args, 4, 0.0)) if has(args, 4) else False
    pad = scalar(args[5].value) if has(args, 5) else CellError.NA
    if not cols_d and not rows_d:
        raise XlError(CellError.VALUE)

    def split(s: str, ds: List[str]) -> List[str]:
        if not ds or all(d == "" for d in ds):
            return [s]
        out, pos = [], 0
        for i, ln in _text_split_points(s, ds, nocase):
            out.append(s[pos:i])
            pos = i + ln
        out.append(s[pos:])
        return out

    rows = split(t, rows_d)
    if ignore:
        rows = [r for r in rows if r != ""]
    grid = [split(r, cols_d) for r in rows]
    if ignore:
        grid = [[c for c in r if c != ""] for r in grid]
        grid = [r for r in grid if r]
    if not grid:
        raise XlError(CellError.CALC)
    w = max(len(r) for r in grid)
    data: List[Value] = []
    for r in grid:
        data.extend(r + [pad] * (w - len(r)))
    return array_val(len(grid), w, data)


def _value_to_text(v: Value, strict: bool) -> str:
    if isinstance(v, str):
        return '"' + v.replace('"', '""') + '"' if strict else v
    if isinstance(v, CellError):
        return v.value
    return to_text(v)


@fn("VALUETOTEXT", 1, 2, T, (False, True), "VALUETOTEXT(value, [format])", "Returns text from any specified value.")
def _valuetotext(args, ctx):
    v = args[0].value
    strict = bool(opt_num(args, 1, 0.0)) if has(args, 1) else False

    def one(x):
        return _value_to_text(x, strict)

    if isinstance(v, Array):
        return Array(v.rows, v.cols, [one(x) for x in v.data])
    return one(v)


@fn("ARRAYTOTEXT", 1, 2, T, (False, True), "ARRAYTOTEXT(array, [format])", "Returns an array of text values from any specified range.")
def _arraytotext(args, ctx):
    a = as_array(args[0].value)
    strict = bool(opt_num(args, 1, 0.0)) if has(args, 1) else False
    if not strict:
        return ", ".join(_value_to_text(x, False) for x in a.data)
    rows = [",".join(_value_to_text(x, True) for x in a.row_list(r)) for r in range(a.rows)]
    return "{" + ";".join(rows) + "}"


def _regex(pattern: str, case_sensitivity: int):
    try:
        return re.compile(pattern, re.IGNORECASE if case_sensitivity == 1 else 0)
    except re.error:
        raise XlError(CellError.VALUE) from None


@fn("REGEXTEST", 2, 3, T, S, "REGEXTEST(text, pattern, [case_sensitivity])", "Determines whether any part of text matches the pattern.")
def _regextest(args, ctx):
    rx = _regex(text(args, 1), int(opt_num(args, 2, 0.0)) if has(args, 2) else 0)
    return rx.search(text(args, 0)) is not None


@fn("REGEXEXTRACT", 2, 4, T, S, "REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity])", "Extracts strings within the provided text that matches the pattern.")
def _regexextract(args, ctx):
    t = text(args, 0)
    mode = int(opt_num(args, 2, 0.0)) if has(args, 2) else 0
    rx = _regex(text(args, 1), int(opt_num(args, 3, 0.0)) if has(args, 3) else 0)
    if mode == 0:
        m = rx.search(t)
        if m is None:
            raise XlError(CellError.NA)
        return m.group(0)
    if mode == 1:
        ms = [m.group(0) for m in rx.finditer(t)]
        if not ms:
            raise XlError(CellError.NA)
        return array_val(len(ms), 1, ms)
    if mode == 2:
        m = rx.search(t)
        if m is None:
            raise XlError(CellError.NA)
        groups = [g if g is not None else "" for g in m.groups()] or [m.group(0)]
        return array_val(1, len(groups), groups)
    raise XlError(CellError.VALUE)


@fn("REGEXREPLACE", 3, 5, T, S, "REGEXREPLACE(text, pattern, replacement, [occurrence], [case_sensitivity])", "Replaces strings within the provided text that matches the pattern with replacement.")
def _regexreplace(args, ctx):
    t = text(args, 0)
    rx = _regex(text(args, 1), int(opt_num(args, 4, 0.0)) if has(args, 4) else 0)
    rep = re.sub(r"\$(\d+)", r"\\\1", text(args, 2))
    occ = int(opt_num(args, 3, 0.0)) if has(args, 3) else 0
    if occ == 0:
        return _limit(rx.sub(rep, t))
    ms = list(rx.finditer(t))
    idx = occ - 1 if occ > 0 else len(ms) + occ
    if not 0 <= idx < len(ms):
        return t
    m = ms[idx]
    return _limit(t[: m.start()] + m.expand(rep) + t[m.end():])


@fn("ENCODEURL", 1, 1, "Web", S, "ENCODEURL(text)", "Returns a URL-encoded string.")
def _encodeurl(args, ctx):
    return urllib.parse.quote(text(args, 0), safe="-_.~")


def _width_map(s: str, to_full: bool) -> str:
    out = []
    for c in s:
        o = ord(c)
        if to_full and 0x21 <= o <= 0x7E:
            out.append(chr(o + 0xFEE0))
        elif to_full and c == " ":
            out.append("　")
        elif not to_full and 0xFF01 <= o <= 0xFF5E:
            out.append(chr(o - 0xFEE0))
        elif not to_full and c == "　":
            out.append(" ")
        else:
            out.append(c)
    return "".join(out)


@fn("ASC", 1, 1, T, S, "ASC(text)", "Changes full-width (double-byte) English letters or katakana within a character string to half-width (single-byte) characters.")
def _asc(args, ctx):
    return _width_map(text(args, 0), False)


@fn("DBCS", 1, 1, T, S, "DBCS(text)", "Changes half-width (single-byte) English letters or katakana within a character string to full-width (double-byte) characters.")
def _dbcs(args, ctx):
    return _width_map(text(args, 0), True)


_alias2("JIS", "DBCS")


@fn("PHONETIC", 1, 1, T, A, "PHONETIC(reference)", "Extracts the phonetic (furigana) characters from a text string.")
def _phonetic(args, ctx):
    v = scalar(args[0].value)
    return to_text(v) if not isinstance(v, CellError) else v
