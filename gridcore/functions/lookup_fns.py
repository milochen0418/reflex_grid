"""Lookup and reference functions over evaluated arrays: VLOOKUP, HLOOKUP, LOOKUP, MATCH,
XLOOKUP, XMATCH, TRANSPOSE, ADDRESS, HYPERLINK.

Reference-returning behaviour (INDEX, OFFSET, ROW…) belongs to the evaluator."""

from __future__ import annotations

import bisect
from typing import List, Optional

from ..core.addr import col_to_letters
from ..core.value import Array, CellError, Value, XlError, compare, scalar
from .criteria import has_wildcards, lookup_equal
from .registry import A, Arg, Ctx, S, fn
from .util import array_val, as_array, has, int_, num, opt_bool, opt_int, opt_num, text

LR = "Lookup & Reference"


def _needle(args, i: int) -> Value:
    v = scalar(args[i].value)
    if isinstance(v, CellError):
        raise XlError(v)
    return 0.0 if v is None else v


def _same_type(a: Value, b: Value) -> bool:
    ta, tb = type(a), type(b)
    return (ta is float and tb is float) or (ta is str and tb is str) or (ta is bool and tb is bool)


def find_exact(x: Value, hay: List[Value], wild: bool, reverse: bool) -> Optional[int]:
    wild = wild and isinstance(x, str) and has_wildcards(x)
    rng = range(len(hay) - 1, -1, -1) if reverse else range(len(hay))
    if type(x) is str and not wild:
        xl = x.lower()
        for i in rng:
            v = hay[i]
            if type(v) is str and v.lower() == xl:
                return i
        return None
    for i in rng:
        if lookup_equal(x, hay[i], wild):
            return i
    return None


def bsearch_last(x: Value, hay: List[Value], desc: bool) -> Optional[int]:
    """Approximate match of VLOOKUP, HLOOKUP, LOOKUP and MATCH: Excel's bisection, skipping
    entries of another type than the needle."""
    if not hay:
        return None
    lo, hi = 0, len(hay) - 1
    best = None
    while lo <= hi:
        mid = lo + (hi - lo) // 2
        j = None
        for k in range(mid, lo - 1, -1):
            if _same_type(x, hay[k]):
                j = k
                break
        if j is None:
            lo = mid + 1
            continue
        o = compare(hay[j], x)
        if o == 0:
            k = j
            while k + 1 < len(hay) and _same_type(x, hay[k + 1]) and compare(hay[k + 1], x) == 0:
                k += 1
            return k
        before = o > 0 if desc else o < 0
        if before:
            best = j
            lo = mid + 1
        elif j == 0:
            break
        else:
            hi = j - 1
    return best


def _binary_order(v: Value, x: Value, desc: bool) -> int:
    if v is None:
        return 1
    o = compare(v, x)
    return -o if desc else o


def _bsearch_first(x: Value, hay: List[Value], desc: bool) -> Optional[int]:
    lo, hi = 0, len(hay)
    while lo < hi:
        mid = (lo + hi) // 2
        if _binary_order(hay[mid], x, desc) < 0:
            lo = mid + 1
        else:
            hi = mid
    return lo if lo < len(hay) else None


def _bsearch_upto(x: Value, hay: List[Value], desc: bool) -> Optional[int]:
    lo, hi = 0, len(hay)
    while lo < hi:
        mid = (lo + hi) // 2
        if _binary_order(hay[mid], x, desc) <= 0:
            lo = mid + 1
        else:
            hi = mid
    return lo - 1 if lo > 0 else None


def _bsearch_exact(x: Value, hay: List[Value], desc: bool) -> Optional[int]:
    p = _bsearch_first(x, hay, desc)
    if p is None:
        return None
    return p if _binary_order(hay[p], x, desc) == 0 else None


def _vector(a: Array) -> Optional[List[Value]]:
    if a.rows == 1 or a.cols == 1:
        return list(a.data)
    return None


def _cell_result(v: Value) -> Value:
    return 0.0 if v is None else v


def _vh_lookup(args, vertical: bool) -> Value:
    x = _needle(args, 0)
    tv = args[1].value
    if isinstance(tv, CellError):
        raise XlError(tv)
    table = as_array(tv)
    idx = num(args, 2)
    if idx < 1:
        raise XlError(CellError.VALUE)
    idx = int(idx) - 1
    approx = opt_bool(args, 3, True)
    if vertical:
        keys, ln = table.col_list(0), table.cols
    else:
        keys, ln = table.row_list(0), table.rows
    if idx >= ln:
        raise XlError(CellError.REF)
    pos = bsearch_last(x, keys, False) if approx else find_exact(x, keys, True, False)
    if pos is None:
        raise XlError(CellError.NA)
    return _cell_result(table.get(pos, idx) if vertical else table.get(idx, pos))


@fn("VLOOKUP", 3, 4, LR, (True, False, True, True), "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])", "Looks in the first column of an array and moves across the row to return the value of a cell.")
def _vlookup(args, ctx):
    return _vh_lookup(args, True)


@fn("HLOOKUP", 3, 4, LR, (True, False, True, True), "HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])", "Looks in the top row of an array and returns the value of the indicated cell.")
def _hlookup(args, ctx):
    return _vh_lookup(args, False)


@fn("LOOKUP", 2, 3, LR, (True, False, False), "LOOKUP(lookup_value, lookup_vector, [result_vector])", "Looks up values in a vector or array.")
def _lookup(args, ctx):
    x = _needle(args, 0)
    lv = args[1].value
    if isinstance(lv, CellError):
        raise XlError(lv)
    la = as_array(lv)
    if len(args) >= 3:
        keys = list(la.data) if (la.rows == 1 or la.cols == 1) else la.col_list(0)
        pos = bsearch_last(x, keys, False)
        if pos is None:
            raise XlError(CellError.NA)
        rv = args[2].value
        if isinstance(rv, CellError):
            raise XlError(rv)
        ra = as_array(rv)
        if ra.rows == 1 or ra.cols == 1:
            if pos >= len(ra.data):
                raise XlError(CellError.NA)
            return _cell_result(ra.data[pos])
        if ra.rows >= len(keys):
            return _cell_result(ra.get(pos, 0))
        raise XlError(CellError.NA)
    if la.cols > la.rows:
        keys = la.row_list(0)
        pos = bsearch_last(x, keys, False)
        if pos is None:
            raise XlError(CellError.NA)
        return _cell_result(la.get(la.rows - 1, pos))
    keys = la.col_list(0)
    pos = bsearch_last(x, keys, False)
    if pos is None:
        raise XlError(CellError.NA)
    return _cell_result(la.get(pos, la.cols - 1))


@fn("MATCH", 2, 3, LR, (True, False, True), "MATCH(lookup_value, lookup_array, [match_type])", "Looks up values in a reference or array.")
def _match(args, ctx):
    x = _needle(args, 0)
    av = args[1].value
    if isinstance(av, CellError):
        raise XlError(av)
    hay = _vector(as_array(av))
    if hay is None:
        raise XlError(CellError.NA)
    mt = opt_num(args, 2, 1.0) if has(args, 2) else 1.0
    if mt == 0:
        pos = find_exact(x, hay, True, False)
    elif mt > 0:
        pos = bsearch_last(x, hay, False)
    else:
        pos = bsearch_last(x, hay, True)
    if pos is None:
        raise XlError(CellError.NA)
    return float(pos + 1)


def xsearch(x: Value, hay: List[Value], match_mode: int, search_mode: int) -> Optional[int]:
    if match_mode not in (-1, 0, 1, 2) or search_mode not in (-2, -1, 1, 2):
        raise XlError(CellError.VALUE)
    if abs(search_mode) == 2:
        desc = search_mode == -2
        if match_mode in (0, 2):
            return _bsearch_exact(x, hay, desc)
        if match_mode == -1:
            return _bsearch_first(x, hay, True) if desc else _bsearch_upto(x, hay, False)
        return _bsearch_upto(x, hay, True) if desc else _bsearch_first(x, hay, False)
    reverse = search_mode == -1
    p = find_exact(x, hay, match_mode == 2, reverse)
    if p is not None:
        return p
    if match_mode in (0, 2):
        return None
    order = range(len(hay) - 1, -1, -1) if reverse else range(len(hay))
    best = None
    for i in order:
        v = hay[i]
        if not _same_type(x, v):
            continue
        o = compare(v, x)
        if not (o < 0 if match_mode == -1 else o > 0):
            continue
        if best is None:
            best = i
        else:
            c = compare(v, hay[best])
            if (c > 0) if match_mode == -1 else (c < 0):
                best = i
    return best


@fn("XLOOKUP", 3, 6, LR, (True, False, False, False, False, False), "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])", "Searches a range or an array, and returns an item corresponding to the first match it finds.")
def _xlookup(args, ctx):
    x = _needle(args, 0)
    lv, rv = args[1].value, args[2].value
    if isinstance(lv, CellError):
        raise XlError(lv)
    if isinstance(rv, CellError):
        raise XlError(rv)
    la, ra = as_array(lv), as_array(rv)
    hay = _vector(la)
    if hay is None:
        raise XlError(CellError.VALUE)
    if la.rows == 1 and la.cols == 1:
        vertical = ra.rows == 1 or ra.cols != 1
    else:
        vertical = la.cols == 1
    if (vertical and ra.rows != la.rows) or (not vertical and ra.cols != la.cols):
        raise XlError(CellError.VALUE)
    mm = opt_int(args, 4, 0) if has(args, 4) else 0
    sm = opt_int(args, 5, 1) if has(args, 5) else 1
    p = xsearch(x, hay, mm, sm)
    if p is None:
        if has(args, 3):
            return args[3].value
        raise XlError(CellError.NA)
    if vertical:
        r = ra.row_list(p)
        return _cell_result(r[0]) if len(r) == 1 else array_val(1, len(r), [_cell_result(v) for v in r])
    c = ra.col_list(p)
    return _cell_result(c[0]) if len(c) == 1 else array_val(len(c), 1, [_cell_result(v) for v in c])


@fn("XMATCH", 2, 4, LR, (True, False, True, True), "XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])", "Returns the relative position of an item in an array or range of cells.")
def _xmatch(args, ctx):
    x = _needle(args, 0)
    lv = args[1].value
    if isinstance(lv, CellError):
        raise XlError(lv)
    hay = _vector(as_array(lv))
    if hay is None:
        raise XlError(CellError.VALUE)
    mm = int_(args, 2) if has(args, 2) else 0
    sm = int_(args, 3) if has(args, 3) else 1
    p = xsearch(x, hay, mm, sm)
    if p is None:
        raise XlError(CellError.NA)
    return float(p + 1)


@fn("TRANSPOSE", 1, 1, LR, A, "TRANSPOSE(array)", "Returns the transpose of an array.")
def _transpose(args, ctx):
    v = args[0].value
    if isinstance(v, Array):
        return v.transpose()
    return v


def _quote_address_sheet(s: str) -> str:
    """Sheet text for ADDRESS, quoted as Excel 16 for Mac quotes it: a bare sheet name follows the
    rule references follow (``'A1'``, ``'RC'``, ``'.x'``, ``'TRUE'``, ``'My Sheet'``; ``XFE1`` and
    ``Sheet.1`` stay bare); ``[Book]Sheet`` is quoted when either part holds a character a name
    can't (``'[Book 1]Sheet1'``), and ``[Book1]A1`` stays bare."""
    from ..core.addr import sheet_needs_quotes

    def word(p: str) -> bool:
        return all(c.isalnum() or c in "_." for c in p)

    if s.startswith("[") and "]" in s:
        book, sheet = s[1:].split("]", 1)
        plain = bool(book) and bool(sheet) and word(book) and word(sheet) and not sheet[0].isdigit()
    else:
        plain = not sheet_needs_quotes(s)
    return s if plain else "'" + s.replace("'", "''") + "'"


@fn("ADDRESS", 2, 5, LR, S, "ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])", "Returns a reference as text to a single cell in a worksheet.")
def _address(args, ctx):
    r, c = int(num(args, 0)), int(num(args, 1))
    ab = int_(args, 2) if has(args, 2) else 1
    a1 = opt_bool(args, 3, True) if has(args, 3) else True
    if not 1 <= r <= 1_048_576 or not 1 <= c <= 16_384 or not 1 <= ab <= 4:
        raise XlError(CellError.VALUE)
    row_abs, col_abs = {1: (True, True), 2: (True, False), 3: (False, True), 4: (False, False)}[ab]
    if a1:
        cell = f"{'$' if col_abs else ''}{col_to_letters(c - 1)}{'$' if row_abs else ''}{r}"
    else:
        rp = f"R{r}" if row_abs else f"R[{r}]"
        cp = f"C{c}" if col_abs else f"C[{c}]"
        cell = rp + cp
    if has(args, 4):
        sh = text(args, 4)
        if sh == "":
            return "!" + cell
        return _quote_address_sheet(sh) + "!" + cell
    return cell


@fn("HYPERLINK", 1, 2, LR, S, "HYPERLINK(link_location, [friendly_name])", "Creates a shortcut or jump that opens a document stored on a network server, an intranet, or the Internet.")
def _hyperlink(args, ctx):
    if has(args, 1):
        v = scalar(args[1].value)
        return 0.0 if v is None else v
    return text(args, 0)
