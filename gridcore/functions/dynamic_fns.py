"""Dynamic-array functions: FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, RANDARRAY, VSTACK, HSTACK,
TOROW, TOCOL, WRAPROWS, WRAPCOLS, TAKE, DROP, CHOOSEROWS, CHOOSECOLS, EXPAND, GROUPBY,
PIVOTBY, PERCENTOF."""

from __future__ import annotations

import functools
import math
from typing import Dict, List, Optional, Tuple

from ..core.value import Array, CellError, Value, XlError, scalar, sort_compare, to_bool, to_number, to_text
from .registry import MAX_CELLS, A, Arg, Ctx, VAR, fn
from .util import array_val, as_array, has, num, opt_bool

LR = "Lookup & Reference"
M = "Math & Trig"


def _arr(args, i: int) -> Array:
    v = args[i].value
    if isinstance(v, CellError):
        raise XlError(v)
    return as_array(v)


def _rows(a: Array) -> List[List[Value]]:
    return a.to_rows()


def _from_rows(rows: List[List[Value]], pad: Value) -> Value:
    h = len(rows)
    w = max((len(r) for r in rows), default=0)
    if h == 0 or w == 0:
        raise XlError(CellError.CALC)
    if h * w > MAX_CELLS:
        raise XlError(CellError.NUM)
    data: List[Value] = []
    for r in rows:
        data.extend(r)
        if len(r) < w:
            data.extend([pad] * (w - len(r)))
    return array_val(h, w, data)


def _opt_count(args, i: int) -> Optional[int]:
    if not has(args, i):
        return None
    n = num(args, i)
    if not math.isfinite(n) or abs(n) > 1e15:
        raise XlError(CellError.NUM)
    return int(n)


def _pad(args, i: int) -> Value:
    if has(args, i):
        return scalar(args[i].value)
    return CellError.NA


def _key(v: Value):
    if v is None:
        return ("e",)
    if isinstance(v, bool):
        return ("b", v)
    if isinstance(v, float):
        return ("n", 0.0 if v == 0 else v)
    if isinstance(v, str):
        return ("t", v.lower())
    if isinstance(v, CellError):
        return ("x", v.code)
    return ("?",)


def _truthy(v: Value) -> bool:
    if isinstance(v, CellError):
        raise XlError(v)
    if v is None:
        return False
    if isinstance(v, str):
        raise XlError(CellError.VALUE)
    return to_bool(v)


@fn("FILTER", 2, 3, LR, (False, False, False), "FILTER(array, include, [if_empty])", "Filters a range of data based on criteria you define.")
def _filter(args, ctx):
    a = _arr(args, 0)
    inc = _arr(args, 1)
    keep_rows = inc.cols == 1 and inc.rows == a.rows
    keep_cols = not keep_rows and inc.rows == 1 and inc.cols == a.cols
    if not keep_rows and not keep_cols:
        raise XlError(CellError.VALUE)
    flags = [_truthy(v) for v in inc.data]
    grid = _rows(a)
    if keep_rows:
        out = [r for r, f in zip(grid, flags) if f]
    else:
        out = [[v for v, f in zip(r, flags) if f] for r in grid]
    if not out or all(not r for r in out):
        if has(args, 2):
            return args[2].value
        raise XlError(CellError.CALC)
    return _from_rows(out, None)


def _cmp_keys(a: List[Value], b: List[Value], orders: List[int]) -> int:
    for i, o in enumerate(orders):
        x, y = a[i], b[i]
        c = sort_compare(x, y)
        if o < 0:
            if x is None and y is None:
                c = 0
            elif x is None:
                c = 1
            elif y is None:
                c = -1
            else:
                c = -c
        if c:
            return c
    return 0


def _int_list(v: Value) -> List[int]:
    out = []
    for x in as_array(v).data:
        n = to_number(x)
        if not math.isfinite(n) or abs(n) > 1e15:
            raise XlError(CellError.VALUE)
        out.append(int(n))
    return out


@fn("SORT", 1, 4, LR, A, "SORT(array, [sort_index], [sort_order], [by_col])", "Sorts the contents of a range or array.")
def _sort(args, ctx):
    a = _arr(args, 0)
    by_col = opt_bool(args, 3, False)
    if by_col:
        a = a.transpose()
    idx = _int_list(args[1].value) if has(args, 1) else [1]
    ords = _int_list(args[2].value) if has(args, 2) else [1]
    if not idx or any(i < 1 or i > a.cols for i in idx) or any(o not in (1, -1) for o in ords):
        raise XlError(CellError.VALUE)
    if len(ords) > 1 and len(ords) != len(idx):
        raise XlError(CellError.VALUE)
    orders = [ords[i] if i < len(ords) else ords[0] for i in range(len(idx))]
    lines = [([r[i - 1] for i in idx], r) for r in _rows(a)]
    lines.sort(key=functools.cmp_to_key(lambda x, y: _cmp_keys(x[0], y[0], orders)))
    out = Array(a.rows, a.cols, [v for _, r in lines for v in r])
    return out.transpose() if by_col else out


@fn("SORTBY", 2, VAR, LR, A, "SORTBY(array, by_array1, [sort_order1], …)", "Sorts the contents of a range or array based on the values in a corresponding range or array.")
def _sortby(args, ctx):
    a = _arr(args, 0)
    keys: List[Array] = []
    orders: List[int] = []
    i = 1
    while i < len(args):
        keys.append(_arr(args, i))
        o = int(num(args, i + 1)) if has(args, i + 1) else 1
        if o not in (1, -1):
            raise XlError(CellError.VALUE)
        orders.append(o)
        i += 2
    first = keys[0]
    by_rows = first.cols == 1 and first.rows == a.rows
    by_cols = not by_rows and first.rows == 1 and first.cols == a.cols
    if not by_rows and not by_cols:
        raise XlError(CellError.VALUE)
    for k in keys:
        ok = (k.cols == 1 and k.rows == a.rows) if by_rows else (k.rows == 1 and k.cols == a.cols)
        if not ok:
            raise XlError(CellError.VALUE)
    base = a if by_rows else a.transpose()
    lines = [([k.data[r] for k in keys], row) for r, row in enumerate(_rows(base))]
    lines.sort(key=functools.cmp_to_key(lambda x, y: _cmp_keys(x[0], y[0], orders)))
    out = Array(base.rows, base.cols, [v for _, r in lines for v in r])
    return out if by_rows else out.transpose()


@fn("UNIQUE", 1, 3, LR, A, "UNIQUE(array, [by_col], [exactly_once])", "Returns a list of unique values in a list or range.")
def _unique(args, ctx):
    a = _arr(args, 0)
    by_col = opt_bool(args, 1, False)
    once = opt_bool(args, 2, False)
    base = a.transpose() if by_col else a
    lines = _rows(base)
    keys = [tuple(_key(v) for v in l) for l in lines]
    counts: Dict[tuple, int] = {}
    for k in keys:
        counts[k] = counts.get(k, 0) + 1
    seen = set()
    out = []
    for k, l in zip(keys, lines):
        if once and counts[k] != 1:
            continue
        if k not in seen:
            seen.add(k)
            out.append(l)
    if not out:
        raise XlError(CellError.CALC)
    v = _from_rows(out, None)
    return v.transpose() if by_col and isinstance(v, Array) else v


def _dims(args, ri: int, ci: int) -> Tuple[int, int]:
    r = math.trunc(num(args, ri)) if has(args, ri) else 1
    c = math.trunc(num(args, ci)) if has(args, ci) else 1
    if r < 0 or c < 0:
        raise XlError(CellError.VALUE)
    if r == 0 or c == 0:
        raise XlError(CellError.CALC)
    if r > 1_048_576 or c > 16_384 or r * c > MAX_CELLS:
        raise XlError(CellError.NUM)
    return r, c


@fn("SEQUENCE", 1, 4, M, A, "SEQUENCE(rows, [columns], [start], [step])", "Generates a list of sequential numbers in an array.")
def _sequence(args, ctx):
    r, c = _dims(args, 0, 1)
    start = num(args, 2) if has(args, 2) else 1.0
    step = num(args, 3) if has(args, 3) else 1.0
    return array_val(r, c, [start + step * i for i in range(r * c)])


@fn("RANDARRAY", 0, 5, M, A, "RANDARRAY([rows], [columns], [min], [max], [whole_number])", "Returns an array of random numbers.", volatile=True)
def _randarray(args, ctx):
    r, c = _dims(args, 0, 1)
    lo = num(args, 2) if has(args, 2) else 0.0
    hi = num(args, 3) if has(args, 3) else 1.0
    whole = opt_bool(args, 4, False) if has(args, 4) else False
    if lo > hi:
        raise XlError(CellError.VALUE)
    data = []
    if whole:
        lo, hi = math.ceil(lo), math.floor(hi)
        if lo > hi:
            raise XlError(CellError.VALUE)
        for _ in range(r * c):
            data.append(float(min(lo + math.floor(ctx.random() * (hi - lo + 1)), hi)))
    else:
        for _ in range(r * c):
            data.append(lo + ctx.random() * (hi - lo))
    return array_val(r, c, data)


@fn("VSTACK", 1, VAR, LR, A, "VSTACK(array1, [array2], …)", "Appends arrays vertically and in sequence to return a larger array.")
def _vstack(args, ctx):
    rows: List[List[Value]] = []
    total = 0
    widest = 0
    for i in range(len(args)):
        a = _arr(args, i)
        total += len(a.data)
        widest = max(widest, a.cols)
        # The padded result (every row as wide as the widest array) is what must fit, not the
        # inputs: VSTACK(SEQUENCE(1,16384),SEQUENCE(1048576)) would build 17 billion cells.
        if total > MAX_CELLS or (len(rows) + a.rows) * widest > MAX_CELLS:
            raise XlError(CellError.NUM)
        rows.extend(_rows(a))
    return _from_rows(rows, CellError.NA)


@fn("HSTACK", 1, VAR, LR, A, "HSTACK(array1, [array2], …)", "Appends arrays horizontally and in sequence to return a larger array.")
def _hstack(args, ctx):
    arrays = [_arr(args, i) for i in range(len(args))]
    # The padded result (the tallest array's rows times every column) is what must fit, not the
    # inputs: HSTACK(SEQUENCE(1048576),SEQUENCE(1,16384)) would build 17 billion cells.
    if max(a.rows for a in arrays) * sum(a.cols for a in arrays) > MAX_CELLS:
        raise XlError(CellError.NUM)
    h = max(a.rows for a in arrays)
    rows: List[List[Value]] = [[] for _ in range(h)]
    for a in arrays:
        for r in range(h):
            for c in range(a.cols):
                rows[r].append(a.get(r, c, CellError.NA))
    return _from_rows(rows, CellError.NA)


def _flat_filtered(args) -> List[Value]:
    a = _arr(args, 0)
    ignore = int(num(args, 1)) if has(args, 1) else 0
    if not 0 <= ignore <= 3:
        raise XlError(CellError.VALUE)
    if opt_bool(args, 2, False):
        a = a.transpose()
    out = []
    for v in a.data:
        if v is None and ignore & 1:
            continue
        if isinstance(v, CellError) and ignore & 2:
            continue
        out.append(v)
    return out


@fn("TOROW", 1, 3, LR, A, "TOROW(array, [ignore], [scan_by_column])", "Returns the array in a single row.")
def _torow(args, ctx):
    v = _flat_filtered(args)
    if not v:
        raise XlError(CellError.CALC)
    return array_val(1, len(v), v)


@fn("TOCOL", 1, 3, LR, A, "TOCOL(array, [ignore], [scan_by_column])", "Returns the array in a single column.")
def _tocol(args, ctx):
    v = _flat_filtered(args)
    if not v:
        raise XlError(CellError.CALC)
    return array_val(len(v), 1, v)


def _wrap(args, rows: bool) -> Value:
    a = _arr(args, 0)
    if a.rows != 1 and a.cols != 1:
        raise XlError(CellError.VALUE)
    n = num(args, 1)
    if n < 1 or not math.isfinite(n):
        raise XlError(CellError.NUM)
    pad = _pad(args, 2)
    ln = len(a.data)
    n = min(int(n), max(ln, 1))
    out = []
    for i in range(0, ln, n):
        line = list(a.data[i:i + n])
        line += [pad] * (n - len(line))
        out.append(line)
    v = _from_rows(out, pad)
    return v.transpose() if not rows and isinstance(v, Array) else v


@fn("WRAPROWS", 2, 3, LR, A, "WRAPROWS(vector, wrap_count, [pad_with])", "Wraps the provided row or column of values by rows after a specified number of elements.")
def _wraprows(args, ctx):
    return _wrap(args, True)


@fn("WRAPCOLS", 2, 3, LR, A, "WRAPCOLS(vector, wrap_count, [pad_with])", "Wraps the provided row or column of values by columns after a specified number of elements.")
def _wrapcols(args, ctx):
    return _wrap(args, False)


def _span(ln: int, n: Optional[int], take: bool) -> Tuple[int, int]:
    if n is None:
        return 0, ln
    k = min(abs(n), ln)
    if take:
        if n == 0:
            raise XlError(CellError.CALC)
        return (0, k) if n > 0 else (ln - k, ln)
    r = (k, ln) if n >= 0 else (0, ln - k)
    if r[0] >= r[1]:
        raise XlError(CellError.CALC)
    return r


def _take_drop(args, take: bool) -> Value:
    a = _arr(args, 0)
    r0, r1 = _span(a.rows, _opt_count(args, 1), take)
    c0, c1 = _span(a.cols, _opt_count(args, 2), take)
    rows = [[a.get(r, c) for c in range(c0, c1)] for r in range(r0, r1)]
    return _from_rows(rows, None)


@fn("TAKE", 2, 3, LR, A, "TAKE(array, rows, [columns])", "Returns a specified number of contiguous rows or columns from the start or end of an array.")
def _take(args, ctx):
    return _take_drop(args, True)


@fn("DROP", 2, 3, LR, A, "DROP(array, rows, [columns])", "Excludes a specified number of rows or columns from the start or end of an array.")
def _drop(args, ctx):
    return _take_drop(args, False)


def _choose_lines(args, rows: bool) -> Value:
    a = _arr(args, 0)
    base = a if rows else a.transpose()
    grid = _rows(base)
    out = []
    for i in range(1, len(args)):
        for k in _int_list(args[i].value):
            ln = len(grid)
            if 0 < k <= ln:
                idx = k - 1
            elif k < 0 and -k <= ln:
                idx = ln + k
            else:
                raise XlError(CellError.VALUE)
            out.append(grid[idx])
            if len(out) * base.cols > MAX_CELLS:
                raise XlError(CellError.NUM)
    v = _from_rows(out, None)
    return v.transpose() if not rows and isinstance(v, Array) else v


@fn("CHOOSEROWS", 2, VAR, LR, A, "CHOOSEROWS(array, row_num1, [row_num2], …)", "Returns the specified rows from an array.")
def _chooserows(args, ctx):
    return _choose_lines(args, True)


@fn("CHOOSECOLS", 2, VAR, LR, A, "CHOOSECOLS(array, col_num1, [col_num2], …)", "Returns the specified columns from an array.")
def _choosecols(args, ctx):
    return _choose_lines(args, False)


@fn("EXPAND", 2, 4, LR, A, "EXPAND(array, rows, [columns], [pad_with])", "Expands or pads an array to specified row and column dimensions.")
def _expand(args, ctx):
    a = _arr(args, 0)
    r = _opt_count(args, 1)
    c = _opt_count(args, 2)
    r = a.rows if r is None else r
    c = a.cols if c is None else c
    if r < a.rows or c < a.cols:
        raise XlError(CellError.VALUE)
    if r > 1_048_576 or c > 16_384 or r * c > MAX_CELLS:
        raise XlError(CellError.NUM)
    pad = _pad(args, 3)
    data = [a.get(i, j, pad) if i < a.rows and j < a.cols else pad for i in range(r) for j in range(c)]
    return array_val(r, c, data)


# ---------------------------------------------------------------- GROUPBY / PIVOTBY

_AGG_NAMES = {
    "SUM": "sum", "AVERAGE": "average", "COUNT": "count", "COUNTA": "counta", "MAX": "max", "MIN": "min",
    "PRODUCT": "product", "MEDIAN": "median", "STDEV.S": "stdev", "STDEV.P": "stdevp", "VAR.S": "var",
    "VAR.P": "varp", "PERCENTOF": "percentof", "ARRAYTOTEXT": "arraytotext", "CONCAT": "concat",
}
_AGG_CODES = {1: "average", 2: "count", 3: "counta", 4: "max", 5: "min", 6: "product", 7: "stdev", 8: "stdevp", 9: "sum", 10: "var", 11: "varp"}


def _agg_kind(v: Value) -> str:
    v = scalar(v)
    if isinstance(v, float):
        k = _AGG_CODES.get(int(v))
        if k is None:
            raise XlError(CellError.VALUE)
        return k
    if isinstance(v, str):
        k = _AGG_NAMES.get(v.strip().upper().replace("_XLFN.", ""))
        if k is None:
            raise XlError(CellError.VALUE)
        return k
    raise XlError(CellError.VALUE)


def _aggregate(kind: str, vals: List[Value], total: Optional[List[Value]] = None) -> Value:
    from .stat import mean, median, var_p, var_s

    nums = [v for v in vals if type(v) is float]
    for v in vals:
        if isinstance(v, CellError):
            return v
    try:
        if kind == "sum":
            return math.fsum(nums)
        if kind == "average":
            return mean(nums)
        if kind == "count":
            return float(len(nums))
        if kind == "counta":
            return float(sum(1 for v in vals if v is not None))
        if kind == "max":
            return max(nums) if nums else 0.0
        if kind == "min":
            return min(nums) if nums else 0.0
        if kind == "product":
            p = 1.0
            for x in nums:
                p *= x
            return p if nums else 0.0
        if kind == "median":
            return median(nums)
        if kind == "stdev":
            return math.sqrt(var_s(nums))
        if kind == "stdevp":
            return math.sqrt(var_p(nums))
        if kind == "var":
            return var_s(nums)
        if kind == "varp":
            return var_p(nums)
        if kind == "percentof":
            t = math.fsum(v for v in (total or []) if type(v) is float)
            if t == 0:
                return CellError.DIV0
            return math.fsum(nums) / t
        if kind in ("arraytotext", "concat"):
            return (", " if kind == "arraytotext" else "").join(to_text(v) for v in vals if v is not None)
    except XlError as e:
        return e.error
    return CellError.VALUE


def _group_keys(a: Array, r: int) -> tuple:
    return tuple(_key(v) for v in a.row_list(r))


@fn("GROUPBY", 3, 8, LR, A, "GROUPBY(row_fields, values, function, [field_headers], [total_depth], [sort_order], [filter_array], [field_relationship])", "Allows you to create a summary of your data via a formula.")
def _groupby(args, ctx):
    keys = _arr(args, 0)
    vals = _arr(args, 1)
    kind = _agg_kind(args[2].value)
    if keys.rows != vals.rows:
        raise XlError(CellError.VALUE)
    headers = int(num(args, 3)) if has(args, 3) else 0
    total_depth = int(num(args, 4)) if has(args, 4) else 1
    sort_order = int(num(args, 5)) if has(args, 5) else 1
    filt = _arr(args, 6) if has(args, 6) else None
    start = 1 if headers in (1, 3) else 0
    groups: Dict[tuple, Tuple[List[Value], List[List[Value]]]] = {}
    order: List[tuple] = []
    all_vals: List[List[Value]] = [[] for _ in range(vals.cols)]
    for r in range(start, keys.rows):
        if filt is not None and r < len(filt.data) and not _truthy(filt.data[r]):
            continue
        k = _group_keys(keys, r)
        if k not in groups:
            groups[k] = (keys.row_list(r), [[] for _ in range(vals.cols)])
            order.append(k)
        for c in range(vals.cols):
            groups[k][1][c].append(vals.get(r, c))
            all_vals[c].append(vals.get(r, c))
    if not order:
        raise XlError(CellError.CALC)
    if sort_order:
        col = abs(sort_order) - 1
        if col < keys.cols:
            order.sort(key=functools.cmp_to_key(lambda x, y: sort_compare(groups[x][0][col], groups[y][0][col])), reverse=sort_order < 0)
        else:
            vc = col - keys.cols
            agg = {k: _aggregate(kind, groups[k][1][vc] if vc < vals.cols else [], all_vals[vc] if vc < vals.cols else []) for k in order}
            order.sort(key=functools.cmp_to_key(lambda x, y: sort_compare(agg[x], agg[y])), reverse=sort_order < 0)
    out: List[List[Value]] = []
    if headers in (1, 3) and start == 1:
        out.append(keys.row_list(0) + vals.row_list(0))
    for k in order:
        gk, gv = groups[k]
        out.append(list(gk) + [_aggregate(kind, gv[c], all_vals[c]) for c in range(vals.cols)])
    if total_depth:
        out.append(["Total"] + [None] * (keys.cols - 1) + [_aggregate(kind, all_vals[c], all_vals[c]) for c in range(vals.cols)])
    return _from_rows(out, None)


@fn("PIVOTBY", 4, 11, LR, A, "PIVOTBY(row_fields, col_fields, values, function, [field_headers], [row_total_depth], [row_sort_order], [col_total_depth], [col_sort_order], [filter_array], [relative_to])", "Allows you to create a summary of your data via a formula.")
def _pivotby(args, ctx):
    rk = _arr(args, 0)
    ck = _arr(args, 1)
    vals = _arr(args, 2)
    kind = _agg_kind(args[3].value)
    if not (rk.rows == ck.rows == vals.rows):
        raise XlError(CellError.VALUE)
    headers = int(num(args, 4)) if has(args, 4) else 0
    start = 1 if headers in (1, 3) else 0
    rows_order: List[tuple] = []
    cols_order: List[tuple] = []
    rlab: Dict[tuple, Value] = {}
    clab: Dict[tuple, Value] = {}
    cells: Dict[Tuple[tuple, tuple], List[Value]] = {}
    for r in range(start, rk.rows):
        a, b = _group_keys(rk, r), _group_keys(ck, r)
        if a not in rlab:
            rlab[a] = rk.get(r, 0)
            rows_order.append(a)
        if b not in clab:
            clab[b] = ck.get(r, 0)
            cols_order.append(b)
        cells.setdefault((a, b), []).append(vals.get(r, 0))
    if not rows_order:
        raise XlError(CellError.CALC)
    rows_order.sort(key=functools.cmp_to_key(lambda x, y: sort_compare(rlab[x], rlab[y])))
    cols_order.sort(key=functools.cmp_to_key(lambda x, y: sort_compare(clab[x], clab[y])))
    out = [[None] + [clab[c] for c in cols_order] + ["Total"]]
    for rkey in rows_order:
        line: List[Value] = [rlab[rkey]]
        allv: List[Value] = []
        for ckey in cols_order:
            vs = cells.get((rkey, ckey))
            if vs:
                allv.extend(vs)
            line.append(_aggregate(kind, vs) if vs else None)
        line.append(_aggregate(kind, allv))
        out.append(line)
    total = ["Total"]
    everything: List[Value] = []
    for ckey in cols_order:
        vs = [v for rkey in rows_order for v in cells.get((rkey, ckey), [])]
        everything.extend(vs)
        total.append(_aggregate(kind, vs))
    total.append(_aggregate(kind, everything))
    out.append(total)
    return _from_rows(out, None)


@fn("PERCENTOF", 2, 2, M, A, "PERCENTOF(data_subset, data_all)", "Returns the percentage that a subset makes up of a given data set.")
def _percentof(args, ctx):
    sub = [v for v in as_array(args[0].value).data if type(v) is float]
    allv = [v for v in as_array(args[1].value).data if type(v) is float]
    t = math.fsum(allv)
    if t == 0:
        raise XlError(CellError.DIV0)
    return math.fsum(sub) / t
