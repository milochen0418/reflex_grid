"""Data commands: sort, AutoFilter (and table filters), remove duplicates, data validation and
text to columns."""

from __future__ import annotations

import functools
import math
from dataclasses import replace
from typing import Any, Dict, List, Optional, Tuple

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef, letters_to_col
from ..core.parse import parse_number_text
from ..core.value import CellError, compare, sort_compare
from ..formula.adjust import shift_relative
from ..model.cell import Cell, Formula
from ..model.features import AutoFilter, FilterCriterion, Validation
from ..model.style import Color
from .cmd_edit import _set, check_arrays
from .cmd_format import _ranges, parse_color
from .display import value_text
from .entry import input_to_cell, list_items, validation_at
from .ranges import remove_rules_in
from .selection import Selection, current_region
from .session import EngineError, Session, command, edit, need, p_bool, p_int, p_str, param_sheet

MAX_SORT_ROWS = 1_048_576


def _data_region(s: Session, p: Dict[str, Any]) -> Tuple[int, RangeRef, Optional[object]]:
    """The range a data command works on: a table or AutoFilter around the active cell, the
    multi-cell selection, or the current region."""
    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)
    if p.get("range"):
        return si, _ranges(s, p)[0], None
    a = d.selection.active
    sel = d.selection.current()
    t = sh.table_at(a)
    if t is not None and (sel.is_single() or t.range.contains_range(sel)):
        return si, t.range, t
    af = sh.autofilter
    if af is not None and af.range.contains(a) and sel.is_single():
        return si, af.range, af
    if not sel.is_single():
        used = sh.used_range()
        if (sel.is_full_cols() or sel.is_full_rows()) and used is not None:
            inter = sel.intersection(used)
            if inter is not None:
                return si, inter, None
        return si, sel, None
    region = current_region(sh, a)
    return si, region, None


def guess_header(sh, rng: RangeRef, wb) -> bool:
    if rng.height < 2:
        return False
    for c in range(rng.c0, min(rng.c1, rng.c0 + 200) + 1):
        v0 = sh.value((rng.r0, c))
        v1 = sh.value((rng.r0 + 1, c))
        if isinstance(v0, str) and v1 is not None and not isinstance(v1, str):
            return True
        b0 = wb.styles.get(sh.style_id((rng.r0, c))).font.bold
        b1 = wb.styles.get(sh.style_id((rng.r0 + 1, c))).font.bold
        if b0 and not b1 and v0 is not None:
            return True
    return False


def _col_param(v: Any, rng: RangeRef) -> int:
    """A sort/filter column: an absolute letter (``"B"``), an absolute index or an offset."""
    if isinstance(v, str):
        c = letters_to_col(v.strip().lstrip("$"))
        if c is None:
            raise EngineError(f"not a column: `{v}`")
        if not rng.c0 <= c <= rng.c1:
            raise EngineError(f"column {v} is outside the range {rng.a1()}")
        return c
    c = int(v)
    if rng.c0 <= c <= rng.c1:
        return c
    if 0 <= c < rng.width:
        return rng.c0 + c
    raise EngineError("column outside the range")


# ---------------------------------------------------------------- sort


def _cell_sort_value(wb, sh, rc, by: str, case: bool):
    if by == "cellColor":
        f = wb.styles.get(sh.style_id(rc)).fill
        return f.fg.hex(wb.theme) if f.pattern != "none" else None
    if by == "fontColor":
        return wb.styles.get(sh.style_id(rc)).font.color.hex(wb.theme)
    v = sh.value(rc)
    if isinstance(v, str) and v == "":
        return None
    return v


def sort_range(ctx, si: int, rng: RangeRef, keys: List[Dict[str, Any]], header: bool, by_cols: bool = False, case: bool = False) -> None:
    wb = ctx.wb
    sh = wb.sheet(si)
    if any(m.intersects(rng) for m in sh.merges):
        raise EngineError("To do this, all the merged cells need to be the same size.")
    data = RangeRef(rng.r0 + (1 if header else 0), rng.c0, rng.r1, rng.c1) if not by_cols else RangeRef(rng.r0, rng.c0 + (1 if header else 0), rng.r1, rng.c1)
    if data.r0 > data.r1 or data.c0 > data.c1:
        return
    if not by_cols:
        used = sh.used_range()
        if used is not None and data.r1 > used.r1:
            data = RangeRef(data.r0, data.c0, max(used.r1, data.r0), data.c1)
        lines = [r for r in range(data.r0, data.r1 + 1) if not sh.is_row_hidden(r)]
    else:
        used = sh.used_range()
        if used is not None and data.c1 > used.c1:
            data = RangeRef(data.r0, data.c0, data.r1, max(used.c1, data.c0))
        lines = [c for c in range(data.c0, data.c1 + 1) if not sh.is_col_hidden(c)]
    if len(lines) > MAX_SORT_ROWS:
        raise EngineError("The range is too large to sort.")
    # Each line's key values.
    specs = []
    for k in keys:
        col = k["col"]
        desc = k.get("order", "asc") in ("desc", "descending")
        by = k.get("by", "values")
        color = k.get("color")
        custom = k.get("customList")
        specs.append((col, desc, by, color, custom))

    def key_values(line: int):
        out = []
        for col, desc, by, color, custom in specs:
            rc = (line, col) if not by_cols else (col, line)
            out.append(_cell_sort_value(wb, sh, rc, by, case))
        return out

    rows = [(line, key_values(line)) for line in lines]

    def cmp(a, b) -> int:
        for i, (col, desc, by, color, custom) in enumerate(specs):
            x, y = a[1][i], b[1][i]
            if by in ("cellColor", "fontColor"):
                want = (color or "").upper() if color else None
                ix = 0 if (want is not None and (x or "").upper() == want) else 1
                iy = 0 if (want is not None and (y or "").upper() == want) else 1
                if ix != iy:
                    return (ix - iy) if not desc else (iy - ix)
                continue
            if x is None and y is None:
                continue
            if x is None:
                return 1
            if y is None:
                return -1
            if custom:
                lst = [str(e).lower() for e in custom]
                xi = lst.index(str(x).lower()) if isinstance(x, str) and str(x).lower() in lst else len(lst)
                yi = lst.index(str(y).lower()) if isinstance(y, str) and str(y).lower() in lst else len(lst)
                if xi != yi:
                    return (xi - yi) if not desc else (yi - xi)
            if case and isinstance(x, str) and isinstance(y, str):
                c = compare(x, y)
                if c == 0 and x != y:
                    # Lowercase sorts before uppercase in a case-sensitive sort.
                    c = -1 if x.swapcase() < y.swapcase() else 1
            else:
                c = sort_compare(x, y)
            if c != 0:
                return -c if desc else c
        return 0

    order = sorted(rows, key=functools.cmp_to_key(cmp))
    perm = {new_line: old_line for new_line, (old_line, _) in zip(lines, order)}
    if all(a == b for a, b in perm.items()):
        return
    sm = ctx.sheet_mut(si)
    # Take every moved line's cells, notes and links, then put them back permuted.
    taken: Dict[int, List[Tuple[Tuple[int, int], Cell]]] = {}
    for new_line, old_line in perm.items():
        if new_line == old_line:
            continue
        span = RangeRef(old_line, data.c0, old_line, data.c1) if not by_cols else RangeRef(data.r0, old_line, data.r1, old_line)
        taken[old_line] = list(sm.cells.iter_range(span))
    for old_line, cells in taken.items():
        for rc, _ in cells:
            sm.cells.remove(rc)
            ctx.touch(si, rc)
    moved_notes = {}
    moved_links = {}
    for name, store in (("comments", moved_notes), ("hyperlinks", moved_links)):
        d = getattr(sm, name)
        for rc in list(d):
            line = rc[0] if not by_cols else rc[1]
            inside = data.contains(rc)
            if inside and line in taken:
                store[rc] = d.pop(rc)
    for new_line, old_line in perm.items():
        if new_line == old_line:
            continue
        delta = new_line - old_line
        for rc, cell in taken[old_line]:
            nrc = (new_line, rc[1]) if not by_cols else (rc[0], new_line)
            if cell.formula is not None:
                e = cell.formula.get_expr()
                if e is not None:
                    ne = shift_relative(e, delta, 0) if not by_cols else shift_relative(e, 0, delta)
                    cell = Cell(cell.value, Formula.from_expr(ne), cell.style)
            sm.cells.set(nrc, cell)
            ctx.touch(si, nrc)
        for store, name in ((moved_notes, "comments"), (moved_links, "hyperlinks")):
            d = getattr(sm, name)
            for rc, x in list(store.items()):
                line = rc[0] if not by_cols else rc[1]
                if line == old_line:
                    d[(new_line, rc[1]) if not by_cols else (rc[0], new_line)] = x
    ctx.structural = True


def _sort(s: Session, p: Dict[str, Any], desc: bool):
    d = s.doc()
    si, rng, owner = _data_region(s, p)
    sh = d.wb.sheet(si)
    if sh.protection is not None and not sh.protection.sort:
        raise EngineError("The cell or chart you're trying to change is on a protected sheet.")
    if hasattr(owner, "header_row"):
        header = owner.header_row
        if owner.totals_row:
            rng = RangeRef(rng.r0, rng.c0, rng.r1 - 1, rng.c1)
    elif owner is not None:
        header = True
    else:
        header = p_bool(p, "header", guess_header(sh, rng, d.wb)) if "header" in p else guess_header(sh, rng, d.wb)
    a = d.selection.active
    col = a[1] if rng.c0 <= a[1] <= rng.c1 else rng.c0
    if p.get("col") is not None:
        col = _col_param(p["col"], rng)
    check_arrays(sh, rng)

    def f(ctx):
        sort_range(ctx, si, rng, [{"col": col, "order": "desc" if desc else "asc"}], header)

    edit(s, f)
    return {"range": rng.a1(), "header": header}


@command("data.sortAscending", "Sort A to Z", menu=("Data", "Sort & Filter"), params="{range?, col?, header?}")
def sort_asc(s, p):
    return _sort(s, p, False)


@command("data.sortDescending", "Sort Z to A", menu=("Data", "Sort & Filter"), params="{range?, col?, header?}")
def sort_desc(s, p):
    return _sort(s, p, True)


@command("data.sort", "Custom Sort…", menu=("Data", "Sort & Filter"),
         params='{range?, keys: [{col: "B" | offset, order?: asc|desc, by?: values|cellColor|fontColor, color?, customList?: [..]}], header?: bool, matchCase?: bool, byColumns?: bool}')
def sort_custom(s, p):
    d = s.doc()
    si, rng, owner = _data_region(s, p)
    sh = d.wb.sheet(si)
    if sh.protection is not None and not sh.protection.sort:
        raise EngineError("The cell or chart you're trying to change is on a protected sheet.")
    keys_in = need(p, "keys", "data.sort")
    if not isinstance(keys_in, list) or not keys_in:
        raise EngineError("data.sort: `keys` must be a non-empty list")
    by_cols = p_bool(p, "byColumns")
    if hasattr(owner, "header_row"):
        header = owner.header_row
        if owner.totals_row:
            rng = RangeRef(rng.r0, rng.c0, rng.r1 - 1, rng.c1)
    elif owner is not None:
        header = True
    else:
        header = p_bool(p, "header") if "header" in p else guess_header(sh, rng, d.wb)
    keys = []
    for k in keys_in[:64]:
        if not isinstance(k, dict) or k.get("col") is None:
            raise EngineError("data.sort: each key needs `col`")
        if by_cols:
            v = k["col"]
            if isinstance(v, str) and v.isdigit():
                r = int(v) - 1
            else:
                r = int(v) if int(v) >= rng.r0 else rng.r0 + int(v)
            keys.append({**k, "col": r})
        else:
            keys.append({**k, "col": _col_param(k["col"], rng)})
    check_arrays(sh, rng)
    edit(s, lambda ctx: sort_range(ctx, si, rng, keys, header, by_cols, p_bool(p, "matchCase")))
    return {"range": rng.a1(), "header": header}


# ---------------------------------------------------------------- filter


def _num(v) -> Optional[float]:
    if isinstance(v, bool):
        return None
    if isinstance(v, float):
        return v
    if isinstance(v, str):
        return parse_number_text(v)
    return None


def _wild_match(pattern: str, text: str) -> bool:
    from ..functions.criteria import wildcard_match

    return wildcard_match(pattern, text)


def _condition(op: str, arg: str, v, shown: str) -> bool:
    if op in ("beginsWith", "endsWith", "contains", "notContains"):
        t = shown.lower()
        a = arg.lower()
        return {
            "beginsWith": t.startswith(a),
            "endsWith": t.endswith(a),
            "contains": a in t,
            "notContains": a not in t,
        }[op]
    n_arg = _num(arg)
    n_v = v if isinstance(v, float) else None
    if n_arg is not None and n_v is not None:
        x, y = n_v, n_arg
        return {
            "equal": x == y,
            "notEqual": x != y,
            "greaterThan": x > y,
            "greaterThanOrEqual": x >= y,
            "lessThan": x < y,
            "lessThanOrEqual": x <= y,
        }.get(op, False)
    if op in ("equal", "notEqual"):
        if any(ch in arg for ch in "*?"):
            m = _wild_match(arg, shown)
        else:
            m = shown.lower() == arg.lower()
        return m if op == "equal" else not m
    if v is None or isinstance(v, float):
        return False
    c = compare(shown, arg)
    return {
        "greaterThan": c > 0,
        "greaterThanOrEqual": c >= 0,
        "lessThan": c < 0,
        "lessThanOrEqual": c <= 0,
    }.get(op, False)


def _matcher(wb, sh, crit: FilterCriterion, col: int, rows: List[int]):
    k = crit.kind
    if k == "values":
        wanted = {x.lower() for x in crit.values}

        def m(r):
            v = sh.value((r, col))
            if v is None or v == "":
                return crit.blanks
            return value_text(wb, sh, (r, col)).lower() in wanted

        return m
    if k == "custom":
        conds = crit.conditions

        def m(r):
            v = sh.value((r, col))
            shown = value_text(wb, sh, (r, col)) if v is not None else ""
            res = [_condition(op, arg, v, shown) for op, arg in conds]
            if not res:
                return True
            return all(res) if crit.and_ else any(res)

        return m
    if k in ("top10", "aboveAverage", "belowAverage"):
        nums = [sh.value((r, col)) for r in rows]
        nums = [x for x in nums if isinstance(x, float)]
        if k == "top10":
            if not nums:
                return lambda r: False
            n = crit.count
            if crit.percent:
                n = max(1, int(math.floor(len(nums) * crit.count / 100.0)))
            n = max(1, min(n, len(nums)))
            srt = sorted(nums, reverse=not crit.bottom)
            th = srt[n - 1]

            def m(r):
                v = sh.value((r, col))
                if not isinstance(v, float):
                    return False
                return v <= th if crit.bottom else v >= th

            return m
        avg = sum(nums) / len(nums) if nums else 0.0

        def m(r):
            v = sh.value((r, col))
            if not isinstance(v, float):
                return False
            return v > avg if k == "aboveAverage" else v < avg

        return m
    if k in ("fillColor", "fontColor"):
        want = crit.color.hex(wb.theme)

        def m(r):
            st = wb.styles.get(sh.style_id((r, col)))
            if k == "fillColor":
                got = st.fill.fg.hex(wb.theme) if st.fill.pattern != "none" else None
            else:
                got = st.font.color.hex(wb.theme)
            return got == want

        return m
    return lambda r: True


def apply_filter(ctx, si: int, rng: RangeRef, criteria: List[Tuple[int, FilterCriterion]], header: bool = True) -> int:
    """Hides the rows of ``rng`` (below its header) that fail any criterion; returns the
    number of visible data rows."""
    wb = ctx.wb
    sh = wb.sheet(si)
    r0 = rng.r0 + (1 if header else 0)
    used = sh.used_range()
    r1 = rng.r1 if used is None else min(rng.r1, max(used.r1, r0))
    rows = list(range(r0, r1 + 1))
    matchers = [(_matcher(wb, sh, crit, rng.c0 + off, rows), off) for off, crit in criteria]
    sm = ctx.sheet_mut(si)
    shown = 0
    for r in rows:
        ok = all(m(r) for m, _ in matchers)
        info = sm.row_info(r)
        if ok:
            shown += 1
            if info.hidden:
                sm.set_row(r, replace(info, hidden=False))
        elif not info.hidden:
            sm.set_row(r, replace(info, hidden=True))
    return shown


def _filter_owner(sh, a):
    t = sh.table_at(a)
    if t is not None:
        return t
    af = sh.autofilter
    if af is not None and af.range.contains(a):
        return af
    return None


@command("data.filter", "Filter", menu=("Data", "Sort & Filter"), shortcut="Ctrl+Shift+L", params="{range?, on?: bool}")
def toggle_filter(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)
    if sh.protection is not None and not sh.protection.autofilter:
        raise EngineError("The cell or chart you're trying to change is on a protected sheet.")
    a = d.selection.active
    t = sh.table_at(a)
    if t is not None:
        on = p_bool(p, "on", not t.filter_button)

        def ft(ctx):
            sm = ctx.sheet_mut(si)
            tt = sm.table_at(a)
            tt.filter_button = on
            if not on and tt.filter:
                tt.filter = []
                for r in range(tt.range.r0, tt.range.r1 + 1):
                    info = sm.row_info(r)
                    if info.hidden:
                        sm.set_row(r, replace(info, hidden=False))

        edit(s, ft)
        return {"on": on, "table": t.name}
    on = p_bool(p, "on", sh.autofilter is None)
    if not on:
        if sh.autofilter is None:
            return {"on": False}
        rng = sh.autofilter.range

        def off(ctx):
            sm = ctx.sheet_mut(si)
            sm.autofilter = None
            for r in range(rng.r0, rng.r1 + 1):
                info = sm.row_info(r)
                if info.hidden:
                    sm.set_row(r, replace(info, hidden=False))

        edit(s, off)
        return {"on": False}
    _, rng, _owner = _data_region(s, {k: v for k, v in p.items() if k != "on"})
    if rng.is_single() and sh.value(rng.start) is None:
        raise EngineError("This can't be applied to the selected range. Select a single cell in a range and try again.")

    def f(ctx):
        ctx.sheet_mut(si).autofilter = AutoFilter(rng)

    edit(s, f)
    return {"on": True, "range": rng.a1()}


def parse_criterion(c: Dict[str, Any]) -> FilterCriterion:
    kind = str(c.get("kind", "values"))
    if kind == "values":
        return FilterCriterion("values", values=[str(x) for x in c.get("values", [])], blanks=bool(c.get("blanks", False)))
    if kind == "custom":
        conds = []
        for x in c.get("conditions", [])[:2]:
            if isinstance(x, dict):
                conds.append((str(x.get("op", "equal")), str(x.get("value", ""))))
            else:
                conds.append((str(x[0]), str(x[1])))
        return FilterCriterion("custom", conditions=conds, and_=bool(c.get("and", False)))
    if kind == "top10":
        return FilterCriterion("top10", count=int(c.get("count", 10)), percent=bool(c.get("percent", False)), bottom=bool(c.get("bottom", False)))
    if kind in ("aboveAverage", "belowAverage"):
        return FilterCriterion(kind)
    if kind in ("fillColor", "fontColor"):
        return FilterCriterion(kind, color=parse_color(c.get("color")))
    raise EngineError(f"unknown filter kind `{kind}`")


@command("data.filterColumn", "Filter Column", menu=("Data", "Sort & Filter"), params='{col: "B" | offset, criterion: {kind: values, values: [..], blanks?} | {kind: custom, conditions: [{op, value}], and?} | {kind: top10, count, percent?, bottom?} | {kind: aboveAverage|belowAverage} | {kind: fillColor|fontColor, color} | null (clear)}')
def filter_column(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)
    a = d.selection.active
    owner = None
    if p.get("table"):
        loc = d.wb.table(str(p["table"]))
        if loc is None:
            raise EngineError(f"no table `{p['table']}`")
        owner = sh.tables[loc[1]] if loc[0] == si else None
    if owner is None:
        owner = _filter_owner(sh, a)
    if owner is None and sh.autofilter is not None:
        owner = sh.autofilter
    if owner is None:
        raise EngineError("Turn on the filter first (Data › Filter).")
    rng = owner.range
    col = _col_param(need(p, "col", "data.filterColumn"), rng)
    off = col - rng.c0
    crit = p.get("criterion")
    new = None if crit is None else parse_criterion(crit)
    is_table = hasattr(owner, "header_row")
    tname = owner.name if is_table else None

    def f(ctx):
        sm = ctx.sheet_mut(si)
        o = next(t for t in sm.tables if t.name == tname) if is_table else sm.autofilter
        crits = [(k, c) for k, c in (o.filter if is_table else o.criteria) if k != off]
        if new is not None:
            crits.append((off, new))
            crits.sort(key=lambda x: x[0])
        if is_table:
            o.filter = crits
            body = o.range if not o.totals_row else RangeRef(o.range.r0, o.range.c0, o.range.r1 - 1, o.range.c1)
            apply_filter(ctx, si, body, crits, o.header_row)
        else:
            o.criteria = crits
            apply_filter(ctx, si, o.range, crits, True)

    edit(s, f)
    return None


@command("data.clearFilter", "Clear", menu=("Data", "Sort & Filter"), params="{}")
def clear_filter(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)
    owner = _filter_owner(sh, d.selection.active) or sh.autofilter
    if owner is None:
        return None
    is_table = hasattr(owner, "header_row")
    tname = owner.name if is_table else None

    def f(ctx):
        sm = ctx.sheet_mut(si)
        o = next(t for t in sm.tables if t.name == tname) if is_table else sm.autofilter
        if is_table:
            o.filter = []
        else:
            o.criteria = []
        for r in range(o.range.r0, o.range.r1 + 1):
            info = sm.row_info(r)
            if info.hidden:
                sm.set_row(r, replace(info, hidden=False))

    edit(s, f)
    return None


@command("data.reapply", "Reapply", menu=("Data", "Sort & Filter"), shortcut="Ctrl+Alt+L")
def reapply(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)

    def f(ctx):
        sm = ctx.sheet_mut(si)
        if sm.autofilter is not None and sm.autofilter.criteria:
            apply_filter(ctx, si, sm.autofilter.range, sm.autofilter.criteria, True)
        for t in sm.tables:
            if t.filter:
                body = t.range if not t.totals_row else RangeRef(t.range.r0, t.range.c0, t.range.r1 - 1, t.range.c1)
                apply_filter(ctx, si, body, t.filter, t.header_row)

    edit(s, f)
    return None


@command("data.filterValues", "Filter Values", params='{col: "B" | offset, table?}', undoable=False, journal=False)
def filter_values(s: Session, p: Dict[str, Any]):
    """The distinct values of a filter column, for the filter dropdown."""
    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)
    owner = None
    if p.get("table"):
        loc = d.wb.table(str(p["table"]))
        if loc is not None and loc[0] == si:
            owner = sh.tables[loc[1]]
    if owner is None:
        owner = _filter_owner(sh, d.selection.active) or sh.autofilter
    if owner is None:
        raise EngineError("No filter here.")
    rng = owner.range
    col = _col_param(need(p, "col", "data.filterValues"), rng)
    off = col - rng.c0
    is_table = hasattr(owner, "header_row")
    header = owner.header_row if is_table else True
    r0 = rng.r0 + (1 if header else 0)
    r1 = rng.r1 - (1 if is_table and owner.totals_row else 0)
    used = sh.used_range()
    if used is not None:
        r1 = min(r1, max(used.r1, r0))
    crit = dict(owner.filter if is_table else owner.criteria).get(off)
    # Rows hidden by other columns' criteria don't offer their values (Excel's cascading lists).
    others = [(k, c) for k, c in (owner.filter if is_table else owner.criteria) if k != off]
    rows = list(range(r0, r1 + 1))
    matchers = [_matcher(d.wb, sh, c, rng.c0 + k, rows) for k, c in others]
    seen: Dict[str, Tuple[int, Any]] = {}
    blanks = False
    for r in rows:
        if matchers and not all(m(r) for m in matchers):
            continue
        v = sh.value((r, col))
        if v is None or v == "":
            blanks = True
            continue
        t = value_text(d.wb, sh, (r, col))
        if t.lower() not in seen:
            seen[t.lower()] = (len(seen), t, v)
        if len(seen) >= 10000:
            break

    items = sorted(seen.values(), key=lambda x: functools.cmp_to_key(lambda a, b: sort_compare(a, b))(x[2]))
    checked = None
    if crit is not None and crit.kind == "values":
        checked = {x.lower() for x in crit.values}
    out = [{"text": t, "checked": checked is None or t.lower() in checked} for _, t, _v in items]
    return {
        "values": out,
        "blanks": blanks,
        "blanksChecked": checked is None or (crit is not None and crit.blanks),
        "active": crit is not None,
        "kind": crit.kind if crit is not None else None,
        "numeric": sum(1 for _, _, v in items if isinstance(v, float)) > len(items) / 2 if items else False,
    }


# ---------------------------------------------------------------- remove duplicates


@command("data.removeDuplicates", "Remove Duplicates", menu=("Data", "Data Tools"), params="{range?, columns?: [offsets], header?: bool}")
def remove_duplicates(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si, rng, owner = _data_region(s, p)
    sh = d.wb.sheet(si)
    if hasattr(owner, "header_row"):
        header = owner.header_row
        if owner.totals_row:
            rng = RangeRef(rng.r0, rng.c0, rng.r1 - 1, rng.c1)
    else:
        header = p_bool(p, "header") if "header" in p else guess_header(sh, rng, d.wb)
    cols = p.get("columns")
    offs = [int(c) for c in cols] if cols else list(range(rng.width))
    for o in offs:
        if not 0 <= o < rng.width:
            raise EngineError("column outside the range")
    r0 = rng.r0 + (1 if header else 0)
    used = sh.used_range()
    r1 = rng.r1 if used is None else min(rng.r1, max(used.r1, r0))
    check_arrays(sh, rng)
    keep: List[int] = []
    seen = set()
    for r in range(r0, r1 + 1):
        key = []
        for o in offs:
            v = sh.value((r, rng.c0 + o))
            if isinstance(v, str):
                key.append(("s", v.lower()))
            elif isinstance(v, CellError):
                key.append(("e", v.code))
            else:
                key.append((type(v).__name__, v))
        key = tuple(key)
        if key in seen:
            continue
        seen.add(key)
        keep.append(r)
    removed = (r1 - r0 + 1) - len(keep)
    if removed == 0:
        return {"removed": 0, "remaining": len(keep)}

    def f(ctx):
        sm = ctx.sheet_mut(si)
        body = RangeRef(r0, rng.c0, r1, rng.c1)
        rows = {r: [(rc, c) for rc, c in sm.cells.iter_range(RangeRef(r, rng.c0, r, rng.c1))] for r in range(r0, r1 + 1)}
        for rc, _ in list(sm.cells.iter_range(body)):
            sm.cells.remove(rc)
            ctx.touch(si, rc)
        for i, r in enumerate(keep):
            nr = r0 + i
            for rc, cell in rows[r]:
                if cell.formula is not None and nr != r:
                    e = cell.formula.get_expr()
                    if e is not None:
                        cell = Cell(cell.value, Formula.from_expr(shift_relative(e, nr - r, 0)), cell.style)
                sm.cells.set((nr, rc[1]), cell)
                ctx.touch(si, (nr, rc[1]))
        if hasattr(owner, "header_row"):
            for t in sm.tables:
                if t.range == owner.range:
                    t.range = RangeRef(t.range.r0, t.range.c0, t.range.r1 - removed, t.range.c1)
        ctx.structural = True

    edit(s, f)
    return {"removed": removed, "remaining": len(keep)}


# ---------------------------------------------------------------- data validation


@command("data.validation", "Data Validation…", menu=("Data", "Data Tools"),
         params='{range?, kind: any|whole|decimal|list|date|time|textLength|custom, operator?: between|notBetween|equal|notEqual|greaterThan|lessThan|greaterThanOrEqual|lessThanOrEqual, f1?: "1" | "a,b,c" (list) | "=$A$1:$A$5", f2?, allowBlank?, inCellDropdown?, inputTitle?, inputMessage?, showInput?, errorTitle?, errorMessage?, errorStyle?: stop|warning|information, showError?}')
def set_validation(s: Session, p: Dict[str, Any]):
    si = param_sheet(s, p)
    ranges = _ranges(s, p)
    kind = p_str(p, "kind", "any")
    if kind not in ("any", "whole", "decimal", "list", "date", "time", "textLength", "custom"):
        raise EngineError(f"unknown validation kind `{kind}`")
    op = p_str(p, "operator", "between")
    f1 = p_str(p, "f1", "") or ""
    f2 = p_str(p, "f2")
    if kind == "list" and f1 and not f1.startswith("="):
        f1 = '"' + f1.replace('"', '""') + '"' if not (f1.startswith('"') and f1.endswith('"')) else f1
    elif f1.startswith("="):
        f1 = f1[1:]
    if f2 is not None and f2.startswith("="):
        f2 = f2[1:]
    if kind in ("whole", "decimal", "date", "time", "textLength") and not f1:
        raise EngineError("You must enter a value in the Minimum/Value box.")
    if kind == "list" and not f1:
        raise EngineError("You must enter a value in the Source box.")
    if kind == "custom" and not f1:
        raise EngineError("You must enter a formula.")
    rule = Validation(
        ranges=list(ranges),
        kind=kind,
        operator=op,
        f1=f1,
        f2=f2 if op in ("between", "notBetween") else None,
        allow_blank=p_bool(p, "allowBlank", True),
        in_cell_dropdown=p_bool(p, "inCellDropdown", True),
        input_title=p_str(p, "inputTitle", "") or "",
        input_message=p_str(p, "inputMessage", "") or "",
        show_input=p_bool(p, "showInput", True),
        error_title=p_str(p, "errorTitle", "") or "",
        error_message=p_str(p, "errorMessage", "") or "",
        error_style=p_str(p, "errorStyle", "stop") or "stop",
        show_error=p_bool(p, "showError", True),
    )

    def f(ctx):
        sm = ctx.sheet_mut(si)
        rules = sm.validations
        for r in ranges:
            rules = remove_rules_in(rules, r)
        if kind != "any" or rule.input_message or rule.input_title:
            rules = rules + [rule]
        sm.validations = rules

    edit(s, f)
    return None


@command("data.clearValidation", "Clear Validation", menu=("Data", "Data Tools"), params="{range?}")
def clear_validation(s: Session, p: Dict[str, Any]):
    si = param_sheet(s, p)
    ranges = _ranges(s, p)

    def f(ctx):
        sm = ctx.sheet_mut(si)
        rules = sm.validations
        for r in ranges:
            rules = remove_rules_in(rules, r)
        sm.validations = rules

    edit(s, f)
    return None


@command("data.validationAt", "Validation Rule", params="{cell?}", undoable=False, journal=False)
def validation_query(s: Session, p: Dict[str, Any]):
    from .session import param_cell

    d = s.doc()
    si = param_sheet(s, p)
    at = param_cell(s, p)
    sh = d.wb.sheet(si)
    v = validation_at(sh, at)
    if v is None:
        return None
    out = {
        "kind": v.kind,
        "operator": v.operator,
        "f1": v.f1,
        "f2": v.f2,
        "allowBlank": v.allow_blank,
        "inCellDropdown": v.in_cell_dropdown,
        "inputTitle": v.input_title,
        "inputMessage": v.input_message,
        "showInput": v.show_input,
        "errorTitle": v.error_title,
        "errorMessage": v.error_message,
        "errorStyle": v.error_style,
        "showError": v.show_error,
        "ranges": [r.a1() for r in v.ranges],
    }
    if v.kind == "list":
        out["items"] = list_items(d.wb, si, v, at)
    return out


@command("data.circleInvalid", "Circle Invalid Data", menu=("Data", "Data Tools"), undoable=False, journal=False)
def circle_invalid(s: Session, p: Dict[str, Any]):
    from ..core.addr import a1
    from .entry import is_valid

    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)
    out = []
    for v in sh.validations:
        for r in v.ranges:
            used = sh.used_range()
            inter = r.intersection(used) if used is not None else None
            if inter is None:
                continue
            for rc in inter.cells():
                val = sh.value(rc)
                if not is_valid(d.wb, si, rc, val, v):
                    out.append(a1(*rc))
                    if len(out) >= 1000:
                        return {"invalid": out}
    return {"invalid": out}


# ---------------------------------------------------------------- text to columns


@command("data.textToColumns", "Text to Columns…", menu=("Data", "Data Tools"),
         params='{range?, delimiters?: [",", "\\t", ";", " "], other?: "|", consecutive?: bool, qualifier?: "\\"", destination?: "B1"}')
def text_to_columns(s: Session, p: Dict[str, Any]):
    from ..core.addr import parse_cell
    from ..xlsx.csv_io import parse_records

    d = s.doc()
    si = param_sheet(s, p)
    rng = _ranges(s, p)[0]
    if rng.width != 1:
        raise EngineError("Text to Columns can convert only one column at a time.")
    sh = d.wb.sheet(si)
    used = sh.used_range()
    if used is None:
        return None
    rng = RangeRef(rng.r0, rng.c0, min(rng.r1, used.r1), rng.c1)
    delims = list(p.get("delimiters") or ["\t"])
    if p.get("other"):
        delims.append(str(p["other"])[:1])
    consecutive = p_bool(p, "consecutive")
    q = p_str(p, "qualifier", '"') or ""
    dest = parse_cell(str(p["destination"])) if p.get("destination") else rng.start
    if dest is None:
        raise EngineError("Destination reference is not valid.")
    rows = []
    for r in range(rng.r0, rng.r1 + 1):
        v = sh.value((r, rng.c0))
        t = value_text(d.wb, sh, (r, rng.c0)) if v is not None else ""
        if not isinstance(v, str):
            rows.append([t] if t else [])
            continue
        main = delims[0]
        line = t
        for x in delims[1:]:
            line = line.replace(x, main)
        parts = next(iter(parse_records(line, main, q or '"')), [line]) if q else line.split(main)
        if consecutive:
            parts = [x for x in parts if x != ""]
        rows.append(parts)
    width = max((len(x) for x in rows), default=1)
    if dest[1] + width > MAX_COLS:
        raise EngineError("The data doesn't fit on the sheet.")

    def f(ctx):
        wb = ctx.wb
        for i, parts in enumerate(rows):
            r = dest[0] + i
            for j, t in enumerate(parts):
                rc = (r, dest[1] + j)
                _set(ctx, si, rc, input_to_cell(wb, wb.sheet(si), rc, t))

    edit(s, f)
    return {"columns": width}
