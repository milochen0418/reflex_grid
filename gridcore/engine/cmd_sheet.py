"""Rows, columns, cells and sheets: insert, delete, rename, move, copy, hide, tab colour,
activation and protection."""

from __future__ import annotations

import copy
from dataclasses import replace
from typing import Any, Dict, List, Optional

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef
from ..formula import ParseError, parse, print_formula
from ..formula.adjust import delete_sheet as adjust_delete_sheet
from ..formula.adjust import rename_sheet as adjust_rename_sheet
from ..model.cell import Cell, Formula
from ..model.features import SheetProtection, legacy_password_hash
from ..model.sheet import Sheet
from .cmd_edit import _activate, check_arrays
from .cmd_format import _ranges, parse_color
from .selection import Selection
from .session import EngineError, Session, command, edit, need, p_bool, p_int, p_str, param_sheet
from .structure import shift_cells, shift_lines


def _sel_lines(s: Session, p: Dict[str, Any], axis: str) -> List[tuple]:
    """Blocks of rows/columns (start, count) from a parameter or the selection, bottom-up."""
    key = "rows" if axis == "rows" else "cols"
    spans = []
    t = p.get(key)
    if t is not None:
        for part in str(t).split(","):
            r = RangeRef.parse(part if ":" in part else f"{part}:{part}")
            if r is None:
                raise EngineError(f"not a reference: `{part}`")
            spans.append((r.r0, r.height) if axis == "rows" else (r.c0, r.width))
    else:
        for r in _ranges(s, p):
            spans.append((r.r0, r.height) if axis == "rows" else (r.c0, r.width))
    n = p_int(p, "count")
    if n is not None and spans:
        spans = [(spans[0][0], max(n, 1))]
    # Merge overlapping spans; apply from the bottom so earlier edits don't move later ones.
    spans.sort()
    merged: List[List[int]] = []
    for a, c in spans:
        if merged and a <= merged[-1][0] + merged[-1][1]:
            end = max(merged[-1][0] + merged[-1][1], a + c)
            merged[-1][1] = end - merged[-1][0]
        else:
            merged.append([a, c])
    return [tuple(x) for x in reversed(merged)]


def _check_protect(sh, what: str) -> None:
    pr = sh.protection
    if pr is None:
        return
    if not getattr(pr, what):
        raise EngineError("The cell or chart you're trying to change is on a protected sheet.")


def _insert_lines(s: Session, p: Dict[str, Any], axis: str):
    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)
    _check_protect(sh, "insert_rows" if axis == "rows" else "insert_columns")
    spans = _sel_lines(s, p, axis)
    mx = MAX_ROWS if axis == "rows" else MAX_COLS
    total = sum(c for _, c in spans)
    used = sh.used_range()
    if used is not None and (used.r1 if axis == "rows" else used.c1) + total >= mx:
        raise EngineError("To prevent possible loss of data, ReflexGrid can't shift nonblank cells off of the worksheet.")

    def f(ctx):
        for at, count in spans:
            shift_lines(ctx, si, axis, at, count, True)

    edit(s, f)
    return None


def _delete_lines(s: Session, p: Dict[str, Any], axis: str):
    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)
    _check_protect(sh, "delete_rows" if axis == "rows" else "delete_columns")
    spans = _sel_lines(s, p, axis)
    for at, count in spans:
        rng = RangeRef.rows(at, at + count - 1) if axis == "rows" else RangeRef.cols(at, at + count - 1)
        check_arrays(sh, rng)

    def f(ctx):
        for at, count in spans:
            shift_lines(ctx, si, axis, at, count, False)
        sm = ctx.wb.sheet(si)
        a = ctx.sel.active
        ctx.sel = Selection.at((min(a[0], MAX_ROWS - 1), min(a[1], MAX_COLS - 1)))
        if spans:
            at = spans[-1][0]
            ctx.sel = Selection.of_range(RangeRef.rows(at, at) if axis == "rows" else RangeRef.cols(at, at), (at, a[1]) if axis == "rows" else (a[0], at))

    edit(s, f)
    return None


@command("home.insertRows", "Insert Sheet Rows", menu=("Home", "Cells", "Insert"), shortcut="Ctrl+Shift+=", params='{rows?: "3:5", count?}')
def insert_rows(s, p):
    return _insert_lines(s, p, "rows")


@command("home.insertColumns", "Insert Sheet Columns", menu=("Home", "Cells", "Insert"), params='{cols?: "B:C", count?}')
def insert_cols(s, p):
    return _insert_lines(s, p, "cols")


@command("home.deleteRows", "Delete Sheet Rows", menu=("Home", "Cells", "Delete"), shortcut="Ctrl+-", params='{rows?: "3:5"}')
def delete_rows(s, p):
    return _delete_lines(s, p, "rows")


@command("home.deleteColumns", "Delete Sheet Columns", menu=("Home", "Cells", "Delete"), params='{cols?: "B:C"}')
def delete_cols(s, p):
    return _delete_lines(s, p, "cols")


@command("home.insertCells", "Insert Cells…", menu=("Home", "Cells", "Insert"), params="{shift: down|right|row|col, range?}")
def insert_cells(s, p):
    shift = p_str(p, "shift", "down")
    if shift == "row":
        return _insert_lines(s, p, "rows")
    if shift == "col":
        return _insert_lines(s, p, "cols")
    if shift not in ("down", "right"):
        raise EngineError(f"unknown shift `{shift}`")
    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)
    _check_protect(sh, "insert_rows" if shift == "down" else "insert_columns")
    ranges = _ranges(s, p)
    axis = "rows" if shift == "down" else "cols"
    for rng in ranges:
        for m in sh.merges:
            if m.intersects(rng) and not rng.contains_range(m):
                raise EngineError("This operation will cause some merged cells to unmerge.")
        for t in sh.tables:
            if t.range.intersects(rng) and axis == "cols":
                raise EngineError("This operation isn't allowed. The operation is attempting to shift cells in a table on your worksheet.")

    def f(ctx):
        for rng in sorted(ranges, reverse=True):
            shift_cells(ctx, si, rng, axis, True)

    edit(s, f)
    return None


@command("home.deleteCells", "Delete Cells…", menu=("Home", "Cells", "Delete"), params="{shift: up|left|row|col, range?}")
def delete_cells(s, p):
    shift = p_str(p, "shift", "up")
    if shift == "row":
        return _delete_lines(s, p, "rows")
    if shift == "col":
        return _delete_lines(s, p, "cols")
    if shift not in ("up", "left"):
        raise EngineError(f"unknown shift `{shift}`")
    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)
    _check_protect(sh, "delete_rows" if shift == "up" else "delete_columns")
    ranges = _ranges(s, p)
    axis = "rows" if shift == "up" else "cols"
    for rng in ranges:
        check_arrays(sh, rng)
        for m in sh.merges:
            if m.intersects(rng) and not rng.contains_range(m):
                raise EngineError("This operation will cause some merged cells to unmerge.")

    def f(ctx):
        for rng in sorted(ranges, reverse=True):
            shift_cells(ctx, si, rng, axis, False)

    edit(s, f)
    return None


# ---------------------------------------------------------------- sheets


def _check_structure(wb) -> None:
    if wb.protected_structure:
        raise EngineError("The workbook is protected: unprotect it (Review › Protect Workbook) to change sheets.")


def _remap_scopes(wb, mapping: Dict[int, Optional[int]]) -> None:
    names = []
    for dn in wb.names:
        if dn.scope is None:
            names.append(dn)
            continue
        n = mapping.get(dn.scope, dn.scope)
        if n is None:
            continue
        dn.scope = n
        names.append(dn)
    wb.names = names


@command("home.insertSheet", "Insert Sheet", menu=("Home", "Cells", "Insert"), shortcut="Shift+F11", params="{name?, index?}")
def insert_sheet(s, p):
    d = s.doc()
    _check_structure(d.wb)
    name = p_str(p, "name") or d.wb.next_sheet_name()
    err = d.wb.check_sheet_name(name)
    if err:
        raise EngineError(err)
    index = p_int(p, "index")
    if index is None:
        index = d.wb.active_sheet
    index = max(0, min(index, len(d.wb.sheets)))

    def f(ctx):
        wb = ctx.wb
        mapping = {i: (i + 1 if i >= index else i) for i in range(len(wb.sheets))}
        wb.add_sheet(Sheet(name), index)
        _remap_scopes(wb, mapping)
        wb.active_sheet = index
        ctx.sel = Selection.at((0, 0))
        ctx.structural = True

    edit(s, f)
    d.sheet_selections = {}
    return {"sheet": name, "index": index}


@command("home.deleteSheet", "Delete Sheet", menu=("Home", "Cells", "Delete"), params="{sheet?}")
def delete_sheet(s, p):
    d = s.doc()
    _check_structure(d.wb)
    si = param_sheet(s, p)
    visible = [i for i, sh in enumerate(d.wb.sheets) if sh.visibility == "visible"]
    if len(d.wb.sheets) <= 1 or (visible == [si]):
        raise EngineError("A workbook must contain at least one visible worksheet.")
    name = d.wb.sheets[si].name
    order = [sh.name for sh in d.wb.sheets]

    def f(ctx):
        wb = ctx.wb
        del wb.sheets[si]
        mapping = {i: (None if i == si else (i - 1 if i > si else i)) for i in range(len(order))}
        _remap_scopes(wb, mapping)
        for i, sh in enumerate(wb.sheets):
            changes = []
            for rc, cell in sh.cells.items():
                if cell.formula is None:
                    continue
                e = cell.formula.get_expr()
                if e is None:
                    continue
                ne = adjust_delete_sheet(e, name, order)
                if ne != e:
                    changes.append((rc, Cell(cell.value, Formula.from_expr(ne, cell.formula.array), cell.style)))
            if changes:
                sm = wb.sheet_mut(i)
                for rc, c in changes:
                    sm.cells.set(rc, c)
            if any(ser for ch in sh.charts for ser in ch.series):
                sm = wb.sheet_mut(i)
                for ch in sm.charts:
                    for ser in ch.series:
                        ser.values = _adjust_text(ser.values, lambda e: adjust_delete_sheet(e, name, order))
                        ser.categories = _adjust_text(ser.categories, lambda e: adjust_delete_sheet(e, name, order))
        for dn in wb.names:
            dn.formula = _adjust_text(dn.formula, lambda e: adjust_delete_sheet(e, name, order)) or dn.formula
        n = len(wb.sheets)
        a = min(si, n - 1)
        while wb.sheets[a].visibility != "visible":
            a = (a + 1) % n
        wb.active_sheet = a
        ctx.sel = Selection.at(wb.sheets[a].view_active)
        ctx.structural = True

    edit(s, f)
    d.sheet_selections = {}
    return {"deleted": name}


def _adjust_text(text: Optional[str], fn) -> Optional[str]:
    if not text:
        return text
    eq = text.startswith("=")
    body = text[1:] if eq else text
    try:
        e = parse(body)
    except ParseError:
        return text
    ne = fn(e)
    if ne == e:
        return text
    out = print_formula(ne)
    return ("=" + out) if eq else out


@command("sheet.rename", "Rename Sheet", menu=("Home", "Cells", "Format"), params="{sheet?, name}")
def rename_sheet(s, p):
    d = s.doc()
    _check_structure(d.wb)
    si = param_sheet(s, {k: v for k, v in p.items() if k != "name"})
    new = str(need(p, "name", "sheet.rename")).strip()
    old = d.wb.sheets[si].name
    if new == old:
        return None
    err = d.wb.check_sheet_name(new, si)
    if err:
        raise EngineError(err)

    def f(ctx):
        wb = ctx.wb
        sm = wb.sheet_mut(si)
        sm.name = new
        for i, sh in enumerate(wb.sheets):
            changes = []
            for rc, cell in sh.cells.items():
                if cell.formula is None:
                    continue
                e = cell.formula.get_expr()
                if e is None:
                    continue
                ne = adjust_rename_sheet(e, old, new)
                if ne != e:
                    changes.append((rc, Cell(cell.value, Formula.from_expr(ne, cell.formula.array), cell.style)))
            if changes:
                m = wb.sheet_mut(i)
                for rc, c in changes:
                    m.cells.set(rc, c)
            if sh.charts or sh.cond_formats or sh.validations:
                m = wb.sheet_mut(i)
                fn = lambda e: adjust_rename_sheet(e, old, new)
                for ch in m.charts:
                    for ser in ch.series:
                        ser.values = _adjust_text(ser.values, fn) or ser.values
                        ser.categories = _adjust_text(ser.categories, fn)
                        if ser.name and not ser.name.startswith('"'):
                            ser.name = _adjust_text(ser.name, fn)
                for cf in m.cond_formats:
                    cf.rule.formulas = [_adjust_text(x, fn) or x for x in cf.rule.formulas]
                for v in m.validations:
                    if v.f1 and not v.f1.startswith('"'):
                        v.f1 = _adjust_text(v.f1, fn) or v.f1
                    if v.f2:
                        v.f2 = _adjust_text(v.f2, fn)
        for dn in wb.names:
            dn.formula = _adjust_text(dn.formula, lambda e: adjust_rename_sheet(e, old, new)) or dn.formula
        ctx.structural = True

    edit(s, f)
    return {"sheet": new}


@command("sheet.move", "Move or Copy Sheet…", menu=("Home", "Cells", "Format"), params="{sheet?, to: index, copy?: bool, name?}")
def move_sheet(s, p):
    d = s.doc()
    _check_structure(d.wb)
    si = param_sheet(s, {k: v for k, v in p.items() if k not in ("name",)})
    to = p_int(p, "to")
    if to is None:
        raise EngineError("sheet.move: missing `to`")
    n = len(d.wb.sheets)
    if p_bool(p, "copy"):
        return _copy_sheet(s, si, max(0, min(to, n)), p_str(p, "name"))
    to = max(0, min(to, n - 1))
    if to == si:
        return None

    def f(ctx):
        wb = ctx.wb
        sh = wb.sheets.pop(si)
        wb.sheets.insert(to, sh)
        order = list(range(n))
        moved = order.pop(si)
        order.insert(to, moved)
        mapping = {old: new for new, old in enumerate(order)}
        _remap_scopes(wb, mapping)
        wb.active_sheet = to
        ctx.structural = True

    edit(s, f)
    d.sheet_selections = {}
    return {"index": to}


def _unique_sheet_name(wb, base: str) -> str:
    stem = base
    k = 2
    if base.endswith(")") and " (" in base:
        head, num = base.rsplit(" (", 1)
        if num[:-1].isdigit():
            stem = head
    while True:
        cand = f"{stem} ({k})"
        if len(cand) > 31:
            cand = f"{stem[: 31 - len(str(k)) - 3]} ({k})"
        if wb.sheet_index(cand) is None:
            return cand
        k += 1


def _copy_sheet(s: Session, si: int, to: int, name: Optional[str]):
    d = s.doc()
    src = d.wb.sheets[si]
    new_name = name or _unique_sheet_name(d.wb, src.name)
    err = d.wb.check_sheet_name(new_name)
    if err:
        raise EngineError(err)

    def f(ctx):
        wb = ctx.wb
        sh = src.clone()
        sh.name = new_name
        next_id = wb.next_object_id()
        used = {t.name.lower() for x in wb.sheets for t in x.tables}
        for t in sh.tables:
            t.id = next_id
            next_id += 1
            k = 2
            base = t.name
            while f"{base}{k}".lower() in used:
                k += 1
            t.name = f"{base}{k}"
            used.add(t.name.lower())
        for ch in sh.charts:
            ch.id = next_id
            next_id += 1
            fn = lambda e: adjust_rename_sheet(e, src.name, new_name)
            for ser in ch.series:
                ser.values = _adjust_text(ser.values, fn) or ser.values
                ser.categories = _adjust_text(ser.categories, fn)
                if ser.name and not ser.name.startswith('"'):
                    ser.name = _adjust_text(ser.name, fn)
        for im in sh.images:
            im.id = next_id
            next_id += 1
        mapping = {i: (i + 1 if i >= to else i) for i in range(len(wb.sheets))}
        scoped = [copy.deepcopy(dn) for dn in wb.names if dn.scope == si]
        wb.add_sheet(sh, to)
        _remap_scopes(wb, mapping)
        for dn in scoped:
            dn.scope = to
            wb.names.append(dn)
        wb.active_sheet = to
        ctx.sel = Selection.at(sh.view_active)
        ctx.structural = True

    edit(s, f)
    d.sheet_selections = {}
    return {"sheet": new_name, "index": to}


@command("sheet.duplicate", "Duplicate Sheet", params="{sheet?}")
def duplicate_sheet(s, p):
    d = s.doc()
    _check_structure(d.wb)
    si = param_sheet(s, p)
    return _copy_sheet(s, si, si + 1, p_str(p, "name"))


@command("sheet.tabColor", "Tab Color", menu=("Home", "Cells", "Format"), params='{sheet?, color: "#FF0000" | null}')
def tab_color(s, p):
    si = param_sheet(s, p)
    c = parse_color(p.get("color"))

    def f(ctx):
        ctx.sheet_mut(si).tab_color = None if c.is_auto else c

    edit(s, f)
    return None


@command("sheet.hide", "Hide Sheet", menu=("Home", "Cells", "Format", "Hide & Unhide"), params="{sheet?}")
def hide_sheet(s, p):
    d = s.doc()
    _check_structure(d.wb)
    si = param_sheet(s, p)
    visible = [i for i, sh in enumerate(d.wb.sheets) if sh.visibility == "visible"]
    if visible == [si]:
        raise EngineError("A workbook must contain at least one visible worksheet.")

    def f(ctx):
        wb = ctx.wb
        wb.sheet_mut(si).visibility = "hidden"
        if wb.active_sheet == si:
            n = len(wb.sheets)
            a = si
            for k in range(1, n + 1):
                cand = (si + k) % n
                if wb.sheets[cand].visibility == "visible":
                    a = cand
                    break
            wb.active_sheet = a
            ctx.sel = Selection.at(wb.sheets[a].view_active)

    edit(s, f)
    return None


@command("sheet.unhide", "Unhide Sheet…", menu=("Home", "Cells", "Format", "Hide & Unhide"), params="{sheet}")
def unhide_sheet(s, p):
    d = s.doc()
    _check_structure(d.wb)
    si = param_sheet(s, p)

    def f(ctx):
        ctx.wb.sheet_mut(si).visibility = "visible"
        ctx.wb.active_sheet = si
        ctx.sel = Selection.at(ctx.wb.sheets[si].view_active)

    edit(s, f)
    return None


@command("sheet.activate", "Activate Sheet", params="{sheet: name | index}", undoable=False, journal=False)
def activate_sheet(s, p):
    d = s.doc()
    si = param_sheet(s, p)
    if d.wb.sheets[si].visibility != "visible":
        raise EngineError("That sheet is hidden.")
    _activate(s, si)
    return {"sheet": d.wb.sheets[si].name}


def _step_sheet(s: Session, step: int):
    d = s.doc()
    n = len(d.wb.sheets)
    i = d.wb.active_sheet
    for _ in range(n):
        i = (i + step) % n
        if d.wb.sheets[i].visibility == "visible":
            break
    _activate(s, i)
    return {"sheet": d.wb.sheets[i].name}


@command("sheet.next", "Next Sheet", shortcut="Ctrl+PageDown", undoable=False, journal=False)
def next_sheet(s, p):
    return _step_sheet(s, 1)


@command("sheet.previous", "Previous Sheet", shortcut="Ctrl+PageUp", undoable=False, journal=False)
def prev_sheet(s, p):
    return _step_sheet(s, -1)


# ---------------------------------------------------------------- protection


@command("review.protectSheet", "Protect Sheet…", menu=("Review", "Protect"),
         params="{password?, selectLocked?, selectUnlocked?, formatCells?, formatColumns?, formatRows?, insertColumns?, insertRows?, deleteColumns?, deleteRows?, sort?, autoFilter?}")
def protect_sheet(s, p):
    si = param_sheet(s, p)
    pw = p_str(p, "password") or None
    prot = SheetProtection(
        password_hash=legacy_password_hash(pw) if pw else None,
        select_locked=p_bool(p, "selectLocked", True),
        select_unlocked=p_bool(p, "selectUnlocked", True),
        format_cells=p_bool(p, "formatCells"),
        format_columns=p_bool(p, "formatColumns"),
        format_rows=p_bool(p, "formatRows"),
        insert_columns=p_bool(p, "insertColumns"),
        insert_rows=p_bool(p, "insertRows"),
        delete_columns=p_bool(p, "deleteColumns"),
        delete_rows=p_bool(p, "deleteRows"),
        sort=p_bool(p, "sort"),
        autofilter=p_bool(p, "autoFilter"),
    )

    def f(ctx):
        ctx.sheet_mut(si).protection = prot

    edit(s, f)
    return None


@command("review.unprotectSheet", "Unprotect Sheet", menu=("Review", "Protect"), params="{password?}")
def unprotect_sheet(s, p):
    d = s.doc()
    si = param_sheet(s, p)
    pr = d.wb.sheet(si).protection
    if pr is None:
        return None
    if pr.password_hash or pr.hash_attrs:
        pw = p_str(p, "password") or ""
        if pr.hash_attrs and not pr.password_hash:
            if not _check_modern_hash(pr.hash_attrs, pw):
                raise EngineError("The password you supplied is not correct.")
        elif legacy_password_hash(pw) != (pr.password_hash or "").upper():
            raise EngineError("The password you supplied is not correct.")

    def f(ctx):
        ctx.sheet_mut(si).protection = None

    edit(s, f)
    return None


def _check_modern_hash(attrs: Dict[str, str], pw: str) -> bool:
    """ECMA-376 agile password hash (``algorithmName``, ``hashValue``, ``saltValue``, ``spinCount``)."""
    import base64
    import hashlib

    algo = attrs.get("algorithmName", "SHA-512").replace("-", "").lower()
    try:
        salt = base64.b64decode(attrs.get("saltValue", ""))
        want = base64.b64decode(attrs.get("hashValue", ""))
        spin = int(attrs.get("spinCount", "100000"))
    except Exception:
        return False
    if algo not in hashlib.algorithms_available:
        return False
    h = hashlib.new(algo, salt + pw.encode("utf-16-le")).digest()
    for i in range(min(spin, 10_000_000)):
        h = hashlib.new(algo, h + i.to_bytes(4, "little")).digest()
    return h == want


@command("review.protectWorkbook", "Protect Workbook…", menu=("Review", "Protect"), params="{password?}")
def protect_workbook(s, p):
    pw = p_str(p, "password") or None

    def f(ctx):
        ctx.wb.protected_structure = True
        ctx.wb.structure_password = legacy_password_hash(pw) if pw else None

    edit(s, f)
    return None


@command("review.unprotectWorkbook", "Unprotect Workbook", menu=("Review", "Protect"), params="{password?}")
def unprotect_workbook(s, p):
    d = s.doc()
    if not d.wb.protected_structure:
        return None
    if d.wb.structure_password and legacy_password_hash(p_str(p, "password") or "") != d.wb.structure_password.upper():
        raise EngineError("The password you supplied is not correct.")

    def f(ctx):
        ctx.wb.protected_structure = False
        ctx.wb.structure_password = None

    edit(s, f)
    return None
