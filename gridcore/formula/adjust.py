"""Reference adjustment: copying formulas, inserting/deleting rows and columns, moving cells,
renaming and deleting sheets."""

from __future__ import annotations

from dataclasses import dataclass, replace
from typing import List, Optional, Sequence

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef
from ..core.value import CellError
from .ast import Err, Expr, Name, Ref


def _shift_line(v: int, absolute: bool, d: int, mx: int) -> Optional[int]:
    if absolute:
        return v
    n = v + d
    return n if 0 <= n < mx else None


def shift_ref(r: Ref, dr: int, dc: int) -> Optional[Ref]:
    """Copy/fill: relative parts move by ``(dr, dc)``; ``None`` when pushed off the sheet."""
    k = r.kind
    if k in ("cell", "range"):
        r0 = _shift_line(r.r0, r.r0_abs, dr, MAX_ROWS)
        c0 = _shift_line(r.c0, r.c0_abs, dc, MAX_COLS)
        if r0 is None or c0 is None:
            return None
        if k == "cell":
            return replace(r, r0=r0, c0=c0, r1=r0, c1=c0)
        r1 = _shift_line(r.r1, r.r1_abs, dr, MAX_ROWS)
        c1 = _shift_line(r.c1, r.c1_abs, dc, MAX_COLS)
        if r1 is None or c1 is None:
            return None
        return replace(r, r0=r0, c0=c0, r1=r1, c1=c1)
    if k == "rows":
        r0 = _shift_line(r.r0, r.r0_abs, dr, MAX_ROWS)
        r1 = _shift_line(r.r1, r.r1_abs, dr, MAX_ROWS)
        if r0 is None or r1 is None:
            return None
        return replace(r, r0=r0, r1=r1)
    c0 = _shift_line(r.c0, r.c0_abs, dc, MAX_COLS)
    c1 = _shift_line(r.c1, r.c1_abs, dc, MAX_COLS)
    if c0 is None or c1 is None:
        return None
    return replace(r, c0=c0, c1=c1)


def shift_relative(e: Expr, dr: int, dc: int) -> Expr:
    """Copy/fill: relative parts move by ``(dr, dc)``. References pushed off the sheet become ``#REF!``."""
    if dr == 0 and dc == 0:
        return e

    def f(x: Expr) -> Expr:
        if isinstance(x, Ref):
            s = shift_ref(x, dr, dc)
            return s if s is not None else Err(CellError.REF)
        return x

    return e.map(f)


def shift_relative_wrapping(e: Expr, dr: int, dc: int) -> Expr:
    """Defined names: relative parts move by ``(dr, dc)``, wrapping around the sheet's edges."""

    def w(v: int, a: bool, d: int, mx: int) -> int:
        return v if a else (v + d) % mx

    def f(x: Expr) -> Expr:
        if not isinstance(x, Ref):
            return x
        return replace(
            x,
            r0=w(x.r0, x.r0_abs, dr, MAX_ROWS),
            r1=w(x.r1, x.r1_abs, dr, MAX_ROWS),
            c0=w(x.c0, x.c0_abs, dc, MAX_COLS),
            c1=w(x.c1, x.c1_abs, dc, MAX_COLS),
        )

    return e.map(f)


def has_relative_refs(e: Expr) -> bool:
    for r in e.references():
        if r.kind == "cell" and (not r.r0_abs or not r.c0_abs):
            return True
        if r.kind == "range" and (not r.r0_abs or not r.c0_abs or not r.r1_abs or not r.c1_abs):
            return True
        if r.kind == "rows" and (not r.r0_abs or not r.r1_abs):
            return True
        if r.kind == "cols" and (not r.c0_abs or not r.c1_abs):
            return True
    return False


@dataclass(frozen=True)
class Edit:
    """A structural edit on one sheet.

    * ``insert``/``delete``: ``axis`` ("rows"/"cols"), ``at``, ``count``
    * ``insert_cells``/``delete_cells``: ``axis`` and ``range`` (shift down/right, up/left)
    * ``move``: ``range`` (from) to ``to_row``/``to_col``
    * ``move_to_sheet``: like move, plus ``to_sheet``
    """

    kind: str
    axis: str = "rows"
    at: int = 0
    count: int = 0
    range: Optional[RangeRef] = None
    to_row: int = 0
    to_col: int = 0
    to_sheet: str = ""


def adjust(e: Expr, host: str, target: str, edit: Edit) -> Expr:
    """Adjusts a formula living on sheet ``host`` after ``edit`` on sheet ``target``."""
    host_l, target_l = host.lower(), target.lower()

    def f(x: Expr) -> Expr:
        if not isinstance(x, Ref):
            return x
        if x.sheet is None:
            on_target = host_l == target_l
        elif isinstance(x.sheet, str):
            on_target = x.sheet.lower() == target_l
        else:
            on_target = False
        if not on_target:
            return x
        r = adjust_ref(x, edit)
        return r if r is not None else Err(CellError.REF)

    return e.map(f)


def _map_index(v: int, at: int, count: int, insert: bool, mx: int) -> Optional[int]:
    if insert:
        if v >= at:
            n = v + count
            return None if n >= mx else n
        return v
    if v < at:
        return v
    if v >= at + count:
        return v - count
    return None


def _map_span(lo: int, hi: int, at: int, count: int, insert: bool, mx: int) -> Optional[tuple[int, int]]:
    if insert:
        lo2 = lo + count if lo >= at else lo
        hi2 = hi + count if hi >= at else hi
        hi2 = min(hi2, mx - 1)
        if lo2 > hi2:
            return None
        return lo2, hi2
    end = at + count
    if lo >= at and hi < end:
        return None
    lo2 = lo if lo < at else (lo - count if lo >= end else at)
    hi2 = hi if hi < at else (hi - count if hi >= end else max(at - 1, 0))
    return lo2, max(hi2, lo2)


def adjust_ref(r: Ref, edit: Edit) -> Optional[Ref]:
    k = edit.kind
    if k in ("insert", "delete"):
        insert = k == "insert"
        axis, at, count = edit.axis, edit.at, edit.count
        if r.kind == "cell":
            if axis == "rows":
                v = _map_index(r.r0, at, count, insert, MAX_ROWS)
                return None if v is None else replace(r, r0=v, r1=v)
            v = _map_index(r.c0, at, count, insert, MAX_COLS)
            return None if v is None else replace(r, c0=v, c1=v)
        if r.kind == "range":
            if axis == "rows":
                lo, hi = min(r.r0, r.r1), max(r.r0, r.r1)
                sp = _map_span(lo, hi, at, count, insert, MAX_ROWS)
                if sp is None:
                    return None
                # Keep each end's anchor flags with the end it belongs to.
                if r.r0 <= r.r1:
                    return replace(r, r0=sp[0], r1=sp[1])
                return replace(r, r0=sp[1], r1=sp[0])
            lo, hi = min(r.c0, r.c1), max(r.c0, r.c1)
            sp = _map_span(lo, hi, at, count, insert, MAX_COLS)
            if sp is None:
                return None
            if r.c0 <= r.c1:
                return replace(r, c0=sp[0], c1=sp[1])
            return replace(r, c0=sp[1], c1=sp[0])
        if r.kind == "rows" and axis == "rows":
            sp = _map_span(min(r.r0, r.r1), max(r.r0, r.r1), at, count, insert, MAX_ROWS)
            return None if sp is None else replace(r, r0=sp[0], r1=sp[1])
        if r.kind == "cols" and axis == "cols":
            sp = _map_span(min(r.c0, r.c1), max(r.c0, r.c1), at, count, insert, MAX_COLS)
            return None if sp is None else replace(r, c0=sp[0], c1=sp[1])
        return r
    if k in ("insert_cells", "delete_cells"):
        # Like inserting/deleting lines, for references inside the shifted columns (rows).
        span = r.range()
        rng = edit.range
        assert rng is not None
        if edit.axis == "rows":
            inside = rng.c0 <= span.c0 and span.c1 <= rng.c1
            at, count = rng.r0, rng.height
        else:
            inside = rng.r0 <= span.r0 and span.r1 <= rng.r1
            at, count = rng.c0, rng.width
        if not inside:
            return r
        lines = Edit("insert" if k == "insert_cells" else "delete", axis=edit.axis, at=at, count=count)
        return adjust_ref(r, lines)
    if k == "move":
        frm = edit.range
        assert frm is not None
        dr = edit.to_row - frm.r0
        dc = edit.to_col - frm.c0
        if not frm.contains_range(r.range()):
            return r
        if r.kind in ("cell", "range"):
            r0, c0, r1, c1 = r.r0 + dr, r.c0 + dc, r.r1 + dr, r.c1 + dc
            if not (0 <= r0 < MAX_ROWS and 0 <= r1 < MAX_ROWS and 0 <= c0 < MAX_COLS and 0 <= c1 < MAX_COLS):
                return None
            return replace(r, r0=r0, c0=c0, r1=r1, c1=c1)
        if r.kind == "rows":
            a, b = _shift_line(r.r0, False, dr, MAX_ROWS), _shift_line(r.r1, False, dr, MAX_ROWS)
            return None if a is None or b is None else replace(r, r0=a, r1=b)
        a, b = _shift_line(r.c0, False, dc, MAX_COLS), _shift_line(r.c1, False, dc, MAX_COLS)
        return None if a is None or b is None else replace(r, c0=a, c1=b)
    if k == "move_to_sheet":
        frm = edit.range
        assert frm is not None
        if not frm.contains_range(r.range()):
            return r
        moved = adjust_ref(r, Edit("move", range=frm, to_row=edit.to_row, to_col=edit.to_col))
        if moved is None:
            return None
        return replace(moved, sheet=edit.to_sheet)
    return r


def rename_sheet(e: Expr, old: str, new: str) -> Expr:
    """Renames sheet references ``old`` → ``new``."""
    ol = old.lower()

    def f(x: Expr) -> Expr:
        if isinstance(x, Ref):
            s = x.sheet
            if isinstance(s, str) and s.lower() == ol:
                return replace(x, sheet=new)
            if isinstance(s, tuple):
                a, b = s
                a = new if a.lower() == ol else a
                b = new if b.lower() == ol else b
                return replace(x, sheet=(a, b))
            return x
        if isinstance(x, Name):
            n = x.name
            head, rest = n[: len(old)], n[len(old):]
            # Names keep their sheet bare, as the parser does; the printer quotes it.
            if head.lower() == ol and rest.startswith("!"):
                return Name(new + rest)
            if n.startswith("'") and "!" in n:
                q, rest2 = n.rsplit("!", 1)
                if q.strip("'").replace("''", "'").lower() == ol:
                    return Name(new + "!" + rest2)
        return x

    return e.map(f)


def delete_sheet(e: Expr, name: str, order: Sequence[str]) -> Expr:
    """References to a deleted sheet become ``#REF!``. ``order`` lists the sheet names (before
    the deletion) in tab order: a 3-D span that loses an end shrinks to the neighbouring sheet."""
    nl = name.lower()
    lower = [o.lower() for o in order]

    def pos(n: str) -> Optional[int]:
        try:
            return lower.index(n.lower())
        except ValueError:
            return None

    def f(x: Expr) -> Expr:
        if not isinstance(x, Ref):
            return x
        s = x.sheet
        if isinstance(s, str) and s.lower() == nl:
            return Err(CellError.REF)
        if isinstance(s, tuple):
            a, b = s
            if a.lower() == nl or b.lower() == nl:
                ia, ib = pos(a), pos(b)
                if ia is None or ib is None:
                    return x
                if ia == ib:
                    return Err(CellError.REF)

                def inward(frm: int, to: int) -> Optional[str]:
                    i = frm + 1 if frm < to else frm - 1
                    return order[i] if 0 <= i < len(order) else None

                if a.lower() == nl:
                    na, nb = inward(ia, ib), order[ib]
                else:
                    na, nb = order[ia], inward(ib, ia)
                if na is None or nb is None:
                    return Err(CellError.REF)
                return replace(x, sheet=na if na.lower() == nb.lower() else (na, nb))
        return x

    return e.map(f)
