"""Range-list arithmetic: subtracting a rectangle, turning cell sets into ranges, and copying
range-based features (conditional formats, validation) to a new place."""

from __future__ import annotations

import copy
from typing import Iterable, List, Optional, Tuple

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef
from ..formula import ParseError, parse, print_formula
from ..formula.adjust import shift_relative


def subtract(r: RangeRef, cut: RangeRef) -> List[RangeRef]:
    """``r`` minus ``cut``: up to four rectangles."""
    i = r.intersection(cut)
    if i is None:
        return [r]
    out = []
    if r.r0 < i.r0:
        out.append(RangeRef(r.r0, r.c0, i.r0 - 1, r.c1))
    if i.r1 < r.r1:
        out.append(RangeRef(i.r1 + 1, r.c0, r.r1, r.c1))
    if r.c0 < i.c0:
        out.append(RangeRef(i.r0, r.c0, i.r1, i.c0 - 1))
    if i.c1 < r.c1:
        out.append(RangeRef(i.r0, i.c1 + 1, i.r1, r.c1))
    return out


def subtract_all(ranges: Iterable[RangeRef], cut: RangeRef) -> List[RangeRef]:
    out: List[RangeRef] = []
    for r in ranges:
        out.extend(subtract(r, cut))
    return out


def cells_to_ranges(cells: Iterable[Tuple[int, int]], limit: int = 20000) -> List[RangeRef]:
    """Groups cells into rectangles: runs within a row, then identical runs on consecutive rows."""
    by_row = {}
    for r, c in cells:
        by_row.setdefault(r, []).append(c)
    runs: List[RangeRef] = []
    for r in sorted(by_row):
        cols = sorted(set(by_row[r]))
        start = prev = cols[0]
        for c in cols[1:]:
            if c == prev + 1:
                prev = c
                continue
            runs.append(RangeRef(r, start, r, prev))
            start = prev = c
        runs.append(RangeRef(r, start, r, prev))
    # Merge vertically: a run continues the rectangle that ended on the row above with the same span.
    open_: dict = {}
    out: List[RangeRef] = []
    for run in runs:
        key = (run.c0, run.c1)
        prev = open_.get(key)
        if prev is not None and prev.r1 == run.r0 - 1:
            open_[key] = RangeRef(prev.r0, prev.c0, run.r1, prev.c1)
        else:
            if prev is not None:
                out.append(prev)
            open_[key] = run
    out.extend(open_.values())
    out.sort()
    return out[:limit]


def shift_formula_text(text: str, dr: int, dc: int) -> str:
    """A formula's relative references moved by ``(dr, dc)`` (text with or without ``=``)."""
    if dr == 0 and dc == 0:
        return text
    eq = text.startswith("=")
    body = text[1:] if eq else text
    if body.startswith('"'):
        return text
    try:
        e = parse(body)
    except ParseError:
        return text
    out = print_formula(shift_relative(e, dr, dc))
    return ("=" + out) if eq else out


def top_left(ranges: Iterable[RangeRef]) -> Optional[Tuple[int, int]]:
    """The top-left cell of a list of areas: the lowest row and the lowest column across all of them,
    whatever order they are listed in. Excel writes the relative formulas of a multi-area validation
    or conditional format for this cell (GridCraft #3152). ``None`` for no areas."""
    ranges = list(ranges)
    if not ranges:
        return None
    return (min(r.r0 for r in ranges), min(r.c0 for r in ranges))


def rebase_rule(rule, frm: Optional[Tuple[int, int]], to: Optional[Tuple[int, int]]) -> None:
    """Moves the relative references of a rule's formulas, written for the cell ``frm``, to be
    written for ``to``, so each of its cells keeps reading the same cells (a validation's
    ``f1``/``f2``, a conditional format's ``rule.formulas``)."""
    if frm is None or to is None or frm == to:
        return
    dr, dc = to[0] - frm[0], to[1] - frm[1]
    if hasattr(rule, "rule"):
        rule.rule.formulas = [shift_formula_text(f, dr, dc) for f in rule.rule.formulas]
    else:
        if rule.f1 and not rule.f1.startswith('"'):
            rule.f1 = shift_formula_text(rule.f1, dr, dc)
        if rule.f2:
            rule.f2 = shift_formula_text(rule.f2, dr, dc)


def copy_rules(rules, src: RangeRef, dr: int, dc: int, transpose: bool = False) -> list:
    """Conditional formats or validations (objects with ``ranges``) that cover part of ``src``,
    as new rules for the copied block placed ``(dr, dc)`` away."""
    out = []
    for rule in rules:
        parts = [r.intersection(src) for r in rule.ranges]
        parts = [p for p in parts if p is not None]
        if not parts:
            continue
        top = top_left(rule.ranges)
        new = copy.deepcopy(rule)
        moved = []
        for p in parts:
            if transpose:
                a = (src.r0 + (p.c0 - src.c0), src.c0 + (p.r0 - src.r0))
                b = (src.r0 + (p.c1 - src.c0), src.c0 + (p.r1 - src.r0))
                p = RangeRef(a[0], a[1], b[0], b[1])
            q = RangeRef(p.r0 + dr, p.c0 + dc, p.r1 + dr, p.c1 + dc)
            if 0 <= q.r0 and q.r1 < MAX_ROWS and 0 <= q.c0 and q.c1 < MAX_COLS:
                moved.append(q)
        if not moved:
            continue
        new.ranges = moved
        # The copy's formulas are written for its own top-left cell: a relative reference moves
        # with the cells, as a copied formula's does.
        rebase_rule(new, top, moved[0].start if transpose else top_left(moved))
        out.append(new)
    return out


def remove_rules_in(rules, area: RangeRef) -> list:
    """Rules with ``area`` cut out of their ranges (rules left without ranges are dropped)."""
    out = []
    for rule in rules:
        if not any(r.intersects(area) for r in rule.ranges):
            out.append(rule)
            continue
        rest = subtract_all(rule.ranges, area)
        if rest:
            before = top_left(rule.ranges)
            rule.ranges = rest
            # When the rule's top-left cell goes, its relative formulas are rewritten for the new
            # one, so the cells left keep reading the same cells (GridCraft PR 3173).
            rebase_rule(rule, before, top_left(rest))
            out.append(rule)
    return out


def clamp_range(r: RangeRef) -> Optional[RangeRef]:
    r0, c0 = max(r.r0, 0), max(r.c0, 0)
    r1, c1 = min(r.r1, MAX_ROWS - 1), min(r.c1, MAX_COLS - 1)
    if r0 > r1 or c0 > c1:
        return None
    return RangeRef(r0, c0, r1, c1)
