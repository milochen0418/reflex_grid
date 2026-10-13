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


def copy_rules(rules, src: RangeRef, dr: int, dc: int, transpose: bool = False) -> list:
    """Conditional formats or validations (objects with ``ranges``) that cover part of ``src``,
    as new rules for the copied block placed ``(dr, dc)`` away."""
    out = []
    for rule in rules:
        parts = [r.intersection(src) for r in rule.ranges]
        parts = [p for p in parts if p is not None]
        if not parts:
            continue
        top = rule.ranges[0].start
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
        ndr = moved[0].r0 - top[0]
        ndc = moved[0].c0 - top[1]
        if hasattr(new, "rule"):
            new.rule.formulas = [shift_formula_text(f, ndr, ndc) for f in new.rule.formulas]
        else:
            if new.f1 and not new.f1.startswith('"'):
                new.f1 = shift_formula_text(new.f1, ndr, ndc)
            if new.f2:
                new.f2 = shift_formula_text(new.f2, ndr, ndc)
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
            rule.ranges = rest
            out.append(rule)
    return out


def clamp_range(r: RangeRef) -> Optional[RangeRef]:
    r0, c0 = max(r.r0, 0), max(r.c0, 0)
    r1, c1 = min(r.r1, MAX_ROWS - 1), min(r.c1, MAX_COLS - 1)
    if r0 > r1 or c0 > c1:
        return None
    return RangeRef(r0, c0, r1, c1)
