"""Conditional formatting evaluation: the extra style, colour-scale fill, data bar or icon a
cell gets."""

from __future__ import annotations

import math
from dataclasses import dataclass, replace
from typing import Dict, List, Optional, Tuple

from ..calc import evaluate_expr
from ..core.addr import RangeRef
from ..core.date import datetime_from_serial, now_serial
from ..core.value import Array, CellError, compare, display, to_bool, XlError
from ..formula import ParseError, parse
from ..formula.adjust import shift_relative
from ..model.features import CfValue, CondFormat, default_icon_thresholds, icon_count
from ..model.style import AUTO, BorderLine, Style


@dataclass
class CfLook:
    style: Optional[Style] = None
    scale_fill: Optional[Tuple[int, int, int]] = None
    # (fraction 0..1, colour, gradient)
    bar: Optional[Tuple[float, Tuple[int, int, int], bool]] = None
    # (set name, index from 0 = lowest)
    icon: Optional[Tuple[str, int]] = None
    hide_value: bool = False


@dataclass
class _Stats:
    sorted: List[float]
    avg: float
    std: float
    counts: Dict[str, int]


def _numbers_in(sh, ranges: List[RangeRef]) -> Tuple[List[float], Dict[str, int]]:
    nums: List[float] = []
    counts: Dict[str, int] = {}
    used = sh.used_range()
    if used is None:
        return nums, counts
    for r in ranges:
        inter = r.intersection(used)
        if inter is None:
            continue
        n = 0
        for rc, cell in sh.cells.iter_range(inter):
            v = sh.value(rc)
            n += 1
            if isinstance(v, float):
                nums.append(v)
            if v is not None:
                k = display(v).lower()
                counts[k] = counts.get(k, 0) + 1
            if n >= 2_000_000:
                break
        for rc, v in sh.spill.items():
            if inter.contains(rc) and sh.cell(rc) is None:
                if isinstance(v, float):
                    nums.append(v)
                if v is not None:
                    k = display(v).lower()
                    counts[k] = counts.get(k, 0) + 1
    nums.sort()
    return nums, counts


def _percentile(xs: List[float], p: float) -> float:
    if not xs:
        return 0.0
    rank = min(max(p / 100.0, 0.0), 1.0) * (len(xs) - 1)
    lo, hi = int(math.floor(rank)), int(math.ceil(rank))
    a = xs[lo]
    b = xs[hi]
    return a + (b - a) * (rank - lo)


def _rel_expr(text: str, first: Tuple[int, int], c: Tuple[int, int]):
    body = text[1:] if text.startswith("=") else text
    try:
        e = parse(body)
    except ParseError:
        return None
    if first == c:
        return e
    return shift_relative(e, c[0] - first[0], c[1] - first[1])


def _eval(wb, si: int, c, text: str, first) -> object:
    e = _rel_expr(text, first, c)
    if e is None:
        return CellError.NAME
    v = evaluate_expr(wb, si, c, e)
    if isinstance(v, Array):
        v = v.get(0, 0)
    return v


def _num(v) -> float:
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, float):
        return v
    if isinstance(v, str):
        try:
            return float(v)
        except ValueError:
            return 0.0
    return 0.0


def _threshold(wb, si: int, c, st: _Stats, t: CfValue, first) -> float:
    mn = st.sorted[0] if st.sorted else 0.0
    mx = st.sorted[-1] if st.sorted else 0.0
    k = t.kind
    if k in ("min", "autoMin"):
        return mn
    if k in ("max", "autoMax"):
        return mx
    if k == "num":
        v = t.value
        if isinstance(v, str):
            try:
                return float(v)
            except ValueError:
                return _num(_eval(wb, si, c, v, first))
        return _num(v)
    if k == "percent":
        return mn + (mx - mn) * _num(t.value) / 100.0
    if k == "percentile":
        return _percentile(st.sorted, _num(t.value))
    if k == "formula":
        return _num(_eval(wb, si, c, str(t.value), first))
    return mn


def _lerp(a, b, t: float):
    t = min(max(t, 0.0), 1.0)
    return tuple(int(round(a[i] + (b[i] - a[i]) * t)) for i in range(3))


def _layer_under(style: Style, lower: Style) -> Style:
    """Adds the parts ``lower`` sets that ``style`` leaves unset."""
    font = style.font
    if font.color.is_auto:
        font = replace(font, color=lower.font.color)
    font = replace(font, bold=font.bold or lower.font.bold, italic=font.italic or lower.font.italic, strike=font.strike or lower.font.strike)
    if font.underline == "none":
        font = replace(font, underline=lower.font.underline)
    fill = style.fill if style.fill.pattern != "none" else lower.fill
    b = style.borders
    kw = {}
    for side in ("left", "right", "top", "bottom", "diag_down", "diag_up"):
        if getattr(b, side).is_none:
            kw[side] = getattr(lower.borders, side)
    borders = replace(b, **kw)
    num_fmt = style.num_fmt if style.num_fmt != "General" else lower.num_fmt
    return replace(style, font=font, fill=fill, borders=borders, num_fmt=num_fmt)


class CfCache:
    """Evaluates conditional formats for many cells of one sheet, sharing per-rule statistics."""

    def __init__(self, wb, si: int):
        self.wb = wb
        self.si = si
        self.sh = wb.sheet(si)
        self.stats: Dict[int, _Stats] = {}
        self.today = math.floor(now_serial())

    def _stats(self, idx: int, cf: CondFormat) -> _Stats:
        st = self.stats.get(idx)
        if st is None:
            nums, counts = _numbers_in(self.sh, cf.ranges)
            n = max(len(nums), 1)
            avg = sum(nums) / n if nums else 0.0
            std = math.sqrt(sum((x - avg) ** 2 for x in nums) / n) if nums else 0.0
            st = _Stats(nums, avg, std, counts)
            self.stats[idx] = st
        return st

    def look(self, c: Tuple[int, int]) -> Optional[CfLook]:
        sh = self.sh
        if sh is None or not sh.cond_formats:
            return None
        rules = [(i, cf) for i, cf in enumerate(sh.cond_formats) if any(r.contains(c) for r in cf.ranges)]
        if not rules:
            return None
        rules.sort(key=lambda x: x[1].priority)
        wb = self.wb
        si = self.si
        v = sh.value(c)
        if isinstance(v, Array):
            v = v.get(0, 0)
        out = CfLook()
        hit = False
        for idx, cf in rules:
            rule = cf.rule
            first = cf.ranges[0].start if cf.ranges else c
            t = rule.type
            matched: Optional[Style] = None
            ok = False
            try:
                if t == "cellIs":
                    if not isinstance(v, CellError):
                        a = _eval(wb, si, c, rule.formulas[0], first) if rule.formulas else None
                        cv = v
                        op = rule.operator
                        if op in ("between", "notBetween"):
                            b = _eval(wb, si, c, rule.formulas[1], first) if len(rule.formulas) > 1 else None
                            lo, hi = (b, a) if compare(a, b) > 0 else (a, b)
                            inside = compare(cv, lo) >= 0 and compare(cv, hi) <= 0
                            ok = inside if op == "between" else not inside
                        elif not isinstance(a, CellError):
                            k = compare(cv, a)
                            ok = {
                                "equal": k == 0,
                                "notEqual": k != 0,
                                "greaterThan": k > 0,
                                "lessThan": k < 0,
                                "greaterThanOrEqual": k >= 0,
                                "lessThanOrEqual": k <= 0,
                            }.get(op, False)
                elif t == "expression":
                    r = _eval(wb, si, c, rule.formulas[0], first) if rule.formulas else False
                    try:
                        ok = not isinstance(r, CellError) and to_bool(r)
                    except XlError:
                        ok = False
                elif t in ("containsText", "notContainsText", "beginsWith", "endsWith"):
                    text = display(v).lower() if v is not None else ""
                    needle = (rule.text or "").lower()
                    ok = {
                        "containsText": needle in text,
                        "notContainsText": needle not in text,
                        "beginsWith": text.startswith(needle),
                        "endsWith": text.endswith(needle),
                    }[t]
                elif t == "containsBlanks":
                    ok = v is None or (isinstance(v, str) and v.strip() == "")
                elif t == "notContainsBlanks":
                    ok = not (v is None or (isinstance(v, str) and v.strip() == ""))
                elif t == "containsErrors":
                    ok = isinstance(v, CellError)
                elif t == "notContainsErrors":
                    ok = not isinstance(v, CellError)
                elif t in ("duplicateValues", "uniqueValues"):
                    if v is not None:
                        n = self._stats(idx, cf).counts.get(display(v).lower(), 0)
                        ok = n > 1 if t == "duplicateValues" else n == 1
                elif t == "top10":
                    st = self._stats(idx, cf)
                    if isinstance(v, float) and st.sorted:
                        k = int(math.floor(len(st.sorted) * rule.rank / 100.0)) if rule.percent else rule.rank
                        k = max(1, min(max(k, 1), len(st.sorted)))
                        ok = v <= st.sorted[k - 1] if rule.bottom else v >= st.sorted[len(st.sorted) - k]
                elif t == "aboveAverage":
                    st = self._stats(idx, cf)
                    if isinstance(v, float):
                        thr = st.avg + (-1.0 if rule.below else 1.0) * st.std * rule.std_dev
                        ok = (v < thr or (rule.equal_average and v == thr)) if rule.below else (v > thr or (rule.equal_average and v == thr))
                elif t == "timePeriod":
                    if isinstance(v, float):
                        ok = self._period(rule.period, math.floor(v))
                elif t == "colorScale":
                    if isinstance(v, float) and out.scale_fill is None:
                        st = self._stats(idx, cf)
                        pts = []
                        for tv, col in rule.stops:
                            rgb = col.resolve(wb.theme) or (255, 255, 255)
                            pts.append((_threshold(wb, si, c, st, tv, first), rgb))
                        if len(pts) == 2:
                            (a, ca), (b, cb) = pts
                            out.scale_fill = _lerp(ca, cb, (v - a) / (b - a) if b > a else 0.0)
                        elif len(pts) >= 3:
                            (a, ca), (m, cm), (b, cb) = pts[:3]
                            if v <= m:
                                out.scale_fill = _lerp(ca, cm, (v - a) / (m - a) if m > a else 0.0)
                            else:
                                out.scale_fill = _lerp(cm, cb, (v - m) / (b - m) if b > m else 1.0)
                        hit = hit or out.scale_fill is not None
                elif t == "dataBar":
                    if isinstance(v, float) and out.bar is None:
                        st = self._stats(idx, cf)
                        lo = _threshold(wb, si, c, st, rule.bar_min, first)
                        hi = _threshold(wb, si, c, st, rule.bar_max, first)
                        lo = min(lo, 0.0)
                        frac = min(max((v - lo) / (hi - lo), 0.0), 1.0) if hi > lo else 1.0
                        color = rule.bar_color.resolve(wb.theme) if not rule.bar_color.is_auto else None
                        out.bar = (frac * 0.9 + (0.1 if frac > 0 else 0.0), color or (0x63, 0x8E, 0xC6), True)
                        out.hide_value = out.hide_value or not rule.show_value
                        hit = True
                elif t == "iconSet":
                    if isinstance(v, float) and out.icon is None:
                        st = self._stats(idx, cf)
                        k = icon_count(rule.icon_set)
                        th = rule.thresholds
                        if len(th) == k:
                            lower = th[1:]
                        elif len(th) + 1 == k:
                            lower = th
                        else:
                            lower = default_icon_thresholds(rule.icon_set)[1:]
                        cuts = [_threshold(wb, si, c, st, x, first) for x in lower]
                        i = sum(1 for cut in cuts if v >= cut)
                        if rule.reverse:
                            i = k - 1 - min(i, k - 1)
                        out.icon = (rule.icon_set, min(i, k - 1))
                        out.hide_value = out.hide_value or not rule.show_value
                        hit = True
            except Exception:
                ok = False
            if ok and t not in ("colorScale", "dataBar", "iconSet"):
                matched = rule.style or Style()
            if matched is not None:
                hit = True
                out.style = matched if out.style is None else _layer_under(out.style, matched)
                if cf.stop_if_true:
                    break
        return out if hit else None

    def _period(self, period: str, d: float) -> bool:
        wb = self.wb
        today = float(self.today)
        dt = datetime_from_serial(wb.date_system, today)
        if dt is None:
            return False
        wd = (int(today) + 6) % 7  # 0 = Sunday for the 1900 system
        week_start = today - wd
        ym = datetime_from_serial(wb.date_system, d)

        def month_off(k: int) -> bool:
            m = dt.month + k
            y = dt.year + (m - 1) // 12
            m = (m - 1) % 12 + 1
            return ym is not None and (ym.year, ym.month) == (y, m)

        return {
            "yesterday": d == today - 1,
            "today": d == today,
            "tomorrow": d == today + 1,
            "last7Days": today - 7 < d <= today,
            "lastWeek": week_start - 7 <= d < week_start,
            "thisWeek": week_start <= d < week_start + 7,
            "nextWeek": week_start + 7 <= d < week_start + 14,
            "lastMonth": month_off(-1),
            "thisMonth": month_off(0),
            "nextMonth": month_off(1),
        }.get(period, False)
