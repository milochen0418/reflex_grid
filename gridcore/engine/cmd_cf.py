"""Conditional formatting commands: add rules (highlight, top/bottom, data bars, colour scales,
icon sets, formulas), clear, list, reorder and delete them."""

from __future__ import annotations

from dataclasses import replace
from typing import Any, Dict, List, Optional

from ..core.addr import RangeRef
from ..formula import ParseError, parse
from ..model.features import CfRule, CfValue, CondFormat, default_icon_thresholds
from ..model.style import BorderLine, Borders, Color, Fill, Style
from .cmd_format import _ranges, parse_color, style_from_patch
from .ranges import remove_rules_in
from .session import EngineError, Session, command, edit, need, p_bool, p_int, p_str, param_sheet

PRESETS = {
    "lightRed": lambda: Style(fill=Fill.solid(Color.rgb(0xFFC7CE)), font=replace(Style().font, color=Color.rgb(0x9C0006))),
    "yellow": lambda: Style(fill=Fill.solid(Color.rgb(0xFFEB9C)), font=replace(Style().font, color=Color.rgb(0x9C5700))),
    "green": lambda: Style(fill=Fill.solid(Color.rgb(0xC6EFCE)), font=replace(Style().font, color=Color.rgb(0x006100))),
    "lightRedFill": lambda: Style(fill=Fill.solid(Color.rgb(0xFFC7CE))),
    "redText": lambda: Style(font=replace(Style().font, color=Color.rgb(0x9C0006))),
    "redBorder": lambda: Style(borders=Borders(*(BorderLine("thin", Color.rgb(0xC00000)) for _ in range(4)))),
}
PRESETS["lightRedFillDarkRedText"] = PRESETS["lightRed"]

OPERATORS = {
    "between": "between",
    "notBetween": "notBetween",
    "equal": "equal",
    "equalTo": "equal",
    "notEqual": "notEqual",
    "greaterThan": "greaterThan",
    "greater": "greaterThan",
    "lessThan": "lessThan",
    "less": "lessThan",
    "greaterThanOrEqual": "greaterThanOrEqual",
    "lessThanOrEqual": "lessThanOrEqual",
}

COLOR_SCALES = {
    "greenYellowRed": ["#63BE7B", "#FFEB84", "#F8696B"],
    "redYellowGreen": ["#F8696B", "#FFEB84", "#63BE7B"],
    "greenWhiteRed": ["#63BE7B", "#FCFCFF", "#F8696B"],
    "redWhiteGreen": ["#F8696B", "#FCFCFF", "#63BE7B"],
    "blueWhiteRed": ["#5A8AC6", "#FCFCFF", "#F8696B"],
    "redWhiteBlue": ["#F8696B", "#FCFCFF", "#5A8AC6"],
    "whiteRed": ["#FCFCFF", "#F8696B"],
    "redWhite": ["#F8696B", "#FCFCFF"],
    "greenWhite": ["#63BE7B", "#FCFCFF"],
    "whiteGreen": ["#FCFCFF", "#63BE7B"],
    "greenYellow": ["#63BE7B", "#FFEF9C"],
    "yellowGreen": ["#FFEF9C", "#63BE7B"],
}

ICON_SETS = [
    "3Arrows", "3ArrowsGray", "3Flags", "3TrafficLights1", "3TrafficLights2", "3Signs", "3Symbols", "3Symbols2",
    "3Stars", "3Triangles", "4Arrows", "4ArrowsGray", "4RedToBlack", "4Rating", "4TrafficLights", "5Arrows",
    "5ArrowsGray", "5Rating", "5Quarters", "5Boxes",
]

TYPES = (
    "cellIs", "expression", "containsText", "notContainsText", "beginsWith", "endsWith", "containsBlanks",
    "notContainsBlanks", "containsErrors", "notContainsErrors", "duplicateValues", "uniqueValues", "top10",
    "aboveAverage", "timePeriod", "colorScale", "dataBar", "iconSet",
)


def _formula_text(v: Any) -> str:
    """A rule value: a number, quoted text or ``=formula``."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return repr(float(v)).rstrip("0").rstrip(".") if isinstance(v, float) else str(v)
    t = str(v).strip()
    if t.startswith("="):
        body = t[1:]
        try:
            parse(body)
        except ParseError as e:
            raise EngineError(f"There's a problem with this formula: {e}")
        return body
    from ..core.parse import parse_number_text

    if parse_number_text(t) is not None and not any(ch.isalpha() for ch in t.replace("E", "").replace("e", "")):
        n = parse_number_text(t)
        return repr(n).rstrip("0").rstrip(".") if n != int(n) else str(int(n))
    return '"' + t.replace('"', '""') + '"'


def _cfv(v: Any, default_kind: str) -> CfValue:
    if isinstance(v, dict):
        kind = str(v.get("kind", default_kind))
        val = v.get("value")
        if kind == "formula" and isinstance(val, str) and val.startswith("="):
            val = val[1:]
        return CfValue(kind, val)
    if v is None:
        return CfValue(default_kind)
    return CfValue("num", v)


def build_rule(r: Dict[str, Any]) -> CfRule:
    t = str(r.get("type", "cellIs"))
    if t not in TYPES:
        raise EngineError(f"unknown rule type `{t}`")
    style = None
    if t not in ("colorScale", "dataBar", "iconSet"):
        preset = r.get("preset")
        if isinstance(r.get("style"), dict):
            style = style_from_patch(Style(), r["style"])
        elif preset:
            f = PRESETS.get(str(preset))
            if f is None:
                raise EngineError(f"unknown preset `{preset}`")
            style = f()
        else:
            style = PRESETS["lightRed"]()
    rule = CfRule(type=t, style=style)
    if t == "cellIs":
        op = OPERATORS.get(str(r.get("operator", "greaterThan")))
        if op is None:
            raise EngineError(f"unknown operator `{r.get('operator')}`")
        rule.operator = op
        rule.formulas = [_formula_text(r.get("value", r.get("formula")))]
        if op in ("between", "notBetween"):
            rule.formulas.append(_formula_text(r.get("value2")))
        if rule.formulas[0] == "":
            raise EngineError("Enter a value for the rule.")
    elif t == "expression":
        f = str(need(r, "formula", "home.conditionalFormat")).strip()
        body = f[1:] if f.startswith("=") else f
        try:
            parse(body)
        except ParseError as e:
            raise EngineError(f"There's a problem with this formula: {e}")
        rule.formulas = [body]
    elif t in ("containsText", "notContainsText", "beginsWith", "endsWith"):
        rule.text = str(need(r, "text", "home.conditionalFormat"))
        rule.operator = {"containsText": "containsText", "notContainsText": "notContains", "beginsWith": "beginsWith", "endsWith": "endsWith"}[t]
    elif t == "top10":
        rule.rank = max(1, min(int(r.get("rank", 10)), 1000))
        rule.bottom = bool(r.get("bottom", False))
        rule.percent = bool(r.get("percent", False))
    elif t == "aboveAverage":
        rule.below = bool(r.get("below", False))
        rule.equal_average = bool(r.get("equal", False))
        rule.std_dev = int(r.get("stdDev", 0))
    elif t == "timePeriod":
        rule.period = str(r.get("period", "today"))
        if rule.period not in ("yesterday", "today", "tomorrow", "last7Days", "lastWeek", "thisWeek", "nextWeek", "lastMonth", "thisMonth", "nextMonth"):
            raise EngineError(f"unknown period `{rule.period}`")
    elif t == "colorScale":
        colors = r.get("colors")
        if colors is None:
            colors = COLOR_SCALES.get(str(r.get("scale", "greenYellowRed")))
            if colors is None:
                raise EngineError(f"unknown colour scale `{r.get('scale')}`")
        if len(colors) not in (2, 3):
            raise EngineError("A colour scale has two or three colours.")
        kinds = [_cfv(r.get("min"), "min"), _cfv(r.get("mid"), "percentile") if len(colors) == 3 else None, _cfv(r.get("max"), "max")]
        if len(colors) == 3 and kinds[1].value is None:
            kinds[1] = CfValue("percentile", 50)
        stops = [kinds[0]] + ([kinds[1]] if len(colors) == 3 else []) + [kinds[2]]
        rule.stops = [(k, parse_color(c)) for k, c in zip(stops, colors)]
    elif t == "dataBar":
        rule.bar_color = parse_color(r.get("color", "#638EC6"))
        rule.bar_min = _cfv(r.get("min"), "min")
        rule.bar_max = _cfv(r.get("max"), "max")
        rule.show_value = bool(r.get("showValue", True))
    elif t == "iconSet":
        rule.icon_set = str(r.get("iconSet", "3TrafficLights1"))
        if rule.icon_set not in ICON_SETS:
            raise EngineError(f"unknown icon set `{rule.icon_set}`")
        th = r.get("thresholds")
        rule.thresholds = [_cfv(x, "percent") for x in th] if th else default_icon_thresholds(rule.icon_set)
        rule.reverse = bool(r.get("reverse", False))
        rule.show_value = bool(r.get("showValue", True))
    return rule


@command("home.conditionalFormat", "Conditional Formatting", menu=("Home", "Styles"),
         params='{range?, rule: {type: cellIs|expression|containsText|notContainsText|beginsWith|endsWith|containsBlanks|notContainsBlanks|containsErrors|notContainsErrors|duplicateValues|uniqueValues|top10|aboveAverage|timePeriod|colorScale|dataBar|iconSet, operator?, value?, value2?, formula?, text?, preset?: lightRed|yellow|green|lightRedFill|redText|redBorder, style?: {font, fill, borders}, rank?, bottom?, percent?, below?, equal?, stdDev?, period?, scale?|colors?, min?, mid?, max?, color?, showValue?, iconSet?, thresholds?, reverse?}, stopIfTrue?}')
def add_rule(s: Session, p: Dict[str, Any]):
    si = param_sheet(s, p)
    ranges = _ranges(s, p)
    d = s.doc()
    sh = d.wb.sheet(si)
    used = sh.used_range()
    ranges = [r if not (r.is_full_cols() and r.is_full_rows()) or used is None else (r.intersection(used) or r) for r in ranges]
    rule = build_rule(dict(need(p, "rule", "home.conditionalFormat")))
    stop = p_bool(p, "stopIfTrue")

    def f(ctx):
        sm = ctx.sheet_mut(si)
        # The new rule goes on top.
        for cf in sm.cond_formats:
            cf.priority += 1
        sm.cond_formats.insert(0, CondFormat(list(ranges), rule, 1, stop))

    edit(s, f)
    return {"rules": len(s.doc().wb.sheet(si).cond_formats)}


@command("home.clearRules", "Clear Rules", menu=("Home", "Styles", "Conditional Formatting"), params="{scope?: selection|sheet}")
def clear_rules(s: Session, p: Dict[str, Any]):
    si = param_sheet(s, p)
    scope = p_str(p, "scope", "selection")
    ranges = _ranges(s, p)

    def f(ctx):
        sm = ctx.sheet_mut(si)
        if scope == "sheet":
            sm.cond_formats = []
            return
        rules = sm.cond_formats
        for r in ranges:
            rules = remove_rules_in(rules, r)
        sm.cond_formats = rules

    edit(s, f)
    return None


def describe(cf: CondFormat) -> str:
    r = cf.rule
    t = r.type
    if t == "cellIs":
        ops = {"between": "between", "notBetween": "not between", "equal": "=", "notEqual": "<>", "greaterThan": ">", "lessThan": "<", "greaterThanOrEqual": ">=", "lessThanOrEqual": "<="}
        vals = " and ".join(r.formulas)
        return f"Cell Value {ops.get(r.operator, r.operator)} {vals}"
    if t == "expression":
        return "Formula: =" + (r.formulas[0] if r.formulas else "")
    if t in ("containsText", "notContainsText", "beginsWith", "endsWith"):
        return {"containsText": "Cell Value contains", "notContainsText": "Cell Value does not contain", "beginsWith": "Cell Value begins with", "endsWith": "Cell Value ends with"}[t] + f' "{r.text}"'
    if t == "top10":
        return f"{'Bottom' if r.bottom else 'Top'} {r.rank}{'%' if r.percent else ''}"
    if t == "aboveAverage":
        return ("Below" if r.below else "Above") + (" or equal to" if r.equal_average else "") + " Average"
    if t == "colorScale":
        return f"Graded Color Scale ({len(r.stops)} colors)"
    if t == "dataBar":
        return "Data Bar"
    if t == "iconSet":
        return f"Icon Set ({r.icon_set})"
    return {
        "containsBlanks": "Blanks", "notContainsBlanks": "No Blanks", "containsErrors": "Errors", "notContainsErrors": "No Errors",
        "duplicateValues": "Duplicate Values", "uniqueValues": "Unique Values", "timePeriod": f"Dates Occurring ({r.period})",
    }.get(t, t)


@command("home.manageRules", "Manage Rules…", menu=("Home", "Styles", "Conditional Formatting"), params="{scope?: selection|sheet}", undoable=False, journal=False)
def manage_rules(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)
    scope = p_str(p, "scope", "sheet")
    sel = d.selection.ranges
    out = []
    for i, cf in sorted(enumerate(sh.cond_formats), key=lambda x: x[1].priority):
        if scope == "selection" and not any(r.intersects(x) for r in cf.ranges for x in sel):
            continue
        st = cf.rule.style
        out.append({
            "index": i,
            "type": cf.rule.type,
            "description": describe(cf),
            "appliesTo": ",".join(r.a1() for r in cf.ranges),
            "priority": cf.priority,
            "stopIfTrue": cf.stop_if_true,
            "fill": st.fill.fg.hex(d.wb.theme) if st is not None and st.fill.pattern != "none" else None,
            "color": st.font.color.hex(d.wb.theme) if st is not None else None,
        })
    return {"rules": out}


@command("home.deleteRule", "Delete Rule", menu=("Home", "Styles", "Conditional Formatting"), params="{index}")
def delete_rule(s: Session, p: Dict[str, Any]):
    si = param_sheet(s, p)
    i = p_int(p, "index")
    sh = s.doc().wb.sheet(si)
    if i is None or not 0 <= i < len(sh.cond_formats):
        raise EngineError("no such rule")

    def f(ctx):
        sm = ctx.sheet_mut(si)
        del sm.cond_formats[i]

    edit(s, f)
    return None


@command("home.moveRule", "Move Rule", menu=("Home", "Styles", "Conditional Formatting"), params="{index, up?: bool}")
def move_rule(s: Session, p: Dict[str, Any]):
    si = param_sheet(s, p)
    i = p_int(p, "index")
    sh = s.doc().wb.sheet(si)
    if i is None or not 0 <= i < len(sh.cond_formats):
        raise EngineError("no such rule")
    up = p_bool(p, "up", True)

    def f(ctx):
        sm = ctx.sheet_mut(si)
        order = sorted(range(len(sm.cond_formats)), key=lambda k: sm.cond_formats[k].priority)
        pos = order.index(i)
        npos = pos - 1 if up else pos + 1
        if 0 <= npos < len(order):
            order[pos], order[npos] = order[npos], order[pos]
        for prio, k in enumerate(order, 1):
            sm.cond_formats[k].priority = prio

    edit(s, f)
    return None


@command("home.editRule", "Edit Rule", menu=("Home", "Styles", "Conditional Formatting"), params="{index, rule?, appliesTo?: \"A1:B10\", stopIfTrue?}")
def edit_rule(s: Session, p: Dict[str, Any]):
    si = param_sheet(s, p)
    i = p_int(p, "index")
    sh = s.doc().wb.sheet(si)
    if i is None or not 0 <= i < len(sh.cond_formats):
        raise EngineError("no such rule")
    rule = build_rule(dict(p["rule"])) if p.get("rule") else None
    ranges = None
    if p.get("appliesTo"):
        ranges = []
        for part in str(p["appliesTo"]).replace(" ", ",").split(","):
            if not part:
                continue
            r = RangeRef.parse(part.rsplit("!", 1)[-1])
            if r is None:
                raise EngineError(f"not a range: `{part}`")
            ranges.append(r)

    def f(ctx):
        cf = ctx.sheet_mut(si).cond_formats[i]
        if rule is not None:
            cf.rule = rule
        if ranges:
            cf.ranges = ranges
        if "stopIfTrue" in p:
            cf.stop_if_true = p_bool(p, "stopIfTrue")

    edit(s, f)
    return None
