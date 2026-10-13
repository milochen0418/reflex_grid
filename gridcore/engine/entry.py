"""Turning what someone types into a cell, and checking it against data validation."""

from __future__ import annotations

from dataclasses import replace
from typing import List, Optional, Tuple

from ..calc import evaluate_expr
from ..core.addr import RangeRef
from ..core.parse import parse_input, parse_number_text
from ..core.value import Array, CellError, Value, to_text
from ..formula import Call, ParseError, parse
from ..formula.adjust import shift_relative
from ..functions import lookup
from ..model.cell import Cell, Formula
from ..model.features import Validation
from ..numfmt import parse as parse_fmt
from .session import EngineError

_AUTO_DATE = (
    (("TODAY(", "DATE(", "EDATE(", "EOMONTH(", "WORKDAY(", "WORKDAY.INTL("), "m/d/yyyy"),
    (("NOW(",), "m/d/yyyy h:mm"),
    (("TIME(",), "h:mm AM/PM"),
)


def is_formula_input(text: str) -> bool:
    if text.startswith("="):
        return len(text) > 1
    return (
        len(text) > 1
        and text[0] in "+-"
        and parse_number_text(text) is None
        and (text[1].isascii() and text[1].isalpha() or text[1] == "(")
    )


def check_arg_counts(expr) -> None:
    """A built-in function called with too few or too many arguments is refused when entered."""
    for x in expr.walk():
        if isinstance(x, Call):
            spec = lookup(x.name)
            if spec is not None and not spec.accepts(len(x.args)):
                few = len(x.args) < spec.min_args
                raise EngineError(f"{spec.name} has too {'few' if few else 'many'} arguments. Use {spec.signature}")


def parse_formula_input(text: str) -> Formula:
    body = text[1:] if text.startswith("=") else text
    candidate = body
    err: Optional[ParseError] = None
    for _ in range(9):
        try:
            e = parse(candidate)
        except ParseError as ex:
            err = err or ex
            # Excel closes missing parentheses for you.
            candidate += ")"
            continue
        check_arg_counts(e)
        return Formula.from_expr(e)
    raise EngineError(f"There's a problem with this formula: {err}")


def input_to_cell(wb, sheet, rc: Tuple[int, int], text: str) -> Optional[Cell]:
    """The cell that typing ``text`` into ``rc`` produces (``None``: an empty, unformatted cell).
    May intern new styles (automatic number formats, wrap for line breaks)."""
    style = sheet.style_id(rc)
    st = wb.styles.get(style)
    if text == "":
        c = Cell(None, None, style)
        return None if c.is_blank() else c
    if st.num_fmt == "@":
        return Cell(text, None, style)
    if is_formula_input(text):
        f = parse_formula_input(text)
        if st.num_fmt == "General":
            up = f.text.upper()
            for prefixes, code in _AUTO_DATE:
                if up.startswith(prefixes):
                    style = wb.styles.derive(style, lambda s, code=code: replace(s, num_fmt=code))
                    break
        return Cell(None, f, style)
    parsed = parse_input(text, wb.date_system)
    v = parsed.value
    if parsed.format is not None and st.num_fmt == "General":
        style = wb.styles.derive(style, lambda s: replace(s, num_fmt=parsed.format))
    elif isinstance(v, float) and parsed.format is None and parse_fmt(st.num_fmt).is_percent() and "%" not in text:
        # Automatic percent entry: 10 typed into a percent cell is 10%.
        if abs(v) >= 1:
            v = v / 100.0
    if "\n" in text and not wb.styles.get(style).align.wrap:
        style = wb.styles.derive(style, lambda s: replace(s, align=replace(s.align, wrap=True)))
    return Cell(v, None, style)


# ---------------------------------------------------------------- data validation


def validation_at(sheet, rc: Tuple[int, int]) -> Optional[Validation]:
    for v in sheet.validations:
        if any(r.contains(rc) for r in v.ranges):
            return v
    return None


def _rule_expr(text: str, v: Validation, rc: Tuple[int, int]):
    """A validation formula evaluated for ``rc``: relative parts are relative to the rule's
    top-left cell."""
    body = text[1:] if text.startswith("=") else text
    try:
        e = parse(body)
    except ParseError:
        return None
    top = v.ranges[0].start if v.ranges else rc
    return shift_relative(e, rc[0] - top[0], rc[1] - top[1])


def _eval_rule(wb, si: int, text: Optional[str], v: Validation, rc: Tuple[int, int]) -> Value:
    if text is None or text == "":
        return None
    e = _rule_expr(text, v, rc)
    if e is None:
        return CellError.NAME
    return evaluate_expr(wb, si, rc, e)


def _flatten(x: Value) -> List[Value]:
    if isinstance(x, Array):
        return [x.get(r, c) for r in range(x.rows) for c in range(x.cols)]
    return [x]


def list_items(wb, si: int, v: Validation, rc: Tuple[int, int]) -> List[str]:
    """Entries of a list validation (the in-cell dropdown)."""
    f1 = v.f1.strip()
    if f1.startswith('"') and f1.endswith('"') and len(f1) >= 2:
        return [x.strip() for x in f1[1:-1].replace('""', '"').split(",") if x.strip() != ""]
    val = _eval_rule(wb, si, f1, v, rc)
    out = []
    from .display import fmt

    for x in _flatten(val):
        if x is None:
            continue
        if isinstance(x, CellError):
            continue
        out.append(to_text(x) if not isinstance(x, float) else fmt(x, "General", wb).text)
    return out[:10000]


def _compare(op: str, x: float, a: Value, b: Value) -> bool:
    def n(y):
        if isinstance(y, float):
            return y
        if isinstance(y, bool):
            return float(y)
        if isinstance(y, str):
            p = parse_number_text(y)
            return p
        return None

    a1, b1 = n(a), n(b)
    if a1 is None:
        return False
    if op == "between":
        return b1 is not None and min(a1, b1) <= x <= max(a1, b1)
    if op == "notBetween":
        return b1 is not None and not (min(a1, b1) <= x <= max(a1, b1))
    return {
        "equal": x == a1,
        "notEqual": x != a1,
        "greaterThan": x > a1,
        "lessThan": x < a1,
        "greaterThanOrEqual": x >= a1,
        "lessThanOrEqual": x <= a1,
    }.get(op, True)


def is_valid(wb, si: int, rc: Tuple[int, int], value: Value, v: Validation) -> bool:
    if value is None or value == "":
        return v.allow_blank or v.kind == "any"
    k = v.kind
    if k == "any":
        return True
    if k == "list":
        items = list_items(wb, si, v, rc)
        if isinstance(value, float):
            from .display import fmt

            t = fmt(value, "General", wb).text
        else:
            t = to_text(value) if not isinstance(value, CellError) else value.code
        low = t.lower()
        return any(x.lower() == low for x in items)
    if k == "custom":
        r = _eval_rule(wb, si, v.f1, v, rc)
        if isinstance(r, Array):
            r = r.get(0, 0)
        return r is True or (isinstance(r, float) and r != 0)
    a = _eval_rule(wb, si, v.f1, v, rc)
    b = _eval_rule(wb, si, v.f2, v, rc)
    if isinstance(a, Array):
        a = a.get(0, 0)
    if isinstance(b, Array):
        b = b.get(0, 0)
    if k == "textLength":
        if isinstance(value, float):
            from .display import fmt

            t = fmt(value, "General", wb).text
        else:
            t = to_text(value) if not isinstance(value, CellError) else value.code
        return _compare(v.operator, float(len(t)), a, b)
    if not isinstance(value, float):
        return False
    if k == "whole" and value != int(value):
        return False
    if k == "time" and not 0 <= value < 1:
        # A time alone; a date-time is checked on its time part.
        value = value - int(value)
    return _compare(v.operator, value, a, b)


def check_entry(wb, si: int, rc: Tuple[int, int], value: Value) -> Optional[Tuple[Validation, str]]:
    """The rule an entered value breaks (and its message), or ``None``."""
    sh = wb.sheet(si)
    v = validation_at(sh, rc) if sh is not None else None
    if v is None or not v.show_error or is_valid(wb, si, rc, value, v):
        return None
    msg = v.error_message or "This value doesn't match the data validation restrictions defined for this cell."
    if v.error_title:
        msg = f"{v.error_title}: {msg}"
    return v, msg
