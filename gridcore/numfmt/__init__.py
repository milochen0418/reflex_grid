"""Number formats: Excel format codes (``#,##0.00;[Red](#,##0.00)``, dates, fractions,
scientific…), the General display format and the ``TEXT()`` worksheet function. en-US only."""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from functools import lru_cache
from typing import List, Optional, Tuple

from ..core.date import DateSystem
from ..core.parse import parse_number_text
from ..core.value import Array, CellError, XlError, Value, scalar
from . import render
from .builtin import builtin_format, builtin_id
from .decimal import format_general_fit
from .parse import Section, is_calendar, parse_sections

GENERAL_WIDTH = 11
OVERFLOW = "########"

# The 56-colour legacy palette for `[ColorN]` (RRGGBB), index 1-based.
PALETTE56 = [
    0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
    0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
    0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
    0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
    0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
    0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
    0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333,
]

NAMED_COLORS = {
    "black": "#000000",
    "blue": "#0000FF",
    "cyan": "#00FFFF",
    "green": "#00FF00",
    "magenta": "#FF00FF",
    "red": "#FF0000",
    "white": "#FFFFFF",
    "yellow": "#FFFF00",
}


def color_hex(color: Optional[str]) -> Optional[str]:
    """CSS colour of a ``[Red]``/``[Color10]`` section colour."""
    if color is None:
        return None
    if color in NAMED_COLORS:
        return NAMED_COLORS[color]
    if color.startswith("color"):
        try:
            i = int(color[5:])
        except ValueError:
            return None
        if 1 <= i <= 56:
            return f"#{PALETTE56[i - 1]:06X}"
    return None


@dataclass
class Formatted:
    text: str = ""
    color: Optional[str] = None
    # Repeat-fill char from `*x` and its position in `text`, for the renderer to expand.
    fill: Optional[Tuple[str, int]] = None
    # True when the result is right-aligned by default (numbers).
    numeric: bool = False


SPECIAL_CODES = ("00000", "00000-0000", "[<=9999999]###-####;(###) ###-####", "000-00-0000")


class NumberFormat:
    """A parsed number format code."""

    __slots__ = ("code", "general", "sections", "num_sections", "text_section")

    def __init__(self, code: str):
        self.code = code
        trimmed = code.strip()
        if not trimmed or trimmed.lower() == "general" or trimmed == "checkbox":
            self.general = True
            self.sections: List[Section] = []
            self.num_sections = 0
            self.text_section: Optional[int] = None
            return
        self.general = False
        self.sections = parse_sections(code)
        n = len(self.sections)
        last_is_text = n > 0 and self.sections[-1].kind == "text"
        if n >= 4:
            self.text_section = 3
            self.num_sections = 3
        elif last_is_text:
            self.text_section = n - 1
            self.num_sections = n - 1
        else:
            self.text_section = None
            self.num_sections = n

    def _nums(self) -> List[Section]:
        return self.sections[: self.num_sections]

    def is_general(self) -> bool:
        return self.general

    def is_date(self) -> bool:
        return any(s.kind == "date" for s in self._nums())

    def is_time_only(self) -> bool:
        secs = self._nums()
        return bool(secs) and all(
            s.kind == "date" and not any(t[0] == "date" and is_calendar(t[1]) for t in s.toks) for s in secs
        )

    def is_percent(self) -> bool:
        return any(s.kind == "number" and s.percent > 0 for s in self._nums())

    def is_text(self) -> bool:
        return self.num_sections == 0 and self.text_section is not None

    def kind(self) -> str:
        """Best guess at the ribbon's number-format category."""
        if self.general:
            return "General"
        if self.is_text():
            return "Text"
        if self.code in SPECIAL_CODES:
            return "Special"
        if not self.sections:
            return "Custom"
        s = self.sections[0]
        if s.kind == "general":
            return "General"
        if s.kind == "text":
            return "Text"
        if s.kind == "date":
            cal = any(t[0] == "date" and is_calendar(t[1]) for t in s.toks)
            return "Date" if cal else "Time"
        if s.fraction:
            return "Fraction"
        if s.exp:
            return "Scientific"
        if s.percent > 0:
            return "Percentage"
        if s.fill and s.has_digits:
            return "Accounting"
        if s.currency and s.has_digits:
            return "Currency"
        if s.has_digits:
            return "Number"
        return "Custom"

    def decimals(self) -> int:
        """Number of decimal places of the first number section (for Increase/Decrease Decimal)."""
        if not self.sections:
            return 0
        s = self.sections[0]
        seen = False
        n = 0
        for t in s.toks:
            if t[0] == "point":
                seen = True
            elif seen and t[0] == "digit":
                n += 1
            elif seen and t[0] not in ("digit",):
                if n:
                    break
        return n

    def _select(self, v: float) -> Optional[Tuple[Section, bool]]:
        secs = self._nums()
        k = len(secs)
        if not secs:
            return None
        s0 = secs[0]
        c0 = s0.cond
        c1 = secs[1].cond if k > 1 else None
        if c0 is not None or c1 is not None:
            if c0 is not None and c0.test(v):
                idx = 0
            elif c1 is not None and c1.test(v):
                idx = 1
            elif c0 is not None and c1 is not None:
                idx = 2
            elif c0 is None:
                idx = 0
            else:
                idx = 1
            if idx < k:
                # Excel never adds a minus sign in a multi-section conditional format.
                return secs[idx], k >= 2
            return None
        if v < 0 and k >= 2:
            return secs[1], True
        if v == 0 and k >= 3:
            return secs[2], False
        return s0, False

    def format_number(self, n: float, sys: DateSystem = DateSystem.D1900, width: int = GENERAL_WIDTH) -> Tuple[Formatted, bool]:
        """Formats a number; the flag is true when the value cannot be shown (date out of range)."""

        def general(x: float) -> Formatted:
            g = format_general_fit(x, width)
            return Formatted(g if g is not None else OVERFLOW, None, None, True)

        if self.general:
            return general(n), False
        sel = self._select(n)
        if sel is None:
            return general(n), False
        sec, drop_sign = sel
        minus = n < 0 and not drop_sign
        out: Optional[render.Out]
        if sec.kind == "text":
            out = render.render_text(sec, format_general_fit(n, GENERAL_WIDTH) or "")
        elif sec.kind == "general":
            out = render.render_general(sec, abs(n), width - 1 if minus else width)
        elif sec.kind == "date":
            out = None if minus else render.render_date(sec, abs(n), sys)
        else:
            out = render.render_number(sec, abs(n))
        if out is None:
            return Formatted(OVERFLOW, sec.color, None, True), True
        if minus and sec.kind != "text":
            out.prepend("-")
        return Formatted(out.text, sec.color, out.fill, True), False

    def format_text(self, t: str) -> Formatted:
        if self.text_section is not None and self.text_section < len(self.sections):
            sec = self.sections[self.text_section]
            out = render.render_text(sec, t)
            return Formatted(out.text, sec.color, out.fill, False)
        return Formatted(t, None, None, False)


@lru_cache(maxsize=4096)
def parse(code: str) -> NumberFormat:
    """Parses (and caches) a format code. Never fails: malformed parts degrade to literal text."""
    return NumberFormat(code)


GENERAL = parse("General")


def format_value(v: Value, fmt: NumberFormat, sys: DateSystem = DateSystem.D1900, width: int = GENERAL_WIDTH) -> Formatted:
    """Formats a cell value for display."""
    if v is None:
        return Formatted()
    t = type(v)
    if t is float or t is int:
        if not math.isfinite(v):  # type: ignore[arg-type]
            return Formatted(CellError.NUM.value)
        return fmt.format_number(float(v), sys, width)[0]  # type: ignore[arg-type]
    if t is str:
        return fmt.format_text(v)  # type: ignore[arg-type]
    if t is bool:
        return Formatted("TRUE" if v else "FALSE")
    if t is CellError:
        return Formatted(v.value)  # type: ignore[union-attr]
    if isinstance(v, Array):
        first = v.data[0] if v.data else None
        if isinstance(first, Array) or first is None:
            return Formatted()
        return format_value(first, fmt, sys, width)
    return Formatted(str(v))


def text_function(v: Value, code: str, sys: DateSystem = DateSystem.D1900) -> str:
    """The ``TEXT()`` worksheet function. Raises :class:`XlError`."""
    if code == "":
        return ""
    fmt = parse(code)
    x = scalar(v)
    if isinstance(x, CellError):
        raise XlError(x)
    if x is None:
        n = 0.0
    elif isinstance(x, bool):
        return fmt.format_text("TRUE" if x else "FALSE").text
    elif isinstance(x, (int, float)):
        n = float(x)
    elif isinstance(x, str):
        p = parse_number_text(x)
        if p is None:
            return fmt.format_text(x).text
        n = p
    else:
        raise XlError(CellError.VALUE)
    if not math.isfinite(n):
        raise XlError(CellError.NUM)
    f, overflow = fmt.format_number(n, sys)
    if overflow:
        raise XlError(CellError.VALUE)
    return f.text


__all__ = [
    "Formatted",
    "NumberFormat",
    "builtin_format",
    "builtin_id",
    "color_hex",
    "format_general_fit",
    "format_value",
    "parse",
    "text_function",
]
