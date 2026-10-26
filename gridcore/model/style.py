"""Cell styles (font, fill, borders, alignment, number format, protection), interned per
workbook. Everything here is immutable (frozen dataclasses) so styles can be shared freely."""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Callable, Dict, List, Optional, Tuple

DEFAULT_FONT = "Calibri"
DEFAULT_FONT_SIZE = 11.0


@dataclass(frozen=True)
class Color:
    """A colour: ``auto`` (black text / no fill), explicit ``rgb`` (0xRRGGBB) or a ``theme`` slot
    with a tint in thousandths (-1000..=1000). Theme slots: 0 = Background 1 (lt1), 1 = Text 1
    (dk1), 2 = lt2, 3 = dk2, 4–9 = Accent 1–6, 10 hyperlink, 11 followed hyperlink."""

    kind: str = "auto"
    value: int = 0
    tint: int = 0

    @staticmethod
    def rgb(v: int) -> "Color":
        return Color("rgb", v & 0xFFFFFF)

    @staticmethod
    def theme(i: int, tint: int = 0) -> "Color":
        return Color("theme", i, tint)

    @staticmethod
    def from_hex(s: str) -> Optional["Color"]:
        h = s.strip().lstrip("#")
        if len(h) == 8:
            h = h[2:]
        if len(h) == 3:
            h = "".join(c * 2 for c in h)
        if len(h) != 6:
            return None
        try:
            return Color.rgb(int(h, 16))
        except ValueError:
            return None

    @property
    def is_auto(self) -> bool:
        return self.kind == "auto"

    def resolve(self, theme: "Theme") -> Optional[Tuple[int, int, int]]:
        if self.kind == "auto":
            return None
        if self.kind == "rgb":
            v = self.value
            return ((v >> 16) & 255, (v >> 8) & 255, v & 255)
        base = theme.colors[self.value] if 0 <= self.value < len(theme.colors) else 0
        return apply_tint(((base >> 16) & 255, (base >> 8) & 255, base & 255), self.tint / 1000.0)

    def hex(self, theme: "Theme") -> Optional[str]:
        rgb = self.resolve(theme)
        if rgb is None:
            return None
        return "#{:02X}{:02X}{:02X}".format(*rgb)

    def to_json(self) -> Optional[str]:
        """Compact text form: ``None`` (auto), ``#RRGGBB`` or ``theme:<i>:<tint>``."""
        if self.kind == "auto":
            return None
        if self.kind == "rgb":
            return f"#{self.value:06X}"
        return f"theme:{self.value}:{self.tint}"

    @staticmethod
    def from_json(s) -> "Color":
        if not s:
            return AUTO
        if isinstance(s, str) and s.startswith("theme:"):
            parts = s.split(":")
            try:
                return Color.theme(int(parts[1]), int(parts[2]) if len(parts) > 2 else 0)
            except (ValueError, IndexError):
                return AUTO
        c = Color.from_hex(str(s))
        return c if c is not None else AUTO


AUTO = Color()


def apply_tint(rgb: Tuple[int, int, int], tint: float) -> Tuple[int, int, int]:
    """Tint as in ECMA-376: lighten toward white (positive) or darken (negative) in HSL luminance."""
    if tint == 0.0:
        return rgb
    h, s, l = _rgb_to_hsl(rgb)
    l = l * (1.0 + tint) if tint < 0 else l * (1.0 - tint) + tint
    return _hsl_to_rgb(h, s, min(max(l, 0.0), 1.0))


def _rgb_to_hsl(rgb: Tuple[int, int, int]) -> Tuple[float, float, float]:
    r, g, b = (x / 255.0 for x in rgb)
    mx, mn = max(r, g, b), min(r, g, b)
    l = (mx + mn) / 2.0
    if mx == mn:
        return 0.0, 0.0, l
    d = mx - mn
    s = d / (2.0 - mx - mn) if l > 0.5 else d / (mx + mn)
    if mx == r:
        h = (g - b) / d + (6.0 if g < b else 0.0)
    elif mx == g:
        h = (b - r) / d + 2.0
    else:
        h = (r - g) / d + 4.0
    return h / 6.0, s, l


def _hsl_to_rgb(h: float, s: float, l: float) -> Tuple[int, int, int]:
    def rnd(x: float) -> int:
        return int(x * 255.0 + 0.5)

    if s == 0.0:
        v = rnd(l)
        return v, v, v
    q = l * (1.0 + s) if l < 0.5 else l + s - l * s
    p = 2.0 * l - q

    def hue(t: float) -> float:
        if t < 0:
            t += 1
        if t > 1:
            t -= 1
        if t < 1 / 6:
            return p + (q - p) * 6 * t
        if t < 0.5:
            return q
        if t < 2 / 3:
            return p + (q - p) * (2 / 3 - t) * 6
        return p

    return rnd(hue(h + 1 / 3)), rnd(hue(h)), rnd(hue(h - 1 / 3))


@dataclass(frozen=True)
class Theme:
    """Workbook theme: 12 colours (lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink) and the
    heading/body fonts. The default palette is ReflexGrid's own theme."""

    name: str = "ReflexGrid"
    colors: Tuple[int, ...] = (
        0xFFFFFF, 0x000000, 0xE8E8E8, 0x0E2841, 0x156082, 0xE97132,
        0x196B24, 0x0F9ED5, 0xA02B93, 0x4EA72E, 0x467886, 0x96607D,
    )
    major_font: str = DEFAULT_FONT
    minor_font: str = DEFAULT_FONT
    # The heading and body fonts' East Asian (<a:ea>) and complex-script (<a:cs>) typefaces, and
    # their per-script typefaces (<a:font script="Arab" typeface="…"/>, as (script, typeface) in
    # file order); empty when the file did not set them. A Japanese workbook's theme fonts
    # (游ゴシック, ＭＳ Ｐゴシック) used to be deleted by an open and save.
    major_font_ea: str = ""
    minor_font_ea: str = ""
    major_font_cs: str = ""
    minor_font_cs: str = ""
    major_font_scripts: Tuple[Tuple[str, str], ...] = ()
    minor_font_scripts: Tuple[Tuple[str, str], ...] = ()


@dataclass(frozen=True)
class Font:
    name: str = DEFAULT_FONT
    size: float = DEFAULT_FONT_SIZE
    bold: bool = False
    italic: bool = False
    underline: str = "none"  # none single double singleAccounting doubleAccounting
    strike: bool = False
    color: Color = AUTO
    vert: str = "baseline"  # baseline superscript subscript
    # Windows character set the font was chosen for (178 = Arabic), kept so Excel substitutes a
    # font for the same script; None = the file gave none.
    charset: Optional[int] = None
    # Font family class (1 roman, 2 swiss, 3 modern…); None = swiss, the default.
    family: Optional[int] = None
    # The font follows the theme's body ("minor") or heading ("major") font while its face is that
    # font's (`<scheme val="minor"/>`); "" = no link. New fonts follow the body font, as Excel's do,
    # and fonts derived from them (bold, a colour, a size) keep the link.
    scheme: str = "minor"


@dataclass(frozen=True)
class Fill:
    pattern: str = "none"  # none solid mediumGray darkGray lightGray gray125 gray0625 darkHorizontal …
    fg: Color = AUTO
    bg: Color = AUTO

    @staticmethod
    def solid(c: Color) -> "Fill":
        return Fill("solid", c, AUTO)


BORDER_WIDTHS = {
    "none": 0,
    "hair": 1,
    "thin": 1,
    "dashed": 1,
    "dotted": 1,
    "dashDot": 1,
    "dashDotDot": 1,
    "medium": 2,
    "mediumDashed": 2,
    "mediumDashDot": 2,
    "mediumDashDotDot": 2,
    "slantDashDot": 2,
    "thick": 3,
    "double": 3,
}


@dataclass(frozen=True)
class BorderLine:
    style: str = "none"
    color: Color = AUTO

    @property
    def is_none(self) -> bool:
        return self.style == "none"


NO_LINE = BorderLine()


@dataclass(frozen=True)
class Borders:
    left: BorderLine = NO_LINE
    right: BorderLine = NO_LINE
    top: BorderLine = NO_LINE
    bottom: BorderLine = NO_LINE
    diag_down: BorderLine = NO_LINE
    diag_up: BorderLine = NO_LINE


@dataclass(frozen=True)
class Alignment:
    h: str = "general"  # general left center right fill justify centerContinuous distributed
    v: str = "bottom"  # top center bottom justify distributed
    wrap: bool = False
    shrink: bool = False
    indent: int = 0
    # Degrees -90..=90, or 255 for vertical stacked text.
    rotation: int = 0


@dataclass(frozen=True)
class Protection:
    locked: bool = True
    hidden: bool = False


@dataclass(frozen=True)
class Style:
    font: Font = Font()
    fill: Fill = Fill()
    borders: Borders = Borders()
    align: Alignment = Alignment()
    num_fmt: str = "General"
    protection: Protection = Protection()


DEFAULT_STYLE = Style()


class StyleTable:
    """Interned styles. Index 0 is the default (Normal) style."""

    __slots__ = ("_styles", "_index")

    def __init__(self, default: Style = DEFAULT_STYLE) -> None:
        self._styles: List[Style] = [default]
        self._index: Dict[Style, int] = {default: 0}

    def get(self, sid: int) -> Style:
        if 0 <= sid < len(self._styles):
            return self._styles[sid]
        return self._styles[0]

    def intern(self, s: Style) -> int:
        i = self._index.get(s)
        if i is not None:
            return i
        i = len(self._styles)
        self._styles.append(s)
        self._index[s] = i
        return i

    def derive(self, sid: int, f: Callable[[Style], Style]) -> int:
        """The style ``sid`` changed by ``f`` (which returns a new Style), interned."""
        return self.intern(f(self.get(sid)))

    def __len__(self) -> int:
        return len(self._styles)

    def items(self):
        return enumerate(self._styles)

    def clone(self) -> "StyleTable":
        t = StyleTable.__new__(StyleTable)
        t._styles = list(self._styles)
        t._index = dict(self._index)
        return t


def font_with(st: Style, **kw) -> Style:
    return replace(st, font=replace(st.font, **kw))


def align_with(st: Style, **kw) -> Style:
    return replace(st, align=replace(st.align, **kw))
