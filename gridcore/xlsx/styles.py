"""``styles.xml``: reading cell formats and differential formats into model styles, and writing
the workbook's style table."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

from ..model.style import (
    AUTO,
    Alignment,
    BorderLine,
    Borders,
    Color,
    Fill,
    Font,
    Protection,
    Style,
)
from .tables import INDEXED, builtin_format, builtin_id
from .xml import XML_DECL, attr, attr_float, attr_int, child, child_val, esc_attr, first_child, flag, kids, num, path

# ---------------------------------------------------------------- reading


@dataclass
class StylesIn:
    xfs: List[Style] = field(default_factory=list)
    dxfs: List[Style] = field(default_factory=list)
    palette: List[int] = field(default_factory=lambda: list(INDEXED))
    named: List[Tuple[str, Style]] = field(default_factory=list)
    num_fmts: Dict[int, str] = field(default_factory=dict)


def read_color(e: ET.Element, palette: List[int]) -> Color:
    if flag(e, "auto", False):
        return AUTO
    rgb = attr(e, "rgb")
    if rgb is not None:
        c = Color.from_hex(rgb)
        return c if c is not None else AUTO
    t = attr_int(e, "theme")
    if t is not None:
        if not 0 <= t <= 11:
            return AUTO
        tint = attr_float(e, "tint") or 0.0
        return Color.theme(t, int(max(min(round(tint * 1000), 1000), -1000)))
    i = attr_int(e, "indexed")
    if i is not None:
        if 0 <= i < len(palette):
            return Color.rgb(palette[i])
        return AUTO
    return AUTO


def _opt_color(e: Optional[ET.Element], palette: List[int]) -> Color:
    return read_color(e, palette) if e is not None else AUTO


def _bool_prop(e: ET.Element, name: str) -> bool:
    c = child(e, name)
    return c is not None and flag(c, "val", True)


def read_font(e: ET.Element, palette: List[int]) -> Font:
    name = None
    n = first_child(e, "name", "rFont")
    if n is not None:
        name = attr(n, "val")
    size = None
    sz = child(e, "sz")
    if sz is not None:
        size = attr_float(sz, "val")
    underline = "none"
    u = child(e, "u")
    if u is not None:
        val = attr(u, "val") or "single"
        underline = val if val in ("double", "singleAccounting", "doubleAccounting", "none") else "single"
    vert = child_val(e, "vertAlign") or "baseline"
    if vert not in ("superscript", "subscript"):
        vert = "baseline"
    color = _opt_color(child(e, "color"), palette)
    # Text 1 (dk1) is what "automatic" font colour means in practice.
    if color == Color.theme(1, 0):
        color = AUTO
    f = Font()
    return Font(
        name=name or f.name,
        size=float(min(max(size, 1.0), 409.0)) if size is not None else f.size,
        bold=_bool_prop(e, "b"),
        italic=_bool_prop(e, "i"),
        underline=underline,
        strike=_bool_prop(e, "strike"),
        color=color,
        vert=vert,
    )


_PATTERNS = {
    "solid", "mediumGray", "darkGray", "lightGray", "gray125", "gray0625", "darkHorizontal", "darkVertical", "darkDown",
    "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis",
}


def read_fill(e: ET.Element, palette: List[int], dxf: bool) -> Fill:
    p = child(e, "patternFill")
    if p is not None:
        fg = _opt_color(child(p, "fgColor"), palette)
        bg = _opt_color(child(p, "bgColor"), palette)
        pt = attr(p, "patternType")
        if dxf and (pt is None or pt == "solid"):
            c = bg if not bg.is_auto else fg
            if c.is_auto and pt is None:
                return Fill()
            return Fill.solid(c)
        pattern = pt if pt in _PATTERNS else "none"
        if pattern == "none":
            return Fill()
        return Fill(pattern, fg, bg)
    g = child(e, "gradientFill")
    if g is not None:
        for stop in kids(g, "stop"):
            c = child(stop, "color")
            if c is not None:
                return Fill.solid(read_color(c, palette))
    return Fill()


_BORDER_STYLES = {
    "thin", "medium", "thick", "dashed", "dotted", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot",
    "dashDotDot", "mediumDashDotDot", "slantDashDot",
}


def _read_line(e: Optional[ET.Element], palette: List[int]) -> BorderLine:
    if e is None:
        return BorderLine()
    st = attr(e, "style") or "none"
    if st not in _BORDER_STYLES:
        return BorderLine()
    return BorderLine(st, _opt_color(child(e, "color"), palette))


def read_border(e: ET.Element, palette: List[int]) -> Borders:
    diag = _read_line(child(e, "diagonal"), palette)
    return Borders(
        left=_read_line(first_child(e, "left", "start"), palette),
        right=_read_line(first_child(e, "right", "end"), palette),
        top=_read_line(child(e, "top"), palette),
        bottom=_read_line(child(e, "bottom"), palette),
        diag_down=diag if flag(e, "diagonalDown", False) else BorderLine(),
        diag_up=diag if flag(e, "diagonalUp", False) else BorderLine(),
    )


def _read_alignment(a: ET.Element) -> Alignment:
    h = attr(a, "horizontal") or "general"
    if h not in ("left", "center", "right", "fill", "justify", "centerContinuous", "distributed"):
        h = "general"
    v = attr(a, "vertical") or "bottom"
    if v not in ("top", "center", "justify", "distributed"):
        v = "bottom"
    rot = attr_int(a, "textRotation") or 0
    if 0 <= rot <= 90:
        rotation = rot
    elif 91 <= rot <= 180:
        rotation = -(rot - 90)
    elif rot == 255:
        rotation = 255
    else:
        rotation = 0
    return Alignment(
        h=h,
        v=v,
        wrap=flag(a, "wrapText", False),
        shrink=flag(a, "shrinkToFit", False),
        indent=min(max(attr_int(a, "indent") or 0, 0), 250),
        rotation=rotation,
    )


def _read_protection(p: ET.Element) -> Protection:
    return Protection(locked=flag(p, "locked", True), hidden=flag(p, "hidden", False))


def _num_fmt_code(fid: int, custom: Dict[int, str]) -> str:
    if fid in custom:
        return custom[fid]
    return builtin_format(fid) or "General"


def _read_xf(xf: ET.Element, fonts, fills, borders, fmts) -> Style:
    fi = attr_int(xf, "fontId")
    font = fonts[fi] if fi is not None and 0 <= fi < len(fonts) else (fonts[0] if fonts else Font())
    li = attr_int(xf, "fillId")
    fill = fills[li] if li is not None and 0 <= li < len(fills) else Fill()
    bi = attr_int(xf, "borderId")
    border = borders[bi] if bi is not None and 0 <= bi < len(borders) else Borders()
    a = child(xf, "alignment")
    p = child(xf, "protection")
    return Style(
        font=font,
        fill=fill,
        borders=border,
        align=_read_alignment(a) if a is not None else Alignment(),
        num_fmt=_num_fmt_code(attr_int(xf, "numFmtId") or 0, fmts),
        protection=_read_protection(p) if p is not None else Protection(),
    )


def read_styles(root: ET.Element) -> StylesIn:
    out = StylesIn()
    ix = path(root, "colors", "indexedColors")
    if ix is not None:
        custom = []
        for c in kids(ix, "rgbColor"):
            col = read_color(c, INDEXED)
            custom.append(col.value if col.kind == "rgb" else 0)
        if custom:
            out.palette = custom
    nf = child(root, "numFmts")
    if nf is not None:
        for f in kids(nf, "numFmt"):
            i, code = attr_int(f, "numFmtId"), attr(f, "formatCode")
            if i is not None and code is not None:
                out.num_fmts[i] = code
    fonts_el, fills_el, borders_el = child(root, "fonts"), child(root, "fills"), child(root, "borders")
    fonts = [read_font(e, out.palette) for e in kids(fonts_el, "font")] if fonts_el is not None else []
    fills = [read_fill(e, out.palette, False) for e in kids(fills_el, "fill")] if fills_el is not None else []
    borders = [read_border(e, out.palette) for e in kids(borders_el, "border")] if borders_el is not None else []
    cx = child(root, "cellXfs")
    if cx is not None:
        out.xfs = [_read_xf(x, fonts, fills, borders, out.num_fmts) for x in kids(cx, "xf")]
    csx = child(root, "cellStyleXfs")
    style_xfs = [_read_xf(x, fonts, fills, borders, out.num_fmts) for x in kids(csx, "xf")] if csx is not None else []
    cs = child(root, "cellStyles")
    if cs is not None:
        for c in kids(cs, "cellStyle"):
            name, xi = attr(c, "name"), attr_int(c, "xfId")
            if name and name != "Normal" and xi is not None and 0 <= xi < len(style_xfs):
                out.named.append((name, style_xfs[xi]))
    dx = child(root, "dxfs")
    if dx is not None:
        for x in kids(dx, "dxf"):
            st = Style()
            f = child(x, "font")
            font = st.font
            if f is not None:
                rf = read_font(f, out.palette)
                # Differential fonts rarely name a face; keep the default face and size.
                font = Font(st.font.name, st.font.size, rf.bold, rf.italic, rf.underline, rf.strike, rf.color, rf.vert)
            num_fmt = st.num_fmt
            n = child(x, "numFmt")
            if n is not None:
                code = attr(n, "formatCode")
                if code is not None:
                    num_fmt = code
                else:
                    fid = attr_int(n, "numFmtId")
                    if fid is not None:
                        num_fmt = _num_fmt_code(fid, out.num_fmts)
            fl = child(x, "fill")
            b = child(x, "border")
            a = child(x, "alignment")
            p = child(x, "protection")
            out.dxfs.append(
                Style(
                    font=font,
                    fill=read_fill(fl, out.palette, True) if fl is not None else Fill(),
                    borders=read_border(b, out.palette) if b is not None else Borders(),
                    align=_read_alignment(a) if a is not None else Alignment(),
                    num_fmt=num_fmt,
                    protection=_read_protection(p) if p is not None else Protection(),
                )
            )
    return out


# ---------------------------------------------------------------- writing


def color_attrs(c: Color) -> str:
    if c.kind == "auto":
        return 'auto="1"'
    if c.kind == "rgb":
        return f'rgb="FF{c.value & 0xFFFFFF:06X}"'
    if c.tint == 0:
        return f'theme="{c.value}"'
    return f'theme="{c.value}" tint="{num(c.tint / 1000.0)}"'


def _font_xml(f: Font, dxf: bool) -> str:
    s = ["<font>"]
    if f.bold:
        s.append("<b/>")
    if f.italic:
        s.append("<i/>")
    if f.strike:
        s.append("<strike/>")
    if f.underline == "single":
        s.append("<u/>")
    elif f.underline in ("double", "singleAccounting", "doubleAccounting"):
        s.append(f'<u val="{f.underline}"/>')
    if f.vert in ("superscript", "subscript"):
        s.append(f'<vertAlign val="{f.vert}"/>')
    if not dxf:
        s.append(f'<sz val="{num(f.size)}"/>')
    if not f.color.is_auto:
        s.append(f"<color {color_attrs(f.color)}/>")
    elif not dxf:
        s.append('<color theme="1"/>')
    if not dxf:
        s.append(f'<name val="{esc_attr(f.name)}"/><family val="2"/>')
        if f.name == "Calibri":
            s.append('<scheme val="minor"/>')
    s.append("</font>")
    return "".join(s)


def _fill_xml(f: Fill, dxf: bool) -> str:
    if f.pattern == "none":
        return '<fill><patternFill patternType="none"/></fill>'
    if dxf and f.pattern == "solid":
        c = color_attrs(f.fg)
        return f'<fill><patternFill patternType="solid"><fgColor {c}/><bgColor {c}/></patternFill></fill>'
    s = [f'<fill><patternFill patternType="{f.pattern}">']
    if not f.fg.is_auto:
        s.append(f"<fgColor {color_attrs(f.fg)}/>")
    elif f.pattern != "gray125":
        s.append('<fgColor indexed="64"/>')
    if not f.bg.is_auto:
        s.append(f"<bgColor {color_attrs(f.bg)}/>")
    elif f.pattern != "gray125":
        s.append('<bgColor indexed="64"/>')
    s.append("</patternFill></fill>")
    return "".join(s)


def _line_xml(tag: str, l: BorderLine) -> str:
    if l.is_none:
        return f"<{tag}/>"
    color = '<color indexed="64"/>' if l.color.is_auto else f"<color {color_attrs(l.color)}/>"
    return f'<{tag} style="{l.style}">{color}</{tag}>'


def _border_xml(b: Borders) -> str:
    s = ["<border"]
    if not b.diag_up.is_none:
        s.append(' diagonalUp="1"')
    if not b.diag_down.is_none:
        s.append(' diagonalDown="1"')
    s.append(">")
    s.append(_line_xml("left", b.left))
    s.append(_line_xml("right", b.right))
    s.append(_line_xml("top", b.top))
    s.append(_line_xml("bottom", b.bottom))
    s.append(_line_xml("diagonal", b.diag_down if not b.diag_down.is_none else b.diag_up))
    s.append("</border>")
    return "".join(s)


def _alignment_xml(a: Alignment) -> Optional[str]:
    if a == Alignment():
        return None
    s = ["<alignment"]
    if a.h != "general":
        s.append(f' horizontal="{a.h}"')
    if a.v != "bottom":
        s.append(f' vertical="{a.v}"')
    if a.rotation == 255:
        rot = 255
    elif 0 <= a.rotation <= 90:
        rot = a.rotation
    elif -90 <= a.rotation < 0:
        rot = 90 - a.rotation
    else:
        rot = 0
    if rot:
        s.append(f' textRotation="{rot}"')
    if a.wrap:
        s.append(' wrapText="1"')
    if a.indent > 0:
        s.append(f' indent="{a.indent}"')
    if a.shrink:
        s.append(' shrinkToFit="1"')
    s.append("/>")
    return "".join(s)


def _protection_xml(p: Protection) -> Optional[str]:
    if p == Protection():
        return None
    return f'<protection locked="{int(p.locked)}" hidden="{int(p.hidden)}"/>'


class _Interner:
    def __init__(self, initial=()):
        self.items: list = list(initial)
        self.index = {x: i for i, x in enumerate(self.items)}

    def id(self, x) -> int:
        i = self.index.get(x)
        if i is None:
            i = len(self.items)
            self.items.append(x)
            self.index[x] = i
        return i


class _NumFmts:
    def __init__(self) -> None:
        self.custom: List[Tuple[int, str]] = []
        self.map: Dict[str, int] = {}

    def id(self, code: str) -> int:
        b = builtin_id(code)
        if b is not None:
            return b
        if code in self.map:
            return self.map[code]
        i = 164 + len(self.custom)
        self.custom.append((i, code))
        self.map[code] = i
        return i


def write_styles(wb, dxfs: List[Style]) -> str:
    """``styles.xml``: one cell format per model style, in style-id order (so a cell's ``s``
    attribute is its style id), plus the given differential formats."""
    nf = _NumFmts()
    fonts = _Interner([wb.styles.get(0).font])
    fills = _Interner([Fill(), Fill("gray125")])
    borders = _Interner([Borders()])
    xfs = []
    for _, st in wb.styles.items():
        font = fonts.id(st.font)
        fill = fills.id(st.fill)
        border = borders.id(st.borders)
        fmt = 0 if st.num_fmt == "checkbox" else nf.id(st.num_fmt)
        x = [f'<xf numFmtId="{fmt}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0"']
        if fmt:
            x.append(' applyNumberFormat="1"')
        if font:
            x.append(' applyFont="1"')
        if fill:
            x.append(' applyFill="1"')
        if border:
            x.append(' applyBorder="1"')
        al = _alignment_xml(st.align)
        pr = _protection_xml(st.protection)
        if al:
            x.append(' applyAlignment="1"')
        if pr:
            x.append(' applyProtection="1"')
        if not al and not pr:
            x.append("/>")
        else:
            x.append(">" + (al or "") + (pr or "") + "</xf>")
        xfs.append("".join(x))
    style_xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>']
    cell_styles = ['<cellStyle name="Normal" xfId="0" builtinId="0"/>']
    named = 1
    for name, st in wb.cell_styles:
        if not name or name == "Normal" or not isinstance(st, Style):
            continue
        font, fill, border = fonts.id(st.font), fills.id(st.fill), borders.id(st.borders)
        fmt = nf.id(st.num_fmt)
        al, pr = _alignment_xml(st.align), _protection_xml(st.protection)
        if not al and not pr:
            style_xfs.append(f'<xf numFmtId="{fmt}" fontId="{font}" fillId="{fill}" borderId="{border}"/>')
        else:
            style_xfs.append(f'<xf numFmtId="{fmt}" fontId="{font}" fillId="{fill}" borderId="{border}">{al or ""}{pr or ""}</xf>')
        cell_styles.append(f'<cellStyle name="{esc_attr(name)}" xfId="{named}" customBuiltin="1"/>')
        named += 1
    dx = []
    default = Style()
    for d in dxfs:
        parts = ["<dxf>"]
        f = d.font
        if f.bold or f.italic or f.strike or f.underline != "none" or not f.color.is_auto or f.vert != "baseline":
            parts.append(_font_xml(f, True))
        if d.num_fmt != default.num_fmt:
            parts.append(f'<numFmt numFmtId="{nf.id(d.num_fmt)}" formatCode="{esc_attr(d.num_fmt)}"/>')
        if d.fill != default.fill:
            parts.append(_fill_xml(d.fill, True))
        al = _alignment_xml(d.align)
        if al:
            parts.append(al)
        if d.borders != default.borders:
            parts.append(_border_xml(d.borders))
        parts.append("</dxf>")
        dx.append("".join(parts))
    s = [XML_DECL, '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
    if nf.custom:
        s.append(f'<numFmts count="{len(nf.custom)}">')
        for i, code in nf.custom:
            s.append(f'<numFmt numFmtId="{i}" formatCode="{esc_attr(code)}"/>')
        s.append("</numFmts>")
    s.append(f'<fonts count="{len(fonts.items)}">' + "".join(_font_xml(f, False) for f in fonts.items) + "</fonts>")
    s.append(f'<fills count="{len(fills.items)}">' + "".join(_fill_xml(f, False) for f in fills.items) + "</fills>")
    s.append(f'<borders count="{len(borders.items)}">' + "".join(_border_xml(b) for b in borders.items) + "</borders>")
    s.append(f'<cellStyleXfs count="{named}">' + "".join(style_xfs) + "</cellStyleXfs>")
    s.append(f'<cellXfs count="{len(xfs)}">' + "".join(xfs) + "</cellXfs>")
    s.append(f'<cellStyles count="{named}">' + "".join(cell_styles) + "</cellStyles>")
    s.append(f'<dxfs count="{len(dx)}">' + "".join(dx) + "</dxfs>")
    s.append('<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>')
    return "".join(s)


def dxf_id(dxfs: List[Style], st: Style) -> int:
    for i, d in enumerate(dxfs):
        if d == st:
            return i
    dxfs.append(st)
    return len(dxfs) - 1
