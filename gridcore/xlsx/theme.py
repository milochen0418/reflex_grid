"""``theme1.xml``: colour scheme and font scheme.

In the file the colour scheme is ordered dk1, lt1, dk2, lt2, accent1–6, hlink, folHlink, while
cell colour ``theme`` indices (and the model's ``Theme.colors``) use lt1, dk1, lt2, dk2, …."""

from __future__ import annotations

from typing import Optional
from xml.etree import ElementTree as ET

from ..model.style import Theme
from .xml import XML_DECL, attr, child, esc_attr, path

SLOTS = [
    ("dk1", 1), ("lt1", 0), ("dk2", 3), ("lt2", 2), ("accent1", 4), ("accent2", 5), ("accent3", 6),
    ("accent4", 7), ("accent5", 8), ("accent6", 9), ("hlink", 10), ("folHlink", 11),
]


def _scheme_rgb(e: ET.Element) -> Optional[int]:
    c = child(e, "srgbClr")
    if c is not None and attr(c, "val"):
        try:
            return int(attr(c, "val").strip(), 16) & 0xFFFFFF
        except ValueError:
            return None
    c = child(e, "sysClr")
    if c is not None:
        last = attr(c, "lastClr")
        if last:
            try:
                return int(last.strip(), 16) & 0xFFFFFF
            except ValueError:
                pass
        v = attr(c, "val")
        if v == "windowText":
            return 0x000000
        if v == "window":
            return 0xFFFFFF
    return None


def read_theme(root: ET.Element) -> Theme:
    t = Theme()
    name = attr(root, "name") or t.name
    colors = list(t.colors)
    major, minor = t.major_font, t.minor_font
    el = child(root, "themeElements")
    if el is not None:
        cs = child(el, "clrScheme")
        if cs is not None:
            for slot, idx in SLOTS:
                c = child(cs, slot)
                if c is not None:
                    v = _scheme_rgb(c)
                    if v is not None:
                        colors[idx] = v
        fs = child(el, "fontScheme")
        if fs is not None:
            mj = path(fs, "majorFont", "latin")
            if mj is not None and attr(mj, "typeface"):
                major = attr(mj, "typeface")
            mn = path(fs, "minorFont", "latin")
            if mn is not None and attr(mn, "typeface"):
                minor = attr(mn, "typeface")
    return Theme(name, tuple(colors), major, minor)


def write_theme(t: Theme) -> str:
    """A minimal, schema-complete theme part (an original palette, not Office's)."""
    name = esc_attr(t.name or "ReflexGrid")
    s = [XML_DECL, f'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="{name}"><a:themeElements>']
    s.append(f'<a:clrScheme name="{name}">')
    for slot, idx in SLOTS:
        rgb = t.colors[idx] & 0xFFFFFF
        s.append(f'<a:{slot}><a:srgbClr val="{rgb:06X}"/></a:{slot}>')
    s.append("</a:clrScheme>")
    s.append(
        f'<a:fontScheme name="{name}"><a:majorFont><a:latin typeface="{esc_attr(t.major_font)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
        f'<a:minorFont><a:latin typeface="{esc_attr(t.minor_font)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
    )
    fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    line = '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    s.append(
        f'<a:fmtScheme name="{name}"><a:fillStyleLst>{fill}{fill}{fill}</a:fillStyleLst>'
        f"<a:lnStyleLst>{line}{line}{line}</a:lnStyleLst>"
        "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>"
        f"<a:bgFillStyleLst>{fill}{fill}{fill}</a:bgFillStyleLst></a:fmtScheme>"
    )
    s.append("</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>")
    return "".join(s)
