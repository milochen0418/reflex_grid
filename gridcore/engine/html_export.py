"""A sheet as a styled HTML page: for printing (with the sheet's page setup) and File › Save as
Web Page. Formatting, merges, hidden rows/columns, conditional formats and table styles carry
over; floating objects are not drawn."""

from __future__ import annotations

import html
from datetime import datetime
from typing import Dict, List, Optional

from ..core.addr import RangeRef, a1, col_to_letters
from ..core.value import CellError
from ..model.style import Style
from ..numfmt import color_hex, format_value, parse as parse_fmt
from .cf import CfCache
from .tables import table_style_colors

_BORDER = {
    "thin": "1px solid",
    "hair": "1px dotted",
    "dotted": "1px dotted",
    "dashed": "1px dashed",
    "dashDot": "1px dashed",
    "dashDotDot": "1px dashed",
    "medium": "2px solid",
    "mediumDashed": "2px dashed",
    "mediumDashDot": "2px dashed",
    "mediumDashDotDot": "2px dashed",
    "slantDashDot": "2px dashed",
    "thick": "3px solid",
    "double": "3px double",
}
PAPER_CSS = {"Letter": "letter", "Legal": "legal", "A4": "A4", "A3": "A3", "A5": "A5", "B4": "B4", "B5": "B5", "Tabloid": "11in 17in", "Ledger": "17in 11in", "Executive": "7.25in 10.5in"}


def _font_family(name: str) -> str:
    if name.lower().startswith("calibri"):
        return "Calibri, Carlito, 'Segoe UI', Arial, sans-serif"
    return f"'{name}', Calibri, Carlito, Arial, sans-serif"


def _css(st: Style, theme, look=None, tx: Optional[dict] = None, numeric: bool = False, kind: str = "") -> str:
    f = st.font
    out = [f"font-family:{_font_family(f.name)}", f"font-size:{f.size:g}pt"]
    bold = f.bold or (look is not None and look.style is not None and look.style.font.bold) or bool(tx and tx.get("b"))
    if bold:
        out.append("font-weight:bold")
    if f.italic or (look is not None and look.style is not None and look.style.font.italic):
        out.append("font-style:italic")
    deco = []
    if f.underline != "none":
        deco.append("underline")
    if f.strike:
        deco.append("line-through")
    if deco:
        out.append("text-decoration:" + " ".join(deco))
    color = None
    if look is not None and look.style is not None and not look.style.font.color.is_auto:
        color = look.style.font.color.hex(theme)
    if color is None and not f.color.is_auto:
        color = f.color.hex(theme)
    if color is None and tx and tx.get("c"):
        color = tx["c"]
    if color:
        out.append(f"color:{color}")
    bg = None
    if look is not None and look.scale_fill is not None:
        bg = "#{:02X}{:02X}{:02X}".format(*look.scale_fill)
    elif look is not None and look.style is not None and look.style.fill.pattern != "none":
        bg = look.style.fill.fg.hex(theme) or look.style.fill.bg.hex(theme)
    elif st.fill.pattern != "none":
        bg = st.fill.fg.hex(theme) or st.fill.bg.hex(theme)
    elif tx and tx.get("bg"):
        bg = tx["bg"]
    if bg:
        out.append(f"background:{bg}")
    b = st.borders
    for side, line in (("top", b.top), ("right", b.right), ("bottom", b.bottom), ("left", b.left)):
        if not line.is_none:
            out.append(f"border-{side}:{_BORDER.get(line.style, '1px solid')} {line.color.hex(theme) or '#000000'}")
        elif tx and tx.get({"top": "bt", "right": "br", "bottom": "bb", "left": "bl"}[side]):
            spec = tx[{"top": "bt", "right": "br", "bottom": "bb", "left": "bl"}[side]]
            out.append(f"border-{side}:{_BORDER.get(spec[0], '1px solid')} {spec[1]}")
    a = st.align
    h = a.h
    if h == "general":
        h = "right" if numeric else ("center" if kind in ("bool", "error") else "left")
    elif h in ("centerContinuous", "distributed"):
        h = "center"
    elif h == "fill":
        h = "left"
    out.append(f"text-align:{h}")
    out.append("vertical-align:" + {"top": "top", "center": "middle"}.get(a.v, "bottom"))
    out.append("white-space:" + ("pre-wrap" if a.wrap else "pre"))
    if a.indent:
        out.append(f"padding-{'right' if h == 'right' else 'left'}:{3 + a.indent * 9}px")
    return ";".join(out)


def sheet_html(wb, si: int, rng: Optional[RangeRef] = None, title: str = "", for_print: bool = True) -> str:
    sh = wb.sheet(si)
    pr = sh.print
    if rng is None:
        rng = pr.print_area
    if rng is None:
        used = sh.used_range()
        rng = RangeRef(0, 0, used.r1, used.c1) if used is not None else RangeRef(0, 0, 0, 0)
    rng = RangeRef(rng.r0, rng.c0, min(rng.r1, rng.r0 + 20000), min(rng.c1, rng.c0 + 500))
    theme = wb.theme
    cf = CfCache(wb, si) if sh.cond_formats else None
    looks: Dict[int, dict] = {t.id: table_style_colors(t.style, theme) for t in sh.tables}
    covered = set()
    spans: Dict[tuple, tuple] = {}
    for m in sh.merges:
        if not m.intersects(rng):
            continue
        spans[m.start] = (m.height, m.width)
        for rc in m.cells():
            if rc != m.start:
                covered.add(rc)
    cols = [c for c in range(rng.c0, rng.c1 + 1) if not sh.is_col_hidden(c)]
    rows = [r for r in range(rng.r0, rng.r1 + 1) if not sh.is_row_hidden(r)]
    grid = "1px solid #D0D0D0" if pr.gridlines or not for_print else None
    parts: List[str] = []
    parts.append("<table class='sheet'>")
    parts.append("<colgroup>" + ("<col style='width:32px'>" if pr.headings else "") + "".join(f"<col style='width:{sh.col_width(c):g}px'>" for c in cols) + "</colgroup>")
    if pr.headings:
        parts.append("<tr><th></th>" + "".join(f"<th>{col_to_letters(c)}</th>" for c in cols) + "</tr>")
    from .viewport import _table_cell_look

    for r in rows:
        cells = [f"<tr style='height:{sh.row_height(r):g}px'>"]
        if pr.headings:
            cells.append(f"<th>{r + 1}</th>")
        for c in cols:
            rc = (r, c)
            if rc in covered:
                continue
            st = wb.styles.get(sh.style_id(rc))
            v = sh.value(rc)
            look = cf.look(rc) if cf is not None else None
            fmt_code = st.num_fmt
            if look is not None and look.style is not None and look.style.num_fmt != "General":
                fmt_code = look.style.num_fmt
            fm = format_value(v, parse_fmt(fmt_code), wb.date_system, 255) if v is not None else None
            text = fm.text if fm is not None else ""
            if isinstance(v, float) and v == 0 and not sh.show_zeros:
                text = ""
            if look is not None and look.hide_value:
                text = ""
            tx = None
            for t in sh.tables:
                if t.range.contains(rc):
                    tx = _table_cell_look(t, rc, looks[t.id], st)
                    break
            kind = "bool" if isinstance(v, bool) else ("error" if isinstance(v, CellError) else "")
            css = _css(st, theme, look, tx, isinstance(v, float), kind)
            if fm is not None and fm.color:
                css += f";color:{color_hex(fm.color)}"
            if grid and "border" not in css:
                css += f";border:{grid}"
            span = spans.get(rc)
            attrs = f" rowspan='{span[0]}' colspan='{span[1]}'" if span else ""
            if fm is not None and fm.fill is not None and not st.align.wrap:
                pos = fm.fill[1]
                inner = f"<span class='f'><span>{html.escape(text[:pos])}</span><span>{html.escape(text[pos:])}</span></span>"
            else:
                inner = html.escape(text)
            cells.append(f"<td style=\"{css}\"{attrs}>{inner}</td>")
        cells.append("</tr>")
        parts.append("".join(cells))
    parts.append("</table>")
    body = "\n".join(parts)
    paper = PAPER_CSS.get(pr.paper, "letter")
    left, right, top, bottom, _hdr, _ftr = pr.margins
    page_css = f"@page {{ size: {paper} {pr.orientation}; margin: {top}in {right}in {bottom}in {left}in; }}"
    if pr.fit_to is not None or pr.scale != 100:
        zoom = pr.scale / 100.0 if pr.fit_to is None else 1.0
        page_css += f" .sheet {{ zoom: {zoom:g}; }}"
    header = _header_footer(pr.header, sh.name, title)
    footer = _header_footer(pr.footer, sh.name, title)
    center = "margin-left:auto;margin-right:auto;" if pr.center_h else ""
    doc_title = html.escape(f"{title} - {sh.name}" if title else sh.name)
    return f"""<!doctype html>
<html><head><meta charset="utf-8"><title>{doc_title}</title>
<style>
{page_css}
body {{ margin: 0; font-family: Calibri, Carlito, 'Segoe UI', Arial, sans-serif; }}
.sheet {{ border-collapse: collapse; table-layout: fixed; {center} }}
.sheet td {{ padding: 0 3px; overflow: hidden; line-height: 1.15; }}
.sheet th {{ background: #F3F4F6; border: 1px solid #C0C0C0; font: 9pt Arial, sans-serif; color: #444; }}
.sheet .f {{ display: flex; justify-content: space-between; }}
.hf {{ text-align: center; font: 9pt Calibri, Carlito, Arial, sans-serif; margin: 4px 0; }}
</style></head><body>
{f'<div class="hf">{header}</div>' if header else ''}
{body}
{f'<div class="hf">{footer}</div>' if footer else ''}
</body></html>"""


def _header_footer(text: str, sheet: str, title: str) -> str:
    """Header/footer codes we can render statically (&A sheet, &F file, &D date, &T time)."""
    if not text:
        return ""
    now = datetime.now()
    out = []
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == "&" and i + 1 < len(text):
            code = text[i + 1]
            i += 2
            if code == "A":
                out.append(html.escape(sheet))
            elif code == "F":
                out.append(html.escape(title))
            elif code == "D":
                out.append(now.strftime("%m/%d/%Y"))
            elif code == "T":
                out.append(now.strftime("%H:%M"))
            elif code == "&":
                out.append("&amp;")
            elif code in "LCR":
                out.append(" ")
            elif code == '"':
                end = text.find('"', i)
                i = end + 1 if end >= 0 else len(text)
            continue
        out.append(html.escape(ch))
        i += 1
    return "".join(out).strip()
