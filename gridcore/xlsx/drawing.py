"""DrawingML spreadsheet drawings: anchors, pictures and charts (``xl/drawings``, ``xl/charts``).

Charts are read for the common 2-D chart groups (bar/column, line, pie, doughnut, area, scatter,
radar) and written from scratch with cached series values, so other readers can draw them
without recalculating."""

from __future__ import annotations

import posixpath
from typing import List, Optional, Tuple
from xml.etree import ElementTree as ET

from ..core.addr import MAX_COLS, MAX_ROWS
from ..core.value import Array, CellError
from ..model.features import Anchor, Chart, Image, Series
from ..model.sheet import Sheet
from ..model.style import Color
from .tables import EMU_PER_PX
from .xml import XML_DECL, attr, attr_int, child, child_val, esc, esc_attr, kids, local, num, path, text_of

NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# ---------------------------------------------------------------- reading


def _emu_px(v: Optional[str]) -> float:
    try:
        return max(min(int(v) / EMU_PER_PX, 1e7), -1e7) if v is not None else 0.0
    except ValueError:
        return 0.0


def _marker(e: ET.Element) -> Tuple[int, int, float, float]:
    def n(name: str) -> Optional[str]:
        c = child(e, name)
        return (c.text or "").strip() if c is not None else None

    try:
        col = min(max(int(n("col") or 0), 0), MAX_COLS - 1)
        row = min(max(int(n("row") or 0), 0), MAX_ROWS - 1)
    except ValueError:
        col = row = 0
    return row, col, max(_emu_px(n("colOff")), 0.0), max(_emu_px(n("rowOff")), 0.0)


def _read_anchor(a: ET.Element, sheet: Sheet) -> Optional[Anchor]:
    tag = local(a.tag)
    if tag == "twoCellAnchor":
        fr = child(a, "from")
        if fr is None:
            return None
        r, c, dx, dy = _marker(fr)
        to = child(a, "to")
        tr, tc, tdx, tdy = _marker(to) if to is not None else (r, c, dx, dy)
        x1, y1 = sheet.col_left(c) + dx, sheet.row_top(r) + dy
        x2, y2 = sheet.col_left(tc) + tdx, sheet.row_top(tr) + tdy
        mode = {"absolute": "absolute", "oneCell": "oneCell"}.get(attr(a, "editAs") or "", "twoCell")
        return Anchor(r, c, dx, dy, max(x2 - x1, 0.0), max(y2 - y1, 0.0), mode)
    if tag == "oneCellAnchor":
        fr = child(a, "from")
        if fr is None:
            return None
        r, c, dx, dy = _marker(fr)
        ext = child(a, "ext")
        w = max(_emu_px(attr(ext, "cx")), 0.0) if ext is not None else 0.0
        h = max(_emu_px(attr(ext, "cy")), 0.0) if ext is not None else 0.0
        return Anchor(r, c, dx, dy, w, h, "oneCell")
    if tag == "absoluteAnchor":
        pos = child(a, "pos")
        x = max(_emu_px(attr(pos, "x")), 0.0) if pos is not None else 0.0
        y = max(_emu_px(attr(pos, "y")), 0.0) if pos is not None else 0.0
        col = sheet.col_index().at(x)
        row = sheet.row_index().at(y)
        ext = child(a, "ext")
        w = max(_emu_px(attr(ext, "cx")), 0.0) if ext is not None else 0.0
        h = max(_emu_px(attr(ext, "cy")), 0.0) if ext is not None else 0.0
        return Anchor(row, col, max(x - sheet.col_left(col), 0.0), max(y - sheet.row_top(row), 0.0), w, h, "absolute")
    return None


_SCHEME = {"bg1": 0, "lt1": 0, "tx1": 1, "dk1": 1, "bg2": 2, "lt2": 2, "tx2": 3, "dk2": 3, "accent1": 4, "accent2": 5,
           "accent3": 6, "accent4": 7, "accent5": 8, "accent6": 9, "hlink": 10, "folHlink": 11}


def dml_color(e: Optional[ET.Element]) -> Optional[Color]:
    if e is None:
        return None
    c = child(e, "srgbClr")
    if c is not None and attr(c, "val"):
        return Color.from_hex(attr(c, "val"))
    c = child(e, "sysClr")
    if c is not None and attr(c, "lastClr"):
        return Color.from_hex(attr(c, "lastClr"))
    c = child(e, "schemeClr")
    if c is not None and attr(c, "val") in _SCHEME:
        return Color.theme(_SCHEME[attr(c, "val")])
    return None


def _mime(name: str) -> str:
    ext = name.rsplit(".", 1)[-1].lower()
    return {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif", "bmp": "image/bmp",
            "svg": "image/svg+xml", "webp": "image/webp", "emf": "image/x-emf", "wmf": "image/x-wmf"}.get(ext, "application/octet-stream")


def read_drawing(cx, part: str, sheet: Sheet) -> None:
    x = cx.optional_xml(part)
    if x is None:
        return
    rels = cx.pkg.rels(part)
    for a in x:
        anchor = _read_anchor(a, sheet)
        if anchor is None:
            continue
        pic = child(a, "pic")
        gf = child(a, "graphicFrame")
        if pic is not None:
            blip = path(pic, "blipFill", "blip")
            rid = attr(blip, "embed") if blip is not None else None
            rel = next((r for r in rels if r.id == rid), None)
            if rel is None:
                continue
            data = cx.pkg.read(rel.target)
            if not data:
                continue
            nv = path(pic, "nvPicPr", "cNvPr")
            alt = (attr(nv, "descr") or "") if nv is not None else ""
            sheet.images.append(Image(cx.next_object_id(), anchor, data, _mime(rel.target), alt))
        elif gf is not None:
            gd = path(gf, "graphic", "graphicData")
            ch = child(gd, "chart") if gd is not None else None
            if ch is None:
                if gd is not None:
                    cx.warn(f"a chart of a newer type (chartex) or an embedded object (SmartArt, OLE…) was skipped ({part})")
                continue
            rid = attr(ch, "id")
            rel = next((r for r in rels if r.id == rid), None)
            if rel is None:
                continue
            cxml = cx.optional_xml(rel.target)
            if cxml is None:
                continue
            chart = read_chart(cxml)
            if chart is None:
                cx.warn(f"chart {rel.target} has no supported plot and was skipped ({part})")
                continue
            chart.id = cx.next_object_id()
            chart.anchor = anchor
            sheet.charts.append(chart)
        elif child(a, "sp") is not None or child(a, "grpSp") is not None or child(a, "cxnSp") is not None:
            cx.warn(f"a shape, text box or group of shapes was skipped ({part})")


def _group_kind(g: ET.Element) -> Optional[str]:
    tag = local(g.tag)
    grouping = child_val(g, "grouping") or "standard"
    stacked, pct = grouping == "stacked", grouping == "percentStacked"
    if tag in ("barChart", "bar3DChart"):
        bar = child_val(g, "barDir") == "bar"
        base = "bar" if bar else "column"
        return base + ("Stacked" if stacked else "Stacked100" if pct else "Clustered")
    if tag in ("lineChart", "line3DChart"):
        if stacked or pct:
            return "lineStacked"
        mk = child_val(g, "marker")
        sers = list(kids(g, "ser"))
        all_none = bool(sers) and all((path(s, "marker", "symbol") is not None and attr(path(s, "marker", "symbol"), "val") == "none") for s in sers)
        if (mk is not None and mk not in ("1", "true")) or all_none:
            return "line"
        return "lineMarkers"
    if tag in ("pieChart", "pie3DChart", "ofPieChart"):
        return "pie"
    if tag == "doughnutChart":
        return "doughnut"
    if tag in ("areaChart", "area3DChart"):
        return "areaStacked100" if pct else ("areaStacked" if stacked else "area")
    if tag == "scatterChart":
        # Excel writes "lineMarker" for every scatter chart: one without lines has series lines with no fill.
        sers = list(kids(g, "ser"))
        no_lines = bool(sers) and all(path(s, "spPr", "ln", "noFill") is not None for s in sers)
        return "scatter" if child_val(g, "scatterStyle") == "marker" or no_lines else "scatterLines"
    if tag == "radarChart":
        return "radar"
    return None


def _data_ref(e: Optional[ET.Element]) -> Optional[str]:
    if e is None:
        return None
    for k in ("numRef", "strRef", "multiLvlStrRef"):
        r = child(e, k)
        if r is not None:
            f = child(r, "f")
            if f is not None:
                t = (f.text or "").strip()
                return t[1:] if t.startswith("=") else t
    for k in ("numLit", "strLit"):
        l = child(e, k)
        if l is not None:
            pts = []
            for p in kids(l, "pt"):
                v = child(p, "v")
                if v is None:
                    continue
                t = v.text or ""
                pts.append(t.strip() if k == "numLit" else '"' + t.replace('"', '""') + '"')
            return "{" + ",".join(pts) + "}"
    return None


def _read_series(s: ET.Element) -> Series:
    name = None
    tx = child(s, "tx")
    if tx is not None:
        f = path(tx, "strRef", "f")
        if f is not None:
            t = (f.text or "").strip()
            name = t[1:] if t.startswith("=") else t
        else:
            v = child(tx, "v")
            if v is not None:
                name = '"' + (v.text or "").replace('"', '""') + '"'
    sp = child(s, "spPr")
    color = None
    if sp is not None:
        color = dml_color(child(sp, "solidFill")) or dml_color(path(sp, "ln", "solidFill"))
    if color is None:
        color = dml_color(path(s, "marker", "spPr", "solidFill"))
    return Series(
        name=name,
        categories=_data_ref(child(s, "cat")) or _data_ref(child(s, "xVal")),
        values=_data_ref(child(s, "val")) or _data_ref(child(s, "yVal")) or "",
        color=color,
    )


def _rich_title(t: Optional[ET.Element]) -> Optional[str]:
    if t is None:
        return None
    tx = child(t, "tx")
    if tx is None:
        return None
    rich = child(tx, "rich")
    if rich is not None:
        paras = []
        for p in kids(rich, "p"):
            paras.append("".join(text_of(child(r, "t")) for r in kids(p, "r")))
        s = "\n".join(paras)
        return s or None
    sr = child(tx, "strRef")
    if sr is not None:
        v = path(sr, "strCache", "pt", "v")
        if v is not None:
            return v.text
        f = child(sr, "f")
        return (f.text or "").strip() if f is not None else None
    return None


def read_chart(root: ET.Element) -> Optional[Chart]:
    chart = child(root, "chart")
    if chart is None:
        return None
    plot = child(chart, "plotArea")
    if plot is None:
        return None
    groups = [(g, _group_kind(g)) for g in plot]
    groups = [(g, k) for g, k in groups if k is not None]
    if not groups:
        return None
    first, kind = groups[0]
    series: List[Series] = []
    for g, _k in groups:
        for s in kids(g, "ser"):
            series.append(_read_series(s))
    title = _rich_title(child(chart, "title"))
    legend = "none"
    lg = child(chart, "legend")
    if lg is not None:
        pos = child_val(lg, "legendPos") or "r"
        legend = {"b": "bottom", "t": "top", "l": "left", "r": "right", "tr": "right"}.get(pos, "right")
    data_labels = False
    dl = child(first, "dLbls")
    if dl is not None:
        sv = child(dl, "showVal")
        data_labels = sv is not None and (attr(sv, "val") or "1") in ("1", "true")
    gridlines = any(child(ax, "majorGridlines") is not None for ax in plot if local(ax.tag) == "valAx")
    x_title = y_title = None
    for ax in plot:
        n = local(ax.tag)
        if n in ("catAx", "dateAx") or (n == "valAx" and kind.startswith("scatter") and child_val(ax, "axPos") in ("b", "t")):
            x_title = x_title or _rich_title(child(ax, "title"))
        elif n == "valAx":
            y_title = y_title or _rich_title(child(ax, "title"))
    return Chart(id=0, kind=kind, title=title, series=series, legend=legend, data_labels=data_labels, gridlines=gridlines, x_title=x_title, y_title=y_title)


# ---------------------------------------------------------------- writing


def _anchor_markers(sheet: Sheet, a: Anchor):
    x1 = sheet.col_left(a.col) + a.dx
    y1 = sheet.row_top(a.row) + a.dy
    x2, y2 = x1 + a.width, y1 + a.height
    ci, ri = sheet.col_index(), sheet.row_index()
    c2, r2 = ci.at(x2), ri.at(y2)
    return (a.col, a.dx, a.row, a.dy), (c2, x2 - sheet.col_left(c2), r2, y2 - sheet.row_top(r2))


def _marker_xml(tag: str, m) -> str:
    col, cx, row, cy = m
    return (
        f"<xdr:{tag}><xdr:col>{col}</xdr:col><xdr:colOff>{int(max(cx, 0) * EMU_PER_PX)}</xdr:colOff>"
        f"<xdr:row>{row}</xdr:row><xdr:rowOff>{int(max(cy, 0) * EMU_PER_PX)}</xdr:rowOff></xdr:{tag}>"
    )


def _anchor_open(sheet: Sheet, a: Anchor) -> str:
    fr, to = _anchor_markers(sheet, a)
    edit = {"absolute": ' editAs="absolute"', "oneCell": ' editAs="oneCell"'}.get(a.mode, "")
    return f"<xdr:twoCellAnchor{edit}>" + _marker_xml("from", fr) + _marker_xml("to", to)


def image_ext(mime: str) -> str:
    return {"image/png": "png", "image/jpeg": "jpeg", "image/gif": "gif", "image/bmp": "bmp", "image/svg+xml": "svg", "image/webp": "webp"}.get(mime, "png")


def image_content_type(ext: str) -> str:
    return {"png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg", "gif": "image/gif", "bmp": "image/bmp", "svg": "image/svg+xml", "webp": "image/webp"}.get(ext, "image/png")


def write_drawing(sheet: Sheet, objs: List[tuple]) -> str:
    """``objs``: ("image", Image, rid) or ("chart", Chart, rid)."""
    s = [XML_DECL, f'<xdr:wsDr xmlns:xdr="{NS_XDR}" xmlns:a="{NS_A}">']
    for i, (kind, obj, rid) in enumerate(objs):
        oid = i + 2
        s.append(_anchor_open(sheet, obj.anchor))
        if kind == "image":
            s.append(
                f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{oid}" name="Picture {oid}" descr="{esc_attr(obj.alt)}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
                f'<xdr:blipFill><a:blip xmlns:r="{NS_R}" r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
                f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(obj.anchor.width * EMU_PER_PX)}" cy="{int(obj.anchor.height * EMU_PER_PX)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
            )
        else:
            s.append(
                f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{oid}" name="Chart {oid}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
                '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
                f'<a:graphic><a:graphicData uri="{NS_C}"><c:chart xmlns:c="{NS_C}" xmlns:r="{NS_R}" r:id="{rid}"/></a:graphicData></a:graphic></xdr:graphicFrame>'
            )
        s.append("<xdr:clientData/></xdr:twoCellAnchor>")
    s.append("</xdr:wsDr>")
    return "".join(s)


def _srgb(c: Color, theme) -> str:
    h = c.hex(theme)
    return h[1:] if h else "4472C4"


def _cache(values: List, numeric: bool) -> str:
    if numeric:
        pts = "".join(f'<c:pt idx="{i}"><c:v>{num(v)}</c:v></c:pt>' for i, v in enumerate(values) if isinstance(v, float))
        return f'<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{len(values)}"/>{pts}</c:numCache>'
    pts = "".join(f'<c:pt idx="{i}"><c:v>{esc("" if v is None else str(v))}</c:v></c:pt>' for i, v in enumerate(values))
    return f'<c:strCache><c:ptCount val="{len(values)}"/>{pts}</c:strCache>'


def _series_values(wb, sheet_index: int, ref: Optional[str]) -> List:
    if not ref:
        return []
    from ..calc import evaluate

    v = evaluate(wb, sheet_index, (0, 0), "=" + ref)
    if isinstance(v, Array):
        out = []
        for x in v.data:
            out.append(x)
        return out
    return [v]


def _ref_xml(tag: str, ref: Optional[str], values: List, numeric: bool, from_file) -> str:
    if not ref:
        return ""
    if ref.startswith("{"):
        inner = ref.strip("{}").split(",") if ref.strip("{}") else []
        if numeric:
            pts = "".join(f'<c:pt idx="{i}"><c:v>{esc(p.strip())}</c:v></c:pt>' for i, p in enumerate(inner))
            return f'<c:{tag}><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="{len(inner)}"/>{pts}</c:numLit></c:{tag}>'
        pts = "".join(f'<c:pt idx="{i}"><c:v>{esc(p.strip().strip(chr(34)))}</c:v></c:pt>' for i, p in enumerate(inner))
        return f'<c:{tag}><c:strLit><c:ptCount val="{len(inner)}"/>{pts}</c:strLit></c:{tag}>'
    kind = "numRef" if numeric else "strRef"
    return f"<c:{tag}><c:{kind}><c:f>{esc(from_file(ref))}</c:f>{_cache(values, numeric)}</c:{kind}></c:{tag}>"


_PALETTE = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "264478", "9E480E", "636363", "997300"]


def write_chart(wb, sheet_index: int, ch: Chart) -> str:
    from .fmla import text_to_file

    kind = ch.kind
    k = kind.lower()
    s = [XML_DECL, f'<c:chartSpace xmlns:c="{NS_C}" xmlns:a="{NS_A}" xmlns:r="{NS_R}"><c:roundedCorners val="0"/><c:chart>']
    if ch.title:
        s.append(
            f'<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{esc(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
        )
    else:
        s.append('<c:autoTitleDeleted val="1"/>')
    s.append("<c:plotArea><c:layout/>")
    if k.startswith("column") or k.startswith("bar"):
        grouping = "stacked" if "stacked100" not in k and "stacked" in k else ("percentStacked" if "stacked100" in k else "clustered")
        s.append(f'<c:barChart><c:barDir val="{"bar" if k.startswith("bar") else "col"}"/><c:grouping val="{grouping}"/><c:varyColors val="0"/>')
        group_end = '<c:gapWidth val="150"/>' + ('<c:overlap val="100"/>' if grouping != "clustered" else "") + '<c:axId val="111"/><c:axId val="222"/></c:barChart>'
        axes = "cat"
    elif k.startswith("line"):
        grouping = "stacked" if k == "linestacked" else "standard"
        s.append(f'<c:lineChart><c:grouping val="{grouping}"/><c:varyColors val="0"/>')
        group_end = f'<c:marker val="{"1" if k == "linemarkers" else "0"}"/><c:axId val="111"/><c:axId val="222"/></c:lineChart>'
        axes = "cat"
    elif k in ("pie", "doughnut"):
        s.append(f'<c:{"pieChart" if k == "pie" else "doughnutChart"}><c:varyColors val="1"/>')
        group_end = ('<c:firstSliceAng val="0"/></c:pieChart>' if k == "pie" else '<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>')
        axes = None
    elif k.startswith("area"):
        grouping = "percentStacked" if k == "areastacked100" else ("stacked" if k == "areastacked" else "standard")
        s.append(f'<c:areaChart><c:grouping val="{grouping}"/><c:varyColors val="0"/>')
        group_end = '<c:axId val="111"/><c:axId val="222"/></c:areaChart>'
        axes = "cat"
    elif k.startswith("scatter"):
        s.append('<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>')
        group_end = '<c:axId val="111"/><c:axId val="222"/></c:scatterChart>'
        axes = "val"
    else:  # radar
        s.append('<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>')
        group_end = '<c:axId val="111"/><c:axId val="222"/></c:radarChart>'
        axes = "cat"
    for i, ser in enumerate(ch.series):
        color = _srgb(ser.color, wb.theme) if ser.color is not None else _PALETTE[i % len(_PALETTE)]
        s.append(f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>')
        if ser.name:
            if ser.name.startswith('"'):
                s.append(f"<c:tx><c:v>{esc(ser.name.strip(chr(34)))}</c:v></c:tx>")
            else:
                names = _series_values(wb, sheet_index, ser.name)
                s.append(f"<c:tx><c:strRef><c:f>{esc(text_to_file(ser.name))}</c:f>{_cache(names[:1], False)}</c:strRef></c:tx>")
        if k in ("pie", "doughnut"):
            s.append("")
        elif k == "scatter":
            # Markers only: the series line has no fill, as Excel writes it.
            fill = f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            s.append('<c:spPr><a:ln w="19050" cap="rnd"><a:noFill/><a:round/></a:ln></c:spPr>')
            s.append(f'<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>{fill}<a:ln w="9525">{fill}</a:ln></c:spPr></c:marker>')
        elif k.startswith("line") or k.startswith("scatter") or k == "radar":
            s.append(f'<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:round/></a:ln></c:spPr>')
            if k in ("line", "linestacked"):
                s.append('<c:marker><c:symbol val="none"/></c:marker>')
        else:
            s.append(f'<c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></c:spPr>')
        if ch.data_labels:
            s.append('<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
        cats = _series_values(wb, sheet_index, ser.categories)
        vals = _series_values(wb, sheet_index, ser.values)
        if k.startswith("scatter"):
            numeric_x = all(isinstance(v, float) for v in cats) if cats else True
            s.append(_ref_xml("xVal", ser.categories, cats, numeric_x, text_to_file))
            s.append(_ref_xml("yVal", ser.values, vals, True, text_to_file))
            s.append('<c:smooth val="0"/>')
        else:
            numeric_c = bool(cats) and all(isinstance(v, float) for v in cats)
            s.append(_ref_xml("cat", ser.categories, cats, numeric_c, text_to_file))
            s.append(_ref_xml("val", ser.values, vals, True, text_to_file))
            if k.startswith("line"):
                s.append('<c:smooth val="0"/>')
        s.append("</c:ser>")
    if ch.data_labels and not ch.series:
        pass
    s.append(group_end)
    grid = "<c:majorGridlines/>" if ch.gridlines else ""

    def title_xml(t: Optional[str]) -> str:
        if not t:
            return ""
        return f'<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{esc(t)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'

    if axes == "cat":
        bar_h = k.startswith("bar")
        s.append(
            f'<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{"l" if bar_h else "b"}"/>'
            f'{title_xml(ch.x_title)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
            f'<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{"b" if bar_h else "l"}"/>{grid}'
            f'{title_xml(ch.y_title)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="111"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>'
        )
    elif axes == "val":
        s.append(
            f'<c:valAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>{title_xml(ch.x_title)}'
            '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="222"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>'
            f'<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>{grid}{title_xml(ch.y_title)}'
            '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="111"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>'
        )
    s.append("</c:plotArea>")
    if ch.legend != "none":
        pos = {"bottom": "b", "top": "t", "left": "l", "right": "r"}.get(ch.legend, "b")
        s.append(f'<c:legend><c:legendPos val="{pos}"/><c:overlay val="0"/></c:legend>')
    s.append('<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>')
    return "".join(s)
