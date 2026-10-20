"""Writing worksheet parts and the parts they own (tables, comments, drawings, charts)."""

from __future__ import annotations

from typing import Dict, List, Optional, Set, Tuple

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef, a1
from ..core.value import Array, CellError
from ..formula import quote_sheet
from ..model.features import CondFormat, FilterCriterion, LineInfo, Table, Validation, default_icon_thresholds, icon_count
from ..model.sheet import DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, Sheet
from ..model.style import AUTO, Fill, Style
from .array_formula import needs_array
from .drawing import image_content_type, image_ext, write_chart, write_drawing
from .fmla import text_to_file, to_file
from .styles import color_attrs, dxf_id
from .tables import paper_id, px_to_col_width, px_to_pt
from .xml import XML_DECL, encode_xstring, esc, esc_attr, num, t_el

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"


class Rels:
    def __init__(self) -> None:
        self.items: List[Tuple[str, str, str, bool]] = []

    def add(self, kind: str, target: str, external: bool = False) -> str:
        rid = f"rId{len(self.items) + 1}"
        if "://" in kind:
            ty = kind
        elif kind == "core-properties":
            ty = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
        else:
            ty = f"{NS_REL}/{kind}"
        self.items.append((rid, ty, target, external))
        return rid

    def __bool__(self) -> bool:
        return bool(self.items)

    def xml(self) -> str:
        s = [XML_DECL, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for rid, ty, target, ext in self.items:
            s.append(f'<Relationship Id="{rid}" Type="{ty}" Target="{esc_attr(target)}"{" TargetMode=" + chr(34) + "External" + chr(34) if ext else ""}/>')
        s.append("</Relationships>")
        return "".join(s)


def _sqref(ranges: List[RangeRef]) -> str:
    return " ".join(r.a1() for r in ranges)


def _table_name(name: str, used: Set[str], n: int) -> str:
    s = "".join(c if (c.isalnum() or c in "_." or (i == 0 and c == "\\")) else "_" for i, c in enumerate(name))
    from ..core.addr import parse_cell

    if not s or not (s[0].isalpha() or s[0] in "_\\") or parse_cell(s) is not None:
        s = f"Table{n}"
    out, k = s, 2
    while out.lower() in used:
        out = f"{s}{k}"
        k += 1
    used.add(out.lower())
    return out


def _table_columns(t: Table) -> List[str]:
    used: Set[str] = set()
    out = []
    for i in range(t.range.width):
        base = t.columns[i].name if i < len(t.columns) and t.columns[i].name.strip() else f"Column{i + 1}"
        name, k = base, 2
        while name.lower() in used:
            name = f"{base}{k}"
            k += 1
        used.add(name.lower())
        out.append(name)
    return out


def write_sheet(wb, si: int, selected: bool, out) -> Tuple[str, Rels]:
    rels = Rels()
    sheet: Sheet = wb.sheets[si]
    nstyles = len(wb.styles)

    def sid(s: int) -> int:
        return s if 0 <= s < nstyles else 0

    # Table header cells must hold the column names.
    overrides: Dict[Tuple[int, int], str] = {}
    tables = []
    for t in sheet.tables:
        # Only a table with a data row: Excel finds one without (a header row alone, or a header
        # and a total row) damaged. ReflexGrid's own tables always have one.
        if t.range.height > (1 if t.header_row else 0) + (1 if t.totals_row else 0):
            cols = _table_columns(t)
            tables.append((t, cols))
            if t.header_row:
                for i, name in enumerate(cols):
                    overrides[(t.range.r0, t.range.c0 + i)] = name
    # Dynamic-array spills: anchor -> range, and the spilled values to cache.
    dynamic: Dict[Tuple[int, int], RangeRef] = {}
    for a, r in sheet.spill_ranges.items():
        cell = sheet.cells.get(a)
        if not r.is_single() and cell is not None and cell.formula is not None and cell.formula.array is None:
            dynamic[a] = r
    extra: Dict[Tuple[int, int], object] = {}
    for a, r in sheet.spill_ranges.items():
        if r.is_single():
            continue
        for i, c in enumerate(r.cells()):
            if i > 1_000_000:
                break
            if c != a and c in sheet.spill:
                x = sheet.cells.get(c)
                if x is None or (x.formula is None and x.value is None):
                    extra[c] = sheet.spill[c]

    s = [XML_DECL, f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">']
    fit = sheet.print.fit_to is not None
    names = getattr(out, "code_names", None)
    code_name = names[si] if names is not None and si < len(names) else sheet.code_name
    if sheet.tab_color is not None or fit or code_name:
        s.append(f'<sheetPr codeName="{esc_attr(code_name)}">' if code_name else "<sheetPr>")
        if sheet.tab_color is not None:
            s.append(f"<tabColor {color_attrs(sheet.tab_color)}/>")
        if fit:
            s.append('<pageSetUpPr fitToPage="1"/>')
        s.append("</sheetPr>")
    used = sheet.cells.used_range()
    for c in list(extra) + list(overrides):
        r = RangeRef(c[0], c[1], c[0], c[1])
        used = r if used is None else used.union(r)
    s.append(f'<dimension ref="{used.a1() if used is not None else "A1"}"/>')

    # sheetViews
    v = ["<sheetViews><sheetView"]
    if selected:
        v.append(' tabSelected="1"')
    if not sheet.show_gridlines:
        v.append(' showGridLines="0"')
    if not sheet.show_headings:
        v.append(' showRowColHeaders="0"')
    if sheet.show_formulas:
        v.append(' showFormulas="1"')
    if not sheet.show_zeros:
        v.append(' showZeros="0"')
    if sheet.right_to_left:
        v.append(' rightToLeft="1"')
    if sheet.zoom != 100 and 10 <= sheet.zoom <= 400:
        v.append(f' zoomScale="{sheet.zoom}" zoomScaleNormal="{sheet.zoom}"')
    if sheet.view_top_left != (0, 0) and sheet.freeze is None:
        v.append(f' topLeftCell="{a1(*sheet.view_top_left)}"')
    v.append(' workbookViewId="0">')
    active = a1(*sheet.view_active)
    fz = sheet.freeze
    if fz is not None and (fz[0] > 0 or fz[1] > 0):
        rows, cols = fz
        pane = "bottomRight" if rows and cols else ("bottomLeft" if rows else "topRight")
        v.append("<pane")
        if cols:
            v.append(f' xSplit="{cols}"')
        if rows:
            v.append(f' ySplit="{rows}"')
        tl = (min(max(rows, sheet.view_top_left[0]), MAX_ROWS - 1), min(max(cols, sheet.view_top_left[1]), MAX_COLS - 1))
        v.append(f' topLeftCell="{a1(*tl)}" activePane="{pane}" state="frozen"/>')
        v.append(f'<selection pane="{pane}" activeCell="{active}" sqref="{active}"/>')
    else:
        v.append(f'<selection activeCell="{active}" sqref="{active}"/>')
    v.append("</sheetView></sheetViews>")
    s.append("".join(v))

    max_row_outline = max((i.outline for i in sheet.rows.values()), default=0)
    max_col_outline = max((i.outline for i in sheet.cols.values()), default=0)
    f = ["<sheetFormatPr"]
    if abs(sheet.default_col_width - DEFAULT_COL_WIDTH) > 0.01:
        f.append(f' defaultColWidth="{num(px_to_col_width(sheet.default_col_width))}"')
    f.append(f' defaultRowHeight="{num(px_to_pt(sheet.default_row_height))}"')
    if abs(sheet.default_row_height - DEFAULT_ROW_HEIGHT) > 0.01:
        f.append(' customHeight="1"')
    if max_row_outline:
        f.append(f' outlineLevelRow="{max_row_outline}"')
    if max_col_outline:
        f.append(f' outlineLevelCol="{max_col_outline}"')
    f.append("/>")
    s.append("".join(f))

    if sheet.cols:
        s.append("<cols>")
        run = None

        def flush(r):
            a, b, i = r
            w = i.size if i.size is not None else sheet.default_col_width
            x = [f'<col min="{a + 1}" max="{b + 1}" width="{num(px_to_col_width(w))}"']
            if i.style is not None:
                x.append(f' style="{sid(i.style)}"')
            if i.hidden:
                x.append(' hidden="1"')
            if i.size is not None:
                x.append(' customWidth="1"' if i.custom else ' bestFit="1" customWidth="1"')
            if i.outline:
                x.append(f' outlineLevel="{min(i.outline, 7)}"')
            if i.collapsed:
                x.append(' collapsed="1"')
            x.append("/>")
            s.append("".join(x))

        for c in sorted(sheet.cols):
            if c >= MAX_COLS:
                continue
            info = sheet.cols[c]
            if run is not None and run[1] + 1 == c and run[2] == info:
                run = (run[0], c, info)
            else:
                if run is not None:
                    flush(run)
                run = (c, c, info)
        if run is not None:
            flush(run)
        s.append("</cols>")

    # sheetData
    rows: Dict[int, List[Tuple[int, int]]] = {}
    for (rc, _cell) in sheet.cells.items():
        rows.setdefault(rc[0], []).append(rc)
    for rc in list(extra) + list(overrides):
        rows.setdefault(rc[0], []).append(rc)
    for r in sheet.rows:
        rows.setdefault(r, [])
    s.append("<sheetData>")
    for r in sorted(rows):
        cells = sorted(set(rows[r]))
        info = sheet.rows.get(r, LineInfo())
        if not cells and info.is_default():
            continue
        x = [f'<row r="{r + 1}"']
        if info.style is not None:
            x.append(f' s="{sid(info.style)}" customFormat="1"')
        if info.size is not None:
            x.append(f' ht="{num(px_to_pt(info.size))}"' + (' customHeight="1"' if info.custom else ""))
        if info.hidden:
            x.append(' hidden="1"')
        if info.outline:
            x.append(f' outlineLevel="{min(info.outline, 7)}"')
        if info.collapsed:
            x.append(' collapsed="1"')
        if not cells:
            x.append("/>")
            s.append("".join(x))
            continue
        x.append(">")
        s.append("".join(x))
        for c in cells:
            cell = sheet.cells.get(c)
            arr = dynamic.get(c)
            if arr is None and cell is not None and cell.formula is not None and cell.formula.array is None:
                try:
                    if needs_array(wb, si, c, cell.formula):
                        arr = RangeRef(c[0], c[1], c[0], c[1])
                except RecursionError:
                    arr = RangeRef(c[0], c[1], c[0], c[1])
            _write_cell(s, c, cell, overrides.get(c), extra.get(c, _MISSING), arr, sid, out, sheet)
        s.append("</row>")
    s.append("</sheetData>")

    p = sheet.protection
    if p is not None:
        x = ['<sheetProtection sheet="1" objects="1" scenarios="1"']
        for name, allowed in (
            ("formatCells", p.format_cells), ("formatColumns", p.format_columns), ("formatRows", p.format_rows),
            ("insertColumns", p.insert_columns), ("insertRows", p.insert_rows), ("deleteColumns", p.delete_columns),
            ("deleteRows", p.delete_rows), ("sort", p.sort), ("autoFilter", p.autofilter),
        ):
            if allowed:
                x.append(f' {name}="0"')
        if not p.select_locked:
            x.append(' selectLockedCells="1"')
        if not p.select_unlocked:
            x.append(' selectUnlockedCells="1"')
        if p.hash_attrs:
            for k, val in p.hash_attrs.items():
                x.append(f' {k}="{esc_attr(val)}"')
        elif p.password_hash:
            x.append(f' password="{esc_attr(p.password_hash)}"')
        x.append("/>")
        s.append("".join(x))
    af = sheet.autofilter
    if af is not None and not any(t.range.intersects(af.range) for t in sheet.tables):
        s.append(_autofilter_xml(af.range, af.criteria, out))
    merges = [m for m in sheet.merges if not m.is_single()]
    if merges:
        s.append(f'<mergeCells count="{len(merges)}">' + "".join(f'<mergeCell ref="{m.a1()}"/>' for m in merges) + "</mergeCells>")
    order = sorted(range(len(sheet.cond_formats)), key=lambda i: sheet.cond_formats[i].priority)
    prio = {i: rank + 1 for rank, i in enumerate(order)}
    for i, cf in enumerate(sheet.cond_formats):
        if cf.ranges:
            s.append(f'<conditionalFormatting sqref="{_sqref(cf.ranges)}">{_cf_rule_xml(cf, prio.get(i, 1), out)}</conditionalFormatting>')
    vals = [vv for vv in sheet.validations if vv.ranges]
    if vals:
        s.append(f'<dataValidations count="{len(vals)}">' + "".join(_validation_xml(vv) for vv in vals) + "</dataValidations>")
    if sheet.hyperlinks:
        hs = []
        for c in sorted(sheet.hyperlinks):
            h = sheet.hyperlinks[c]
            t = h.target.strip()
            if not t:
                continue
            tl = t.lower()
            external = h.external or "://" in t or tl.startswith("mailto:") or t.startswith("\\\\") or tl.startswith("file:")
            x = [f'<hyperlink ref="{a1(*c)}"']
            if external:
                if "#" in t and not t.startswith("#"):
                    url, loc = t.split("#", 1)
                else:
                    url, loc = t, None
                rid = rels.add("hyperlink", url, True)
                x.append(f' r:id="{rid}"')
                if loc:
                    x.append(f' location="{esc_attr(loc)}"')
            else:
                loc = t.lstrip("#")
                x.append(f' location="{esc_attr(loc)}" display="{esc_attr(loc)}"')
            if h.tooltip:
                x.append(f' tooltip="{esc_attr(h.tooltip)}"')
            x.append("/>")
            hs.append("".join(x))
        if hs:
            s.append("<hyperlinks>" + "".join(hs) + "</hyperlinks>")
    pr = sheet.print
    if pr.gridlines or pr.headings or pr.center_h or pr.center_v:
        x = ["<printOptions"]
        for name, val in (("horizontalCentered", pr.center_h), ("verticalCentered", pr.center_v), ("headings", pr.headings), ("gridLines", pr.gridlines)):
            if val:
                x.append(f' {name}="1"')
        x.append("/>")
        s.append("".join(x))
    m = pr.margins
    s.append(f'<pageMargins left="{num(m[0])}" right="{num(m[1])}" top="{num(m[2])}" bottom="{num(m[3])}" header="{num(m[4])}" footer="{num(m[5])}"/>')
    x = ["<pageSetup"]
    pid = paper_id(pr.paper)
    if pid is not None and pid != 1:
        x.append(f' paperSize="{pid}"')
    if pr.scale != 100 and 10 <= pr.scale <= 400:
        x.append(f' scale="{pr.scale}"')
    if pr.fit_to is not None:
        x.append(f' fitToWidth="{pr.fit_to[0]}" fitToHeight="{pr.fit_to[1]}"')
    x.append(f' orientation="{"landscape" if pr.orientation == "landscape" else "portrait"}"/>')
    s.append("".join(x))
    if pr.header or pr.footer:
        s.append("<headerFooter>")
        if pr.header:
            s.append(f"<oddHeader>{esc(pr.header)}</oddHeader>")
        if pr.footer:
            s.append(f"<oddFooter>{esc(pr.footer)}</oddFooter>")
        s.append("</headerFooter>")
    for tag, lst, mx in (("rowBreaks", pr.row_breaks, 16383), ("colBreaks", pr.col_breaks, 1048575)):
        lst = [b for b in lst if b > 0]
        if lst:
            s.append(f'<{tag} count="{len(lst)}" manualBreakCount="{len(lst)}">' + "".join(f'<brk id="{b}" max="{mx}" man="1"/>' for b in lst) + f"</{tag}>")

    # Drawing: pictures and charts.
    if sheet.images or sheet.charts:
        out.drawings += 1
        dn = out.drawings
        drels = Rels()
        objs = []
        for img in sheet.images:
            if not img.data:
                continue
            out.images += 1
            ext = image_ext(img.mime)
            name = f"xl/media/image{out.images}.{ext}"
            out.default_type(ext, image_content_type(ext))
            out.part(name, None, img.data)
            objs.append(("image", img, drels.add("image", f"../media/image{out.images}.{ext}")))
        for ch in sheet.charts:
            out.charts += 1
            name = f"xl/charts/chart{out.charts}.xml"
            out.part(name, CT_CHART, write_chart(wb, si, ch).encode("utf-8"))
            objs.append(("chart", ch, drels.add("chart", f"../charts/chart{out.charts}.xml")))
        out.part(f"xl/drawings/drawing{dn}.xml", CT_DRAWING, write_drawing(sheet, objs).encode("utf-8"))
        if drels:
            out.part(f"xl/drawings/_rels/drawing{dn}.xml.rels", None, drels.xml().encode("utf-8"))
        rid = rels.add("drawing", f"../drawings/drawing{dn}.xml")
        s.append(f'<drawing r:id="{rid}"/>')
    if sheet.comments:
        out.comments += 1
        n = out.comments
        out.part(f"xl/comments{n}.xml", CT_COMMENTS, _comments_xml(sheet).encode("utf-8"))
        out.default_type("vml", "application/vnd.openxmlformats-officedocument.vmlDrawing")
        out.part(f"xl/drawings/vmlDrawing{n}.vml", None, _vml_xml(sheet, n).encode("utf-8"))
        rels.add("comments", f"../comments{n}.xml")
        rid = rels.add("vmlDrawing", f"../drawings/vmlDrawing{n}.vml")
        s.append(f'<legacyDrawing r:id="{rid}"/>')
    if tables:
        parts = []
        used_names: Set[str] = set()
        for t, cols in tables:
            out.tables += 1
            out.table_id += 1
            n = out.tables
            name = _table_name(t.name, used_names, out.table_id)
            out.part(f"xl/tables/table{n}.xml", CT_TABLE, _table_xml(t, cols, name, out.table_id, out).encode("utf-8"))
            rid = rels.add("table", f"../tables/table{n}.xml")
            parts.append(f'<tablePart r:id="{rid}"/>')
        s.append(f'<tableParts count="{len(tables)}">' + "".join(parts) + "</tableParts>")
    if sheet.sparklines:
        s.append(
            '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
            '<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
        )
        for sp in sheet.sparklines:
            ty = {"column": ' type="column"', "stacked": ' type="stacked"'}.get(sp.kind, "")
            src = sp.source if "!" in sp.source else f"{quote_sheet(sheet.name)}!{sp.source}"
            color = 'rgb="FF376092"' if sp.color.is_auto else color_attrs(sp.color)
            s.append(
                f'<x14:sparklineGroup displayEmptyCellsAs="gap"{ty}{" markers=" + chr(34) + "1" + chr(34) if sp.markers else ""}>'
                f"<x14:colorSeries {color}/><x14:sparklines><x14:sparkline><xm:f>{esc(src)}</xm:f><xm:sqref>{a1(sp.row, sp.col)}</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>"
            )
        s.append("</x14:sparklineGroups></ext></extLst>")
    s.append("</worksheet>")
    return "".join(s), rels


_MISSING = object()


def _write_cell(s: List[str], c, cell, ov: Optional[str], extra, dynamic: Optional[RangeRef], sid, out, sheet=None) -> None:
    style = sid(cell.style) if cell is not None else 0
    x = [f'<c r="{a1(*c)}"']
    if style:
        x.append(f' s="{style}"')
    formula = None if ov is not None else (cell.formula if cell is not None else None)
    if ov is not None:
        value = ov
    elif extra is not _MISSING:
        value = extra
    else:
        value = cell.value if cell is not None else None
    if isinstance(value, Array):
        value = value.data[0] if value.data else None
    if value is CellError.CIRC:
        value = 0.0
    elif value in (CellError.SPILL, CellError.CALC):
        # #SPILL! and #CALC! are not error codes of the file format; Excel caches them as #VALUE!.
        value = CellError.VALUE
    if dynamic is not None and formula is not None and formula.array is None:
        x.append(' cm="1"')
        out.dynamic_arrays = True
    if formula is not None:
        if isinstance(value, str):
            x.append(' t="str"')
        elif isinstance(value, bool):
            x.append(' t="b"')
        elif isinstance(value, CellError):
            x.append(' t="e"')
        x.append(">")
        table = None
        if sheet is not None and sheet.tables and "[" in formula.text:
            table = next((t.name for t in sheet.tables if t.range.contains(c)), None)
        text = esc(to_file(formula, table))
        if formula.array is not None:
            x.append(f'<f t="array" ref="{formula.array.a1()}">{text}</f>')
        elif dynamic is not None:
            x.append(f'<f t="array" ref="{dynamic.a1()}">{text}</f>')
        else:
            x.append(f"<f>{text}</f>")
        if isinstance(value, bool):
            x.append(f"<v>{int(value)}</v>")
        elif isinstance(value, float):
            x.append(f"<v>{num(value)}</v>")
        elif isinstance(value, str):
            x.append(f"<v>{esc(encode_xstring(value))}</v>")
        elif isinstance(value, CellError):
            x.append(f"<v>{value.value}</v>")
        else:
            out.needs_calc = True
        x.append("</c>")
        s.append("".join(x))
        return
    if value is None:
        x.append("/>")
    elif isinstance(value, bool):
        x.append(f' t="b"><v>{int(value)}</v></c>')
    elif isinstance(value, float):
        x.append(f"><v>{num(value)}</v></c>")
    elif isinstance(value, str):
        x.append(f' t="s"><v>{out.sst.index(value)}</v></c>')
    elif isinstance(value, CellError):
        x.append(f' t="e"><v>{value.value}</v></c>')
    else:
        x.append("/>")
    s.append("".join(x))


def _cfvo_xml(v) -> str:
    if v.kind in ("min", "max"):
        return f'<cfvo type="{v.kind}"/>'
    if v.kind == "formula":
        return f'<cfvo type="formula" val="{esc_attr(text_to_file(str(v.value)))}"/>'
    return f'<cfvo type="{v.kind}" val="{num(float(v.value or 0.0))}"/>'


def _quote(t: str) -> str:
    return '"' + t.replace('"', '""') + '"'


def _cf_rule_xml(cf: CondFormat, priority: int, out) -> str:
    rule = cf.rule
    tl = a1(cf.ranges[0].r0, cf.ranges[0].c0) if cf.ranges else "A1"
    stop = ' stopIfTrue="1"' if cf.stop_if_true else ""
    st = rule.style or Style()

    def fx(t: str) -> str:
        return f"<formula>{esc(text_to_file(t))}</formula>"

    def simple(ty: str, extra: str, body: str) -> str:
        return f'<cfRule type="{ty}" dxfId="{dxf_id(out.dxfs, st)}" priority="{priority}"{stop}{extra}>{body}</cfRule>'

    t = rule.type
    if t == "cellIs":
        body = "".join(fx(f) for f in rule.formulas[:2])
        return simple("cellIs", f' operator="{rule.operator}"', body)
    if t == "expression":
        return simple("expression", "", fx(rule.formulas[0] if rule.formulas else "FALSE"))
    if t == "containsText":
        return simple("containsText", f' operator="containsText" text="{esc_attr(rule.text)}"', fx(f"NOT(ISERROR(SEARCH({_quote(rule.text)},{tl})))"))
    if t == "notContainsText":
        return simple("notContainsText", f' operator="notContains" text="{esc_attr(rule.text)}"', fx(f"ISERROR(SEARCH({_quote(rule.text)},{tl}))"))
    if t == "beginsWith":
        return simple("beginsWith", f' operator="beginsWith" text="{esc_attr(rule.text)}"', fx(f"LEFT({tl},LEN({_quote(rule.text)}))={_quote(rule.text)}"))
    if t == "endsWith":
        return simple("endsWith", f' operator="endsWith" text="{esc_attr(rule.text)}"', fx(f"RIGHT({tl},LEN({_quote(rule.text)}))={_quote(rule.text)}"))
    if t == "containsBlanks":
        return simple("containsBlanks", "", fx(f"LEN(TRIM({tl}))=0"))
    if t == "notContainsBlanks":
        return simple("notContainsBlanks", "", fx(f"LEN(TRIM({tl}))>0"))
    if t == "containsErrors":
        return simple("containsErrors", "", fx(f"ISERROR({tl})"))
    if t == "notContainsErrors":
        return simple("notContainsErrors", "", fx(f"NOT(ISERROR({tl}))"))
    if t in ("duplicateValues", "uniqueValues"):
        return simple(t, "", "")
    if t == "top10":
        extra = (' percent="1"' if rule.percent else "") + (' bottom="1"' if rule.bottom else "") + f' rank="{rule.rank}"'
        return simple("top10", extra, "")
    if t == "aboveAverage":
        extra = (' aboveAverage="0"' if rule.below else "") + (' equalAverage="1"' if rule.equal_average else "") + (f' stdDev="{rule.std_dev}"' if rule.std_dev else "")
        return simple("aboveAverage", extra, "")
    if t == "timePeriod":
        d = f"FLOOR({tl},1)"
        bodies = {
            "yesterday": f"{d}=TODAY()-1",
            "tomorrow": f"{d}=TODAY()+1",
            "last7Days": f"AND(TODAY()-{d}<=6,{d}<=TODAY())",
            "thisMonth": f"AND(MONTH({tl})=MONTH(TODAY()),YEAR({tl})=YEAR(TODAY()))",
            "lastMonth": f"AND(MONTH({tl})=MONTH(EDATE(TODAY(),0-1)),YEAR({tl})=YEAR(EDATE(TODAY(),0-1)))",
            "nextMonth": f"AND(MONTH({tl})=MONTH(EDATE(TODAY(),0+1)),YEAR({tl})=YEAR(EDATE(TODAY(),0+1)))",
            "thisWeek": f"AND(TODAY()-ROUNDDOWN({tl},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({tl},0)-TODAY()<=7-WEEKDAY(TODAY()))",
            "lastWeek": f"AND(TODAY()-ROUNDDOWN({tl},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({tl},0)<(WEEKDAY(TODAY())+7))",
            "nextWeek": f"AND(ROUNDDOWN({tl},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({tl},0)-TODAY()<(15-WEEKDAY(TODAY())))",
        }
        return simple("timePeriod", f' timePeriod="{esc_attr(rule.period)}"', fx(bodies.get(rule.period, f"{d}=TODAY()")))
    if t == "colorScale":
        b = "<colorScale>" + "".join(_cfvo_xml(v) for v, _ in rule.stops) + "".join(f"<color {color_attrs(c)}/>" for _, c in rule.stops) + "</colorScale>"
        return f'<cfRule type="colorScale" priority="{priority}"{stop}>{b}</cfRule>'
    if t == "dataBar":
        color = 'rgb="FF638EC6"' if rule.bar_color.is_auto else color_attrs(rule.bar_color)
        sv = "" if rule.show_value else ' showValue="0"'
        return f'<cfRule type="dataBar" priority="{priority}"{stop}><dataBar{sv}>{_cfvo_xml(rule.bar_min)}{_cfvo_xml(rule.bar_max)}<color {color}/></dataBar></cfRule>'
    if t == "iconSet":
        b = [f'<iconSet iconSet="{esc_attr(rule.icon_set)}"']
        if rule.reverse:
            b.append(' reverse="1"')
        if not rule.show_value:
            b.append(' showValue="0"')
        b.append(">")
        k = icon_count(rule.icon_set)
        th = rule.thresholds
        if len(th) == k:
            cfvos = th
        elif len(th) + 1 == k:
            cfvos = default_icon_thresholds(rule.icon_set)[:1] + list(th)
        else:
            cfvos = default_icon_thresholds(rule.icon_set)
        b.append("".join(_cfvo_xml(x) for x in cfvos))
        b.append("</iconSet>")
        return f'<cfRule type="iconSet" priority="{priority}"{stop}>{"".join(b)}</cfRule>'
    return simple("expression", "", fx("FALSE"))


def _validation_formula(f: str) -> str:
    f = f.strip()
    if f.startswith('"'):
        return esc(f)
    return esc(text_to_file(f[1:] if f.startswith("=") else f))


def _validation_xml(v: Validation) -> str:
    ty = {"any": "none"}.get(v.kind, v.kind)
    x = [f'<dataValidation type="{ty}"']
    if v.error_style in ("warning", "information"):
        x.append(f' errorStyle="{v.error_style}"')
    if v.operator != "between" and v.kind not in ("list", "custom", "any"):
        x.append(f' operator="{v.operator}"')
    if v.allow_blank:
        x.append(' allowBlank="1"')
    if not v.in_cell_dropdown:
        x.append(' showDropDown="1"')
    if v.show_input:
        x.append(' showInputMessage="1"')
    if v.show_error:
        x.append(' showErrorMessage="1"')
    for name, t in (("errorTitle", v.error_title), ("error", v.error_message), ("promptTitle", v.input_title), ("prompt", v.input_message)):
        if t:
            x.append(f' {name}="{esc_attr(t)}"')
    x.append(f' sqref="{_sqref(v.ranges)}">')
    if v.f1:
        x.append(f"<formula1>{_validation_formula(v.f1)}</formula1>")
    if v.f2:
        x.append(f"<formula2>{_validation_formula(v.f2)}</formula2>")
    x.append("</dataValidation>")
    return "".join(x)


_OP_NAMES = {"<": "lessThan", "<=": "lessThanOrEqual", "<>": "notEqual", ">=": "greaterThanOrEqual", ">": "greaterThan"}


def _autofilter_xml(rng: RangeRef, criteria, out) -> str:
    if not criteria:
        return f'<autoFilter ref="{rng.a1()}"/>'
    s = [f'<autoFilter ref="{rng.a1()}">']
    for col, c in criteria:
        s.append(f'<filterColumn colId="{col}">')
        if c.kind == "values":
            s.append(f'<filters{" blank=" + chr(34) + "1" + chr(34) if c.blanks else ""}>')
            s.extend(f'<filter val="{esc_attr(v)}"/>' for v in c.values)
            s.append("</filters>")
        elif c.kind == "custom":
            s.append(f'<customFilters{" and=" + chr(34) + "1" + chr(34) if c.and_ else ""}>')
            for op, val in c.conditions[:2]:
                o = _OP_NAMES.get(op.strip())
                s.append(f'<customFilter{" operator=" + chr(34) + o + chr(34) if o else ""} val="{esc_attr(val)}"/>')
            s.append("</customFilters>")
        elif c.kind == "top10":
            s.append(f'<top10{" top=" + chr(34) + "0" + chr(34) if c.bottom else ""}{" percent=" + chr(34) + "1" + chr(34) if c.percent else ""} val="{c.count}"/>')
        elif c.kind in ("aboveAverage", "belowAverage"):
            s.append(f'<dynamicFilter type="{c.kind}"/>')
        elif c.kind == "fillColor":
            st = Style(fill=Fill.solid(c.color))
            s.append(f'<colorFilter dxfId="{dxf_id(out.dxfs, st)}"/>')
        elif c.kind == "fontColor":
            from dataclasses import replace

            st = replace(Style(), font=replace(Style().font, color=c.color))
            s.append(f'<colorFilter dxfId="{dxf_id(out.dxfs, st)}" cellColor="0"/>')
        s.append("</filterColumn>")
    s.append("</autoFilter>")
    return "".join(s)


def _table_xml(t: Table, cols: List[str], name: str, tid: int, out) -> str:
    s = [XML_DECL, f'<table xmlns="{NS_MAIN}" id="{tid}" name="{esc_attr(name)}" displayName="{esc_attr(name)}" ref="{t.range.a1()}"']
    if not t.header_row:
        s.append(' headerRowCount="0"')
    s.append(' totalsRowCount="1"' if t.totals_row else ' totalsRowShown="0"')
    s.append(">")
    if t.header_row and t.filter_button:
        end_row = max(t.range.r1 - 1, t.range.r0) if t.totals_row else t.range.r1
        s.append(_autofilter_xml(RangeRef(t.range.r0, t.range.c0, end_row, t.range.c1), t.filter, out))
    s.append(f'<tableColumns count="{len(cols)}">')
    for i, cname in enumerate(cols):
        col = t.columns[i] if i < len(t.columns) else None
        x = [f'<tableColumn id="{i + 1}" name="{esc_attr(encode_xstring(cname))}"']
        totals = col.totals if col is not None else "none"
        if t.totals_row:
            if totals != "none":
                x.append(f' totalsRowFunction="{totals}"')
            elif col is not None and col.totals_label:
                x.append(f' totalsRowLabel="{esc_attr(col.totals_label)}"')
        calc = col.formula if col is not None and col.formula else None
        custom = col.totals_label if t.totals_row and totals == "custom" and col is not None else None
        if not calc and not custom:
            x.append("/>")
            s.append("".join(x))
            continue
        x.append(">")
        if calc:
            x.append(f"<calculatedColumnFormula>{esc(text_to_file(calc, name))}</calculatedColumnFormula>")
        if custom:
            x.append(f"<totalsRowFormula>{esc(text_to_file(custom, name))}</totalsRowFormula>")
        x.append("</tableColumn>")
        s.append("".join(x))
    s.append("</tableColumns>")
    style = t.style or "TableStyleMedium2"
    s.append(
        f'<tableStyleInfo name="{esc_attr(style)}" showFirstColumn="{int(t.first_col)}" showLastColumn="{int(t.last_col)}" '
        f'showRowStripes="{int(t.banded_rows)}" showColumnStripes="{int(t.banded_cols)}"/></table>'
    )
    return "".join(s)


def _comments_xml(sheet: Sheet) -> str:
    authors: List[str] = []
    for c in sheet.comments.values():
        if c.author not in authors:
            authors.append(c.author)
    s = [XML_DECL, f'<comments xmlns="{NS_MAIN}"><authors>']
    s.extend(f"<author>{esc(a)}</author>" for a in authors)
    s.append("</authors><commentList>")
    for cell in sorted(sheet.comments):
        c = sheet.comments[cell]
        text = c.text
        for who, reply in c.replies:
            text += f"\n{who}: {reply}"
        s.append(f'<comment ref="{a1(*cell)}" authorId="{authors.index(c.author)}"><text><r>{t_el(text)}</r></text></comment>')
    s.append("</commentList></comments>")
    return "".join(s)


def _vml_xml(sheet: Sheet, n: int) -> str:
    s = [
        '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">'
        f'<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="{n}"/></o:shapelayout>'
        '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
    ]
    for i, cell in enumerate(sorted(sheet.comments)):
        c = sheet.comments[cell]
        sid = n * 1024 + 1 + i
        vis = "visible" if c.visible else "hidden"
        s.append(
            f'<v:shape id="_x0000_s{sid}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:{i + 1};visibility:{vis}" fillcolor="#ffffe1" o:insetmode="auto">'
            '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>'
            '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>'
            f'<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{cell[1] + 1}, 15, {cell[0]}, 2, {cell[1] + 3}, 15, {cell[0] + 4}, 16</x:Anchor>'
            f'<x:AutoFill>False</x:AutoFill><x:Row>{cell[0]}</x:Row><x:Column>{cell[1]}</x:Column>{"<x:Visible/>" if c.visible else ""}</x:ClientData></v:shape>'
        )
    s.append("</xml>")
    return "".join(s)
