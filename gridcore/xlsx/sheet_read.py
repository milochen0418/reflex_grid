"""Reading worksheet parts."""

from __future__ import annotations

import io
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef, parse_cell, parse_ranges
from ..core.date import DateSystem, serial_from_ymd, time_fraction
from ..core.value import CellError, number
from ..formula import ParseError, parse
from ..formula.adjust import shift_relative
from ..model.cell import Cell, Formula
from ..model.features import (
    AutoFilter,
    CfRule,
    CfValue,
    Comment,
    CondFormat,
    FilterCriterion,
    Hyperlink,
    LineInfo,
    SheetProtection,
    Sparkline,
    Table,
    TableColumn,
    Validation,
)
from ..model.sheet import Sheet
from .fmla import from_file
from .package import Rel, XlsxError
from .styles import read_color
from .tables import col_width_to_px, paper_name, pt_to_px
from .xml import attr, attr_float, attr_int, child, decode_xstring, flag, kids, local, path, rich_text, text_of

MAX_HYPERLINK_CELLS = 10_000


class _Shared:
    __slots__ = ("anchor", "expr", "text")

    def __init__(self, anchor, expr, text):
        self.anchor = anchor
        self.expr = expr
        self.text = text


class _RowState:
    def __init__(self) -> None:
        self.next_row = 0
        self.shared: Dict[int, _Shared] = {}
        # Dynamic-array spill ranges and legacy array ranges: cached values inside them (other
        # than the anchor) are not stored, so the formula can fill them again.
        self.dynamic: List[RangeRef] = []
        self.dropped = 0
        self.rows = 0


def read_sheet(cx, part: str, name: str) -> Sheet:
    data = cx.pkg.read(part)
    if data is None:
        raise XlsxError(f"missing part {part}")
    rels = cx.pkg.rels(part)
    sheet = Sheet(name)
    st = _RowState()
    others: List[ET.Element] = []
    depth = 0
    try:
        for ev, el in ET.iterparse(io.BytesIO(data), events=("start", "end")):
            if ev == "start":
                depth += 1
                continue
            depth -= 1
            tag = local(el.tag)
            if depth == 2 and tag == "row":
                _read_row(cx, sheet, st, el)
                el.clear()
            elif depth == 1:
                if tag == "sheetData":
                    el.clear()
                else:
                    others.append(el)
    except ET.ParseError as e:
        cx.warn(f'sheet "{name}" is damaged ({e}); kept {st.rows} rows')
    if st.dropped:
        cx.warn(f'sheet "{name}": {st.dropped} cells or rows outside the 1,048,576 × 16,384 grid were dropped')
    fit_to_page = False
    for e in others:
        tag = local(e.tag)
        if tag == "sheetPr":
            sheet.code_name = attr(e, "codeName") or ""
            tc = child(e, "tabColor")
            if tc is not None:
                sheet.tab_color = read_color(tc, cx.palette)
            ps = child(e, "pageSetUpPr")
            fit_to_page = ps is not None and flag(ps, "fitToPage", False)
        elif tag == "sheetViews":
            _read_views(e, sheet)
        elif tag == "sheetFormatPr":
            h = attr_float(e, "defaultRowHeight")
            if h is not None and pt_to_px(h) > 0:
                sheet.default_row_height = pt_to_px(h)
            w = attr_float(e, "defaultColWidth")
            if w is not None and col_width_to_px(w) > 0:
                sheet.default_col_width = col_width_to_px(w)
            else:
                b = attr_float(e, "baseColWidth")
                if b is not None:
                    import math

                    sheet.default_col_width = float(math.ceil((min(max(b, 0.0), 255.0) * 7.0 + 5.0) / 8.0) * 8.0)
        elif tag == "cols":
            _read_cols(cx, e, sheet)
        elif tag == "mergeCells":
            for m in kids(e, "mergeCell"):
                r = RangeRef.parse(attr(m, "ref") or "")
                if r is not None and not r.is_single():
                    sheet.merges.append(r)
        elif tag == "conditionalFormatting":
            _read_cf(cx, e, sheet)
        elif tag == "dataValidations":
            for v in kids(e, "dataValidation"):
                val = _read_validation(v)
                if val is not None:
                    sheet.validations.append(val)
        elif tag == "hyperlinks":
            _read_hyperlinks(cx, e, rels, sheet, name)
        elif tag == "autoFilter":
            sheet.autofilter = read_autofilter(cx, e)
        elif tag == "printOptions":
            p = sheet.print
            p.gridlines = flag(e, "gridLines", False)
            p.headings = flag(e, "headings", False)
            p.center_h = flag(e, "horizontalCentered", False)
            p.center_v = flag(e, "verticalCentered", False)
        elif tag == "pageMargins":
            m = list(sheet.print.margins)
            for i, k in enumerate(("left", "right", "top", "bottom", "header", "footer")):
                v = attr_float(e, k)
                if v is not None:
                    m[i] = min(max(v, 0.0), 50.0)
            sheet.print.margins = tuple(m)
        elif tag == "pageSetup":
            p = sheet.print
            if attr(e, "orientation") == "landscape":
                p.orientation = "landscape"
            ps = attr_int(e, "paperSize")
            if ps is not None and paper_name(ps):
                p.paper = paper_name(ps)
            sc = attr_int(e, "scale")
            if sc is not None:
                p.scale = min(max(sc, 10), 400)
            if fit_to_page:
                w = attr_int(e, "fitToWidth")
                h = attr_int(e, "fitToHeight")
                p.fit_to = (min(w if w is not None else 1, 32767), min(h if h is not None else 1, 32767))
        elif tag == "headerFooter":
            hdr = child(e, "oddHeader")
            ftr = child(e, "oddFooter")
            sheet.print.header = text_of(hdr)[:255]
            sheet.print.footer = text_of(ftr)[:255]
        elif tag == "rowBreaks":
            sheet.print.row_breaks = [i for i in (attr_int(b, "id") for b in kids(e, "brk")) if i and 0 < i < MAX_ROWS]
        elif tag == "colBreaks":
            sheet.print.col_breaks = [i for i in (attr_int(b, "id") for b in kids(e, "brk")) if i and 0 < i < MAX_COLS]
        elif tag == "sheetProtection" and flag(e, "sheet", False):
            hash_attrs = {k: attr(e, k) for k in ("algorithmName", "hashValue", "saltValue", "spinCount") if attr(e, k)}
            sheet.protection = SheetProtection(
                password_hash=attr(e, "password"),
                hash_attrs=hash_attrs,
                select_locked=not flag(e, "selectLockedCells", False),
                select_unlocked=not flag(e, "selectUnlockedCells", False),
                format_cells=not flag(e, "formatCells", True),
                format_columns=not flag(e, "formatColumns", True),
                format_rows=not flag(e, "formatRows", True),
                insert_columns=not flag(e, "insertColumns", True),
                insert_rows=not flag(e, "insertRows", True),
                delete_columns=not flag(e, "deleteColumns", True),
                delete_rows=not flag(e, "deleteRows", True),
                sort=not flag(e, "sort", True),
                autofilter=not flag(e, "autoFilter", True),
            )
        elif tag == "drawing":
            rid = attr(e, "id")
            rel = next((r for r in rels if r.id == rid), None)
            if rel is not None:
                from .drawing import read_drawing

                read_drawing(cx, rel.target, sheet)
        elif tag == "tableParts":
            for tp in kids(e, "tablePart"):
                rid = attr(tp, "id")
                rel = next((r for r in rels if r.id == rid), None)
                if rel is None:
                    continue
                x = cx.optional_xml(rel.target)
                if x is not None:
                    t = read_table(cx, x)
                    if t is not None:
                        t.id = cx.next_object_id()
                        sheet.tables.append(t)
        elif tag == "extLst":
            _read_ext(cx, e, sheet)
    if sheet.tables:
        _short_table_refs(sheet)
    for rel in rels:
        if rel.kind == "comments":
            x = cx.optional_xml(rel.target)
            if x is not None:
                _read_comments(x, sheet)
    for rel in rels:
        if rel.kind == "threadedComment":
            x = cx.optional_xml(rel.target)
            if x is not None:
                _read_threaded(cx, x, sheet)
    if sheet.comments:
        for rel in rels:
            if rel.kind == "vmlDrawing":
                try:
                    x = cx.pkg.read_xml(rel.target)
                except XlsxError:
                    x = None
                if x is not None:
                    _read_note_visibility(x, sheet, 0)
    return sheet


def _read_views(e: ET.Element, sheet: Sheet) -> None:
    v = child(e, "sheetView")
    if v is None:
        return
    sheet.show_gridlines = flag(v, "showGridLines", True)
    sheet.show_headings = flag(v, "showRowColHeaders", True)
    sheet.show_formulas = flag(v, "showFormulas", False)
    sheet.show_zeros = flag(v, "showZeros", True)
    sheet.right_to_left = flag(v, "rightToLeft", False)
    z = attr_int(v, "zoomScale")
    if z is not None:
        sheet.zoom = min(max(z, 10), 400)
    tl = parse_cell(attr(v, "topLeftCell") or "")
    if tl is not None:
        sheet.view_top_left = tl
    active_pane = None
    p = child(v, "pane")
    if p is not None:
        if attr(p, "state") in ("frozen", "frozenSplit"):
            cols = int(min(max(attr_float(p, "xSplit") or 0.0, 0.0), MAX_COLS - 1))
            rows = int(min(max(attr_float(p, "ySplit") or 0.0, 0.0), MAX_ROWS - 1))
            if rows > 0 or cols > 0:
                sheet.freeze = (rows, cols)
                c = parse_cell(attr(p, "topLeftCell") or "")
                if c is not None:
                    sheet.view_top_left = c
        active_pane = attr(p, "activePane")
    sels = list(kids(v, "selection"))
    sel = next((s for s in sels if attr(s, "pane") == active_pane), sels[-1] if sels else None)
    if sel is not None:
        c = parse_cell(attr(sel, "activeCell") or "")
        if c is not None:
            sheet.view_active = c


def _read_cols(cx, e: ET.Element, sheet: Sheet) -> None:
    for c in kids(e, "col"):
        lo, hi = attr_int(c, "min"), attr_int(c, "max")
        if lo is None or hi is None or lo < 1 or lo > MAX_COLS or hi < lo:
            continue
        hi = min(hi, MAX_COLS)
        size = None
        w = attr_float(c, "width")
        if w is not None:
            px = col_width_to_px(w)
            if flag(c, "customWidth", False) or px != sheet.default_col_width:
                size = px
        style = None
        s = attr_int(c, "style")
        if s is not None:
            sid = cx.style(s)
            if sid != 0:
                style = sid
        info = LineInfo(
            size=size,
            hidden=flag(c, "hidden", False),
            outline=min(attr_int(c, "outlineLevel") or 0, 7),
            collapsed=flag(c, "collapsed", False),
            style=style,
            # Widths set by hand stop automatic widening; best-fit widths don't.
            custom=size is not None and not flag(c, "bestFit", False),
        )
        if info.is_default() or (info.size == sheet.default_col_width and LineInfo(None, info.hidden, info.outline, info.collapsed, info.style).is_default()):
            continue
        if hi - lo > 16_384:
            continue
        for col in range(lo - 1, hi):
            sheet.cols[col] = info


def _read_row(cx, sheet: Sheet, st: _RowState, row: ET.Element) -> None:
    st.rows += 1
    r_attr = attr_int(row, "r")
    if r_attr is not None:
        if r_attr < 1:
            st.dropped += 1
            return
        r = r_attr - 1
    else:
        r = st.next_row
    if r >= MAX_ROWS:
        st.dropped += 1
        return
    st.next_row = r + 1
    size = None
    custom = False
    ht = attr_float(row, "ht")
    if ht is not None:
        custom = flag(row, "customHeight", False)
        if custom or abs(pt_to_px(ht) - sheet.default_row_height) > 0.5:
            size = pt_to_px(ht)
    style = None
    if flag(row, "customFormat", False):
        s = attr_int(row, "s")
        if s is not None:
            style = cx.style(s)
    info = LineInfo(size, flag(row, "hidden", False), min(attr_int(row, "outlineLevel") or 0, 7), flag(row, "collapsed", False), style, custom)
    if not info.is_default():
        sheet.rows[r] = info
    next_col = 0
    for c in row:
        if local(c.tag) != "c":
            continue
        a = attr(c, "r")
        if a is not None:
            pos = parse_cell(a)
            if pos is None:
                st.dropped += 1
                continue
        else:
            pos = (r, next_col)
        if pos[0] >= MAX_ROWS or pos[1] >= MAX_COLS:
            st.dropped += 1
            continue
        next_col = pos[1] + 1
        _read_cell(cx, sheet, st, c, pos)


def _read_cell(cx, sheet: Sheet, st: _RowState, c: ET.Element, pos: Tuple[int, int]) -> None:
    style = cx.style(attr_int(c, "s"))
    v_el = None
    f_el = None
    is_el = None
    for k in c:
        n = local(k.tag)
        if n == "v":
            v_el = k
        elif n == "f":
            f_el = k
        elif n == "is":
            is_el = k
    v = v_el.text if v_el is not None and v_el.text is not None else None
    t = attr(c, "t") or "n"
    value = None
    if t == "s":
        if v is not None:
            try:
                i = int(v.strip())
                if 0 <= i < len(cx.sst):
                    value = cx.sst[i]
                else:
                    cx.warn("some cells refer to missing shared strings")
            except ValueError:
                pass
    elif t == "str":
        value = decode_xstring(v) if v is not None else None
    elif t == "inlineStr":
        if is_el is not None:
            value = rich_text(is_el)
        elif v is not None:
            value = decode_xstring(v)
    elif t == "b":
        if v is not None:
            value = v.strip() in ("1", "true", "TRUE")
    elif t == "e":
        if v is not None:
            value = CellError.parse(v.strip()) or CellError.VALUE
    elif t == "d":
        if v is not None and v.strip():
            n = parse_iso_datetime(v.strip(), cx.date_system)
            if n is None:
                cx.warn("some date cells could not be read")
            else:
                value = number(n)
    else:
        if v is not None and v.strip():
            try:
                fv = float(v.strip())
                value = number(fv)
                if isinstance(value, CellError):
                    value = None
            except ValueError:
                cx.warn("some numeric cells could not be read")
    dynamic = attr(c, "cm") is not None
    formula = _read_formula(cx, st, f_el, pos, dynamic) if f_el is not None else None
    if formula is None and any(rg.contains(pos) and rg.start != pos for rg in st.dynamic):
        value = None
    sheet.cells.set(pos, Cell(value, formula, style))


def _read_formula(cx, st: _RowState, f: ET.Element, pos, dynamic: bool) -> Optional[Formula]:
    text = from_file((f.text or "").strip())
    ft = attr(f, "t") or "normal"
    if ft == "shared":
        si = attr_int(f, "si")
        if si is None:
            return None
        if text:
            try:
                e = parse(text)
            except ParseError:
                e = None
            st.shared[si] = _Shared(pos, e, text)
            return Formula.new(text)
        sh = st.shared.get(si)
        if sh is None:
            cx.warn("a shared formula refers to a missing master cell; kept the cached value")
            return None
        if sh.expr is not None:
            return Formula.from_expr(shift_relative(sh.expr, pos[0] - sh.anchor[0], pos[1] - sh.anchor[1]))
        return Formula.new(sh.text)
    if ft == "array":
        if not text:
            return None
        fm = Formula.new(text)
        rng = RangeRef.parse(attr(f, "ref") or "") or RangeRef.of(pos)
        if not rng.is_single():
            st.dynamic.append(rng)
        if not dynamic:
            fm.array = rng
        return fm
    if ft == "dataTable":
        cx.warn("what-if data tables are not supported; kept their values")
        return None
    return Formula.new(text) if text else None


def parse_iso_datetime(s: str, sys: DateSystem) -> Optional[float]:
    s = s.rstrip("Z")
    if "T" in s:
        date, time = s.split("T", 1)
    elif ":" in s:
        date, time = "", s
    else:
        date, time = s, None
    serial = 0.0
    if date:
        parts = date.split("-")
        if len(parts) != 3:
            return None
        try:
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
        except ValueError:
            return None
        if not 1 <= m <= 12 or not 1 <= d <= 31:
            return None
        sv = serial_from_ymd(sys, y, m, d)
        if sv is None:
            return None
        serial = sv
    if time is not None:
        t = time.split("+")[0]
        parts = t.split(":")
        try:
            h = float(parts[0])
            mi = float(parts[1]) if len(parts) > 1 else 0.0
            se = float(parts[2]) if len(parts) > 2 else 0.0
        except ValueError:
            return None
        if not (0 <= h < 24 and 0 <= mi < 60 and 0 <= se < 61):
            return None
        serial += time_fraction(h, mi, se)
    return serial


_CF_OPS = {
    "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual",
}


def _cfvo(e: ET.Element) -> CfValue:
    val = attr(e, "val")
    if val is None:
        fe = child(e, "f")
        val = fe.text if fe is not None else ""
    val = val or ""
    try:
        n = float(val.strip())
    except ValueError:
        n = None
    ty = attr(e, "type") or "num"
    if ty in ("min", "autoMin"):
        return CfValue("min")
    if ty in ("max", "autoMax"):
        return CfValue("max")
    if ty == "percent":
        return CfValue("percent", n) if n is not None else CfValue("formula", from_file(val))
    if ty == "percentile":
        return CfValue("percentile", n) if n is not None else CfValue("formula", from_file(val))
    if ty == "formula":
        return CfValue("formula", from_file(val))
    return CfValue("num", n) if n is not None else CfValue("formula", from_file(val))


def _read_cf(cx, e: ET.Element, sheet: Sheet) -> None:
    # Every range is kept: the list is no longer than the file's text.
    ranges = parse_ranges(attr(e, "sqref") or "")
    if not ranges:
        return
    for r in kids(e, "cfRule"):
        di = attr_int(r, "dxfId")
        style = cx.dxfs[di] if di is not None and 0 <= di < len(cx.dxfs) else None
        formulas = [from_file((f.text or "").strip()) for f in kids(r, "formula")]
        ty = attr(r, "type") or ""
        rule = CfRule(type=ty, formulas=formulas, style=style, text=attr(r, "text") or "")
        if ty == "cellIs":
            op = attr(r, "operator") or "between"
            rule.operator = op if op in _CF_OPS else "between"
        elif ty == "top10":
            rule.bottom = flag(r, "bottom", False)
            rule.percent = flag(r, "percent", False)
            rule.rank = attr_int(r, "rank") or 10
        elif ty == "aboveAverage":
            rule.below = not flag(r, "aboveAverage", True)
            rule.equal_average = flag(r, "equalAverage", False)
            rule.std_dev = min(attr_int(r, "stdDev") or 0, 3)
        elif ty == "timePeriod":
            rule.period = attr(r, "timePeriod") or "today"
        elif ty == "colorScale":
            cs = child(r, "colorScale")
            if cs is None:
                continue
            vals = [_cfvo(x) for x in kids(cs, "cfvo")]
            cols = [read_color(x, cx.palette) for x in kids(cs, "color")]
            rule.stops = list(zip(vals, cols))
        elif ty == "dataBar":
            db = child(r, "dataBar")
            if db is None:
                continue
            vals = [_cfvo(x) for x in kids(db, "cfvo")]
            rule.bar_min = vals[0] if vals else CfValue("min")
            rule.bar_max = vals[1] if len(vals) > 1 else CfValue("max")
            ce = child(db, "color")
            if ce is not None:
                rule.bar_color = read_color(ce, cx.palette)
            rule.show_value = flag(db, "showValue", True)
        elif ty == "iconSet":
            ic = child(r, "iconSet")
            if ic is None:
                continue
            rule.icon_set = attr(ic, "iconSet") or "3TrafficLights1"
            rule.thresholds = [_cfvo(x) for x in kids(ic, "cfvo")]
            rule.reverse = flag(ic, "reverse", False)
            rule.show_value = flag(ic, "showValue", True)
        elif ty not in (
            "expression", "containsText", "notContainsText", "beginsWith", "endsWith", "containsBlanks",
            "notContainsBlanks", "containsErrors", "notContainsErrors", "duplicateValues", "uniqueValues",
        ):
            cx.warn(f'conditional format type "{ty}" is not supported')
            continue
        sheet.cond_formats.append(CondFormat(list(ranges), rule, attr_int(r, "priority") or 0, flag(r, "stopIfTrue", False)))


def _read_validation(v: ET.Element) -> Optional[Validation]:
    sq = attr(v, "sqref")
    if sq is None:
        se = child(v, "sqref")
        sq = se.text if se is not None else ""
    ranges = parse_ranges(sq or "")
    if not ranges:
        return None

    def formula(n: str) -> Optional[str]:
        fe = child(v, n)
        if fe is None:
            return None
        inner = child(fe, "f")
        t = (inner.text if inner is not None else fe.text) or ""
        t = t.strip()
        return t or None

    ty = attr(v, "type") or "none"
    kind = ty if ty in ("whole", "decimal", "list", "date", "time", "textLength", "custom") else "any"
    f1 = formula("formula1") or ""
    if kind == "list" and f1 and not f1.startswith('"'):
        f1 = "=" + from_file(f1)
    elif f1 and not f1.startswith('"'):
        f1 = from_file(f1)
    f2 = formula("formula2")
    op = attr(v, "operator") or "between"
    return Validation(
        ranges=ranges,
        kind=kind,
        operator=op if op in _CF_OPS or op == "between" else "between",
        f1=f1,
        f2=from_file(f2) if f2 else None,
        allow_blank=flag(v, "allowBlank", False),
        in_cell_dropdown=not flag(v, "showDropDown", False),
        input_title=attr(v, "promptTitle") or "",
        input_message=attr(v, "prompt") or "",
        show_input=flag(v, "showInputMessage", False),
        error_title=attr(v, "errorTitle") or "",
        error_message=attr(v, "error") or "",
        error_style=attr(v, "errorStyle") if attr(v, "errorStyle") in ("warning", "information") else "stop",
        show_error=flag(v, "showErrorMessage", False),
    )


def _read_hyperlinks(cx, e: ET.Element, rels: List[Rel], sheet: Sheet, name: str) -> None:
    cells = 0
    for h in kids(e, "hyperlink"):
        rng = RangeRef.parse(attr(h, "ref") or "")
        if rng is None:
            continue
        rid = attr(h, "id")
        target = next((r.target for r in rels if r.id == rid), "") if rid else ""
        loc = attr(h, "location")
        if loc:
            target = loc.lstrip("#") if not target else f"{target}#{loc}"
        if not target:
            continue
        link = Hyperlink(target, attr(h, "tooltip"), external=bool(rid) and not target.startswith("#"))
        for rc in rng.cells():
            if cells >= MAX_HYPERLINK_CELLS:
                cx.warn(f'sheet "{name}": hyperlinks past the first {MAX_HYPERLINK_CELLS:,} cells were dropped')
                return
            cells += 1
            sheet.hyperlinks[rc] = link


_FILTER_OPS = {"lessThan": "<", "lessThanOrEqual": "<=", "notEqual": "<>", "greaterThanOrEqual": ">=", "greaterThan": ">"}


def read_autofilter(cx, e: ET.Element) -> Optional[AutoFilter]:
    rng = RangeRef.parse(attr(e, "ref") or "")
    if rng is None:
        return None
    criteria = []
    for fc in kids(e, "filterColumn"):
        col = attr_int(fc, "colId")
        if col is None:
            continue
        fl = child(fc, "filters")
        cf = child(fc, "customFilters")
        t10 = child(fc, "top10")
        dyn = child(fc, "dynamicFilter")
        colf = child(fc, "colorFilter")
        if fl is not None:
            crit = FilterCriterion("values", values=[attr(x, "val") or "" for x in kids(fl, "filter")], blanks=flag(fl, "blank", False))
        elif cf is not None:
            conds = [(_FILTER_OPS.get(attr(x, "operator") or "", "="), attr(x, "val") or "") for x in kids(cf, "customFilter")]
            if not conds:
                continue
            crit = FilterCriterion("custom", conditions=conds[:2], and_=flag(cf, "and", False))
        elif t10 is not None:
            crit = FilterCriterion("top10", bottom=not flag(t10, "top", True), percent=flag(t10, "percent", False), count=int(min(max(attr_float(t10, "val") or 10.0, 0.0), 1e9)))
        elif dyn is not None:
            ty = attr(dyn, "type")
            if ty == "aboveAverage":
                crit = FilterCriterion("aboveAverage")
            elif ty == "belowAverage":
                crit = FilterCriterion("belowAverage")
            else:
                cx.warn("date filters in AutoFilters are not supported")
                continue
        elif colf is not None:
            di = attr_int(colf, "dxfId")
            if di is None or not 0 <= di < len(cx.dxfs):
                continue
            d = cx.dxfs[di]
            if flag(colf, "cellColor", True):
                crit = FilterCriterion("fillColor", color=d.fill.fg)
            else:
                crit = FilterCriterion("fontColor", color=d.font.color)
        else:
            continue
        criteria.append((col, crit))
    return AutoFilter(rng, criteria)


_TOTALS = {"average", "count", "countNums", "max", "min", "sum", "stdDev", "var", "custom"}


def _short_table_refs(sheet: Sheet) -> None:
    """Formulas inside a table show their own table's this-row references as ``[@Col]``."""
    from .fmla import unqualify_own_table

    for t in sheet.tables:
        for rc, cell in list(sheet.cells.iter_range(t.range)):
            f = cell.formula
            if f is None or "[" not in f.text:
                continue
            e = f.get_expr()
            if e is None:
                continue
            ne = unqualify_own_table(e, t.name)
            if ne != e:
                sheet.cells.set(rc, Cell(cell.value, Formula.from_expr(ne, f.array), cell.style))


def read_table(cx, x: ET.Element) -> Optional[Table]:
    rng = RangeRef.parse(attr(x, "ref") or "")
    if rng is None:
        return None
    columns = []
    tc = child(x, "tableColumns")
    if tc is not None:
        for c in kids(tc, "tableColumn"):
            tf = attr(c, "totalsRowFunction") or "none"
            totals = tf if tf in _TOTALS else "none"
            if totals == "custom":
                trf = child(c, "totalsRowFormula")
                label = from_file((trf.text or "").strip()) if trf is not None else None
            else:
                label = attr(c, "totalsRowLabel")
            ccf = child(c, "calculatedColumnFormula")
            columns.append(
                TableColumn(
                    name=decode_xstring(attr(c, "name") or ""),
                    totals=totals,
                    totals_label=label,
                    formula=from_file((ccf.text or "").strip()) if ccf is not None else None,
                )
            )
    si = child(x, "tableStyleInfo")
    af = child(x, "autoFilter")
    af_read = read_autofilter(cx, af) if af is not None else None
    return Table(
        id=attr_int(x, "id") or 0,
        name=attr(x, "displayName") or attr(x, "name") or "Table",
        range=rng,
        header_row=(attr_int(x, "headerRowCount") if attr_int(x, "headerRowCount") is not None else 1) > 0,
        totals_row=(attr_int(x, "totalsRowCount") or 0) > 0,
        columns=columns,
        style=(attr(si, "name") if si is not None else None) or "",
        banded_rows=si is not None and flag(si, "showRowStripes", False),
        banded_cols=si is not None and flag(si, "showColumnStripes", False),
        first_col=si is not None and flag(si, "showFirstColumn", False),
        last_col=si is not None and flag(si, "showLastColumn", False),
        filter_button=af is not None,
        filter=af_read.criteria if af_read is not None else [],
    )


def _read_comments(x: ET.Element, sheet: Sheet) -> None:
    authors_el = child(x, "authors")
    authors = [a.text or "" for a in kids(authors_el, "author")] if authors_el is not None else []
    lst = child(x, "commentList")
    if lst is None:
        return
    for c in kids(lst, "comment"):
        cell = parse_cell(attr(c, "ref") or "")
        if cell is None:
            continue
        ai = attr_int(c, "authorId")
        author = authors[ai] if ai is not None and 0 <= ai < len(authors) else ""
        te = child(c, "text")
        text = rich_text(te) if te is not None else ""
        if author and text.startswith(author + ":"):
            text = text[len(author) + 1:]
            if text.startswith("\n"):
                text = text[1:]
        if author.startswith("tc="):
            author = ""
        sheet.comments[cell] = Comment(author, text)


def _read_note_visibility(x: ET.Element, sheet: Sheet, depth: int) -> None:
    if depth > 64:
        return
    for e in x:
        if local(e.tag) != "shape":
            _read_note_visibility(e, sheet, depth + 1)
            continue
        cd = child(e, "ClientData")
        if cd is None or attr(cd, "ObjectType") != "Note" or child(cd, "Visible") is None:
            continue
        try:
            row = int((text_of(child(cd, "Row"))).strip())
            col = int((text_of(child(cd, "Column"))).strip())
        except ValueError:
            continue
        c = sheet.comments.get((row, col))
        if c is not None and not c.threaded:
            from dataclasses import replace

            sheet.comments[(row, col)] = replace(c, visible=True)


def _read_threaded(cx, x: ET.Element, sheet: Sheet) -> None:
    from dataclasses import replace

    by_id: Dict[str, Tuple[int, int]] = {}
    replies = []
    for t in kids(x, "threadedComment"):
        cell = parse_cell(attr(t, "ref") or "")
        if cell is None:
            continue
        author = cx.persons.get(attr(t, "personId") or "", "")
        te = child(t, "text")
        text = (te.text or "") if te is not None else ""
        parent = attr(t, "parentId")
        if parent:
            replies.append((parent, author, text))
        else:
            tid = attr(t, "id")
            if tid:
                by_id[tid] = cell
            sheet.comments[cell] = Comment(author, text, (), True, flag(t, "done", False))
    for parent, author, text in replies:
        cell = by_id.get(parent)
        if cell is None:
            continue
        c = sheet.comments.get(cell)
        if c is not None:
            sheet.comments[cell] = replace(c, replies=c.replies + ((author, text),))


def _read_ext(cx, e: ET.Element, sheet: Sheet) -> None:
    for ext in kids(e, "ext"):
        for c in ext:
            n = local(c.tag)
            if n == "dataValidations":
                for v in kids(c, "dataValidation"):
                    val = _read_validation(v)
                    if val is not None:
                        sheet.validations.append(val)
            elif n == "sparklineGroups":
                for g in kids(c, "sparklineGroup"):
                    ty = attr(g, "type")
                    kind = "column" if ty == "column" else ("stacked" if ty == "stacked" else "line")
                    cs = child(g, "colorSeries")
                    color = read_color(cs, cx.palette) if cs is not None else None
                    markers = flag(g, "markers", False)
                    sl = child(g, "sparklines")
                    if sl is None:
                        continue
                    for sp in kids(sl, "sparkline"):
                        fe, sq = child(sp, "f"), child(sp, "sqref")
                        cell = parse_cell((sq.text or "") if sq is not None else "")
                        if cell is not None:
                            from ..model.style import AUTO

                            sheet.sparklines.append(Sparkline(cell[0], cell[1], (fe.text or "").strip() if fe is not None else "", kind, color or AUTO, markers))
            elif n == "conditionalFormattings":
                cx.warn("some Excel 2010 conditional format extensions were skipped")
