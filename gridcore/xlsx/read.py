"""Reading the workbook part and everything hanging off it."""

from __future__ import annotations

import io
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

from ..core.addr import MAX_ROWS, RangeRef, letters_to_col
from ..core.date import DateSystem
from ..model.features import DefinedName
from ..model.sheet import Sheet
from ..model.style import Style, StyleTable
from ..model.workbook import Workbook
from .fmla import from_file
from .package import Package, Rel, XlsxError
from .sheet_read import read_sheet
from .styles import read_styles
from .tables import INDEXED
from .theme import read_theme
from .xml import attr, attr_float, attr_int, child, flag, kids, local, path, rich_text


class ReadReport:
    def __init__(self) -> None:
        self.warnings: List[str] = []

    def warn(self, m: str) -> None:
        if len(self.warnings) < 1000 and m not in self.warnings:
            self.warnings.append(m)


class Ctx:
    """State shared while reading one package."""

    def __init__(self, pkg: Package):
        self.pkg = pkg
        self.report = ReadReport()
        self.sst: List[str] = []
        self.xf_map: List[int] = []
        self.dxfs: List[Style] = []
        self.palette: List[int] = list(INDEXED)
        self.date_system = DateSystem.D1900
        self.styles = StyleTable()
        self.persons: Dict[str, str] = {}
        self.next_id = 1

    def warn(self, m: str) -> None:
        self.report.warn(m)

    def next_object_id(self) -> int:
        i = self.next_id
        self.next_id += 1
        return i

    def style(self, xf: Optional[int]) -> int:
        if xf is None or not 0 <= xf < len(self.xf_map):
            return 0
        return self.xf_map[xf]

    def optional_xml(self, part: str) -> Optional[ET.Element]:
        try:
            x = self.pkg.read_xml(part)
        except XlsxError as e:
            if "too large" in str(e):
                raise
            self.warn(f"skipped {part}: {e}")
            return None
        if x is None:
            self.warn(f"missing part {part}")
        return x


def _find_rel(rels: List[Rel], kind: str) -> Optional[Rel]:
    for r in rels:
        if r.kind == kind:
            return r
    return None


def read_xlsx(data: bytes) -> Tuple[Workbook, List[str]]:
    pkg = Package(data)
    cx = Ctx(pkg)
    root_rels = pkg.rels("")
    r = _find_rel(root_rels, "officeDocument")
    wb_part = r.target if r is not None else "xl/workbook.xml"
    wb_xml = pkg.read_xml(wb_part)
    if wb_xml is None:
        if pkg.has("xl/workbook.xml") and wb_part != "xl/workbook.xml":
            wb_part = "xl/workbook.xml"
            wb_xml = pkg.read_xml(wb_part)
        elif pkg.has("content.xml"):
            raise XlsxError("this is an OpenDocument file, not XLSX")
    if wb_xml is None:
        raise XlsxError("no workbook part")
    if local(wb_xml.tag) != "workbook":
        raise XlsxError(f"unexpected root element <{local(wb_xml.tag)}> in workbook part")
    wb_rels = pkg.rels(wb_part)
    wb = Workbook(empty=True)

    sr = _find_rel(wb_rels, "styles")
    if sr is not None:
        x = cx.optional_xml(sr.target)
        if x is not None:
            st = read_styles(x)
            cx.palette = st.palette
            cx.dxfs = st.dxfs
            # The default style (xf 0) defines what "Normal" means in this workbook.
            if st.xfs:
                base = st.xfs[0]
                table = StyleTable(base)
                remap = {}
                for i, s in enumerate(st.xfs):
                    remap[i] = table.intern(s)
                cx.styles = table
                cx.xf_map = [remap[i] for i in range(len(st.xfs))]
            wb.cell_styles = st.named
    else:
        cx.warn("no styles part")
    tr = _find_rel(wb_rels, "theme")
    if tr is not None:
        x = cx.optional_xml(tr.target)
        if x is not None:
            wb.theme = read_theme(x)
    ssr = _find_rel(wb_rels, "sharedStrings")
    if ssr is not None:
        b = pkg.read(ssr.target)
        if b is None:
            cx.warn(f"missing part {ssr.target}")
        else:
            sst: List[str] = []
            try:
                for ev, el in ET.iterparse(io.BytesIO(b), events=("end",)):
                    if local(el.tag) == "si":
                        sst.append(rich_text(el))
                        el.clear()
            except ET.ParseError as e:
                cx.warn(f"shared strings are damaged ({e}); kept {len(sst)} strings")
            cx.sst = sst
    pr = _find_rel(wb_rels, "person")
    if pr is not None:
        try:
            x = pkg.read_xml(pr.target)
        except XlsxError:
            x = None
        if x is not None:
            for p in kids(x, "person"):
                if attr(p, "id") and attr(p, "displayName"):
                    cx.persons[attr(p, "id")] = attr(p, "displayName")

    wp = child(wb_xml, "workbookPr")
    if wp is not None and flag(wp, "date1904", False):
        cx.date_system = DateSystem.D1904
    if wp is not None and attr(wp, "codeName"):
        wb.code_name = attr(wp, "codeName") or ""
    wb.date_system = cx.date_system
    cp = child(wb_xml, "calcPr")
    if cp is not None:
        mode = attr(cp, "calcMode")
        wb.calc.mode = "manual" if mode == "manual" else ("autoNoTable" if mode == "autoNoTable" else "auto")
        wb.calc.iterative = flag(cp, "iterate", False)
        n = attr_int(cp, "iterateCount")
        if n is not None:
            wb.calc.max_iterations = min(max(n, 1), 32767)
        d = attr_float(cp, "iterateDelta")
        if d is not None:
            wb.calc.max_change = abs(d)
        wb.calc.precision_as_displayed = not flag(cp, "fullPrecision", True)
    wpr = child(wb_xml, "workbookProtection")
    if wpr is not None:
        wb.protected_structure = flag(wpr, "lockStructure", False)
        wb.structure_password = attr(wpr, "workbookPassword")
    wv = path(wb_xml, "bookViews", "workbookView")
    active_tab = (attr_int(wv, "activeTab") or 0) if wv is not None else 0

    file_to_model: List[Optional[int]] = []
    sheets_el = child(wb_xml, "sheets")
    sheet_els = list(kids(sheets_el, "sheet")) if sheets_el is not None else []
    for i, se in enumerate(sheet_els):
        name = attr(se, "name") or ""
        rid = attr(se, "id")
        rel = next((r for r in wb_rels if r.id == rid), None)
        if rel is None:
            cx.warn(f'sheet "{name}" has no part; skipped')
            file_to_model.append(None)
            continue
        if rel.kind != "worksheet":
            cx.warn(f'sheet "{name}" is a {rel.kind} and was skipped')
            file_to_model.append(None)
            continue
        name = _unique_name(wb, name, i)
        try:
            sheet = read_sheet(cx, rel.target, name)
        except XlsxError as e:
            if "too large" in str(e):
                raise
            cx.warn(f'sheet "{name}" could not be read: {e}')
            sheet = Sheet(name)
        state = attr(se, "state")
        sheet.visibility = state if state in ("hidden", "veryHidden") else "visible"
        file_to_model.append(len(wb.sheets))
        wb.add_sheet(sheet)
    if not wb.sheets:
        cx.warn("the workbook has no worksheets; added an empty sheet")
        wb.add_sheet(Sheet("Sheet1"))
    wb.active_sheet = file_to_model[active_tab] if 0 <= active_tab < len(file_to_model) and file_to_model[active_tab] is not None else 0

    dn = child(wb_xml, "definedNames")
    if dn is not None:
        for d in kids(dn, "definedName"):
            name = attr(d, "name")
            if not name:
                continue
            ls = attr_int(d, "localSheetId")
            scope = None
            if ls is not None:
                scope = file_to_model[ls] if 0 <= ls < len(file_to_model) else None
                if scope is None:
                    continue
            text = (d.text or "").strip()
            lname = name.lower()
            if lname == "_xlnm._filterdatabase" or lname.startswith("_xlchart."):
                continue
            if lname in ("_xlnm.print_area", "_xlnm.print_titles"):
                if scope is not None:
                    sh = wb.sheets[scope]
                    if lname == "_xlnm.print_area":
                        sh.print.print_area = parse_print_area(text)
                    else:
                        sh.print.title_rows, sh.print.title_cols = parse_print_titles(text)
                continue
            wb.names.append(DefinedName(name=name, formula=from_file(text), scope=scope, comment=attr(d, "comment") or "", hidden=flag(d, "hidden", False)))

    cpr = _find_rel(root_rels, "core-properties")
    if cpr is not None:
        try:
            x = pkg.read_xml(cpr.target)
        except XlsxError:
            x = None
        if x is not None:
            def get(n: str) -> str:
                c = child(x, n)
                return (c.text or "").strip() if c is not None else ""

            wb.props.title = get("title")
            wb.props.subject = get("subject")
            wb.props.author = get("creator")
            wb.props.keywords = get("keywords")
            wb.props.description = get("description")
    epr = _find_rel(root_rels, "extended-properties")
    if epr is not None:
        try:
            x = pkg.read_xml(epr.target)
        except XlsxError:
            x = None
        if x is not None:
            c = child(x, "Company")
            wb.props.company = (c.text or "").strip() if c is not None else ""
    vr = _find_rel(wb_rels, "vbaProject")
    if vr is not None:
        _read_vba(cx, wb, vr.target)
    wb.styles = cx.styles
    return wb, cx.report.warnings


# Parts that sign a VBA project: relationship type → content type.
SIGNATURE_TYPES = {
    "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature": "application/vnd.ms-office.vbaProjectSignature",
    "http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile": "application/vnd.ms-office.vbaProjectSignatureAgile",
    "http://schemas.microsoft.com/office/2020/07/relationships/vbaProjectSignatureV3": "application/vnd.ms-office.vbaProjectSignatureV3",
}


def _read_vba(cx: Ctx, wb: Workbook, part: str) -> None:
    """The workbook's VBA project. One that can't be read is still kept, byte for byte."""
    from ..model.vba import VbaProject
    from .ovba import OvbaError, read_project

    data = cx.pkg.read(part)
    if data is None:
        cx.warn(f"missing part {part} (the VBA project)")
        return
    sigs = []
    for r in cx.pkg.rels(part):
        if r.type in SIGNATURE_TYPES and not r.external:
            b = cx.pkg.read(r.target)
            if b is not None:
                sigs.append((r.type, r.target.rsplit("/", 1)[-1], b))
    try:
        project = read_project(data)
    except (OvbaError, ValueError, KeyError, IndexError) as e:
        cx.warn(f"the VBA project can't be read ({e}); it is kept as it is but can't be edited or run")
        project = VbaProject(base=data, unreadable=str(e) or "unreadable")
    for w in project.warnings:
        cx.warn(f"VBA: {w}")
    from dataclasses import replace

    wb.vba = replace(project, signatures=tuple(sigs))


def _unique_name(wb: Workbook, name: str, i: int) -> str:
    base = "".join(c for c in name if c not in ':\\/?*[]')[:31]
    if not base.strip():
        base = f"Sheet{i + 1}"
    if wb.sheet_index(base) is None:
        return base
    n = 2
    while True:
        cand = f"{base[:25]} ({n})"
        if wb.sheet_index(cand) is None:
            return cand
        n += 1


def _strip_sheet(s: str) -> str:
    i = s.rfind("!")
    return s[i + 1:] if i >= 0 else s


def parse_print_area(text: str) -> Optional[RangeRef]:
    first = text.split(",")[0].strip()
    return RangeRef.parse(_strip_sheet(first).replace("$", ""))


def parse_print_titles(text: str):
    rows = cols = None
    for part in text.split(","):
        p = _strip_sheet(part.strip()).replace("$", "")
        if ":" not in p:
            continue
        a, b = p.split(":", 1)
        if a.isdigit() and b.isdigit():
            r0, r1 = int(a), int(b)
            if 1 <= r0 <= r1 <= MAX_ROWS:
                rows = (r0 - 1, r1 - 1)
        else:
            c0, c1 = letters_to_col(a), letters_to_col(b)
            if c0 is not None and c1 is not None and c1 >= c0:
                cols = (c0, c1)
    return rows, cols
