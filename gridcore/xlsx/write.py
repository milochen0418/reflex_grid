"""Writing a workbook as an XLSX package."""

from __future__ import annotations

import io
import zipfile
from typing import Dict, List, Optional, Tuple

from ..core.addr import RangeRef
from ..core.date import DateSystem
from ..formula import quote_sheet
from ..model.style import Style
from ..model.workbook import Workbook
from .fmla import text_to_file
from .sheet_write import NS_MAIN, NS_REL, Rels, write_sheet
from .styles import write_styles
from .theme import write_theme
from .xml import XML_DECL, esc, esc_attr, num, t_el

CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
CT_MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
CT_MAIN_MACRO = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
CT_VBA = "application/vnd.ms-office.vbaProject"
REL_VBA = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"

METADATA = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'
)


class Sst:
    def __init__(self) -> None:
        self.list: List[str] = []
        self.map: Dict[str, int] = {}
        self.count = 0

    def index(self, s: str) -> int:
        self.count += 1
        i = self.map.get(s)
        if i is None:
            i = len(self.list)
            self.list.append(s)
            self.map[s] = i
        return i


class Out:
    def __init__(self) -> None:
        self.parts: List[Tuple[str, bytes]] = []
        self.overrides: List[Tuple[str, str]] = []
        self.defaults: Dict[str, str] = {}
        self.dxfs: List[Style] = []
        self.sst = Sst()
        self.images = 0
        self.charts = 0
        self.drawings = 0
        self.tables = 0
        self.comments = 0
        self.table_id = 0
        self.dynamic_arrays = False
        self.needs_calc = False
        # Each sheet's VBA code name (set when the workbook has a VBA project).
        self.code_names: Optional[List[str]] = None

    def part(self, name: str, content_type: Optional[str], data: bytes) -> None:
        if content_type:
            self.overrides.append(("/" + name, content_type))
        self.parts.append((name, data))

    def default_type(self, ext: str, ct: str) -> None:
        self.defaults.setdefault(ext, ct)


def write_xlsx(wb: Workbook, app_name: str = "ReflexGrid", macros: bool = True) -> bytes:
    """The workbook as a package: ``.xlsm`` (macro-enabled) when it has a VBA project and
    ``macros`` is true, else ``.xlsx`` (a VBA project is then left out)."""
    if not wb.sheets:
        raise ValueError("a workbook needs at least one sheet")
    out = Out()
    project = wb.vba if macros else None
    wb_code = wb.code_name
    if project is not None:
        from ..model.vba import reconcile

        project, wb_code, out.code_names = reconcile(project, wb.code_name, [s.code_name for s in wb.sheets])
    out.default_type("rels", "application/vnd.openxmlformats-package.relationships+xml")
    out.default_type("xml", "application/xml")
    visible = [i for i, s in enumerate(wb.sheets) if s.visibility == "visible"]
    active = wb.active_sheet if wb.active_sheet in visible else (visible[0] if visible else 0)
    wb_rels = Rels()
    sheets_xml = []
    for i, sheet in enumerate(wb.sheets):
        xml, rels = write_sheet(wb, i, i == active, out)
        out.part(f"xl/worksheets/sheet{i + 1}.xml", CT_WORKSHEET, xml.encode("utf-8"))
        if rels:
            out.part(f"xl/worksheets/_rels/sheet{i + 1}.xml.rels", None, rels.xml().encode("utf-8"))
        rid = wb_rels.add("worksheet", f"worksheets/sheet{i + 1}.xml")
        state = ""
        if sheet.visibility != "visible" and not (not visible and i == active):
            state = f' state="{sheet.visibility}"'
        sheets_xml.append(f'<sheet name="{esc_attr(sheet.name[:31])}" sheetId="{i + 1}"{state} r:id="{rid}"/>')

    names = []
    seen = set()
    for n in wb.names:
        key = (n.name.lower(), n.scope)
        if not n.name or key in seen:
            continue
        if n.scope is not None and n.scope >= len(wb.sheets):
            continue
        seen.add(key)
        x = f'<definedName name="{esc_attr(n.name)}"'
        if n.scope is not None:
            x += f' localSheetId="{n.scope}"'
        if n.hidden:
            x += ' hidden="1"'
        if n.comment:
            x += f' comment="{esc_attr(n.comment)}"'
        names.append(x + f">{esc(text_to_file(n.formula))}</definedName>")
    for i, sheet in enumerate(wb.sheets):
        q = quote_sheet(sheet.name)
        pa = sheet.print.print_area
        if pa is not None:
            names.append(f'<definedName name="_xlnm.Print_Area" localSheetId="{i}">{esc(q)}!{pa.abs_a1()}</definedName>')
        titles = []
        if sheet.print.title_cols is not None:
            c0, c1 = sheet.print.title_cols
            titles.append(f"{q}!{RangeRef.cols(min(c0, c1), max(c0, c1)).abs_a1()}")
        if sheet.print.title_rows is not None:
            r0, r1 = sheet.print.title_rows
            titles.append(f"{q}!{RangeRef.rows(min(r0, r1), max(r0, r1)).abs_a1()}")
        if titles:
            names.append(f'<definedName name="_xlnm.Print_Titles" localSheetId="{i}">{esc(",".join(titles))}</definedName>')
        af = sheet.autofilter
        if af is not None:
            names.append(f'<definedName name="_xlnm._FilterDatabase" localSheetId="{i}" hidden="1">{esc(q)}!{af.range.abs_a1()}</definedName>')

    w = [XML_DECL, f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">']
    wp = f' codeName="{esc_attr(wb_code)}"' if wb_code else ""
    w.append(f'<workbookPr{wp} date1904="1"/>' if wb.date_system is DateSystem.D1904 else f"<workbookPr{wp}/>")
    if wb.protected_structure:
        pw = f' workbookPassword="{esc_attr(wb.structure_password)}"' if wb.structure_password else ""
        w.append(f'<workbookProtection{pw} lockStructure="1"/>')
    w.append(f'<bookViews><workbookView activeTab="{active}"/></bookViews><sheets>{"".join(sheets_xml)}</sheets>')
    if names:
        w.append("<definedNames>" + "".join(names) + "</definedNames>")
    calc = ['<calcPr calcId="191029"']
    if wb.calc.mode == "manual":
        calc.append(' calcMode="manual"')
    elif wb.calc.mode == "autoNoTable":
        calc.append(' calcMode="autoNoTable"')
    if wb.calc.iterative:
        calc.append(' iterate="1"')
    if wb.calc.iterative or wb.calc.max_iterations != 100:
        calc.append(f' iterateCount="{wb.calc.max_iterations}"')
    if wb.calc.iterative or wb.calc.max_change != 0.001:
        calc.append(f' iterateDelta="{num(wb.calc.max_change)}"')
    if wb.calc.precision_as_displayed:
        calc.append(' fullPrecision="0"')
    if out.needs_calc:
        calc.append(' fullCalcOnLoad="1"')
    calc.append("/>")
    w.append("".join(calc))
    w.append("</workbook>")

    out.part("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", write_styles(wb, out.dxfs).encode("utf-8"))
    wb_rels.add("styles", "styles.xml")
    out.part("xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml", write_theme(wb.theme).encode("utf-8"))
    wb_rels.add("theme", "theme/theme1.xml")
    sst = [XML_DECL, f'<sst xmlns="{NS_MAIN}" count="{out.sst.count}" uniqueCount="{len(out.sst.list)}">']
    sst.extend(f"<si>{t_el(t)}</si>" for t in out.sst.list)
    sst.append("</sst>")
    out.part("xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", "".join(sst).encode("utf-8"))
    wb_rels.add("sharedStrings", "sharedStrings.xml")
    if out.dynamic_arrays:
        out.part("xl/metadata.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml", METADATA.encode("utf-8"))
        wb_rels.add("sheetMetadata", "metadata.xml")
    if project is not None:
        from .ovba import write_project

        out.part("xl/vbaProject.bin", CT_VBA, write_project(project))
        wb_rels.add(REL_VBA, "vbaProject.bin")
        if project is wb.vba and not project.modified and project.signatures:
            from .read import SIGNATURE_TYPES

            sig_rels = Rels()
            for ty, fname, data in project.signatures:
                out.part(f"xl/{fname}", SIGNATURE_TYPES[ty], data)
                sig_rels.add(ty, fname)
            out.part("xl/_rels/vbaProject.bin.rels", None, sig_rels.xml().encode("utf-8"))
    out.part("xl/workbook.xml", CT_MAIN_MACRO if project is not None else CT_MAIN, "".join(w).encode("utf-8"))
    out.part("xl/_rels/workbook.xml.rels", None, wb_rels.xml().encode("utf-8"))

    p = wb.props
    core = [
        XML_DECL,
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
    ]
    for tag, v in (("dc:title", p.title), ("dc:subject", p.subject), ("dc:creator", p.author), ("cp:keywords", p.keywords), ("dc:description", p.description)):
        if v:
            core.append(f"<{tag}>{esc(v)}</{tag}>")
    core.append("</cp:coreProperties>")
    out.part("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml", "".join(core).encode("utf-8"))
    app = [XML_DECL, '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">']
    app.append(f"<Application>{esc(app_name)}</Application>")
    if p.company:
        app.append(f"<Company>{esc(p.company)}</Company>")
    app.append("</Properties>")
    out.part("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml", "".join(app).encode("utf-8"))
    root = Rels()
    root.add("officeDocument", "xl/workbook.xml")
    root.add("core-properties", "docProps/core.xml")
    root.add("extended-properties", "docProps/app.xml")
    out.part("_rels/.rels", None, root.xml().encode("utf-8"))

    ct = [XML_DECL, '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">']
    for ext, t in sorted(out.defaults.items()):
        ct.append(f'<Default Extension="{esc_attr(ext)}" ContentType="{esc_attr(t)}"/>')
    for part, t in out.overrides:
        ct.append(f'<Override PartName="{esc_attr(part)}" ContentType="{esc_attr(t)}"/>')
    ct.append("</Types>")
    buf = io.BytesIO()
    rank = {"_rels/.rels": 0, "xl/workbook.xml": 2, "xl/_rels/workbook.xml.rels": 3}
    parts = sorted(out.parts, key=lambda kv: rank.get(kv[0], 1 if kv[0].startswith("docProps/") else 4))
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(ct))
        for name, data in parts:
            z.writestr(name, data)
    return buf.getvalue()
