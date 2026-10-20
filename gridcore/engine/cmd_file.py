"""File commands: new, open (XLSX/CSV bytes or path), save (XLSX/CSV), close, properties and
switching between open workbooks."""

from __future__ import annotations

import base64
import os
from typing import Any, Dict, Optional

from ..xlsx import XlsxError, read_csv, read_xlsx, sniff, write_csv, write_xlsx
from ..xlsx.csv_io import CsvTooLarge
from .display import cell_text
from .session import DocState, EngineError, Session, command, edit, p_int, p_str

MAX_FILE = 200 * 1024 * 1024


def open_bytes(s: Session, data: bytes, name: str, path: Optional[str] = None, delimiter: Optional[str] = None) -> DocState:
    if len(data) > MAX_FILE:
        raise EngineError("The file is too large to open.")
    kind = sniff(data)
    lower = name.lower()
    warnings = []
    # The content decides, not the name: a workbook named `.csv` still opens as a workbook.
    if kind == "xlsx":
        try:
            wb, warnings = read_xlsx(data)
        except XlsxError as e:
            raise EngineError(f"ReflexGrid can't open `{name}`: {e}")
    elif kind == "csv" or lower.endswith((".csv", ".tsv", ".txt")):
        if lower.endswith(".tsv") and delimiter is None:
            delimiter = "\t"
        wb = read_csv(data, delimiter)
        base = os.path.splitext(os.path.basename(name))[0] or "Sheet1"
        err = wb.check_sheet_name(base[:31])
        if err is None:
            wb.sheets[0].name = base[:31]
    elif kind == "xls":
        if _encrypted_package(data):
            raise EngineError(f"`{name}` needs a password to open, which ReflexGrid can't open yet. Remove the password in Excel and try again.")
        raise EngineError("Legacy .xls (Excel 97-2003) files aren't supported. Save the file as .xlsx and try again.")
    elif kind == "ods":
        raise EngineError("OpenDocument (.ods) files aren't supported yet. Save the file as .xlsx and try again.")
    else:
        raise EngineError(f"ReflexGrid can't open `{name}`: the file format isn't recognised.")
    d = DocState(wb, os.path.basename(name), path)
    d.warnings = list(warnings)
    s.add_document(d)
    return d


def _encrypted_package(data: bytes) -> bool:
    """A workbook with a password to open is a compound file holding the encrypted package (not an
    .xls file)."""
    from ..xlsx.cfb import read_cfb

    try:
        return read_cfb(data).stream("EncryptedPackage") is not None
    except Exception:  # a damaged compound file is reported as .xls, as before
        return False


@command("file.new", "New Workbook", menu=("File",), shortcut="Ctrl+N", params="{sheets?: 1}", undoable=False)
def new_workbook(s: Session, p: Dict[str, Any]):
    d = s.new_workbook(p_int(p, "sheets", 1))
    return {"name": d.name, "index": s.active}


@command("file.open", "Open…", menu=("File",), shortcut="Ctrl+O", params="{path?, data?: base64, name?, delimiter?}", undoable=False)
def open_file(s: Session, p: Dict[str, Any]):
    path = p_str(p, "path")
    if p.get("data") is not None:
        raw = p["data"]
        data = raw if isinstance(raw, (bytes, bytearray)) else base64.b64decode(str(raw))
        name = p_str(p, "name") or (os.path.basename(path) if path else "Book.xlsx")
    elif path:
        try:
            with open(path, "rb") as fh:
                data = fh.read(MAX_FILE + 1)
        except OSError as e:
            raise EngineError(f"ReflexGrid can't open `{path}`: {e.strerror}")
        name = os.path.basename(path)
    else:
        raise EngineError("file.open: missing `path` or `data`")
    d = open_bytes(s, bytes(data), name, path, p_str(p, "delimiter"))
    return {"name": d.name, "index": s.active, "warnings": d.warnings}


MIMES = {
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
    "csv": "text/csv",
    "tsv": "text/tab-separated-values",
}


def workbook_format(s: Session) -> str:
    """The file format that keeps everything: ``xlsm`` for a workbook with macros, else ``xlsx``."""
    return "xlsm" if s.doc().wb.vba is not None else "xlsx"


def encode(s: Session, fmt: str, sheet: Optional[int] = None, delimiter: str = ",") -> bytes:
    """The workbook as a file. ``xlsx`` leaves a VBA project out (a macro-free workbook);
    ``xlsm`` keeps it."""
    d = s.doc()
    if fmt == "auto":
        fmt = workbook_format(s)
    if fmt in ("xlsx", "xlsm"):
        return write_xlsx(d.wb, macros=fmt == "xlsm")
    if fmt in ("csv", "tsv"):
        si = d.wb.active_sheet if sheet is None else sheet
        sh = d.wb.sheet(si)
        try:
            return write_csv(sh, lambda rc: cell_text(d.wb, sh, rc), "\t" if fmt == "tsv" else delimiter)
        except CsvTooLarge as e:
            raise EngineError(f"ReflexGrid can't save this sheet as {fmt.upper()}: {e}.")
    raise EngineError(f"unknown format `{fmt}`")


@command("file.saveBytes", "Encode Workbook", params="{format?: auto|xlsx|xlsm|csv|tsv (auto: xlsm when the workbook has macros), sheet?} → {name, mime, data (base64)}", undoable=False, journal=False)
def save_bytes(s: Session, p: Dict[str, Any]):
    d = s.doc()
    fmt = p_str(p, "format", "auto")
    if fmt == "auto":
        fmt = workbook_format(s)
    if fmt not in MIMES:
        raise EngineError(f"unknown format `{fmt}`")
    data = encode(s, fmt, p_int(p, "sheet"))
    stem = os.path.splitext(d.name)[0] or "Book1"
    mime = MIMES[fmt]
    if fmt in ("xlsx", "xlsm"):
        d.saved = d.wb
        if not d.name.lower().endswith("." + fmt):
            d.name = stem + "." + fmt
    return {"name": f"{stem}.{fmt}", "mime": mime, "data": base64.b64encode(data).decode("ascii"), "size": len(data)}


@command("file.save", "Save", menu=("File",), shortcut="Ctrl+S", params="{path?}", undoable=False)
def save(s: Session, p: Dict[str, Any]):
    d = s.doc()
    path = p_str(p, "path") or d.path
    if not path:
        raise EngineError("file.save: no path (use file.saveAs or file.saveBytes)")
    return _save_to(s, path)


@command("file.saveAs", "Save As…", menu=("File",), shortcut="F12", params="{path}", undoable=False)
def save_as(s: Session, p: Dict[str, Any]):
    path = p_str(p, "path")
    if not path:
        raise EngineError("file.saveAs: missing `path`")
    return _save_to(s, path)


def _save_to(s: Session, path: str):
    d = s.doc()
    lower = path.lower()
    if not lower.endswith((".xlsx", ".xltx", ".xlsm", ".xltm", ".csv", ".tsv", ".txt")) and path != d.path:
        # An unknown extension would get an .xlsx under another name; a workbook opened from such a
        # file still saves back to it.
        raise EngineError(f"`{os.path.basename(path)}`: unknown file format. Save as .xlsx, .xlsm, .csv or .tsv.")
    fmt = "csv" if lower.endswith(".csv") else ("tsv" if lower.endswith((".tsv", ".txt")) else ("xlsm" if lower.endswith((".xlsm", ".xltm")) else "xlsx"))
    data = encode(s, fmt)
    tmp = path + ".tmp"
    try:
        with open(tmp, "wb") as fh:
            fh.write(data)
        os.replace(tmp, path)
    except OSError as e:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise EngineError(f"ReflexGrid can't save `{path}`: {e.strerror}")
    if fmt in ("xlsx", "xlsm"):
        d.saved = d.wb
        d.path = path
        d.name = os.path.basename(path)
    out: Dict[str, Any] = {"path": path, "size": len(data)}
    if fmt == "xlsx" and d.wb.vba is not None:
        out["warning"] = "The VBA project was left out: an .xlsx workbook can't hold macros. Save as .xlsm to keep them."
        s.messages.append(out["warning"])
    return out


@command("file.close", "Close", menu=("File",), shortcut="Ctrl+W", params="{index?}", undoable=False)
def close(s: Session, p: Dict[str, Any]):
    i = p_int(p, "index", s.active)
    if not 0 <= i < len(s.docs):
        raise EngineError("no such workbook")
    del s.docs[i]
    if s.clipboard is not None and not any(d.uid == s.clipboard.doc_uid for d in s.docs):
        s.clipboard = None
    if not s.docs:
        s.new_workbook()
    s.active = min(s.active, len(s.docs) - 1)
    return {"index": s.active}


@command("window.activate", "Switch Window", menu=("View", "Window"), params="{index}", undoable=False, journal=False)
def activate_window(s: Session, p: Dict[str, Any]):
    i = p_int(p, "index", 0)
    if not 0 <= i < len(s.docs):
        raise EngineError("no such workbook")
    s.active = i
    return {"name": s.docs[i].name}


@command("file.properties", "Properties", menu=("File", "Info"), params="{title?, subject?, author?, company?, keywords?, description?}")
def properties(s: Session, p: Dict[str, Any]):
    keys = ("title", "subject", "author", "company", "keywords", "description")

    def f(ctx):
        for k in keys:
            if k in p:
                setattr(ctx.wb.props, k, str(p[k] or ""))

    edit(s, f)
    d = s.doc()
    return {k: getattr(d.wb.props, k) for k in keys}


@command("file.exportHtml", "Save as Web Page", menu=("File", "Export"), params="{sheet?, range?, print?: bool} → {name, html}", undoable=False, journal=False)
def export_html(s: Session, p: Dict[str, Any]):
    from ..core.addr import RangeRef
    from .html_export import sheet_html
    from .session import param_sheet

    d = s.doc()
    si = param_sheet(s, p)
    rng = RangeRef.parse(str(p["range"])) if p.get("range") else None
    stem = os.path.splitext(d.name)[0] or "Book1"
    page = sheet_html(d.wb, si, rng, stem, for_print=bool(p.get("print", True)))
    return {"name": f"{stem}.html", "html": page}
