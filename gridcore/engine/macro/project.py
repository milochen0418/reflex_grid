"""A workbook's VBA project, compiled.

Compiling happens once per :class:`~gridcore.model.vba.VbaProject` object and the result (with
its module-level variables) is kept while the project is: editing the code makes a new project
and so a fresh state, as in Excel. The cache goes by object identity, so two workbooks with the
same code never share variables.
"""

from __future__ import annotations

import threading
import weakref
from typing import Any, Dict, List, Optional, Tuple

from ...model.vba import CLASS, DOCUMENT, FORM, STANDARD, WORKBOOK_BASE, VbaProject
from ...vba.compiler import ProjectRuntime, compile_project
from ...vba.errors import VbaSyntaxError
from .constants import CONSTANTS, ENUMS
from .host import GLOBAL_NAMES

# Excel's own classes: usable in `Dim x As Range`, never with New.
HOST_TYPES = {
    "range", "worksheet", "workbook", "application", "window", "name", "names", "listobject", "listcolumn", "listrow",
    "hyperlink", "comment", "font", "interior", "border", "borders", "chart", "chartobject", "shape", "validation",
    "worksheetfunction", "sheets", "worksheets", "areas", "pagesetup", "sort", "autofilter", "formatcondition",
}

_lock = threading.Lock()
_cache: Dict[int, Tuple["weakref.ref", ProjectRuntime]] = {}


def module_inputs(p: VbaProject) -> List[Tuple[str, str, str, Tuple[Tuple[str, str], ...], str]]:
    out = []
    for m in p.modules:
        doc_kind = ""
        if m.kind == DOCUMENT:
            doc_kind = "workbook" if any(WORKBOOK_BASE[2:10] in a for a in m.attributes) else "worksheet"
        out.append((m.name, m.kind, m.code, tuple(m.member_attributes), doc_kind))
    return out


def compile_vba(p: VbaProject) -> ProjectRuntime:
    """Compiles a project (no caching): its modules' compile errors are in ``.errors``."""
    rt = compile_project(module_inputs(p), GLOBAL_NAMES, CONSTANTS, p.name or "VBAProject", HOST_TYPES, ENUMS)
    return rt


def runtime_for(p: VbaProject) -> ProjectRuntime:
    """The compiled project, made once per project object."""
    key = id(p)
    with _lock:
        hit = _cache.get(key)
        if hit is not None and hit[0]() is p:
            return hit[1]
    rt = compile_vba(p)
    with _lock:
        hit = _cache.get(key)
        if hit is not None and hit[0]() is p:
            return hit[1]

        def gone(_ref, key=key):
            with _lock:
                cur = _cache.get(key)
                if cur is not None and cur[0]() is None:
                    del _cache[key]

        _cache[key] = (weakref.ref(p, gone), rt)
    return rt


def diagnostics(p: VbaProject) -> List[Dict[str, Any]]:
    """Compile errors, one per module that has one: ``{module, line, message}``."""
    rt = runtime_for(p)
    out = []
    for name, e in rt.errors.items():
        mod = p.module(name)
        out.append({"module": mod.name if mod is not None else name, "line": e.line, "message": e.msg})
    return out


def procedures(p: VbaProject) -> List[Dict[str, Any]]:
    """Every procedure: ``{module, name, kind, macro}`` (``macro``: runnable from the Macros
    dialog — a public Sub without parameters outside class modules)."""
    rt = runtime_for(p)
    out = []
    for m in rt.order:
        for name, d in m.procs.items():
            for kind, proc in d.items():
                out.append({
                    "module": m.name,
                    "name": proc.name,
                    "kind": kind,
                    "line": proc.line,
                    "macro": kind == "sub" and proc.scope != "private" and not proc.params and m.kind not in ("class", "form"),
                    "udf": kind == "function" and proc.scope != "private" and m.kind == "standard",
                    "event": m.kind == "document" and "_" in proc.name,
                })
    return out
