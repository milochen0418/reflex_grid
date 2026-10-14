"""Macro commands: the VBA project (modules, code, import/export), running macros and answering
their dialogs, event macros on or off, and the Immediate window.

A workbook's project is ``Workbook.vba`` (:mod:`gridcore.model.vba`); changing code is an
ordinary undoable edit, so everyone on a shared workbook gets it and undo takes it back. Runs
go through :mod:`.macro.runner`; a run that pauses for a dialog stays in ``session.macro_run``
until ``macro.answer`` (or ``macro.stop``) resumes it.
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import replace
from typing import Any, Dict, List, Optional

from ..model.vba import (
    CLASS,
    CLASS_BASE,
    DOCUMENT,
    FORM,
    STANDARD,
    VbaModule,
    VbaProject,
    default_attributes,
    new_project,
    normalize_code,
    reconcile,
    split_source,
    valid_module_name,
)
from .session import EngineError, Session, command, edit, p_bool, p_str

MAX_MODULE = 2 * 1024 * 1024


def _version(m: Optional[VbaModule]) -> str:
    return hashlib.sha1((m.code if m is not None else "").encode("utf-8")).hexdigest()[:16]


def _project(s: Session) -> VbaProject:
    p = s.doc().wb.vba
    if p is None:
        raise EngineError("This workbook has no macros.")
    if p.unreadable:
        raise EngineError(f"The workbook's VBA project can't be read ({p.unreadable}); it is kept as it is but can't be edited or run.")
    return p


def _editable(s: Session) -> VbaProject:
    p = _project(s)
    if p.locked:
        raise EngineError("The VBA project is locked (protected with a password in Excel); its code can't be viewed or changed.")
    return p


def _reconciled(wb) -> VbaProject:
    p, _, _ = reconcile(wb.vba, wb.code_name, [x.code_name for x in wb.sheets])
    return p  # type: ignore[return-value]


def _change_project(s: Session, f) -> VbaProject:
    """Applies ``f(project) -> project`` as one edit (creating the project when the workbook has
    none), keeps the sheets' code names in step and recalculates (worksheet functions may have
    changed)."""
    from ..xlsx.ovba import codec, pick_code_page

    out: List[VbaProject] = []

    def run(ctx):
        wb = ctx.wb
        p = wb.vba
        codes = [x.code_name for x in wb.sheets]
        if p is None:
            p = new_project([])
        p, wb_code, names = reconcile(p, wb.code_name, codes)
        p = f(p)
        # Every character must fit the project's code page (VBA stores code that way).
        texts = [m.code for m in p.modules] + [m.name for m in p.modules]
        try:
            for t in texts:
                t.encode(codec(p.code_page))
        except UnicodeEncodeError:
            cp = pick_code_page(texts)
            try:
                for t in texts:
                    t.encode(codec(cp))
            except UnicodeEncodeError:
                raise EngineError("Some characters can't be stored in VBA code (it uses one code page per project).")
            p = replace(p, code_page=cp, modified=True, signatures=())
        wb.vba = p
        if wb.code_name != wb_code:
            wb.code_name = wb_code
        for i, n in enumerate(names):
            if wb.sheets[i].code_name != n:
                ctx.sheet_mut(i).code_name = n
        ctx.structural = True
        out.append(p)

    edit(s, run)
    return out[0]


def _doc_sheet_names(wb) -> Dict[str, str]:
    """Document module name (lower case) -> the sheet it belongs to ("(workbook)" for ThisWorkbook)."""
    if wb.vba is None:
        return {}
    _, wb_code, names = reconcile(wb.vba, wb.code_name, [x.code_name for x in wb.sheets])
    out = {wb_code.lower(): "(workbook)"}
    for i, n in enumerate(names):
        out[n.lower()] = wb.sheets[i].name
    return out


# ---------------------------------------------------------------- the project


@command("macro.list", "Macros", menu=("View", "Macros"), shortcut="Alt+F8", params="{} → {modules, macros, functions, diagnostics, …}", undoable=False, journal=False)
def macro_list(s: Session, p: Dict[str, Any]):
    d = s.doc()
    wb = d.wb
    proj = wb.vba
    base = {"hasProject": proj is not None, "eventsEnabled": bool(d.macros_enabled), "running": s.macro_run is not None and s.macro_run.paused}
    if proj is None:
        return {**base, "modules": [], "macros": [], "functions": [], "diagnostics": [], "references": [], "locked": False, "unreadable": "", "codePage": 1252}
    if proj.unreadable:
        return {**base, "modules": [], "macros": [], "functions": [], "diagnostics": [], "references": [], "locked": False, "unreadable": proj.unreadable, "codePage": proj.code_page}
    from .macro.project import diagnostics, procedures

    shown = _reconciled(wb)
    sheets = _doc_sheet_names(wb)
    modules = [{"name": m.name, "kind": m.kind, "lines": len(m.code.splitlines()) if m.code else 0, "sheet": sheets.get(m.name.lower(), "")} for m in shown.modules]
    procs = procedures(proj)
    names: Dict[str, int] = {}
    for x in procs:
        if x["macro"]:
            names[x["name"].lower()] = names.get(x["name"].lower(), 0) + 1
    macros = []
    for x in procs:
        if not x["macro"]:
            continue
        mod = proj.module(x["module"])
        qualified = mod is not None and (mod.kind == DOCUMENT or names[x["name"].lower()] > 1)
        macros.append({"name": f"{x['module']}.{x['name']}" if qualified else x["name"], "module": x["module"], "proc": x["name"], "line": x["line"]})
    return {
        **base,
        "locked": proj.locked,
        "unreadable": "",
        "codePage": proj.code_page,
        "projectName": proj.name,
        "modules": modules,
        "macros": macros,
        "functions": [{"name": x["name"], "module": x["module"]} for x in procs if x["udf"]],
        "events": [{"name": x["name"], "module": x["module"]} for x in procs if x["event"]],
        "diagnostics": diagnostics(proj),
        "references": [{"name": r.name, "description": r.description} for r in proj.references],
    }


@command("macro.getModule", "Read VBA Module", params="{name} → {name, kind, code, version}", undoable=False, journal=False)
def macro_get_module(s: Session, p: Dict[str, Any]):
    proj = _editable(s)
    name = p_str(p, "name") or ""
    m = _reconciled(s.doc().wb).module(name)
    if m is None:
        raise EngineError(f"There is no module named {name}.")
    return {"name": m.name, "kind": m.kind, "code": m.code, "version": _version(m)}


@command("macro.setModule", "Edit VBA Code", menu=("Developer", "Code"), params="{name, code, version?: the version you edited (refused when someone changed it since)}")
def macro_set_module(s: Session, p: Dict[str, Any]):
    name = (p_str(p, "name") or "").strip()
    code = p.get("code")
    if not isinstance(code, str):
        raise EngineError("macro.setModule: missing `code`")
    if len(code) > MAX_MODULE:
        raise EngineError("That module is too large.")
    wb = s.doc().wb
    if wb.vba is not None:
        _editable(s)
    cur = _reconciled(wb).module(name) if wb.vba is not None else None
    if cur is not None and p.get("version") and p["version"] != _version(cur):
        raise EngineError(f"Someone else changed {cur.name} since you opened it. Reopen it to see their changes (yours are still in the editor).")
    text = normalize_code(code)

    def f(proj: VbaProject) -> VbaProject:
        m = proj.module(name)
        if m is None:
            err = valid_module_name(name)
            if err:
                raise EngineError(err)
            return proj.with_module(VbaModule(name, STANDARD, text, default_attributes(name, STANDARD)))
        if m.kind == FORM:
            raise EngineError("UserForms can't be edited in ReflexGrid.")
        return proj.with_module(replace(m, code=text))

    proj = _change_project(s, f)
    from .macro.project import diagnostics

    m = proj.module(name)
    return {"name": m.name if m else name, "version": _version(m), "diagnostics": diagnostics(proj)}


@command("macro.addModule", "Insert Module", menu=("Developer", "Code"), params="{kind?: standard|class, name?, code?}")
def macro_add_module(s: Session, p: Dict[str, Any]):
    kind = p_str(p, "kind", STANDARD) or STANDARD
    if kind not in (STANDARD, CLASS):
        raise EngineError("Only standard and class modules can be added.")
    if s.doc().wb.vba is not None:
        _editable(s)
    made: List[str] = []

    def f(proj: VbaProject) -> VbaProject:
        name = (p_str(p, "name") or "").strip() or proj.new_module_name(kind)
        err = valid_module_name(name)
        if err:
            raise EngineError(err)
        if proj.module(name) is not None:
            raise EngineError(f"A module named {name} already exists.")
        made.append(name)
        code = normalize_code(p_str(p, "code", "") or "")
        return proj.with_module(VbaModule(name, kind, code, default_attributes(name, kind)))

    _change_project(s, f)
    return {"name": made[0], "kind": kind}


@command("macro.deleteModule", "Remove Module", menu=("Developer", "Code"), params="{name}")
def macro_delete_module(s: Session, p: Dict[str, Any]):
    _editable(s)
    name = p_str(p, "name") or ""

    def f(proj: VbaProject) -> VbaProject:
        m = proj.module(name)
        if m is None:
            raise EngineError(f"There is no module named {name}.")
        if m.kind == DOCUMENT:
            raise EngineError(f"{m.name} belongs to a sheet or the workbook and can't be removed; clear its code instead.")
        return proj.without_module(name)

    _change_project(s, f)
    return {"removed": name}


@command("macro.renameModule", "Rename Module", menu=("Developer", "Code"), params="{name, newName}")
def macro_rename_module(s: Session, p: Dict[str, Any]):
    _editable(s)
    old = p_str(p, "name") or ""
    new = (p_str(p, "newName") or "").strip()
    err = valid_module_name(new)
    if err:
        raise EngineError(err)

    def f(proj: VbaProject) -> VbaProject:
        m = proj.module(old)
        if m is None:
            raise EngineError(f"There is no module named {old}.")
        if m.kind in (DOCUMENT, FORM):
            raise EngineError("Sheet, workbook and form modules can't be renamed here.")
        if new.lower() != old.lower() and proj.module(new) is not None:
            raise EngineError(f"A module named {new} already exists.")
        return proj.renamed_module(old, new)

    _change_project(s, f)
    return {"name": new}


def parse_module_file(file_name: str, text: str):
    """(name, kind, code, attributes, member attributes) of an exported module (.bas or .cls)."""
    ext = file_name.lower().rsplit(".", 1)[-1] if "." in file_name else ""
    if ext == "frm":
        raise EngineError("UserForms (.frm) can't be imported: ReflexGrid doesn't show forms.")
    body = text.replace("\r\n", "\n").replace("\r", "\n")
    if body.startswith("﻿"):
        body = body[1:]
    lines = body.split("\n")
    kind = STANDARD
    i = 0
    if lines and lines[0].strip().upper().startswith("VERSION") and "CLASS" in lines[0].upper():
        kind = CLASS
        i = 1
        if i < len(lines) and lines[i].strip().upper() == "BEGIN":
            while i < len(lines) and lines[i].strip().upper() != "END":
                i += 1
            i += 1
    elif ext == "cls":
        kind = CLASS
    code, attrs, members = split_source("\n".join(lines[i:]))
    name = ""
    for a in attrs:
        m = re.match(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"', a, re.I)
        if m:
            name = m.group(1)
    if not name:
        name = re.sub(r"[^A-Za-z0-9_]", "_", file_name.rsplit("/", 1)[-1].rsplit(".", 1)[0]) or "Module1"
        if not name[0].isalpha():
            name = "M" + name
    predeclared = any(re.match(r"^\s*Attribute\s+VB_PredeclaredId\s*=\s*True", a, re.I) for a in attrs)
    return name, kind, code, predeclared, members


@command("macro.importModule", "Import File", menu=("Developer", "Code"), params="{fileName: Module1.bas | Class1.cls, text}")
def macro_import_module(s: Session, p: Dict[str, Any]):
    if s.doc().wb.vba is not None:
        _editable(s)
    file_name = p_str(p, "fileName") or "Module1.bas"
    text = p.get("text")
    if not isinstance(text, str):
        raise EngineError("macro.importModule: missing `text`")
    if len(text) > MAX_MODULE:
        raise EngineError("That file is too large.")
    name, kind, code, predeclared, members = parse_module_file(file_name, text)
    err = valid_module_name(name)
    if err:
        raise EngineError(f"{name}: {err}")
    made: List[str] = []

    def f(proj: VbaProject) -> VbaProject:
        n = name
        k = 1
        while proj.module(n) is not None:
            n = f"{name}{k}"
            k += 1
        attrs = default_attributes(n, kind)
        if kind == CLASS and predeclared:
            attrs = tuple(a.replace("VB_PredeclaredId = False", "VB_PredeclaredId = True") for a in attrs)
        made.append(n)
        return proj.with_module(VbaModule(n, kind, code, attrs, members))

    _change_project(s, f)
    return {"name": made[0], "kind": kind}


@command("macro.exportModule", "Export File", menu=("Developer", "Code"), params="{name} → {fileName, text}", undoable=False, journal=False)
def macro_export_module(s: Session, p: Dict[str, Any]):
    _editable(s)
    name = p_str(p, "name") or ""
    m = _reconciled(s.doc().wb).module(name)
    if m is None:
        raise EngineError(f"There is no module named {name}.")
    if m.kind == STANDARD:
        head = [f'Attribute VB_Name = "{m.name}"']
        ext = "bas"
    else:
        head = ["VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "END", f'Attribute VB_Name = "{m.name}"']
        for a in m.attributes:
            if re.match(r"^\s*Attribute\s+VB_(GlobalNameSpace|Creatable|PredeclaredId|Exposed)\b", a, re.I):
                head.append(a)
        ext = "cls"
    body = replace(m, attributes=()).source().rstrip("\n")
    # Exported files have Windows line ends, as the VBA editor writes them.
    return {"fileName": f"{m.name}.{ext}", "text": "\r\n".join(head + ([body] if body else [])).replace("\r\n", "\n").replace("\n", "\r\n") + "\r\n"}


@command("macro.removeAll", "Remove Macros", menu=("Developer", "Code"), params="{} (the workbook becomes a macro-free .xlsx)")
def macro_remove_all(s: Session, p: Dict[str, Any]):
    _project(s)

    def f(ctx):
        ctx.wb.vba = None
        ctx.structural = True

    edit(s, f)
    s.doc().macros_enabled = False
    return {"removed": True}


@command("macro.check", "Compile VBAProject", menu=("Developer", "Code"), params="{} → {diagnostics}", undoable=False, journal=False)
def macro_check(s: Session, p: Dict[str, Any]):
    from .macro.project import diagnostics

    return {"diagnostics": diagnostics(_project(s))}


@command("macro.enableEvents", "Enable Macros", menu=("Developer", "Code"), params="{on?: bool} (event macros such as Worksheet_Change run when people edit)", undoable=False)
def macro_enable_events(s: Session, p: Dict[str, Any]):
    d = s.doc()
    on = p_bool(p, "on", True)
    if on:
        _project(s)
    d.macros_enabled = on
    return {"eventsEnabled": on}


# ---------------------------------------------------------------- running


def _start(s: Session, job, label: str, p: Dict[str, Any]):
    from .macro.runner import MacroRun

    cur = s.macro_run
    if cur is not None and cur.paused:
        raise EngineError(f"The macro {cur.label} is waiting for an answer; answer it or stop it first.")
    answers = p.get("answers")
    run = MacroRun(s, job, label, answers if isinstance(answers, list) else None, interactive=bool(s.interactive_macros), time_limit=float(p.get("timeLimit") or 60.0), user_name=s.prefs.get("userName", ""))
    s.undo_group = f"macro-{run.id}"
    r = run.step()
    s.macro_run = run if run.paused else None
    return {**r, "run": run.id}


@command("macro.run", "Run Macro", menu=("View", "Macros"), params="{name: Hello | Module1.Hello, args?: [..], answers?: [MsgBox/InputBox answers in order]} → {status: done|error|stopped|prompt|busy, output, messages, error, prompt}")
def macro_run(s: Session, p: Dict[str, Any]):
    from .macro.runner import call_macro_job

    proj = _project(s)
    name = (p_str(p, "name") or "").strip()
    if not name:
        raise EngineError("macro.run: which macro?")
    args = p.get("args") if isinstance(p.get("args"), list) else []
    from .macro.project import runtime_for

    rt = runtime_for(proj)
    if rt.find(name) is None:
        raise EngineError(f"Cannot run the macro '{name}'. The macro may not be available in this workbook.")
    return _start(s, call_macro_job(name, args), name, p)


@command("macro.open", "Workbook Open", params="{} (runs Workbook_Open and Auto_Open when the workbook's macros are enabled)")
def macro_open(s: Session, p: Dict[str, Any]):
    d = s.doc()
    proj = d.wb.vba
    if proj is None or proj.unreadable or not d.macros_enabled:
        return {"status": "skipped"}
    from .macro.project import runtime_for

    rt = runtime_for(proj)
    auto = rt.find("Auto_Open")

    def job(ex, host):
        host.fire(ex, None, "Workbook_Open", [])
        if auto is not None and auto.module.kind == "standard":
            rt.call(ex, auto, [])

    return _start(s, job, "Workbook_Open", p)


@command("macro.immediate", "Immediate Window", params="{code: statements, or ? expression} → {output, error}")
def macro_immediate(s: Session, p: Dict[str, Any]):
    from ..vba.compiler import compile_immediate
    from ..vba.errors import VbaSyntaxError
    from .macro.project import runtime_for

    proj = _project(s)
    code = p_str(p, "code") or ""
    rt = runtime_for(proj)
    try:
        proc = compile_immediate(rt, code)
    except VbaSyntaxError as e:
        return {"status": "error", "error": {"kind": "compile", "description": e.msg, "text": str(e), "line": e.line}, "output": [], "messages": []}

    def job(ex, host):
        from ..vba.runtime import call_procedure

        return call_procedure(ex, proc, [])

    return _start(s, job, "Immediate", p)


def _paused(s: Session):
    run = s.macro_run
    if run is None or not run.paused:
        raise EngineError("No macro is waiting.")
    return run


@command("macro.answer", "Answer Macro", params="{value: the button (vbYes = 6…) or the text typed; false cancels an InputBox}")
def macro_answer(s: Session, p: Dict[str, Any]):
    run = _paused(s)
    if run.state != "prompt":
        raise EngineError("The macro isn't asking anything.")
    value = p.get("value")
    if run.prompt and run.prompt.get("kind") == "msgbox":
        try:
            value = int(value)
        except (TypeError, ValueError):
            value = 1
    elif value is None:
        value = False if (run.prompt or {}).get("cancel_false") else ""
    s.undo_group = f"macro-{run.id}"
    r = run.step(value)
    s.macro_run = run if run.paused else None
    return {**r, "run": run.id}


@command("macro.continue", "Continue Macro", params="{} (after a time slice)")
def macro_continue(s: Session, p: Dict[str, Any]):
    run = _paused(s)
    if run.state != "busy":
        raise EngineError("The macro is waiting for an answer.")
    s.undo_group = f"macro-{run.id}"
    r = run.step()
    s.macro_run = run if run.paused else None
    return {**r, "run": run.id}


@command("macro.stop", "Stop Macro", shortcut="Ctrl+Break", params="{}")
def macro_stop(s: Session, p: Dict[str, Any]):
    run = s.macro_run
    if run is None:
        return {"status": "idle"}
    run.cancel()
    s.undo_group = f"macro-{run.id}"
    r = run.step(None)
    s.macro_run = None
    return {**r, "run": run.id}
