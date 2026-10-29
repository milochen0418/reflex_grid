"""Macro commands: the VBA project (modules, code, import/export), running macros and answering
their dialogs, UserForms (showing them, what the person does on them, their designs), event macros
on or off, and the Immediate window.

A workbook's project is ``Workbook.vba`` (:mod:`gridcore.model.vba`); changing code is an
ordinary undoable edit, so everyone on a shared workbook gets it and undo takes it back. Runs
go through :mod:`.macro.runner`; a run that pauses for a dialog stays in ``session.macro_run``
until ``macro.answer`` (or ``macro.stop``) resumes it.
"""

from __future__ import annotations

import base64
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
from .session import EngineError, Session, command, edit, p_bool, p_int, p_str

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
        return proj.with_module(replace(m, code=text))

    proj = _change_project(s, f)
    from .macro.project import diagnostics

    m = proj.module(name)
    return {"name": m.name if m else name, "version": _version(m), "diagnostics": diagnostics(proj)}


@command("macro.addModule", "Insert Module", menu=("Developer", "Code"), params="{kind?: standard|class|form, name?, code?}")
def macro_add_module(s: Session, p: Dict[str, Any]):
    kind = p_str(p, "kind", STANDARD) or STANDARD
    if kind not in (STANDARD, CLASS, FORM):
        raise EngineError("Only modules, class modules and UserForms can be added.")
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
        designer: tuple = ()
        if kind == FORM:
            # A new form, as the VBA editor makes one (its inside 228 x 150.75 points).
            from ..model.forms import FormDesign
            from ..xlsx.oforms import write_design
            from ..xlsx.ovba import codec

            designer = tuple(write_design(FormDesign(name, caption=name, width=228, height=150.75), codec(proj.code_page)))
        return proj.with_module(VbaModule(name, kind, code, default_attributes(name, kind), designer=designer))

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


def _renamed(proj: VbaProject, old: str, new: str) -> VbaProject:
    """``proj`` with the module ``old`` named ``new``: its code stream, its ``VB_Name`` and, for
    a UserForm, the name in its design's header (the storage follows when the file is written).
    Sheet and workbook modules keep their names. Code that names the module isn't rewritten, as
    in Excel."""
    err = valid_module_name(new)
    if err:
        raise EngineError(err)
    m = proj.module(old)
    if m is None:
        raise EngineError(f"There is no module named {old}.")
    if m.kind == DOCUMENT:
        raise EngineError("Sheet and workbook modules can't be renamed here.")
    if new.lower() != old.lower() and proj.module(new) is not None:
        raise EngineError(f"A module named {new} already exists.")
    proj = proj.renamed_module(old, new)
    if m.kind == FORM and m.designer:
        from ..xlsx.oforms_edit import renamed_designer
        from ..xlsx.ovba import codec

        proj = proj.with_module(replace(proj.module(new), designer=renamed_designer(m.designer, new, codec(proj.code_page))))
    return proj


@command("macro.renameModule", "Rename Module", menu=("Developer", "Code"), params="{name, newName}")
def macro_rename_module(s: Session, p: Dict[str, Any]):
    """Renames a module, a class module or a UserForm (as its (Name) in the VBA editor's
    Properties window does)."""
    _editable(s)
    old = p_str(p, "name") or ""
    new = (p_str(p, "newName") or "").strip()
    err = valid_module_name(new)
    if err:
        raise EngineError(err)
    _change_project(s, lambda proj: _renamed(proj, old, new))
    return {"name": new}


def parse_module_file(file_name: str, text: str):
    """(name, kind, code, attributes, member attributes) of an exported module (.bas or .cls)."""
    ext = file_name.lower().rsplit(".", 1)[-1] if "." in file_name else ""
    if ext == "frm":
        raise EngineError("A UserForm (.frm) comes with its .frx file; import both together.")
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


def parse_form_file(file_name: str, text: str):
    """(name, the form's header lines for its VBFrame, code, member attributes) of an exported
    UserForm (.frm): ``VERSION 5.00``, a ``Begin {…} Name`` … ``End`` block of the form's
    properties, its attributes and its code (the controls are in the .frx)."""
    body = text.replace("\r\n", "\n").replace("\r", "\n").lstrip("\ufeff")
    lines = body.split("\n")
    if len(lines) < 3 or not lines[0].strip().upper().startswith("VERSION") or not lines[1].strip().upper().startswith("BEGIN"):
        raise EngineError(f"{file_name} isn't a UserForm exported from the VBA editor.")
    frame = [lines[1]]
    depth = 0
    i = 2
    while i < len(lines):
        line = lines[i]
        word = line.strip().split(" ", 1)[0].upper() if line.strip() else ""
        i += 1
        if word in ("END", "ENDPROPERTY"):
            if depth == 0:
                frame.append(line)
                break
            depth -= 1
        elif word in ("BEGIN", "BEGINPROPERTY"):
            depth += 1
        elif word == "OLEOBJECTBLOB":
            continue
        frame.append(line)
    else:
        raise EngineError(f"{file_name}: the form's properties don't end.")
    code, attrs, members = split_source("\n".join(lines[i:]))
    name = ""
    for a in attrs:
        m = re.match(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"', a, re.I)
        if m:
            name = m.group(1)
    if not name:
        name = lines[1].split()[-1] if len(lines[1].split()) >= 3 else "UserForm1"
    return name, frame, code, members


def _import_form(s: Session, file_name: str, text: str, frx_b64: Any) -> Dict[str, Any]:
    from ..xlsx.oforms import VBFRAME, FormError, design_streams, frx_items, read_design
    from ..xlsx.ovba import codec

    if not isinstance(frx_b64, str) or not frx_b64:
        raise EngineError("A UserForm (.frm) comes with its .frx file; import both together.")
    try:
        frx = base64.b64decode(frx_b64)
    except (ValueError, TypeError):
        raise EngineError("The .frx file couldn't be read.")
    name, frame, code, members = parse_form_file(file_name, text)
    err = valid_module_name(name)
    if err:
        raise EngineError(f"{name}: {err}")
    try:
        items = [it for it in frx_items(frx) if it[0] != VBFRAME]
    except FormError as e:
        raise EngineError(f"The .frx file can't be read: {e}")
    made: List[str] = []

    def f(proj: VbaProject) -> VbaProject:
        n = name
        k = 1
        while proj.module(n) is not None:
            n = f"{name}{k}"
            k += 1
        head = list(frame)
        if n != name:
            head[0] = re.sub(r"\S+\s*$", n + " ", head[0])
        vbframe = "\r\n".join(["VERSION 5.00"] + head + [""]).encode(codec(proj.code_page), "replace")
        designer = tuple(sorted(items + [(VBFRAME, vbframe)]))
        try:
            read_design(n, design_streams(designer), codec(proj.code_page))
        except FormError as e:
            raise EngineError(f"The form in {file_name} can't be read: {e}")
        made.append(n)
        return proj.with_module(VbaModule(n, FORM, code, default_attributes(n, FORM), members, designer=designer))

    _change_project(s, f)
    return {"name": made[0], "kind": FORM}


@command("macro.importModule", "Import File", menu=("Developer", "Code"), params="{fileName: Module1.bas | Class1.cls | UserForm1.frm, text, frx?: base64 of the form's .frx}")
def macro_import_module(s: Session, p: Dict[str, Any]):
    if s.doc().wb.vba is not None:
        _editable(s)
    file_name = p_str(p, "fileName") or "Module1.bas"
    text = p.get("text")
    if not isinstance(text, str):
        raise EngineError("macro.importModule: missing `text`")
    if len(text) > MAX_MODULE:
        raise EngineError("That file is too large.")
    if file_name.lower().endswith(".frm"):
        return _import_form(s, file_name, text, p.get("frx"))
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


@command("macro.exportModule", "Export File", menu=("Developer", "Code"), params="{name} → {fileName, text, frxName?, frx?: base64 (a UserForm's .frx)}", undoable=False, journal=False)
def macro_export_module(s: Session, p: Dict[str, Any]):
    proj = _editable(s)
    name = p_str(p, "name") or ""
    m = _reconciled(s.doc().wb).module(name)
    if m is None:
        raise EngineError(f"There is no module named {name}.")
    if m.kind == FORM:
        return _export_form(proj, m)
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


def _export_form(proj: VbaProject, m: VbaModule) -> Dict[str, Any]:
    """A UserForm as the VBA editor exports it: the .frm (its properties, attributes and code)
    and the .frx (its controls)."""
    from ..xlsx.oforms import VBFRAME, write_frx
    from ..xlsx.ovba import codec
    from .macro.project import form_design

    if not m.designer:
        raise EngineError(f"{m.name}'s design is missing from the workbook; only its code can be copied.")
    frame = dict(m.designer).get(VBFRAME, b"").decode(codec(proj.code_page), "replace").replace("\r\n", "\n").split("\n")
    head = [ln for ln in frame if ln.strip() and not ln.strip().upper().startswith(("VERSION", "TYPEINFOVER"))]
    if not head or not head[0].strip().upper().startswith("BEGIN"):
        head = [f"Begin {{C62A69F0-16DC-11CE-9E98-00AA00574A4F}} {m.name} ", "End"]
    blob = f'   OleObjectBlob   =   "{m.name}.frx":0000'
    props = head[1:-1]
    at = next((i for i, ln in enumerate(props) if ln.strip().split(" ", 1)[0].lower() > "oleobjectblob"), len(props))
    head = [head[0]] + props[:at] + [blob] + props[at:] + [head[-1]]
    attrs = [f'Attribute VB_Name = "{m.name}"', "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = False",
             "Attribute VB_PredeclaredId = True", "Attribute VB_Exposed = False"]
    body = replace(m, attributes=()).source().rstrip("\n")
    text = "\r\n".join(["VERSION 5.00"] + head + attrs + ([body.replace("\n", "\r\n")] if body else [])) + "\r\n"
    frx = write_frx(m.designer, form_design(proj, m))
    return {"fileName": f"{m.name}.frm", "text": text, "frxName": f"{m.name}.frx", "frx": base64.b64encode(frx).decode("ascii")}


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
    return _stepped(s, run, run.step())


def _stepped(s: Session, run, r) -> Dict[str, Any]:
    """A run's step as the UI gets it: with the forms now on the person's screen."""
    s.macro_run = run if run.paused else None
    if not run.paused:
        _close_stranded_forms(s)
    return {**r, "run": run.id, "forms": _forms(s)}


def _registry(s: Session):
    """This person's forms, if they belong to the workbook's code as it is now."""
    reg = s.vba_forms
    if reg is None or reg.rt is None or not s.docs:
        return None
    proj = s.doc().wb.vba
    if proj is None or proj.unreadable:
        return None
    from .macro.project import runtime_for

    return reg if reg.rt is runtime_for(proj) else None


def _forms(s: Session) -> List[Dict[str, Any]]:
    reg = _registry(s)
    return [f.snapshot() for f in reg.shown()] if reg is not None else []


def _close_stranded_forms(s: Session) -> None:
    """Modal forms nobody waits for any more (the macro showing them stopped, or failed while
    they were up) go away, as ending the macro in Excel unloads them."""
    reg = _registry(s)
    if reg is None:
        return
    for f in reg.shown():
        if f.modal:
            f.loaded = f.shown = f.modal = False
            reg.forget(f)


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
    kind = (run.prompt or {}).get("kind")
    if kind == "msgbox":
        try:
            value = int(value)
        except (TypeError, ValueError):
            value = 1
    elif kind == "form":
        # What the person did on the form (see UserForm.dispatch); anything else closes it.
        value = value if isinstance(value, dict) else None
    elif value is None:
        value = False if (run.prompt or {}).get("cancel_false") else ""
    s.undo_group = f"macro-{run.id}"
    return _stepped(s, run, run.step(value))


@command("macro.continue", "Continue Macro", params="{} (after a time slice)")
def macro_continue(s: Session, p: Dict[str, Any]):
    run = _paused(s)
    if run.state != "busy":
        raise EngineError("The macro is waiting for an answer.")
    s.undo_group = f"macro-{run.id}"
    return _stepped(s, run, run.step())


@command("macro.stop", "Stop Macro", shortcut="Ctrl+Break", params="{}")
def macro_stop(s: Session, p: Dict[str, Any]):
    run = s.macro_run
    if run is None:
        return {"status": "idle"}
    run.cancel()
    s.undo_group = f"macro-{run.id}"
    r = _stepped(s, run, run.step(None))
    s.macro_run = None
    return r


# ---------------------------------------------------------------- UserForms


@command("macro.forms", "UserForms on Screen", params="{} → {forms: [the person's forms on screen], waiting: id of the modal form a macro waits on}", undoable=False, journal=False)
def macro_forms(s: Session, p: Dict[str, Any]):
    run = s.macro_run
    waiting = (run.prompt or {}).get("formId") if run is not None and run.state == "prompt" and (run.prompt or {}).get("kind") == "form" else None
    return {"forms": _forms(s), "waiting": waiting}


@command("macro.showForm", "Run UserForm", menu=("Developer", "Code"), params="{name: UserForm1, modal?: bool (default: the form's ShowModal)} → like macro.run")
def macro_show_form(s: Session, p: Dict[str, Any]):
    proj = _project(s)
    name = (p_str(p, "name") or "").strip()
    m = proj.module(name)
    if m is None or m.kind != FORM:
        raise EngineError(f"There is no UserForm named {name}.")
    modal = p.get("modal")
    from ..vba.values import MISSING

    def job(ex, host):
        mod = ex.rt.module(m.name)
        inst = ex.rt.form_default(ex, mod)
        inst.form.member(ex).g_show(ex, MISSING if modal is None else (1 if p_bool(p, "modal", True) else 0))

    return _start(s, job, m.name, p)


@command("macro.formEvent", "UserForm Event", params="{form: id, event: {event: click|change|dblclick|focus|keypress|keydown|spin|scroll|enter|escape|close, control?, value?, values?: {control: value}, focus?}} → like macro.run")
def macro_form_event(s: Session, p: Dict[str, Any]):
    """What a person did on one of their forms: answers the macro waiting on it (a modal form),
    or runs the form's event code (a modeless one) like a macro of its own."""
    _project(s)
    fid = p_int(p, "form", 0)
    ev = p.get("event") if isinstance(p.get("event"), dict) else {}
    run = s.macro_run
    if run is not None and run.paused:
        prompt = run.prompt or {}
        if run.state == "prompt" and prompt.get("kind") == "form" and prompt.get("formId") == fid:
            s.undo_group = f"macro-{run.id}"
            return _stepped(s, run, run.step(ev))
        raise EngineError(f"The macro {run.label} is waiting for an answer; answer it or stop it first.")
    reg = _registry(s)
    form = reg.get(fid) if reg is not None else None
    if form is None or not form.shown:
        return {"status": "done", "output": [], "messages": [], "error": None, "forms": _forms(s)}

    def job(ex, host):
        form.dispatch(ex, ev)

    return _start(s, job, form.name, p)


def _form_view(proj: VbaProject, m: VbaModule, control: str = "") -> Dict[str, Any]:
    from ..model.forms import parse_ole_color, property_rows
    from ..vba.msforms import css_color, design_snapshot
    from .macro.project import form_design

    d = form_design(proj, m)
    if control and d.find(control) is None:
        control = ""
    rows = property_rows(d, control)
    for row in rows:
        row["css"] = css_color(parse_ole_color(row["value"])) if row["type"] == "color" else ""
        row["value"] = "" if row["value"] is None else str(row["value"]) if row["type"] != "bool" else row["value"]
    return {"name": m.name, "form": design_snapshot(d), "warnings": list(d.warnings), "control": control,
            "props": rows, "editable": bool(m.designer) and not any("can't be read" in w for w in d.warnings)}


@command("macro.formDesign", "UserForm Design", params="{name, control?} → {name, form: the form as designed, warnings, props: the Properties window of control (or of the form), editable}", undoable=False, journal=False)
def macro_form_design(s: Session, p: Dict[str, Any]):
    proj = _editable(s)
    name = p_str(p, "name") or ""
    m = proj.module(name)
    if m is None or m.kind != FORM:
        raise EngineError(f"There is no UserForm named {name}.")
    return _form_view(proj, m, p_str(p, "control") or "")


@command("macro.editForm", "Edit UserForm", menu=("Developer", "Code"), params="{name, ops: [{op: set, control?: '' for the form, key, value} | {op: add, kind, parent?, left, top, width?, height?} | {op: delete, control}], control?} → like macro.formDesign, and added")
def macro_edit_form(s: Session, p: Dict[str, Any]):
    """The form designer's changes: properties (as the Properties window gives them), controls
    added from the toolbox, controls removed. The form's design changes in place: what the
    changes don't touch stays as it was, so a form made in Excel keeps everything."""
    from ..model.forms import property_change
    from ..xlsx.oforms import FormError
    from ..xlsx.oforms_edit import FormEditor
    from ..xlsx.ovba import codec

    _editable(s)
    name = p_str(p, "name") or ""
    ops = p.get("ops")
    if not isinstance(ops, list) or not ops:
        raise EngineError("macro.editForm: which changes (`ops`)?")
    added: List[str] = []
    focus = [p_str(p, "control") or ""]
    proj = _editable(s)
    m = proj.module(name)
    if m is None or m.kind != FORM:
        raise EngineError(f"There is no UserForm named {name}.")
    if not m.designer:
        raise EngineError(f"{m.name}'s design is missing from the workbook.")
    # The form's own (Name) is the module's name: a change to it renames the module (its
    # design's header follows), it isn't a change of the design.
    renamed_to = ""
    for op in ops:
        if isinstance(op, dict) and op.get("op") == "set" and not op.get("control") and str(op.get("key") or "") == "name":
            renamed_to = str(op.get("value") or "").strip()
    if renamed_to == name:
        renamed_to = ""
    if renamed_to:
        err = valid_module_name(renamed_to)
        if err:
            raise EngineError(err)

    def edited() -> tuple:
        try:
            ed = FormEditor(m.name, m.designer, codec(proj.code_page))
            for op in ops[:200]:
                kind = op.get("op") if isinstance(op, dict) else None
                if kind == "set":
                    control = str(op.get("control") or "")
                    key = str(op.get("key") or "")
                    if not control and key == "name":
                        continue
                    if control:
                        c = ed.design.find(control)
                        if c is None:
                            raise EngineError(f"There is no control named {control}.")
                        changes = property_change(c, key, op.get("value"))
                        ed.set(control, changes)
                        if "name" in changes:
                            focus[0] = changes["name"]
                    else:
                        ed.set_form(property_change(ed.design, key, op.get("value"), form=True))
                elif kind == "geometry":
                    control = str(op.get("control") or "")
                    changes = {k: round(float(op[k]), 2) for k in ("left", "top", "width", "height") if k in op}
                    if any(changes.get(k, 1) <= 0 for k in ("width", "height")):
                        raise EngineError("A control's width and height must be more than 0.")
                    if control:
                        ed.set(control, changes)
                    else:
                        ed.set_form({k: v for k, v in changes.items() if k in ("width", "height")})
                elif kind == "add":
                    new = ed.add(str(op.get("kind") or ""), str(op.get("parent") or ""), float(op.get("left") or 6), float(op.get("top") or 6),
                                 float(op["width"]) if op.get("width") else None, float(op["height"]) if op.get("height") else None)
                    added.append(new)
                    focus[0] = new
                elif kind == "delete":
                    ed.delete(str(op.get("control") or ""))
                    focus[0] = ""
                else:
                    raise EngineError(f"macro.editForm: unknown change {kind!r}.")
        except (FormError, ValueError, TypeError, KeyError) as e:
            raise EngineError(str(e).strip("'\""))
        return ed.items()

    designer = edited()
    changed = designer != tuple(sorted(m.designer))
    if changed or renamed_to:
        # One edit of the project (undo takes it back), only when the form did change.
        def f(pr: VbaProject) -> VbaProject:
            if changed:
                pr = pr.with_module(replace(pr.module(name), designer=designer))
            return _renamed(pr, name, renamed_to) if renamed_to else pr

        proj = _change_project(s, f)
    shown = renamed_to or name
    return {**_form_view(proj, proj.module(shown), focus[0]), "added": added}
