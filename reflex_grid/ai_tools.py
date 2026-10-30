"""What the AI Hub's assistant may do in ReflexGrid: tools (``reflex_ddns_auth.tools``) that read
a workbook, and tools that change it: its cells through the engine's commands, its VBA code and
its UserForms.

``room`` is the workbook's id (``/book/<room>``) and the tab that opened the assistant
(``<book>~<peer>``, ``GridState.ai_room``); the assistant passes the one it was opened from.

A tool reads the workbook as it is at that moment (an immutable snapshot taken under the book's
lock), whoever has it open. A tool that changes it is a ``danger`` tool: the AI Hub runs it only
after the person's Accept. It runs as that tab's person, inside ``book.act``, one call being one
act: the change is theirs, and their Ctrl+Z takes the whole call back, while everyone sees it as
the AI's (``ai.AiChange``)."""

from __future__ import annotations

import difflib
import json
from contextlib import contextmanager
from typing import Any, Dict, Iterator, List, Literal, Optional, Tuple

from gridcore.core.addr import RangeRef, a1, col_to_letters
from gridcore.core.value import CellError
from gridcore.engine import COMMANDS, CollabSession, EngineError, View
from gridcore.engine.cmd_edit import resolve_reference
from gridcore.engine.display import value_text
from gridcore.engine.macro.project import form_design as read_form

from . import ai, sessions
from .ai import short
from .books import STORE, valid_id
from .state import push

# What one read may return: enough for a table of a few thousand cells, not a whole sheet.
READ_MAX_CELLS = 3000
READ_MAX_CHARS = 60_000
READ_CHARTS = 5
READ_CELL_CHARS = 200
FIND_MAX = 50
FIND_MAX_CELLS = 400_000
# One grid.edit, and what it reports back.
EDIT_MAX_STEPS = 500
EDIT_MAX_RESULTS = 30
EDIT_MAX_ERRORS = 20
ERROR_SCAN_CELLS = 50_000
# Commands that aren't the AI's: files and the macro machinery (grid.vba_set and
# grid.form_design change code and forms), the person's clipboard, undo, and what only a person
# on screen has a use for.
NOT_FOR_AI_PREFIXES = ("file.", "macro.", "ui.", "window.")
NOT_FOR_AI = frozenset(
    {
        "edit.undo", "edit.redo", "edit.copy", "edit.cut", "edit.paste", "edit.pasteSpecial", "edit.clearClipboard",
        "edit.formatPainter", "edit.applyFormatPainter", "edit.cancelFormatPainter", "view.zoom", "view.scroll",
    }
)

GUIDE = """\
ReflexGrid is an Excel-compatible spreadsheet; each workbook is a room (/book/<room>).
Reading:
- grid_workbook: its sheets with their used ranges, tables and charts, its defined names and its VBA modules. Start here.
- grid_read: cells in an A1 reference with the sheet ("Sales!A1:D20"; also B:B, a table or a defined name). Each cell shows its value as displayed and, for a formula, the formula in brackets.
- grid_find: cells whose text or formula contains some words.
- grid_vba_get: a VBA module's code; for a UserForm, its design too.
- A column named in a question ("the Sales column") is usually a header in row 1, not a sheet: read the first rows (e.g. Sheet1!A1:Z3) to find it.
- Never guess what a cell holds: read it before you compute or explain anything.
Changing: do the change yourself with these tools rather than telling the person what to type. Each call waits for the person's Accept in this chat, and they can take an accepted call back in one click: put all the steps of one request in one call.
- grid_edit: a list of steps, each one of the commands the ribbon runs: {"command": "cell.set", "params": {"cell": "Sales!B7", "input": "=SUM(B2:B6)"}}, {"command": "range.setValues", "params": {"range": "Sales!A1", "values": [["Item", "Qty"], ["Pen", 3]]}}, {"command": "home.bold", "params": {"range": "A1:B1"}}. grid_commands lists the commands and their params: sorting, tables, charts, conditional formats, validation, inserting rows and sheets, and so on. Use only names grid_commands shows (clearing cells is edit.clearContents): a wrong name costs the person an Accept for nothing. A `range` is an A1 reference, not a table's name. All the steps of a call happen or none: when one fails, nothing changed.
- A chart floats over the cells: what is under it can't be seen. grid_read and grid_workbook say which cells each chart covers: put new cells elsewhere. A new chart lands two columns right of its range, about 8 columns wide and 15 rows high; object.move places it ({"id": 2, "row": 0, "col": 12}, counted from 0). When the answer of grid_edit has `hidden`, fix that before you finish.
- grid_vba_set: writes the code of a module, a class module or a UserForm, making it when it is new; its answer has the compile errors, which you fix with another call.
- grid_form_design: makes or changes a UserForm: its caption, size and controls.
- Every change has a summary: one sentence, in the person's language, saying what it does. It is what they read before they accept.
- Put what belongs together in one call. When the person rejects a change, nothing changed: ask what they want instead, don't try the same again.
- ReflexGrid runs Excel VBA itself, in the browser (no desktop Excel needed): modules, class modules and UserForms (also Controls.Add with WithEvents handlers), MsgBox and InputBox, and event macros such as Worksheet_Change once Developer › Macros On is set. A workbook with macros is kept as .xlsm. The person runs a macro with Developer › Macros (Alt+F8); you can't run one. Not available to macros: files, the shell, the network, COM, and the Application object's events.
- A message may begin with what the person's page shows (the sheet, the selection, the active cell and its formula): "this cell", "this formula" or "the selection" mean those."""


def _book(room: str):
    book_id = (room or "").strip().partition("~")[0]
    if not valid_id(book_id) or not STORE.exists(book_id):
        raise ValueError(f"there is no workbook {room!r}: pass the room the assistant was opened from")
    book = STORE.get(book_id) or STORE.open(book_id)
    with book.lock:
        return book, book.doc.wb


def _person(book: Any, room: str) -> View:
    """Who a change is made as: the tab that opened the assistant, so that it is that person's
    change and their undo takes it back. When that tab is gone (it reloaded while the assistant
    stayed open), a participant called AI."""
    peer = (room or "").strip().partition("~")[2]
    view = book.views.get(peer) if peer else None
    if view is None and peer:
        for token in sessions.tokens_on(book.key):
            entry = sessions.peek(token)
            if entry is not None and entry.view is not None and entry.view.key == peer:
                return entry.view
    if view is not None:
        return view
    if book.ai_view is None:
        book.ai_view = book.join(ai.AI_USER, CollabSession(), ai.AI_NAME)
        book.ai_view.info.update(user=ai.AI_USER, color=ai.AI_COLOR)
    return book.ai_view


@contextmanager
def _changing(book: Any, room: str, summary: str) -> Iterator[Tuple[Any, Dict[str, Any]]]:
    """An engine act that is one change of the AI's: made as the tab's person (one undo step),
    remembered as the AI's (``ai.note_change``), and what it changed reported in the dict."""
    done: Dict[str, Any] = {}
    with book.lock:
        view = _person(book, room)
        rev, before = book.rev, book.doc.wb
        try:
            with book.act(view) as s:
                yield s, done
        finally:
            change = ai.note_change(book, view, rev, before, summary)
            if view is book.ai_view:
                book.leave(ai.AI_USER)  # it is on the book only while it acts
            if change is not None:
                done["changed"] = [f"{_quote(name)}!{RangeRef(*box).a1()}" for name, box in change.areas]
                done["hidden"] = _hidden(book.doc.wb, before, change.areas)


def _sheet_index(wb, name: str) -> int:
    i = wb.sheet_index(name)
    if i is None:
        names = ", ".join(repr(s.name) for s in wb.sheets)
        raise ValueError(f"no sheet named {name!r} (the sheets: {names}); if {name!r} is a column's header, read row 1 of a sheet to find its column")
    return i


def _first_visible(wb) -> int:
    return next((i for i, s in enumerate(wb.sheets) if s.visibility == "visible"), 0)


async def guide(ctx: Any) -> str:
    """How to work with ReflexGrid's tools: read this first."""
    return GUIDE


def _chart_box(sh: Any, chart: Any) -> RangeRef:
    """The cells a chart floats over."""
    a = chart.anchor
    bottom, right = sh.row_top(a.row) + a.dy + a.height, sh.col_left(a.col) + a.dx + a.width
    return RangeRef(a.row, a.col, sh.row_index().at(bottom - 1), sh.col_index().at(right - 1))


def _hidden(wb: Any, before: Any, areas: List[Tuple[str, List[int]]]) -> List[str]:
    """What a chart hides after a change: cells it wrote under a chart, and what a new chart
    landed on."""
    known = {ch.id for sh in before.sheets for ch in sh.charts}
    out: List[str] = []
    for sh in wb.sheets:
        written = [RangeRef(*box) for name, box in areas if name == sh.name]
        for chart in sh.charts:
            over = _chart_box(sh, chart)
            look = [box.intersection(over) for box in written if box.intersects(over)] if chart.id in known else [over]
            if any(sh.value((r, c)) is not None for box in look for r in range(box.r0, box.r1 + 1) for c in range(box.c0, box.c1 + 1)):
                out.append(f"chart {chart.id} ({short(chart.title or chart.kind, 60)}) covers {_quote(sh.name)}!{over.a1()}, over cells that hold something")
    return out


async def workbook(ctx: Any, room: str) -> Dict[str, Any]:
    """A workbook at a glance: its title, sheets (used range, tables with their columns, charts
    with the cells they cover),
    defined names and VBA modules (each with its kind: standard, class, form or a sheet's)."""
    book, wb = _book(room)
    sheets = []
    for sh in wb.sheets:
        used = sh.used_range()
        sheets.append(
            {
                "name": sh.name,
                "used": used.a1() if used is not None else "",
                **({"hidden": True} if sh.visibility != "visible" else {}),
                "tables": [{"name": t.name, "range": t.range.a1(), "columns": [c.name for c in t.columns]} for t in sh.tables],
                "charts": [{"id": ch.id, "title": short(ch.title or ch.kind, 60), "covers": _chart_box(sh, ch).a1()} for ch in sh.charts],
            }
        )
    names = [{"name": n.name, "refers_to": "=" + n.formula.lstrip("="), **({"sheet": wb.sheets[n.scope].name} if isinstance(n.scope, int) and 0 <= n.scope < len(wb.sheets) else {})} for n in wb.names if not n.hidden][:60]
    vba = [{"name": m.name, "kind": m.kind} for m in wb.vba.modules] if wb.vba is not None else []
    return {"room": book.key, "title": book.meta.title, "sheets": sheets, "names": names, "vba": vba}


def _cell_text(wb, sh, rc: Tuple[int, int], formulas: bool) -> str:
    cell = sh.cell(rc)
    if cell is None and rc not in sh.spill:
        return ""
    shown = short(value_text(wb, sh, rc), READ_CELL_CHARS)
    if formulas and cell is not None and cell.formula is not None:
        return f"{shown} [={short(cell.formula.text, READ_CELL_CHARS)}]"
    return shown


async def read(ctx: Any, room: str, ref: str, formulas: bool = True) -> str:
    """The cells of an A1 reference, as a tab-separated table with row numbers and column letters;
    formulas in brackets after their values (``formulas``). A big range comes back in part, with
    the reference to read next."""
    book, wb = _book(room)
    text = (ref or "").strip()
    default = _first_visible(wb)
    if "!" in text:
        sheet_part, _, rest = text.rpartition("!")
        si = _sheet_index(wb, sheet_part.strip("'").replace("''", "'"))
        text = rest
    else:
        si = default
    try:
        si, rng = resolve_reference(wb, si, text)
    except EngineError:
        raise ValueError(f"{ref!r} isn't a reference: use A1 notation such as Sales!A1:D20, B:B, a table or a defined name")
    sh = wb.sheet(si)
    used = sh.used_range()
    if used is None:
        return f"{_quote(sh.name)}!{rng.a1()}: the sheet is empty."
    inter = rng.intersection(RangeRef(0, 0, used.r1, used.c1))
    if inter is None:
        return f"{_quote(sh.name)}!{rng.a1()}: empty (the sheet's cells end at {used.a1()})."
    cols = inter.c1 - inter.c0 + 1
    lines = ["\t" + "\t".join(col_to_letters(c) for c in range(inter.c0, inter.c1 + 1))]
    size = 0
    last = inter.r1
    for r in range(inter.r0, inter.r1 + 1):
        row = f"{r + 1}\t" + "\t".join(_cell_text(wb, sh, (r, c), formulas) for c in range(inter.c0, inter.c1 + 1))
        size += len(row)
        if len(lines) > 1 and ((r - inter.r0 + 1) * cols > READ_MAX_CELLS or size > READ_MAX_CHARS):
            last = r - 1
            break
        lines.append(row)
    shown = RangeRef(inter.r0, inter.c0, last, inter.c1)
    head = f"{_quote(sh.name)}!{shown.a1()} ({last - inter.r0 + 1} rows × {cols} columns{', formulas in brackets' if formulas else ''})"
    if last < inter.r1:
        rest = RangeRef(last + 1, inter.c0, inter.r1, inter.c1)
        head += f"; {inter.r1 - last} more rows: read {_quote(sh.name)}!{rest.a1()} next"
    for chart in sh.charts[:READ_CHARTS]:
        head += f"; chart {chart.id} covers {_chart_box(sh, chart).a1()}"
    return head + "\n" + "\n".join(lines)


async def find(ctx: Any, room: str, text: str, sheet: str = "") -> str:
    """Cells whose shown text or formula contains ``text`` (ignoring case), on one sheet or all,
    with their addresses."""
    book, wb = _book(room)
    needle = (text or "").strip().lower()
    if not needle:
        raise ValueError("say what to look for")
    sheets = [_sheet_index(wb, sheet)] if sheet else list(range(len(wb.sheets)))
    found: List[str] = []
    seen = 0
    for si in sheets:
        sh = wb.sheet(si)
        used = sh.used_range()
        if used is None:
            continue
        for r in range(used.r0, used.r1 + 1):
            for c in range(used.c0, used.c1 + 1):
                seen += 1
                if seen > FIND_MAX_CELLS:
                    break
                cell = sh.cell((r, c))
                if cell is None:
                    continue
                shown = value_text(wb, sh, (r, c))
                formula = ("=" + cell.formula.text) if cell.formula is not None else ""
                if needle in shown.lower() or needle in formula.lower():
                    found.append(f"{_quote(sh.name)}!{a1(r, c)}\t{short(shown, 120)}" + (f" [{short(formula, 120)}]" if formula else ""))
                    if len(found) >= FIND_MAX:
                        return "\n".join(found) + f"\n(the first {FIND_MAX}; look on one sheet or for more words to narrow it)"
    if not found:
        return f"No cell contains {text!r}" + (f" on {sheet!r}" if sheet else "") + "."
    return "\n".join(found)


def _for_ai(cid: str) -> bool:
    return not cid.startswith(NOT_FOR_AI_PREFIXES) and cid not in NOT_FOR_AI


async def commands(ctx: Any, search: str = "") -> str:
    """The commands grid_edit can run, one per line: id, what it is, its params. ``search`` keeps
    those with these words in them."""
    words = (search or "").lower().split()
    lines = []
    for cid, spec in sorted(COMMANDS.items()):
        line = f"{cid} — {spec.label}" + (f" — {spec.params}" if spec.params else "")
        if _for_ai(cid) and all(w in line.lower() for w in words):
            lines.append(line)
    return "\n".join(lines) or f"No command matches {search!r}."


def _step_problem(i: int, step: Any) -> str:
    cid = str(step.get("command") or "") if isinstance(step, dict) else ""
    if not cid:
        return f'step {i}: a step is {{"command": "...", "params": {{...}}}}'
    if cid.startswith("macro."):
        return f"step {i}: {cid} isn't a step: grid_vba_set and grid_form_design change code and forms"
    if not _for_ai(cid):
        return f"step {i}: {cid} isn't open to the AI"
    if cid not in COMMANDS:
        near = difflib.get_close_matches(cid, [c for c in COMMANDS if _for_ai(c)], n=3, cutoff=0.5)
        return f"step {i}: there is no command {cid}" + (f" (did you mean {', '.join(near)}?)" if near else "") + "; grid_commands lists them"
    if not isinstance(step.get("params", {}), dict):
        return f"step {i}: params must be an object"
    return ""


def _error_values(book: Any) -> List[str]:
    """The error values in what the latest change of the AI's touched."""
    wb = book.doc.wb
    errors: List[str] = []
    scanned = 0
    for name, box in book.ai_changes[-1].areas if book.ai_changes else ():
        i = wb.sheet_index(name)
        if i is None:
            continue
        sh = wb.sheet(i)
        rng = RangeRef(*box)
        for r in range(rng.r0, rng.r1 + 1):
            for c in range(rng.c0, rng.c1 + 1):
                scanned += 1
                if scanned > ERROR_SCAN_CELLS or len(errors) >= EDIT_MAX_ERRORS:
                    return errors
                v = sh.value((r, c))
                if isinstance(v, CellError):
                    errors.append(f"{_quote(name)}!{a1(r, c)}: {v.value}")
    return errors


async def edit(ctx: Any, room: str, summary: str, steps: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Changes the workbook with the ribbon's commands, in order, as one step of the person's
    undo. Runs once the person has accepted it. Answers with the areas changed and any error
    values in them. All or nothing: at a step that fails nothing stays changed, and the answer
    says why."""
    book, _ = _book(room)
    if not steps:
        raise ValueError("no steps")
    if len(steps) > EDIT_MAX_STEPS:
        raise ValueError(f"at most {EDIT_MAX_STEPS} steps in one call")
    problems = [p for p in (_step_problem(i, step) for i, step in enumerate(steps, 1)) if p]
    if problems:
        raise ValueError("nothing was done: " + "; ".join(problems))
    rev = book.rev
    done = 0
    failed = ""
    results: List[str] = []
    with _changing(book, room, summary) as (s, report):
        for i, step in enumerate(steps, 1):
            cid = step["command"]
            try:
                r = s.execute(cid, dict(step.get("params") or {}))
            except EngineError as ex:
                failed = f"step {i} ({cid}): {str(ex).splitlines()[0]}"
                book.discard()
                break
            done += 1
            if r not in (None, {}) and len(results) < EDIT_MAX_RESULTS:
                results.append(f"{cid}: " + short(json.dumps(r, ensure_ascii=False, default=str), 300))
    await push(book.key)
    if failed:
        return {"done": f"0 of {len(steps)} steps", "failed": f"{failed}; nothing was changed: fix it and send all the steps again"}
    errors = _error_values(book) if book.rev != rev else []
    out: Dict[str, Any] = {"done": f"{done} of {len(steps)} steps", "changed": report.get("changed", [])}
    if errors:
        out["error_values"] = errors
    if report.get("hidden"):
        out["hidden"] = report["hidden"]
    if results:
        out["results"] = results
    return out


def _vba(wb: Any) -> Any:
    project = wb.vba
    if project is None:
        raise ValueError("the workbook has no VBA yet: grid_vba_set adds it")
    if project.locked:
        raise ValueError("the workbook's VBA project is locked with a password")
    return project


def _control_json(c: Any) -> Dict[str, Any]:
    out: Dict[str, Any] = {"name": c.name, "kind": c.kind, "left": c.left, "top": c.top, "width": c.width, "height": c.height}
    if c.caption:
        out["caption"] = c.caption
    if c.children:
        out["children"] = [_control_json(x) for x in c.children]
    return out


async def vba_get(ctx: Any, room: str, module: str) -> Dict[str, Any]:
    """A VBA module's code, by its name (grid_workbook lists them); for a UserForm, its design
    too: caption, size (points) and controls."""
    _, wb = _book(room)
    project = _vba(wb)
    m = project.module(module)
    if m is None:
        raise ValueError(f"no module named {module!r} (the modules: {', '.join(x.name for x in project.modules)})")
    out: Dict[str, Any] = {"name": m.name, "kind": m.kind, "code": m.code}
    if m.kind == "form":
        d = read_form(project, m)
        out["form"] = {"caption": d.caption, "width": d.width, "height": d.height, "controls": [_control_json(c) for c in d.controls]}
    return out


def _problems(s: Any) -> List[str]:
    return [f"{d['module']} line {d['line']}: {d['message']}" for d in s.execute("macro.check", {})["diagnostics"]]


async def vba_set(ctx: Any, room: str, summary: str, module: str, code: str, kind: Literal["standard", "class", "form"] = "standard") -> Dict[str, Any]:
    """Writes a VBA module's whole code, making the module when it is new (``kind``: standard,
    class, or form for a UserForm). Runs once the person has accepted it. Answers with the
    project's compile errors, which the code keeps until they are fixed."""
    book, _ = _book(room)
    module = (module or "").strip()
    with _changing(book, room, summary) as (s, _):
        try:
            known = {m["name"].lower() for m in s.execute("macro.list", {})["modules"]}
            made = module.lower() not in known
            if made:
                s.execute("macro.addModule", {"kind": kind, "name": module, "code": code})
            else:
                s.execute("macro.setModule", {"name": module, "code": code})
            problems = _problems(s)
        except EngineError as ex:
            raise ValueError(str(ex).splitlines()[0])
    await push(book.key)
    return {"module": module, "made": made, "compile_errors": problems}


async def form_design(
    ctx: Any,
    room: str,
    summary: str,
    form: str,
    controls: Optional[List[Dict[str, Any]]] = None,
    caption: str = "",
    width: float = 0.0,
    height: float = 0.0,
    remove: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """Makes a UserForm when it is new and changes its design: caption, inside size (points),
    the ``controls`` given (each by its name: a new one needs ``kind``; the others keep what isn't
    given) and those to ``remove``. Runs once the person has accepted it."""
    book, _ = _book(room)
    form = (form or "").strip()
    with _changing(book, room, summary) as (s, _):
        try:
            if form.lower() not in {m["name"].lower() for m in s.execute("macro.list", {})["modules"]}:
                s.execute("macro.addModule", {"kind": "form", "name": form})

            def change(ops: List[Dict[str, Any]]) -> Dict[str, Any]:
                return s.execute("macro.editForm", {"name": form, "ops": ops})

            own = [{"op": "set", "key": k, "value": v} for k, v in (("caption", caption), ("width", width), ("height", height)) if v]
            if own:
                change(own)
            if remove:
                change([{"op": "delete", "control": name} for name in remove])
            for spec in controls or []:
                props = dict(spec)
                name = str(props.pop("name", "") or "")
                kind = str(props.pop("kind", "") or "")
                parent = str(props.pop("parent", "") or "")
                if "text" in props:
                    props["value"] = props.pop("text")
                if not name:
                    raise ValueError("every control needs its name")
                project = s.doc().wb.vba
                if read_form(project, project.module(form)).find(name) is None:
                    if not kind:
                        raise ValueError(f"{name} is new: say its kind (Label, TextBox, CommandButton, ComboBox, ListBox, CheckBox, OptionButton, Frame…)")
                    box = {k: props.pop(k) for k in ("left", "top", "width", "height") if k in props}
                    added = change([{"op": "add", "kind": kind, "parent": parent, **box}])["added"][0]
                    if added != name:
                        change([{"op": "set", "control": added, "key": "name", "value": name}])
                if props:
                    change([{"op": "set", "control": name, "key": k, "value": v} for k, v in props.items()])
            problems = _problems(s)
            project = s.doc().wb.vba
            design = read_form(project, project.module(form))
        except EngineError as ex:
            raise ValueError(str(ex).splitlines()[0])
    await push(book.key)
    return {"form": form, "caption": design.caption, "width": design.width, "height": design.height, "controls": [c.name for c in design.controls], "compile_errors": problems}


def _quote(name: str) -> str:
    if name.replace("_", "").isalnum() and not name[0].isdigit():
        return name
    return "'" + name.replace("'", "''") + "'"


ROOM = "The workbook: its room, as in /book/<room>."
SUMMARY = "One sentence in the person's language: what this change does. They read it before they accept."


def install(app: Any) -> bool:
    """Offers the tools to the AI Hub (on re_ddns); False without the platform library."""
    try:
        from reflex_ddns_auth.tools import tool
    except ImportError:
        return False
    # The tools that change a workbook are "danger" tools: the AI Hub asks the person first.
    tool(app, "grid.guide", title="How to use ReflexGrid's tools", kind="read")(guide)
    tool(app, "grid.workbook", title="A workbook at a glance", kind="read", params={"room": ROOM})(workbook)
    tool(
        app,
        "grid.read",
        title="Read cells",
        kind="read",
        params={
            "room": ROOM,
            "ref": "An A1 reference with its sheet, e.g. Sales!A1:D20 (also B:B, 3:3, a table's or a defined name's name).",
            "formulas": "Show each formula in brackets after its value (default true).",
        },
    )(read)
    tool(
        app,
        "grid.find",
        title="Find cells",
        kind="read",
        params={"room": ROOM, "text": "The words to look for in the cells' shown text and formulas.", "sheet": "Only this sheet (default: every sheet)."},
    )(find)
    tool(app, "grid.commands", title="List the commands", kind="read", params={"search": "Keep the commands with these words, e.g. sort, chart, row (default: all)."})(commands)
    tool(app, "grid.vba_get", title="Read VBA code", kind="read", params={"room": ROOM, "module": "The module's name."})(vba_get)
    tool(
        app,
        "grid.edit",
        title="Change the workbook",
        kind="danger",
        params={
            "room": ROOM,
            "summary": SUMMARY,
            "steps": 'The commands in order, each {"command": "cell.set", "params": {"cell": "Sales!B7", "input": "=SUM(B2:B6)"}}; grid_commands lists them.',
        },
    )(edit)
    tool(
        app,
        "grid.vba_set",
        title="Write VBA code",
        kind="danger",
        params={
            "room": ROOM,
            "summary": SUMMARY,
            "module": "The module's name (a new one is made).",
            "code": "The module's whole code.",
            "kind": "For a new module: standard, class, or form for a UserForm.",
        },
    )(vba_set)
    tool(
        app,
        "grid.form_design",
        title="Design a UserForm",
        kind="danger",
        params={
            "room": ROOM,
            "summary": SUMMARY,
            "form": "The UserForm's name (a new one is made).",
            "controls": 'Controls to add or change, each {"name": "txtName", "kind": "TextBox", "left": 80, "top": 10, "width": 150, "height": 18}, with "caption", "text", "default", "cancel" or other properties when wanted, and "parent" for one inside a Frame. Sizes are in points.',
            "caption": "The form's title.",
            "width": "The form's inside width, in points.",
            "height": "The form's inside height, in points.",
            "remove": "Names of controls to remove.",
        },
    )(form_design)
    return True
