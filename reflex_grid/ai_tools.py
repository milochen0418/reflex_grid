"""What the AI Hub's assistant may do in ReflexGrid: tools (``reflex_ddns_auth.tools``) that
read a workbook, so the assistant can answer questions about it and explain its formulas.
``room`` is the workbook's id (``/book/<room>``); the assistant passes the one it was opened
from. Nothing here changes a workbook.

A tool reads the workbook as it is at that moment (an immutable snapshot taken under the book's
lock), whoever has it open."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

from gridcore.core.addr import RangeRef, a1, col_to_letters
from gridcore.engine import EngineError
from gridcore.engine.cmd_edit import resolve_reference
from gridcore.engine.display import value_text

from .ai import short
from .books import STORE, valid_id

# What one read may return: enough for a table of a few thousand cells, not a whole sheet.
READ_MAX_CELLS = 3000
READ_MAX_CHARS = 60_000
READ_CELL_CHARS = 200
FIND_MAX = 50
FIND_MAX_CELLS = 400_000

GUIDE = """\
ReflexGrid is an Excel-compatible spreadsheet; each workbook is a room (/book/<room>).
- grid_workbook: its sheets with their used ranges, tables and charts, its defined names and its VBA modules. Start here.
- grid_read: cells in an A1 reference with the sheet ("Sales!A1:D20"; also B:B, a table or a defined name). Each cell shows its value as displayed and, for a formula, the formula in brackets.
- grid_find: cells whose text or formula contains some words.
- A column named in a question ("the Sales column") is usually a header in row 1, not a sheet: read the first rows (e.g. Sheet1!A1:Z3) to find it.
- Never guess what a cell holds: read it before you compute or explain anything.
- You can't change the workbook yourself. For a change, say exactly what to type where (formulas in English with commas, e.g. =SUM(B2:B9) in B10), which command to use (Data › Sort, Home › Conditional Formatting…), or give VBA to paste in.
- ReflexGrid runs Excel VBA itself, in the browser (no desktop Excel needed): modules, class modules and UserForms (also Controls.Add with WithEvents handlers), MsgBox and InputBox, and event macros such as Worksheet_Change once Developer › Macros On is set. The person adds code in Developer › Visual Basic (Module or UserForm in the editor's toolbar, then paste it) and runs a macro with Developer › Macros (Alt+F8); a workbook with macros is kept as .xlsm. Not available to macros: files, the shell, the network, COM, and the Application object's events.
- A message may begin with what the person's page shows (the sheet, the selection, the active cell and its formula): "this cell", "this formula" or "the selection" mean those."""


def _book(room: str):
    room = (room or "").strip()
    if not valid_id(room) or not STORE.exists(room):
        raise ValueError(f"there is no workbook {room!r}: pass the room the assistant was opened from")
    book = STORE.get(room) or STORE.open(room)
    with book.lock:
        return book, book.doc.wb


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


async def workbook(ctx: Any, room: str) -> Dict[str, Any]:
    """A workbook at a glance: its title, sheets (used range, tables with their columns, charts),
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
                "charts": [short(ch.title or ch.kind, 60) for ch in sh.charts],
            }
        )
    names = [{"name": n.name, "refers_to": "=" + n.formula.lstrip("="), **({"sheet": wb.sheets[n.scope].name} if isinstance(n.scope, int) and 0 <= n.scope < len(wb.sheets) else {})} for n in wb.names if not n.hidden][:60]
    vba = [{"name": m.name, "kind": m.kind} for m in wb.vba.modules] if wb.vba is not None else []
    return {"room": book.key, "title": book.meta.title, "sheets": sheets, "names": names, "vba": vba}


def _cell_text(wb, sh, rc: Tuple[int, int], formulas: bool) -> str:
    cell = sh.cell(rc)
    if cell is None:
        return ""
    shown = short(value_text(wb, sh, rc), READ_CELL_CHARS)
    if formulas and cell.formula is not None:
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


def _quote(name: str) -> str:
    if name.replace("_", "").isalnum() and not name[0].isdigit():
        return name
    return "'" + name.replace("'", "''") + "'"


ROOM = "The workbook: its room, as in /book/<room>."


def install(app: Any) -> bool:
    """Offers the tools to the AI Hub (on re_ddns); False without the platform library."""
    try:
        from reflex_ddns_auth.tools import tool
    except ImportError:
        return False
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
    return True
