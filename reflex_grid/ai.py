"""AI in ReflexGrid, through the AI Hub of re_ddns (``reflex_ddns_auth.ai``): AI Fill (a value
for each row of a selection, written into a column as values, in one undo step) and the formula
suggested when a cell's entry starts with ``=``. Which model answers is chosen on the AI Hub's
page (the role ``fast``); the app holds no model, key or address of its own.

The AI Hub's assistant works on a workbook through the tools in ``ai_tools``.

Without the platform library, or with no hub to ask (neither ``RE_DDNS_API_URL``, which re_ddns
sets in its apps, nor ``DDNS_AI_HUB_URL``), these features stay hidden.

This module only words the requests and reads the answers; ``state.py`` runs them."""

from __future__ import annotations

import asyncio
import os
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

from gridcore.core.addr import RangeRef, a1, col_to_letters
from gridcore.engine.display import value_text
from gridcore.engine.entry import is_formula_input

try:
    from reflex_ddns_auth import ai as hub
except ImportError:  # runs without the platform library: no AI
    hub = None

ENABLED = hub is not None and bool(os.environ.get("RE_DDNS_API_URL") or os.environ.get("DDNS_AI_HUB_URL"))
# This app on the platform, as the AI Hub knows it (its tools' prefix is "grid").
APP = (os.environ.get("SERVICE_SUBDOMAIN") or os.environ.get("DDNS_APP_ID") or "grid").strip().lower()

ROLE = "fast"
USES = "AI Fill (a value for each row of a selection) and the formula suggested when you type = in a cell"


def declare(app: Any) -> None:
    """Tells the AI Hub what ReflexGrid asks its models for (shown on the hub's page)."""
    if ENABLED:
        hub.uses(app, fast=USES)


# What asking may raise: the hub isn't there or couldn't answer, or the answer took too long.
FAILED = (hub.HubUnavailable, asyncio.TimeoutError) if hub is not None else (asyncio.TimeoutError,)


async def ask(system: str, prompt: str, *, max_tokens: int, user: str, timeout: float) -> Tuple[str, str]:
    """The answer of the model the hub has for :data:`ROLE`, and that model's name ("" when it
    declined); raises one of :data:`FAILED`."""
    if hub is None:
        raise asyncio.TimeoutError("no AI here")
    reply = hub.stream(ROLE, system, [{"role": "user", "content": prompt}], max_tokens=max_tokens, effort="low", user=user)

    async def collect() -> str:
        return "".join([piece async for piece in reply])

    text = await asyncio.wait_for(collect(), timeout)
    return ("" if reply.stop == "refusal" else text), reply.model


def short(text: Any, limit: int) -> str:
    """One line of at most ``limit`` characters."""
    t = " ".join(str(text if text is not None else "").split())
    return t if len(t) <= limit else t[: limit - 1] + "…"


# ------------------------------------------------------------ AI Fill

FILL_BATCH = 20  # rows per question to the model
FILL_MAX_ROWS = 1000  # rows in one AI Fill
FILL_CELL_CHARS = 300  # of each cell's text sent

FILL_SYSTEM = """\
You fill one column of a spreadsheet. You get an instruction and numbered rows; each row lists \
its cells as "name: value". Apply the instruction to each row on its own, and answer with \
exactly one line per row, in order: the row's number, a dot, a space and the answer, e.g. \
"3. Positive". Each answer is short (a word, a number, a date or a short phrase), on one line, \
without quotes or explanations. When a row gives you nothing to work on, write its number and \
the dot only. Unless the instruction says otherwise, answer in the language of the rows."""

TITLE_LINE = """\
Before the rows' answers, write one line "Title: " and a column title of one to three words \
for these answers, in the same language and script as the answers."""


def fill_prompt(instruction: str, names: Sequence[str], rows: Sequence[Sequence[str]], title: bool = False) -> str:
    """The question for one batch: ``rows`` are the cells' texts, ``names`` the columns'
    names (their headers, else their letters)."""
    lines = [f"Instruction: {short(instruction, 1000)}", ""]
    if title:
        lines += [TITLE_LINE, ""]
    lines.append("Rows:")
    for i, row in enumerate(rows, 1):
        cells = [f"{n}: {short(v, FILL_CELL_CHARS)}" for n, v in zip(names, row) if str(v).strip()]
        lines.append(f"{i}. " + (" | ".join(cells) if cells else "(empty)"))
    return "\n".join(lines)


_NUMBERED = re.compile(r"^\s*(?:row\s*)?(\d{1,4})\s*[.)\]:：、]\s?(.*)$", re.IGNORECASE)
_TITLE = re.compile(r"^\s*\**\s*title\s*\**\s*[:：]\s*(.*)$", re.IGNORECASE)


def _clean(answer: str) -> str:
    a = answer.strip().strip("*`").strip()
    if len(a) >= 2 and a[0] == a[-1] and a[0] in "\"'“”":
        a = a[1:-1].strip()
    if a in ("-", "—", "–", "(empty)", "N/A"):
        return ""
    return a


def parse_fill(text: str, n: int) -> Tuple[str, Dict[int, str]]:
    """The column title (when asked for) and each row's answer (1-based) in a batch's reply.
    Rows the reply leaves out are missing from the dict; answered with nothing, they are ""."""
    title = ""
    answers: Dict[int, str] = {}
    loose: List[str] = []
    for line in (text or "").splitlines():
        if not line.strip() or line.strip().startswith("```"):
            continue
        m = _TITLE.match(line)
        if m and not title and not answers:
            title = short(_clean(m.group(1)), 60)
            continue
        m = _NUMBERED.match(line)
        if m:
            i = int(m.group(1))
            if 1 <= i <= n and i not in answers:
                answers[i] = _clean(m.group(2))
            continue
        loose.append(line)
    if not answers and len(loose) == n:
        # A model that left the numbers out but kept one line per row.
        answers = {i: _clean(v) for i, v in enumerate(loose, 1)}
    return title, answers


def as_input(answer: str) -> str:
    """An answer as typed input that stays the answer: a number, a date or text, never a
    formula (text that would read as one gets the apostrophe Excel uses for text)."""
    a = answer.replace("\r", "").replace("\n", " ").strip()
    if is_formula_input(a) or a.startswith("="):
        return "'" + a
    return a


def column_names(wb: Any, sheet: Any, rng: RangeRef, header: bool) -> List[str]:
    """What the model calls each column of ``rng``: its header (the first row with ``header``,
    else the row above when that holds text), else its letter."""
    out = []
    for c in range(rng.c0, rng.c1 + 1):
        name = ""
        if header:
            name = value_text(wb, sheet, (rng.r0, c))
        elif rng.r0 > 0 and isinstance(sheet.value((rng.r0 - 1, c)), str):
            name = value_text(wb, sheet, (rng.r0 - 1, c))
        out.append(short(name, 40) or col_to_letters(c))
    return out


def looks_like_header(sheet: Any, rng: RangeRef) -> bool:
    """Whether the first row of ``rng`` reads as headers: short text in every column, more rows
    below, and nothing above it to be the headers instead."""
    if rng.r1 <= rng.r0:
        return False
    for c in range(rng.c0, rng.c1 + 1):
        v = sheet.value((rng.r0, c))
        if not isinstance(v, str) or not v.strip() or len(v) > 40:
            return False
        if rng.r0 > 0 and sheet.value((rng.r0 - 1, c)) not in (None, ""):
            return False
    return True


def free_column(sheet: Any, rng: RangeRef, rows: Tuple[int, int], limit: int = 30) -> int:
    """The first column right of ``rng`` with nothing in ``rows`` (first, last), for results."""
    for c in range(rng.c1 + 1, rng.c1 + 1 + limit):
        if all(sheet.value((r, c)) in (None, "") for r in range(rows[0], rows[1] + 1)):
            return c
    return rng.c1 + 1


# ------------------------------------------------------------ formula suggestions

SUGGEST_ROWS_ABOVE = 14
SUGGEST_COLS_LEFT = 6
SUGGEST_CELL_CHARS = 24

SUGGEST_SYSTEM = """\
You suggest one Excel formula for the cell a person is starting to type into. You see the cells \
around it with their addresses; the cell is marked "<<here>>". Suggest the formula they most \
likely want there: a total or average of the numbers above or to the left, the same calculation \
as the rows above, a lookup, a date difference… Let the headers say what the cell holds (a \
Profit column is revenue minus cost, a Total adds up its parts). Use Excel's English function names, commas \
between arguments and references to the cells shown; never refer to the cell itself. Answer \
with exactly two lines and nothing else:
FORMULA: =...
WHY: what it does, in at most twelve words"""


def neighbourhood(wb: Any, si: int, at: Tuple[int, int]) -> str:
    """The cells around ``at`` as the model sees them: a small grid with addresses, formulas in
    brackets after their values, the headers of the first row when it is out of sight."""
    sh = wb.sheet(si)
    r, c = at
    used = sh.used_range()
    last_r = max(r, used.r1 if used is not None else r)
    last_c = max(c, used.c1 if used is not None else c)
    r0, r1 = max(0, r - SUGGEST_ROWS_ABOVE), min(last_r, r + 2)
    c0, c1 = max(0, c - SUGGEST_COLS_LEFT), min(last_c, c + 3)

    def cell(rr: int, cc: int) -> str:
        if (rr, cc) == (r, c):
            return "<<here>>"
        x = sh.cell((rr, cc))
        if x is None:
            return ""
        shown = short(value_text(wb, sh, (rr, cc)), SUGGEST_CELL_CHARS)
        if x.formula is not None:
            return f"{shown} [={short(x.formula.text, 60)}]"
        return shown

    rows_out = []
    if r0 > 0:
        heads = [cell(0, cc) for cc in range(c0, c1 + 1)]
        if any(heads):
            rows_out.append("1\t" + "\t".join(heads))
            rows_out.append("…")
    for rr in range(r0, r1 + 1):
        rows_out.append(f"{rr + 1}\t" + "\t".join(cell(rr, cc) for cc in range(c0, c1 + 1)))
    head = "\t" + "\t".join(col_to_letters(cc) for cc in range(c0, c1 + 1))
    return f'Sheet "{sh.name}", cell {a1(r, c)}. The cells around it (tab-separated):\n{head}\n' + "\n".join(rows_out)


def suggest_prompt(context: str, problem: str = "") -> str:
    """The question for a suggestion (``problem``: why the last answer was refused)."""
    text = context + "\n\nWhich formula goes in the cell marked <<here>>?"
    if problem:
        text += f"\n\nYour last answer can't be used: {short(problem, 200)}. Suggest a corrected formula."
    return text


_FORMULA = re.compile(r"^\s*\**\s*formula\s*\**\s*[:：]\s*\**\s*`*\s*(=.*?)\s*`*\s*$", re.IGNORECASE)
_WHY = re.compile(r"^\s*\**\s*why\s*\**\s*[:：]\s*(.*)$", re.IGNORECASE)


def parse_suggestion(text: str) -> Tuple[str, str]:
    """``(formula, why)`` from a reply ("" when there is no formula in it)."""
    formula = why = ""
    for line in (text or "").splitlines():
        m = _FORMULA.match(line)
        if m and not formula:
            formula = m.group(1)
            continue
        m = _WHY.match(line)
        if m and not why:
            why = short(m.group(1).strip("*` "), 120)
    if not formula:
        # A bare formula on a line of its own (in a code block, say).
        for line in (text or "").splitlines():
            t = line.strip().strip("`").strip()
            if t.startswith("=") and len(t) > 1:
                formula = t
                break
    return formula.strip(), why
