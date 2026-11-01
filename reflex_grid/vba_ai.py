"""AI in the VBA editor, through the AI Hub (see ``ai.py``), as Cellora has for a notebook's cells:
it writes or changes code from what the person types or says (the role ``coder``), fixes compile
and run-time errors, and suggests the rest of a line in grey (the role ``fast``). Which models
answer is chosen on the hub's page.

The AI never touches the workbook. It answers with the procedures it adds or changes, which go
into the module by their names (``merged``), so that what it leaves out stays, whatever the model;
selected lines are replaced as they are. The result is tried by ReflexGrid's own compiler (an
answer that doesn't compile is asked again, with the error), shown as a change to accept or
reject, and goes in like the person's own edit.

This module only words the requests and reads the answers; ``vba_ai_state`` runs them."""

from __future__ import annotations

import re
from dataclasses import dataclass, replace
from typing import Any, Dict, List, Optional, Sequence, Tuple

from gridcore.core.addr import col_to_letters
from gridcore.engine.display import value_text
from gridcore.engine.macro.project import compile_vba, form_design
from gridcore.model.vba import STANDARD, VbaModule, default_attributes, new_project, normalize_code

from .ai import short

WHOLE_MODULE_LINES = 400  # up to this the AI sees the whole module
AROUND_LINES = 60  # of a bigger module, around what it works on
BRIEF_CHARS = 5000  # what it is told of the workbook
BRIEF_COLUMNS = 14
BRIEF_PROCEDURES = 16
ROUNDS = 3  # answers per request: the first, and two with the compile error they had
INLINE_BEFORE = 6000  # of the code around the cursor, for a grey suggestion
INLINE_AFTER = 1500
INLINE_BRIEF = 1500  # and of the workbook
INLINE_LINES = 8
ECHO_MIN = 3  # characters a suggestion repeats of the code before they count as an echo

SYSTEM = """\
You change one module of Excel VBA inside ReflexGrid, an Excel-compatible spreadsheet that runs VBA \
itself, with no desktop Excel: modules, class modules, UserForms, MsgBox and InputBox, \
Scripting.Dictionary, and event macros such as Worksheet_Change. Macros can't reach anything outside \
the workbook: no files, other workbooks, Shell, network, Windows API calls or Application.OnTime.
You get the workbook (its sheets, tables, names, forms and other modules), the module and what the \
person wants. Their words may come from speech recognition: read them for what they must mean, and \
take the names from the workbook.

Reply with one ```vba fenced block and nothing before or after it, holding only what you add or change:
- each procedure you add or change, whole, from its Sub, Function or Property line to its End line;
- before them, any new module-level declarations (Option, Dim, Const, Type).
Procedures you leave out stay as they are: never repeat one that doesn't change. To remove a \
procedure, or the old name of one you rename, add a line: ' REMOVE Name
When the question has a <target>, the block instead holds the complete new text of those lines and \
nothing else. No Attribute lines.
- Use the sheets, tables, columns, controls and procedures that exist, by their exact names. Never \
invent one. Event procedures keep the names VBA gives them (CommandButton1_Click, Worksheet_Change).
- Write comments and messages in the language of the request. No emoji in the code.
- When the person only asks about the code and wants nothing changed, answer that in a few plain \
sentences in their language, with no code block."""

INLINE_SYSTEM = """\
You are a code completion engine inside a VBA editor. You get the workbook's names, then the module \
with the cursor marked <CURSOR>.
Reply with only the text to insert at the cursor: the rest of the current line, and at most a few \
more lines when the code clearly continues. No fences, no explanation, nothing that is already after \
the cursor. Use the sheets, tables and procedures that exist. Reply with nothing when no completion \
is useful."""

_KINDS = {"standard": "module", "class": "class module", "form": "UserForm's code", "document": "code of a sheet or the workbook"}
_HEAD = re.compile(r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\W\d]\w*)", re.IGNORECASE)
_END = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.IGNORECASE)
_PRIVATE = re.compile(r"^\s*Private\b", re.IGNORECASE)
_REMOVE = re.compile(r"^\s*'\s*REMOVE\s+(?:(?:Sub|Function|Property)\s+)?([^\W\d]\w*)", re.IGNORECASE)
_DECLARATION = re.compile(r"^\s*(?:Option|Public|Private|Global|Dim|Const|Declare|Implements|Def[A-Z][a-z]+|#\w+)\b|^\s*#", re.IGNORECASE)
_BLOCK = re.compile(r"^\s*(?:(?:Public|Private)\s+)?(?:Type|Enum)\s+\w+", re.IGNORECASE)
_BLOCK_END = re.compile(r"^\s*End\s+(?:Type|Enum)\b", re.IGNORECASE)
_OPTION = re.compile(r"^\s*Option\b", re.IGNORECASE)
_FENCE = re.compile(r"```[ \t]*[\w+-]*[ \t]*\n(.*?)(?:\n[ \t]*```|\Z)", re.DOTALL)
_ATTRIBUTE = re.compile(r"^\s*Attribute\s+VB_\w+\s*=.*\n?", re.IGNORECASE | re.MULTILINE)
_CODE_START = re.compile(r"^\s*(?:'|Option\s|Sub\s|Function\s|Public\s|Private\s|Friend\s|Dim\s|Const\s|Property\s|Type\s|Enum\s|#If\s)", re.IGNORECASE)
_PROSE = re.compile(r"(Sure|Certainly|Here(?:'s| is| are)|Below|To |I |You can|This (?:code|function))\b")


# ------------------------------------------------------------ a module's procedures


@dataclass
class _Procedure:
    key: str  # a new version takes the place of the one with its key
    first: int  # its first line (from 0), with the comments right above it
    head: int  # its Sub, Function or Property line
    end: int  # the line after its last


def _procedures(lines: Sequence[str]) -> List[_Procedure]:
    """The procedures in ``lines``. One that isn't closed (a compile error) ends where the next one
    starts."""
    out: List[_Procedure] = []
    i = 0
    while i < len(lines):
        m = _HEAD.match(lines[i])
        if not m:
            i += 1
            continue
        first = i
        while first > 0 and lines[first - 1].lstrip().startswith("'"):
            first -= 1
        end = i + 1
        while end < len(lines) and not _END.match(lines[end]) and not _HEAD.match(lines[end]):
            end += 1
        if end < len(lines) and _END.match(lines[end]):
            end += 1
        else:
            while end > i + 1 and not lines[end - 1].strip():
                end -= 1
        kind = " ".join(m.group(1).lower().split())
        out.append(_Procedure((kind + " " if kind.startswith("property") else "") + m.group(2).lower(), first, i, end))
        i = end
    return out


def procedure_at(lines: Sequence[str], line: int) -> Optional[Tuple[int, int]]:
    """The first and last line (from 1) of the procedure that ``line`` (from 1) is in."""
    for p in _procedures(lines):
        if p.first < line <= p.end:
            return p.head + 1, p.end
    return None


def _same(line: str) -> str:
    return " ".join(line.split()).lower()


def _loose_lines(lines: Sequence[str], parts: Sequence[_Procedure]) -> Tuple[List[str], List[str]]:
    """What an answer holds outside its procedures: the declarations, and anything else
    (statements for the place the cursor is)."""
    inside = {i for p in parts for i in range(p.first, p.end)}
    declarations: List[str] = []
    other: List[str] = []
    block = False
    for i, line in enumerate(lines):
        if i in inside or not line.strip():
            continue
        if block or _BLOCK.match(line):
            declarations.append(line)
            block = not _BLOCK_END.match(line)
        elif _DECLARATION.match(line):
            declarations.append(line)
        elif not line.lstrip().startswith("'"):
            other.append(line)
    return declarations, other


def _place(lines: Sequence[str], known: Sequence[_Procedure], line: int) -> int:
    """Where a new procedure goes: after the procedure at the cursor, or at a cursor below the
    first procedure; else (a cursor among the declarations, where a module opens) at the end."""
    at = line - 1
    if at < 0 or not known or at < known[0].first:
        return len(lines)
    for p in known:
        if p.first <= at < p.end:
            return p.end
    return min(line, len(lines))


def _splice(lines: List[str], start: int, end: int, block: List[str], apart: bool) -> List[str]:
    """``lines`` with ``block`` for its lines ``start`` to ``end``; ``apart``: one blank line on
    each side."""
    before, after = lines[:start], lines[end:]
    if not apart:
        return before + block + after
    while before and not before[-1].strip():
        before.pop()
    while after and not after[0].strip():
        after.pop(0)
    if not block:
        return before + ([""] if before and after else []) + after
    return before + ([""] if before else []) + block + ([""] if after else []) + after


def merged(code: str, answer: str, line: int = 0) -> str:
    """The module ``code`` with an answer in it (the cursor on ``line``). A procedure replaces the
    one of its name; a new one goes after the cursor; new declarations go up with the others;
    ``' REMOVE Name`` takes a procedure out; statements alone go after the cursor's line. All that
    the answer leaves out stays."""
    new = answer.replace("\r\n", "\n").strip("\n").split("\n")
    removed = [m.group(1).lower() for m in map(_REMOVE.match, new) if m]
    new = [text for text in new if not _REMOVE.match(text)]
    if not code.strip():
        return "\n".join(new)
    lines = code.split("\n")
    known = _procedures(lines)
    parts = _procedures(new)
    declarations, other = _loose_lines(new, parts)
    edits: List[Tuple[int, int, List[str], bool]] = []
    fresh: List[List[str]] = []
    by_key = {p.key: p for p in parts}
    for p in by_key.values():
        old = next((k for k in known if k.key == p.key), None)
        block = new[p.first : p.end]
        if old is None:
            fresh.append(block)
        else:
            edits.append((old.first if p.first < p.head else old.head, old.end, block, False))
    for k in known:
        if k.key not in by_key and k.key.rpartition(" ")[2] in removed:
            edits.append((k.first, k.end, [], True))
    if fresh:
        block = [text for i, one in enumerate(fresh) for text in ([""] if i else []) + one]
        at = _place(lines, known, line)
        edits.append((at, at, block, True))
    if other and not parts:
        # Statements alone, for the place the cursor is.
        at = max(0, min(line, len(lines)))
        blank = at > 0 and not lines[at - 1].strip()
        edits.append((at - 1, at, new, False) if blank else (at, at, new, False))
    else:
        end = known[0].first if known else len(lines)
        have = {_same(text) for text in lines[:end]}
        add = [text for text in declarations if _same(text) not in have]
        options = [text for text in add if _OPTION.match(text)]
        if options:
            edits.append((0, 0, options, False))
        if len(add) > len(options):
            at = end
            while at > 0 and not lines[at - 1].strip():
                at -= 1
            rest = [text for text in add if not _OPTION.match(text)]
            edits.append((at, at, rest + ([""] if at == end and known else []), False))
    for start, stop, block, apart in sorted(edits, key=lambda e: (e[0], e[1]), reverse=True):
        lines = _splice(lines, start, stop, block, apart)
    out = "\n".join(lines)
    return out + "\n" if code.endswith("\n") and not out.endswith("\n") else out


# ------------------------------------------------------------ what to write


def _indent(line: str) -> int:
    return len(line) - len(line.lstrip(" "))


def _indented(new: List[str], old: Sequence[str]) -> List[str]:
    """Lines put in the place of ``old``, as far in as those were: a model writes them from the
    margin."""
    full = [text for text in new if text.strip()]
    was = [text for text in old if text.strip()]
    if not full or not was:
        return new
    more = min(_indent(text) for text in was) - min(_indent(text) for text in full)
    return [" " * more + text if more > 0 and text.strip() else text for text in new]


@dataclass
class Target:
    """What a request is about: the module ``code``, the cursor on ``line`` (from 1). With ``first``,
    its lines ``first`` to ``last``, which the answer replaces as it is; else the answer goes into
    the module by the names of its procedures."""

    code: str
    line: int = 0
    first: int = 0
    last: int = 0

    def result(self, answer: str) -> str:
        """The module with ``answer`` in it."""
        if not self.first:
            return merged(self.code, answer, self.line)
        lines = self.code.split("\n")
        new = answer.replace("\r\n", "\n").strip("\n").split("\n") if answer.strip() else []
        return "\n".join(lines[: self.first - 1] + _indented(new, lines[self.first - 1 : self.last]) + lines[self.last :])

    def after(self, answer: str) -> "Target":
        """The same target with ``answer`` in it, for a request that goes on from it."""
        count = len(answer.strip("\n").split("\n")) if answer.strip() else 1
        return replace(self, code=self.result(answer), last=self.first + count - 1 if self.first else 0)


def target_of(code: str, line: int = 0, first: int = 0, last: int = 0) -> Target:
    """What a request is about: the selected lines ``first`` to ``last``, else the module."""
    lines = code.split("\n")
    count = len(lines) - 1 if lines[-1] == "" and len(lines) > 1 else len(lines)
    if not first:
        return Target(code, line)
    first = max(1, min(first, count))
    return Target(code, line, first, max(first, min(last or first, count)))


def fix_target(code: str, line: int, compile_error: bool = True) -> Target:
    """What fixing an error on ``line`` is about. A compile error is where it shows: the procedure
    around it, whole (and the one before, when it is the next one's first line that shows a
    procedure left open), or outside the procedures the lines around it, which an answer of
    procedures couldn't reach. A run-time error may come from anywhere in the module."""
    lines = code.split("\n")
    known = _procedures(lines)
    if not line:
        return Target(code, line)
    at = next((i for i, p in enumerate(known) if p.first < line <= p.end), None)
    if at is None:
        return target_of(code, line, max((p.end for p in known if p.end < line), default=0) + 1, min((p.first for p in known if p.first >= line), default=len(lines)))
    if not compile_error:
        return Target(code, line)
    start = known[at - 1] if at and line == known[at].head + 1 else known[at]
    return target_of(code, line, start.first + 1, known[at].end)


# ------------------------------------------------------------ what the AI is told


def _controls(project: Any, module: Any) -> str:
    out: List[str] = []

    def walk(controls: Sequence[Any]) -> None:
        for c in controls:
            out.append(f'{c.name} ({c.kind}' + (f', "{short(c.caption, 30)}")' if c.caption else ")"))
            walk(c.children)

    walk(form_design(project, module).controls)
    return ", ".join(out) or "(none)"


def _signatures(code: str, private: bool = False, limit: int = BRIEF_PROCEDURES) -> List[str]:
    found = [" ".join(line.split()) for line in code.split("\n") if _HEAD.match(line) and (private or not _PRIVATE.match(line))]
    if len(found) > limit:
        found[limit:] = ["…"]
    return found


def workbook_brief(wb: Any, module: str = "") -> str:
    """The workbook as the AI sees it: sheets with their first rows, tables, names, and the other
    modules with their public procedures; a UserForm's controls."""
    lines: List[str] = []
    for sh in wb.sheets:
        used = sh.used_range()
        code = f", code name {sh.code_name}" if sh.code_name else ""
        lines.append(f'Sheet "{sh.name}"{code}' + (f", cells {used.a1()}" if used is not None else ", empty") + (", hidden" if sh.visibility != "visible" else ""))
        for t in sh.tables:
            lines.append(f"  table {t.name} at {t.range.a1()}, columns: " + ", ".join(c.name for c in t.columns))
        if used is None:
            continue
        for r in range(used.r0, min(used.r1, used.r0 + 1) + 1):
            cells = [f"{col_to_letters(c)}{r + 1}: {short(value_text(wb, sh, (r, c)), 24)}" for c in range(used.c0, min(used.c1, used.c0 + BRIEF_COLUMNS - 1) + 1) if sh.cell((r, c)) is not None]
            if cells:
                lines.append("  " + " | ".join(cells))
    names = [f"{n.name} = {short(n.formula, 60)}" for n in wb.names if not n.hidden][:30]
    if names:
        lines.append("Defined names: " + "; ".join(names))
    project = wb.vba
    if project is not None and not project.locked:
        for m in project.modules:
            if m.kind == "form":
                lines.append(f"UserForm {m.name}, controls: {_controls(project, m)}")
            if m.name.lower() != module.lower() and m.code.strip():
                procs = _signatures(m.code)
                if procs:
                    lines.append(f"Module {m.name} ({m.kind}): " + "; ".join(procs))
    return "\n".join(lines)[:BRIEF_CHARS]


def write_prompt(brief: str, name: str, kind: str, target: Target, request: str) -> str:
    """The question: the workbook, the module (around the cursor in a big one) and the request."""
    about = f'name="{name}" kind="{_KINDS.get(kind, kind)}"'
    lines = target.code.split("\n")
    parts = [f"<workbook>\n{brief}\n</workbook>"]
    if target.first:
        before = "\n".join(lines[: target.first - 1][-AROUND_LINES:])
        after = "\n".join(lines[target.last :][:AROUND_LINES])
        parts.append(f"<module_before {about}>\n{before}\n</module_before>\n\n<module_after>\n{after}\n</module_after>")
        parts.append(f'<target lines="{target.first}-{target.last}">\n' + "\n".join(lines[target.first - 1 : target.last]) + "\n</target>")
    else:
        where = ""
        if target.line and target.code.strip():
            box = procedure_at(lines, target.line)
            name_at = _HEAD.match(lines[box[0] - 1]).group(2) if box else ""
            where = f"The person's cursor is on line {target.line}" + (f', in {name_at}: "this" and "here" mean {name_at}.' if box else ".")
        if len(lines) <= WHOLE_MODULE_LINES:
            parts.append(f"<module {about}>\n{target.code}\n</module>")
        else:
            middle = procedure_at(lines, target.line) or (target.line, target.line)
            first = max(1, middle[0] - AROUND_LINES // 2)
            last = min(len(lines), middle[1] + AROUND_LINES // 2)
            shown = "\n".join(lines[first - 1 : last])
            others = "\n".join(_signatures(target.code, private=True, limit=200))
            parts.append(f'<module {about} shown="lines {first}-{last} of {len(lines)}">\n{shown}\n</module>\n\n<its_procedures>\n{others}\n</its_procedures>')
        if where:
            parts.append(where)
    parts.append(f"<request>\n{request.strip()}\n</request>")
    return "\n\n".join(parts)


def _line_text(code: str, line: int) -> str:
    lines = code.split("\n")
    return lines[line - 1].strip() if 1 <= line <= len(lines) else ""


def fix_request(message: str, line: int, code: str) -> str:
    """What to ask for an error the person sees."""
    shown = _line_text(code, line)
    return f"Fix this error: {message}" + (f"\nLine {line} of the module: {shown}" if shown else "") + "\nKeep what the code is meant to do; change only what is needed."


def retry_request(problems: Sequence[Dict[str, Any]], code: str, target: Target) -> str:
    """What to say when an answer doesn't compile (``code``: the module with the answer in it)."""
    lines = [f"That doesn't compile. Line {p['line']} of the module: {p['message']}" + (f" ({_line_text(code, int(p['line']))})" if _line_text(code, int(p["line"])) else "") for p in problems]
    again = "the complete corrected text of the target's lines" if target.first else "each procedure you add or change, whole, from its Sub, Function or Property line to its End line"
    return "\n".join(lines) + f"\nReply again, in one fenced block, with {again}."


def inline_prompt(prefix: str, suffix: str, brief: str) -> str:
    return f"<workbook>\n{brief}\n</workbook>\n\n<module>\n{prefix[-INLINE_BEFORE:]}<CURSOR>{suffix[:INLINE_AFTER]}\n</module>"


# ------------------------------------------------------------ reading the answers


def code_of(answer: str) -> str:
    """The code in an answer, also while it is still being written: its first fenced block, else
    the answer when it reads as code, else "" (a model that left the fence out; prose is an
    answer in words)."""
    m = _FENCE.search(answer)
    if m:
        text = m.group(1)
    elif "```" in answer:
        return ""
    elif _CODE_START.match(answer.lstrip("\n")):
        text = answer
    else:
        return ""
    return _ATTRIBUTE.sub("", text).rstrip("`").strip("\n")


def _past(code: str, text: str) -> int:
    """How far ``text`` runs over ``code``, when it starts with it, spaces aside; else -1."""
    at = 0
    for ch in code:
        if ch.isspace():
            continue
        while at < len(text) and text[at].isspace():
            at += 1
        if at == len(text) or text[at] != ch:
            return -1
        at += 1
    return at


def _without_echo(text: str, prefix: str, suffix: str) -> str:
    """What a model adds, when it writes out code around the cursor again (some write out the whole
    module, with its spaces put right)."""
    lines = prefix[-INLINE_BEFORE:].split("\n")
    for start in range(len(lines)):
        code = "\n".join(lines[start:])
        at = _past(code, text)
        if at >= 0 and len("".join(code.split())) >= ECHO_MIN:
            text = text[at:].lstrip() if prefix[-1:].isspace() else text[at:]
            break
    else:
        tail = prefix[-2000:]
        for size in range(min(len(tail), len(text)), ECHO_MIN - 1, -1):
            if text.startswith(tail[-size:]):
                text = text[size:]
                break
    body = text.rstrip()
    rest = suffix[:INLINE_AFTER].rstrip().split("\n")
    for end in range(len(rest), 0, -1):
        code = "\n".join(rest[:end])
        at = _past(code[::-1], body[::-1])
        if at >= 0 and code.strip():
            return body[: len(body) - at]
    for size in range(min(len(rest[0]), len(body)), 0, -1):
        if body.endswith(rest[0][:size]):
            return body[:-size]
    return text


def clean_inline(text: str, prefix: str, suffix: str) -> str:
    """A suggestion trimmed to what helps: no fences, not the code that is already there, at most
    a handful of lines."""
    text = text.replace("<CURSOR>", "")
    if "```" in text and not text.lstrip().startswith("```"):
        return ""  # prose around a block: a chat answer, not a completion
    text = re.sub(r"^```[\w-]*\n?", "", text.strip("\n") if text.startswith("```") else text)
    text = re.sub(r"\n?```\s*$", "", text)
    if not text.strip() or _PROSE.match(text.lstrip()):
        return ""
    text = _without_echo(text, prefix, suffix)
    if not text.strip():
        return ""
    text = "\n".join(text.split("\n")[:INLINE_LINES])
    return text.rstrip() if "\n" in text else text


# ------------------------------------------------------------ trying an answer


def compile_problems(project: Any, name: str, code: str) -> List[Dict[str, Any]]:
    """The compile errors of the module ``name`` if its code were ``code`` (a new standard module
    when there is none). The workbook isn't touched: a project is an immutable value."""
    name = name or "Module1"
    base = project if project is not None else new_project([])
    if base.module(name) is None:
        trial = base.with_module(VbaModule(name, STANDARD, normalize_code(code), default_attributes(name, STANDARD)))
    else:
        trial = base.with_code(name, code)
    return [{"line": e.line, "message": e.msg} for n, e in compile_vba(trial).errors.items() if n.lower() == name.lower()]


def new_problems(project: Any, name: str, before: str, after: str, fixing: bool = False) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """The compile errors of the code ``after``, and those of them that the change brought (when
    ``fixing``, all): an error the code ``before`` had already isn't an answer's to answer for."""
    problems = compile_problems(project, name, after)
    if fixing or not problems:
        return problems, problems
    known = {(p["message"], _line_text(before, int(p["line"]))) for p in compile_problems(project, name, before)}
    return problems, [p for p in problems if (p["message"], _line_text(after, int(p["line"]))) not in known]


def breaks(project: Any, name: str, prefix: str, text: str, suffix: str) -> bool:
    """Whether a suggestion is what the compiler stumbles on. Code being typed often doesn't
    compile: only an error on the suggestion's lines counts, and not one on the cursor's line that
    was there without it."""
    problems = compile_problems(project, name, prefix + text + suffix)
    if not problems:
        return False
    line, first = int(problems[0]["line"]), prefix.count("\n") + 1
    if not first <= line <= first + text.count("\n"):
        return False
    return line != first or not any(int(p["line"]) == first for p in compile_problems(project, name, prefix + suffix))
