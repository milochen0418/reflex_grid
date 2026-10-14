"""Splitting VBA source into tokens.

Physical lines joined by a line continuation (`` _`` at the end of a line, also inside a
comment) form one logical line; tokens remember the physical line they start on. Comments
(``'`` and ``Rem``) are dropped. Words are not classified: the parser decides when a word is a
keyword, because most VBA keywords are also valid member names (``.Name``, ``.Text``,
``.End``…)."""

from __future__ import annotations

import re
from typing import List, Optional, Tuple

from .errors import VbaSyntaxError


class Tok:
    __slots__ = ("kind", "text", "up", "value", "suffix", "line", "col", "space_before")

    def __init__(self, kind: str, text: str, line: int, col: int, value=None, suffix: str = "", space_before: bool = False):
        self.kind = kind  # word num str date op bracket nl eof
        self.text = text
        self.up = text.upper() if kind == "word" else text
        self.value = value
        self.suffix = suffix
        self.line = line
        self.col = col
        self.space_before = space_before

    def __repr__(self) -> str:
        return f"Tok({self.kind}, {self.text!r}{', ' + repr(self.value) if self.value is not None else ''}, L{self.line})"


_SUFFIXES = "%&!#@$^"
_NUM_RE = re.compile(r"(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?")
_DATE_RE = re.compile(r"#\s*([0-9A-Za-z][0-9A-Za-z/\-.:, ]*?)\s*#")
_TWO = ("<=", ">=", "<>", ":=", "=<", "=>")
_OPS = set("+-*/\\^&=<>(),.!;:#[]{}")


def _is_ident_char(ch: str) -> bool:
    return ch.isalnum() or ch == "_"


def tokenize(src: str) -> List[Tok]:
    """Tokens of a whole module. Raises :class:`VbaSyntaxError` on an unterminated string or a
    character VBA doesn't know."""
    from .values import parse_date_literal

    text = src.replace("\r\n", "\n").replace("\r", "\n")
    toks: List[Tok] = []
    i = 0
    n = len(text)
    line = 1
    line_start = 0
    space = False

    def at_statement_start() -> bool:
        if not toks:
            return True
        t = toks[-1]
        return t.kind == "nl" or (t.kind == "op" and t.text == ":") or (t.kind == "word" and t.up in ("THEN", "ELSE"))

    def operand_expected() -> bool:
        if not toks:
            return True
        t = toks[-1]
        if t.kind in ("nl",):
            return True
        if t.kind == "op":
            return t.text not in (")", "]", "}")
        if t.kind == "word":
            return t.up in ("TO", "STEP", "IS", "THEN", "ELSE", "AND", "OR", "NOT", "XOR", "EQV", "IMP", "MOD", "LIKE", "CASE", "WHILE", "UNTIL", "IF", "ELSEIF", "RETURN")
        return False

    def continuation_at(j: int) -> bool:
        """``j`` is a newline preceded by `` _``: the line continues."""
        k = j - 1
        if k < 0 or text[k] != "_":
            return False
        return k - 1 < 0 or text[k - 1] in " \t"

    while i < n:
        ch = text[i]
        if ch == "\n":
            if toks and toks[-1].kind != "nl":
                toks.append(Tok("nl", "\n", line, i - line_start))
            line += 1
            i += 1
            line_start = i
            space = False
            continue
        if ch in " \t\x0c\xa0":
            i += 1
            space = True
            continue
        if ch == "_" and (i + 1 >= n or text[i + 1] == "\n") and (i == 0 or text[i - 1] in " \t"):
            # Line continuation.
            i += 2
            line += 1
            line_start = i
            space = True
            continue
        if ch == "'":
            j = _skip_comment(text, i, n)
            extra = text.count("\n", i, j)
            if extra:
                line += extra
                line_start = text.rfind("\n", i, j) + 1
            i = j
            continue
        start = i
        col = i - line_start
        if ch == '"':
            j = i + 1
            buf = []
            while True:
                if j >= n or text[j] == "\n":
                    raise VbaSyntaxError("Expected: end of string", line)
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                buf.append(text[j])
                j += 1
            toks.append(Tok("str", text[i:j + 1], line, col, "".join(buf), space_before=space))
            i = j + 1
            space = False
            continue
        if ch.isdigit() or (ch == "." and i + 1 < n and text[i + 1].isdigit() and operand_expected()):
            m = _NUM_RE.match(text, i)
            lit = m.group(0)
            j = m.end()
            suffix = ""
            if j < n and text[j] in _SUFFIXES and text[j] != "$" and not (j + 1 < n and _is_ident_char(text[j + 1])):
                suffix = text[j]
                j += 1
            is_float = "." in lit or m.group(2) is not None
            if is_float:
                val = float(lit.replace("d", "e").replace("D", "e"))
            else:
                val = int(lit)
            toks.append(Tok("num", text[i:j], line, col, (val, suffix, is_float, 10), suffix, space))
            i = j
            space = False
            continue
        if ch == "&" and i + 1 < n and text[i + 1] in "HhOo" and i + 2 < n and text[i + 2].isalnum() and operand_expected():
            base = 16 if text[i + 1] in "Hh" else 8
            j = i + 2
            digits = "0123456789abcdefABCDEF" if base == 16 else "01234567"
            while j < n and text[j] in digits:
                j += 1
            if j == i + 2:
                raise VbaSyntaxError("Expected: expression", line)
            suffix = ""
            if j < n and text[j] in "%&^":
                suffix = text[j]
                j += 1
            val = int(text[i + 2:j - len(suffix)], base)
            toks.append(Tok("num", text[i:j], line, col, (val, suffix, False, base), suffix, space))
            i = j
            space = False
            continue
        if ch == "#" and operand_expected():
            m = _DATE_RE.match(text, i)
            if m is not None and "\n" not in m.group(0):
                v = parse_date_literal(m.group(1))
                if v is not None:
                    toks.append(Tok("date", m.group(0), line, col, v, space_before=space))
                    i = m.end()
                    space = False
                    continue
        if ch.isalpha():
            j = i + 1
            while j < n and _is_ident_char(text[j]):
                j += 1
            word = text[i:j]
            suffix = ""
            if j < n and text[j] in _SUFFIXES:
                c2 = text[j + 1] if j + 1 < n else ""
                if text[j] == "!" and (c2.isalpha() or c2 == "["):
                    pass  # the bang operator: obj!Name
                elif text[j] == "#" and c2.isdigit():
                    pass
                elif text[j] == "&" and (c2 in "Hh" or c2 == "O"):
                    pass
                elif not _is_ident_char(c2):
                    suffix = text[j]
                    j += 1
            up = word.upper()
            if up == "REM" and not suffix and at_statement_start() and not (toks and toks[-1].kind == "op" and toks[-1].text in (".", "!")):
                k = _skip_comment(text, i, n)
                extra = text.count("\n", i, k)
                if extra:
                    line += extra
                    line_start = text.rfind("\n", i, k) + 1
                i = k
                continue
            t = Tok("word", word, line, col, suffix=suffix, space_before=space)
            toks.append(t)
            i = j
            space = False
            continue
        if ch == "[":
            j = text.find("]", i)
            if j < 0 or "\n" in text[i:j]:
                raise VbaSyntaxError("Expected: ]", line)
            toks.append(Tok("bracket", text[i:j + 1], line, col, text[i + 1:j], space_before=space))
            i = j + 1
            space = False
            continue
        two = text[i:i + 2]
        if two in _TWO:
            norm = {"=<": "<=", "=>": ">="}.get(two, two)
            toks.append(Tok("op", norm, line, col, space_before=space))
            i += 2
            space = False
            continue
        if ch in _OPS:
            toks.append(Tok("op", ch, line, col, space_before=space))
            i += 1
            space = False
            continue
        raise VbaSyntaxError(f"Invalid character: {ch!r}", line)
    if toks and toks[-1].kind != "nl":
        toks.append(Tok("nl", "\n", line, i - line_start))
    toks.append(Tok("eof", "", line, 0))
    return toks


def _skip_comment(text: str, i: int, n: int) -> int:
    """From a comment's start to the newline that ends it (a comment ending in `` _``
    continues on the next line, as in VBA)."""
    while True:
        j = text.find("\n", i)
        if j < 0:
            return n
        k = j - 1
        if k >= 0 and text[k] == "_" and (k == 0 or text[k - 1] in " \t"):
            i = j + 1
            continue
        return j
