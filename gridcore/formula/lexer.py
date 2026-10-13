"""Tokenizer for Excel formulas (A1 style, canonical en-US separators).

Tokens are ``(kind, value, start, end)`` tuples; spans are character offsets into the source,
which the editor uses to colour references.

Kinds: ``num text bool error word sheet func struct op lparen rparen lbrace rbrace comma
semicolon colon trim space eof``.
"""

from __future__ import annotations

from typing import List, Optional, Tuple

from ..core.value import ERROR_LITERALS, CellError

Token = Tuple[str, object, int, int]


class LexError(Exception):
    def __init__(self, msg: str, pos: int):
        super().__init__(f"{msg} at {pos}")
        self.msg = msg
        self.pos = pos


def is_word_char(c: str) -> bool:
    return c.isalnum() or c in "_.$\\?"


def _is_range_end(w: str) -> bool:
    """``A1``, ``$A$1``, ``A``, ``$A``, ``1``, ``$1``: a word that can end a range."""
    s = w[1:] if w.startswith("$") else w
    i = 0
    while i < len(s) and s[i].isascii() and s[i].isalpha():
        i += 1
    letters, rest = s[:i], s[i:]
    digits = rest if not letters else (rest[1:] if rest.startswith("$") else rest)
    if len(letters) > 3:
        return False
    if not digits:
        return bool(letters) and not rest
    return digits.isdigit() and digits.isascii()


def _starts_range_end(c: Optional[str]) -> bool:
    return c is not None and ((c.isascii() and c.isalnum()) or c == "$")


def tokenize(src: str) -> List[Token]:
    out: List[Token] = []
    n = len(src)
    i = 0
    array_depth = 0

    def ch(j: int) -> Optional[str]:
        return src[j] if 0 <= j < n else None

    while i < n:
        start = i
        c = src[i]
        if c in " \t\n\r":
            j = i
            while j < n and src[j] in " \t\n\r":
                j += 1
            out.append(("space", None, start, j))
            i = j
            continue
        if c == '"':
            buf = []
            j = i + 1
            while True:
                d = ch(j)
                if d is None:
                    raise LexError("unterminated text", start)
                if d == '"' and ch(j + 1) == '"':
                    buf.append('"')
                    j += 2
                elif d == '"':
                    j += 1
                    break
                else:
                    buf.append(d)
                    j += 1
            out.append(("text", "".join(buf), start, j))
            i = j
            continue
        if c == "'":
            buf = []
            j = i + 1
            while True:
                d = ch(j)
                if d is None:
                    raise LexError("unterminated sheet name", start)
                if d == "'" and ch(j + 1) == "'":
                    buf.append("'")
                    j += 2
                elif d == "'":
                    j += 1
                    break
                else:
                    buf.append(d)
                    j += 1
            if ch(j) != "!":
                raise LexError("expected ! after sheet name", j)
            s = "".join(buf)
            if ":" in s:
                a, b = s.split(":", 1)
                out.append(("sheet", (a, b), start, j + 1))
            else:
                out.append(("sheet", (s, None), start, j + 1))
            i = j + 1
            continue
        if c == "#" and out and out[-1][3] == start and out[-1][0] == "word":
            # Spill range operator right after a reference: `A1#`.
            out.append(("op", "#", start, i + 1))
            i += 1
            continue
        if c == "#":
            rest = src[i:]
            found = None
            for e in ERROR_LITERALS:
                t = e.value
                if rest[: len(t)].upper() == t:
                    found = e
                    break
            if found is None:
                raise LexError("unknown error literal", start)
            out.append(("error", found, start, i + len(found.value)))
            i += len(found.value)
            continue
        if c == "[":
            depth = 0
            j = i
            while True:
                d = ch(j)
                if d is None:
                    raise LexError("unbalanced [", start)
                if d == "'":
                    j += 2  # escape char inside structured references
                elif d == "[":
                    depth += 1
                    j += 1
                elif d == "]":
                    depth -= 1
                    j += 1
                    if depth == 0:
                        break
                else:
                    j += 1
            out.append(("struct", src[i + 1:j - 1], start, j))
            i = j
            continue
        if c == "." and ch(i + 1) == ":":
            # Trim reference `A1.:E10` / `A1.:.E10` (the word before stopped at the dot).
            if ch(i + 2) == "." and _starts_range_end(ch(i + 3)):
                out.append(("trim", ".:.", start, i + 3))
                i += 3
            else:
                out.append(("trim", ".:", start, i + 2))
                i += 2
            continue
        if c.isdigit() and c.isascii() or (c == "." and (ch(i + 1) or "").isdigit()):
            j = i
            while j < n and src[j].isdigit() and src[j].isascii():
                j += 1
            # Row ranges `1:3`, and the trim forms `1.:3`, `1:.3`, `1.:.3`.
            k = j
            if ch(k) == "." and ch(k + 1) == ":":
                k += 1
            is_row_ref = False
            if ch(k) == ":":
                k += 1
                if ch(k) == ".":
                    k += 1
                if ch(k) == "$":
                    k += 1
                is_row_ref = (ch(k) or "").isdigit()
            if is_row_ref and j > i:
                out.append(("word", src[i:j], start, j))
                i = j
                continue
            if ch(j) == ".":
                j += 1
                while j < n and src[j].isdigit() and src[j].isascii():
                    j += 1
            if ch(j) in ("e", "E"):
                k = j + 1
                if ch(k) in ("+", "-"):
                    k += 1
                if (ch(k) or "").isdigit():
                    while (ch(k) or "").isdigit():
                        k += 1
                    j = k
            text = src[start:j]
            try:
                v = float(text)
            except ValueError:
                raise LexError(f"bad number `{text}`", start) from None
            out.append(("num", v, start, j))
            i = j
            continue
        if c == "(":
            out.append(("lparen", None, start, i + 1))
            i += 1
            continue
        if c == ")":
            out.append(("rparen", None, start, i + 1))
            i += 1
            continue
        if c == "{":
            array_depth += 1
            out.append(("lbrace", None, start, i + 1))
            i += 1
            continue
        if c == "}":
            array_depth = max(array_depth - 1, 0)
            out.append(("rbrace", None, start, i + 1))
            i += 1
            continue
        if c == ",":
            out.append(("comma", None, start, i + 1))
            i += 1
            continue
        if c == ";":
            out.append(("semicolon", None, start, i + 1))
            i += 1
            continue
        if c == ":":
            if ch(i + 1) == "." and _starts_range_end(ch(i + 2)):
                out.append(("trim", ":.", start, i + 2))
                i += 2
            else:
                out.append(("colon", None, start, i + 1))
                i += 1
            continue
        if c in "+-*/^&%@=":
            out.append(("op", c, start, i + 1))
            i += 1
            continue
        if c in "<>":
            two = ch(i + 1)
            if c == "<" and two == "=":
                op, ln = "<=", 2
            elif c == "<" and two == ">":
                op, ln = "<>", 2
            elif c == ">" and two == "=":
                op, ln = ">=", 2
            else:
                op, ln = c, 1
            out.append(("op", op, start, i + ln))
            i += ln
            continue
        if is_word_char(c):
            j = i
            while j < n and is_word_char(src[j]):
                j += 1
            w = src[i:j]
            if ch(j) == "!":
                out.append(("sheet", (w, None), start, j + 1))
                i = j + 1
            elif ch(j) == ":" and _span_sheet_end(src, j) is not None:
                k = _span_sheet_end(src, j)
                out.append(("sheet", (w, src[j + 1:k]), start, k + 1))
                i = k + 1
            elif ch(j) == "(":
                out.append(("func", w, start, j + 1))
                i = j + 1
            elif w.upper() == "TRUE" and ch(j) != "[":
                out.append(("bool", True, start, j))
                i = j
            elif w.upper() == "FALSE" and ch(j) != "[":
                out.append(("bool", False, start, j))
                i = j
            elif len(w) > 1 and w.endswith(".") and ch(j) == ":" and _is_range_end(w[:-1]):
                # `A1.:E10`: the dot belongs to the trim operator, lexed next.
                out.append(("word", w[:-1], start, j - 1))
                i = j - 1
            else:
                out.append(("word", w, start, j))
                i = j
            continue
        raise LexError(f"unexpected character `{c}`", start)
    out.append(("eof", None, n, n))
    return out


def _span_sheet_end(src: str, j: int) -> Optional[int]:
    """For ``Sheet1:Sheet3!A1``: the index of the ``!`` after the second sheet name, else None."""
    k = j + 1
    n = len(src)
    while k < n and is_word_char(src[k]):
        k += 1
    if k > j + 1 and k < n and src[k] == "!":
        return k
    return None
