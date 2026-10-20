"""Pratt parser: tokens → :class:`~gridcore.formula.ast.Expr`."""

from __future__ import annotations

from functools import lru_cache
from typing import List, Optional, Tuple

from ..core.addr import MAX_ROWS, letters_to_col, parse_a1_prefix
from ..core.value import CellError
from .ast import (
    PRECEDENCE,
    ArrayLit,
    Binary,
    Bool,
    Call,
    Err,
    Expr,
    Invoke,
    Missing,
    Name,
    Num,
    Paren,
    Ref,
    Str,
    StructRef,
    Unary,
)
from .lexer import LexError, Token, tokenize

MAX_DEPTH = 256
# Excel's limit on formula text, in characters, the leading `=` included (so files, which leave
# the `=` out, hold at most 8,191).
MAX_FORMULA_LEN = 8192


class ParseError(Exception):
    def __init__(self, msg: str, pos: int = 0):
        super().__init__(msg)
        self.msg = msg
        self.pos = pos


def normalize_function_name(name: str) -> str:
    """``_xlfn.XLOOKUP`` → ``XLOOKUP``. ``_xludf.`` stays: it marks a user-defined function that
    Excel must not take for the built-in of the same name (``_xludf.CONCAT``)."""
    u = name.upper()
    for p in ("_XLFN._XLWS.", "_XLFN.", "_XLWS."):
        if u.startswith(p):
            return u[len(p):]
    return u


@lru_cache(maxsize=65536)
def parse(src: str) -> Expr:
    """Parses formula text. A leading ``=`` is accepted and skipped. Raises :class:`ParseError`."""
    body = src[1:] if src.startswith("=") else src
    if len(body) + 1 > MAX_FORMULA_LEN:
        raise ParseError(f"a formula is at most {MAX_FORMULA_LEN:,} characters long", MAX_FORMULA_LEN)
    try:
        toks = tokenize(body)
    except LexError as e:
        raise ParseError(e.msg, e.pos) from None
    toks = _significant_spaces(toks)
    p = _Parser(toks)
    e = p.expr(0, False)
    if p.peek()[0] != "eof":
        raise p.err(f"unexpected {_describe(p.peek())}")
    return e


def try_parse(src: str) -> Optional[Expr]:
    try:
        return parse(src)
    except ParseError:
        return None


def _describe(t: Token) -> str:
    k = t[0]
    if k == "rparen":
        return "`)`"
    if k == "comma":
        return "`,`"
    if k == "eof":
        return "end of formula"
    if t[1] is not None:
        return f"`{t[1]}`"
    return k


def _significant_spaces(toks: List[Token]) -> List[Token]:
    """Keeps a space token only where it is the intersection operator: between the end of an
    operand and the start of another. An error literal counts as an operand: deleting the cells of
    one side of ``A1:B2 D1:E2`` leaves ``A1:B2 #REF!``, which must read back."""
    out: List[Token] = []
    for i, t in enumerate(toks):
        if t[0] == "space":
            prev_end = bool(out) and out[-1][0] in ("word", "rparen", "struct", "error")
            nxt = toks[i + 1] if i + 1 < len(toks) else None
            next_start = nxt is not None and nxt[0] in ("word", "sheet", "lparen", "func", "error")
            if prev_end and next_start:
                out.append(t)
            continue
        out.append(t)
    return out


def _word_is_cols(w: str) -> Optional[Tuple[int, bool]]:
    a = w.startswith("$")
    c = letters_to_col(w[1:] if a else w)  # one $ at most: $$A is no column (Excel)
    return None if c is None else (c, a)


def _word_is_row(w: str) -> Optional[Tuple[int, bool]]:
    a = w.startswith("$")
    s = w[1:] if a else w
    if not s.isdigit() or not s.isascii():
        return None
    n = int(s)
    if n == 0 or n > MAX_ROWS:
        return None
    return n - 1, a


def _word_is_cell(w: str) -> Optional[Tuple[int, int, bool, bool]]:
    p = parse_a1_prefix(w)
    if p is None or p[4]:
        return None
    col, row, col_abs, row_abs, _ = p
    return row, col, row_abs, col_abs


_BINOPS = {"+", "-", "*", "/", "^", "&", "=", "<>", "<", ">", "<=", ">="}


class _Parser:
    def __init__(self, toks: List[Token]):
        self.toks = toks
        self.i = 0
        self.depth = 0

    def peek(self, k: int = 0) -> Token:
        j = self.i + k
        return self.toks[j] if j < len(self.toks) else ("eof", None, 0, 0)

    def pos(self) -> int:
        return self.peek()[2]

    def next(self) -> Token:
        t = self.peek()
        if self.i < len(self.toks):
            self.i += 1
        return t

    def err(self, msg: str) -> ParseError:
        return ParseError(msg, self.pos())

    def expect(self, kind: str) -> None:
        if self.peek()[0] == kind:
            self.next()
            return
        raise self.err(f"expected {kind}, found {_describe(self.peek())}")

    def binop(self, allow_union: bool) -> Optional[str]:
        t = self.peek()
        k = t[0]
        if k == "op" and t[1] in _BINOPS:
            return t[1]  # type: ignore[return-value]
        if k == "colon":
            return ":"
        if k == "space":
            return " "
        if k == "comma" and allow_union:
            return ","
        return None

    def expr(self, min_bp: int, allow_union: bool) -> Expr:
        self.depth += 1
        if self.depth > MAX_DEPTH:
            raise self.err("formula is nested too deeply")
        lhs = self.prefix(allow_union)
        nested = 1
        while True:
            t = self.peek()
            if t[0] == "op" and t[1] == "%":
                if 70 < min_bp:
                    break
                self.next()
                lhs = Unary("%", lhs)
                # Each postfix `%` nests one more level, and so does each call of a result.
                nested += 1
                self.depth += 1
                if self.depth > MAX_DEPTH:
                    raise self.err("formula is nested too deeply")
                continue
            if t[0] == "lparen" and isinstance(lhs, (Call, Invoke, Paren, Name)):
                # Invoking a LAMBDA result.
                self.next()
                lhs = Invoke(lhs, tuple(self.args()))
                nested += 1
                self.depth += 1
                if self.depth > MAX_DEPTH:
                    raise self.err("formula is nested too deeply")
                continue
            if t[0] == "trim":
                bp = PRECEDENCE[":"]
                if bp < min_bp:
                    break
                self.next()
                rhs = self.expr(bp + 1, allow_union)
                lhs = Unary(t[1], _combine(":", lhs, rhs))  # type: ignore[arg-type]
                continue
            op = self.binop(allow_union)
            if op is None:
                break
            bp = PRECEDENCE[op]
            if bp < min_bp:
                break
            self.next()
            rhs = self.expr(bp + 1, allow_union)
            lhs = _combine(op, lhs, rhs)
        self.depth -= nested
        return lhs

    def args(self) -> List[Expr]:
        args: List[Expr] = []
        if self.peek()[0] == "rparen":
            self.next()
            return args
        while True:
            k = self.peek()[0]
            a = Missing() if k in ("comma", "rparen") else self.expr(0, False)
            args.append(a)
            t = self.next()
            if t[0] == "comma":
                continue
            if t[0] == "rparen":
                break
            raise self.err(f"expected `,` or `)`, found {_describe(t)}")
        if len(args) > 255:
            raise self.err("too many arguments")
        return args

    def prefix(self, allow_union: bool) -> Expr:
        t = self.next()
        k, v = t[0], t[1]
        if k == "num":
            return Num(v)  # type: ignore[arg-type]
        if k == "text":
            return Str(v)  # type: ignore[arg-type]
        if k == "bool":
            return Bool(v)  # type: ignore[arg-type]
        if k == "error":
            return Err(v)  # type: ignore[arg-type]
        if k == "op" and v == "-":
            return Unary("-", self.expr(60, allow_union))
        if k == "op" and v == "+":
            return Unary("+", self.expr(60, allow_union))
        if k == "op" and v == "@":
            return Unary("@", self.expr(95, allow_union))
        if k == "lparen":
            e = self.expr(0, True)
            self.expect("rparen")
            return Paren(e)
        if k == "lbrace":
            return self.array()
        if k == "func":
            name = normalize_function_name(v)  # type: ignore[arg-type]
            args = self.args()
            # Files write `A1#` as `_xlfn.ANCHORARRAY(A1)`.
            if name == "ANCHORARRAY" and len(args) == 1 and isinstance(args[0], Ref) and args[0].kind == "cell":
                return Unary("#", args[0])
            # Files write `@x` as `_xlfn.SINGLE(x)`.
            if name == "SINGLE" and len(args) == 1:
                return Unary("@", args[0])
            trim = {"_TRO_LEADING": ".:", "_TRO_TRAILING": ":.", "_TRO_ALL": ".:."}.get(name)
            if trim is not None and len(args) == 1:
                return Unary(trim, args[0])
            return Call(name, tuple(args))
        if k == "sheet":
            a, b = v  # type: ignore[misc]
            sheet = (a, b) if b is not None else a
            t2 = self.next()
            if t2[0] == "word":
                return self.reference_from_word(sheet, t2[1])  # type: ignore[arg-type]
            if t2[0] == "error" and t2[1] is CellError.REF:
                return Err(CellError.REF)
            raise self.err(f"expected a reference after the sheet name, found {_describe(t2)}")
        if k == "word":
            w: str = v  # type: ignore[assignment]
            if self.peek()[0] == "struct" and _word_is_cell(w) is None:
                body = self.next()[1]
                try:
                    return parse_struct(w, body)  # type: ignore[arg-type]
                except ValueError as e:
                    raise self.err(str(e)) from None
            return self.reference_from_word(None, w)
        if k == "struct":
            try:
                return parse_struct("", v)  # type: ignore[arg-type]
            except ValueError as e:
                raise self.err(str(e)) from None
        raise ParseError(f"unexpected {_describe(t)}", t[2])

    def reference_from_word(self, sheet, w: str) -> Expr:
        """A word after an optional sheet: a cell, a column/row range (``A:C``, ``3:5``) or a
        name. A trim operator in place of the colon wraps the range in a trim ``Unary``."""
        nt = self.peek()
        trim: Optional[str] = None
        has_sep = False
        if nt[0] == "colon":
            has_sep = True
        elif nt[0] == "trim":
            has_sep = True
            trim = nt[1]  # type: ignore[assignment]

        def wrap(r: Ref) -> Expr:
            return Unary(trim, r) if trim is not None else r

        cell = _word_is_cell(w)
        if cell is not None:
            r0, c0, ra, ca = cell
            if has_sep:
                t2 = self.peek(1)
                if t2[0] == "word":
                    other = _word_is_cell(t2[1])  # type: ignore[arg-type]
                    if other is not None:
                        self.next()
                        self.next()
                        r1, c1, rb, cb = other
                        return wrap(Ref(sheet, "range", r0, c0, r1, c1, ra, ca, rb, cb))
            ref = Ref.cell(r0, c0, ra, ca, sheet)
            if self.peek()[0] == "op" and self.peek()[1] == "#":
                self.next()
                return Unary("#", ref)
            return ref
        t2 = self.peek(1)
        second: Optional[str] = None
        if t2[0] == "word":
            second = t2[1]  # type: ignore[assignment]
        elif t2[0] == "num" and float(t2[1]).is_integer() and t2[1] >= 1:  # type: ignore[arg-type]
            second = str(int(t2[1]))  # type: ignore[arg-type]
        if has_sep and second is not None:
            a, b = _word_is_cols(w), _word_is_cols(second)
            if a is not None and b is not None:
                self.next()
                self.next()
                return wrap(Ref(sheet, "cols", 0, a[0], 0, b[0], False, a[1], False, b[1]))
            a, b = _word_is_row(w), _word_is_row(second)
            if a is not None and b is not None:
                self.next()
                self.next()
                return wrap(Ref(sheet, "rows", a[0], 0, b[0], 0, a[1], False, b[1], False))
        if sheet is not None:
            if isinstance(sheet, str):
                # Sheet-scoped name `Sheet1!MyName`.
                return Name(f"{sheet}!{w}")
            raise self.err("invalid 3-D reference")
        if w[:1].isdigit():
            raise self.err(f"unexpected `{w}`")
        return Name(w)

    def array(self) -> Expr:
        rows: List[List[Expr]] = [[]]
        while True:
            t = self.next()
            k = t[0]
            if k == "num":
                el: Expr = Num(t[1])  # type: ignore[arg-type]
            elif k == "op" and t[1] in ("-", "+"):
                t2 = self.next()
                if t2[0] != "num":
                    raise self.err(f"expected a number in array, found {_describe(t2)}")
                el = Num(-t2[1] if t[1] == "-" else t2[1])  # type: ignore[operator]
            elif k == "text":
                el = Str(t[1])  # type: ignore[arg-type]
            elif k == "bool":
                el = Bool(t[1])  # type: ignore[arg-type]
            elif k == "error":
                el = Err(t[1])  # type: ignore[arg-type]
            else:
                raise self.err(f"array constants may only contain constants, found {_describe(t)}")
            rows[-1].append(el)
            t = self.next()
            if t[0] == "comma":
                continue
            if t[0] == "semicolon":
                rows.append([])
                continue
            if t[0] == "rbrace":
                break
            raise self.err(f"expected `,`, `;` or `}}`, found {_describe(t)}")
        w = len(rows[0])
        if any(len(r) != w for r in rows):
            raise self.err("array rows must have the same length")
        return ArrayLit(tuple(tuple(r) for r in rows))


def _combine(op: str, lhs: Expr, rhs: Expr) -> Expr:
    """Folds ``A1:B2``-like range operations of two plain refs into a single reference."""
    if (
        op == ":"
        and isinstance(lhs, Ref)
        and isinstance(rhs, Ref)
        and rhs.sheet is None
        and lhs.kind == "cell"
        and rhs.kind == "cell"
    ):
        return Ref(lhs.sheet, "range", lhs.r0, lhs.c0, rhs.r0, rhs.c0, lhs.r0_abs, lhs.c0_abs, rhs.r0_abs, rhs.c0_abs)
    return Binary(op, lhs, rhs)


_ITEMS = {"#all": "#All", "#data": "#Data", "#headers": "#Headers", "#totals": "#Totals", "#this row": "#This Row"}


def parse_struct(table: str, body: str) -> StructRef:
    """Parses the inside of ``Table[...]``. Raises ``ValueError`` on malformed text."""
    specifiers: List[str] = []
    cols: List[str] = []
    raw = body
    body = body.strip()

    def unescape(s: str) -> str:
        out = []
        it = iter(s)
        for c in it:
            if c == "'":
                nxt = next(it, None)
                if nxt is not None:
                    out.append(nxt)
            else:
                out.append(c)
        return "".join(out)

    def item(s: str) -> Optional[str]:
        return _ITEMS.get(s.strip().lower())

    def unexpected(gap: str) -> ValueError:
        return ValueError(f"unexpected `{gap}` in structured reference")

    def column(s: str) -> str:
        esc = False
        for c in s:
            if esc:
                esc = False
            elif c == "'":
                esc = True
            elif c in "[]":
                raise unexpected(s)
        return unescape(s)

    if not body.startswith("["):
        if body.startswith("@"):
            specifiers.append("#This Row")
            rest = body[1:].strip()
            if rest.startswith("[") and rest.endswith("]"):
                rest = rest[1:-1]
            if rest:
                cols.append(column(rest))
        elif item(body) is not None:
            specifiers.append(item(body))  # type: ignore[arg-type]
        elif body:
            # The column keeps the spaces at its edges: `Table1[Qty ]` names `Qty `.
            cols.append(column(raw))
    else:
        parts: List[str] = []
        depth = 0
        cur: List[str] = []
        gap: List[str] = []
        esc = False
        for c in body:
            if esc:
                cur.append(c)
                esc = False
                continue
            if c == "'" and depth > 0:
                cur.append(c)
                esc = True
            elif c == "[":
                if depth == 0:
                    g = "".join(gap).strip()
                    if g not in ("", ",", ":"):
                        raise unexpected(g)
                    gap = []
                depth += 1
                if depth > 1:
                    cur.append(c)
            elif c == "]":
                if depth == 0:
                    raise unexpected("]")
                depth -= 1
                if depth > 0:
                    cur.append(c)
                else:
                    parts.append("".join(cur))
                    cur = []
            elif depth > 0:
                cur.append(c)
            else:
                gap.append(c)
        if depth > 0:
            raise ValueError("unclosed `[` in structured reference")
        for p in parts:
            it = item(p)
            if it is not None:
                specifiers.append(it)
            else:
                cols.append(column(p))
        tail = "".join(gap).strip()
        if tail:
            bare = tail[1:].strip() if tail.startswith(",") else ""
            simple = bool(bare) and not any(ch in "[]#',:" for ch in bare)
            if not (simple and specifiers == ["#This Row"] and not cols):
                raise unexpected(tail)
            cols.append(bare)
    if len(cols) > 2:
        raise ValueError("a structured reference names at most two columns")
    return StructRef(table, tuple(specifiers), cols[0] if cols else None, cols[1] if len(cols) > 1 else None)
