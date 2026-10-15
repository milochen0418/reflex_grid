"""Parsing VBA modules into :mod:`.ast` trees ([MS-VBAL] section 5, as the VBA editor accepts
it). Compile errors carry the module's physical line number."""

from __future__ import annotations

from typing import Any, Callable, Dict, List, Mapping, Optional, Set, Tuple

from . import ast as A
from .errors import VbaSyntaxError
from .lexer import Tok, tokenize
from .values import Byte, Currency, Date, Integer, LongLong, Single

# Words that can't name a variable or start an expression.
RESERVED = {
    "AND", "AS", "BYREF", "BYVAL", "CALL", "CASE", "CONST", "DECLARE", "DIM", "DO", "EACH", "ELSE", "ELSEIF",
    "END", "ENDIF", "ENUM", "EQV", "ERASE", "EVENT", "EXIT", "FALSE", "FOR", "FRIEND", "FUNCTION", "GET", "GLOBAL",
    "GOSUB", "GOTO", "IF", "IMP", "IMPLEMENTS", "IN", "IS", "LET", "LIKE", "LOOP", "LSET", "ME", "MOD", "NEW",
    "NEXT", "NOT", "NOTHING", "NULL", "ON", "OPTION", "OPTIONAL", "OR", "PARAMARRAY", "PRESERVE", "PRIVATE",
    "PUBLIC", "RAISEEVENT", "REDIM", "RESUME", "RETURN", "RSET", "SELECT", "SET", "STATIC", "STEP", "STOP", "SUB",
    "THEN", "TO", "TRUE", "TYPE", "TYPEOF", "UNTIL", "WEND", "WHILE", "WITH", "WITHEVENTS", "XOR", "EMPTY",
}
# Statements that are also common names (`.Name`, `Width`…): a statement only when not followed
# by `=`, `.` or `(`.
SOFT_STATEMENTS = {
    "NAME", "OPEN", "CLOSE", "PRINT", "WRITE", "INPUT", "LINE", "GET", "PUT", "SEEK", "LOCK", "UNLOCK", "WIDTH",
    "KILL", "MKDIR", "RMDIR", "CHDIR", "CHDRIVE", "FILECOPY", "SAVESETTING", "DELETESETTING", "SENDKEYS",
    "APPACTIVATE", "LOAD", "UNLOAD", "RESET", "DATE", "TIME", "SETATTR", "BEEP", "RANDOMIZE", "ERROR",
}
_COMPARE = {"=", "<>", "<", ">", "<=", ">="}
_PREC = {
    "IMP": 1, "EQV": 2, "XOR": 3, "OR": 4, "AND": 5,
    "=": 7, "<>": 7, "<": 7, ">": 7, "<=": 7, ">=": 7, "LIKE": 7, "IS": 7,
    "&": 8, "+": 9, "-": 9, "MOD": 10, "\\": 11, "*": 12, "/": 12, "^": 14,
}
_SUFFIX_TYPE = {"%": "Integer", "&": "Long", "^": "LongLong", "!": "Single", "#": "Double", "@": "Currency", "$": "String"}
_BLOCK_ENDS = {"IF", "SUB", "FUNCTION", "PROPERTY", "SELECT", "WITH", "TYPE", "ENUM"}


def parse_module(src: str, name: str = "Module", constants: Optional[Mapping[str, Any]] = None) -> A.ModuleAst:
    """A module's AST, after its conditional compilation directives (``#If``, ``#Const``…) are
    applied with the project's ``constants`` (see :mod:`.conditional`)."""
    from .conditional import preprocess

    try:
        src = preprocess(src, constants, name)
        toks = tokenize(src)
    except VbaSyntaxError as e:
        e.module = name
        raise
    p = Parser(toks, name)
    m = p.module()
    m.attributes = list(p.attributes)
    return m


def parse_expression(src: str) -> A.Expr:
    """One expression (the Immediate window's ``? expr``)."""
    toks = tokenize(src)
    p = Parser(toks, "Immediate")
    e = p.expr()
    if not p.at_end():
        p.error("Expected: end of statement")
    return e


def parse_statements(src: str) -> List[A.Stmt]:
    """Statements outside any procedure (the Immediate window)."""
    toks = tokenize(src)
    p = Parser(toks, "Immediate")
    out: List[A.Stmt] = []
    while p.peek().kind != "eof":
        if p.peek().kind == "nl" or p.at_op(":"):
            p.i += 1
            continue
        out.extend(p.statement())
        p.end_stmt()
    return out


class Parser:
    def __init__(self, toks: List[Tok], module: str):
        self.t = toks
        self.i = 0
        self.mod = module
        # `Next j, i` closes two loops: the extra closings still to be taken.
        self.pending_next = 0
        # `Attribute X.VB_… = …` lines met in the code (default members, descriptions).
        self.attributes: List[str] = []

    # ------------------------------------------------------------ tokens

    def peek(self, k: int = 0) -> Tok:
        j = self.i + k
        return self.t[j] if j < len(self.t) else self.t[-1]

    def next(self) -> Tok:
        t = self.t[self.i]
        if self.i < len(self.t) - 1:
            self.i += 1
        return t

    def error(self, msg: str, tok: Optional[Tok] = None):
        tok = tok or self.peek()
        raise VbaSyntaxError(msg, tok.line, self.mod)

    def at_kw(self, *words: str) -> bool:
        t = self.peek()
        return t.kind == "word" and t.up in words and not t.suffix

    def at_kw_at(self, k: int, *words: str) -> bool:
        t = self.peek(k)
        return t.kind == "word" and t.up in words and not t.suffix

    def accept_kw(self, *words: str) -> Optional[Tok]:
        if self.at_kw(*words):
            return self.next()
        return None

    def expect_kw(self, word: str, msg: Optional[str] = None) -> Tok:
        if not self.at_kw(word):
            self.error(msg or f"Expected: {word.title()}")
        return self.next()

    def at_op(self, op: str) -> bool:
        t = self.peek()
        return t.kind == "op" and t.text == op

    def accept_op(self, op: str) -> bool:
        if self.at_op(op):
            self.next()
            return True
        return False

    def expect_op(self, op: str, msg: Optional[str] = None) -> None:
        if not self.accept_op(op):
            self.error(msg or f"Expected: {op}")

    def at_end(self) -> bool:
        t = self.peek()
        return t.kind in ("nl", "eof") or (t.kind == "op" and t.text == ":")

    def end_stmt(self) -> None:
        t = self.peek()
        if t.kind == "eof":
            return
        if t.kind == "nl" or (t.kind == "op" and t.text == ":"):
            self.next()
            return
        self.error("Expected: end of statement")

    def skip_newlines(self) -> None:
        while self.peek().kind == "nl" or self.at_op(":"):
            self.next()

    def at_line_start(self) -> bool:
        return self.i == 0 or self.t[self.i - 1].kind == "nl"

    def ident(self, what: str = "identifier") -> Tok:
        t = self.peek()
        if t.kind == "bracket":
            self.next()
            t2 = Tok("word", t.value, t.line, t.col)
            return t2
        if t.kind != "word" or t.up in RESERVED:
            self.error(f"Expected: {what}")
        return self.next()

    def rest_of_statement(self) -> None:
        depth = 0
        while not (self.at_end() and depth == 0):
            t = self.next()
            if t.kind == "eof":
                return
            if t.kind == "op" and t.text == "(":
                depth += 1
            elif t.kind == "op" and t.text == ")":
                depth -= 1

    # ------------------------------------------------------------ module

    def module(self) -> A.ModuleAst:
        m = A.ModuleAst(self.mod)
        while True:
            self.skip_newlines()
            t = self.peek()
            if t.kind == "eof":
                break
            if t.kind != "word":
                self.error("Invalid outside procedure")
            up = t.up
            if up == "OPTION":
                self.next()
                self.option(m)
            elif up == "ATTRIBUTE":
                self.st_attribute(t.line)
            elif up.startswith("DEF") and up in ("DEFBOOL", "DEFBYTE", "DEFINT", "DEFLNG", "DEFLNGLNG", "DEFLNGPTR", "DEFCUR", "DEFSNG", "DEFDBL", "DEFDEC", "DEFDATE", "DEFSTR", "DEFOBJ", "DEFVAR"):
                self.next()
                self.deftype(m, up)
            else:
                self.declaration(m)
            self.end_stmt()
        return m

    def option(self, m: A.ModuleAst) -> None:
        if self.accept_kw("EXPLICIT"):
            m.option_explicit = True
        elif self.accept_kw("BASE"):
            t = self.next()
            if t.kind != "num" or t.value[0] not in (0, 1):
                self.error("Expected: 0 or 1")
            m.option_base = t.value[0]
        elif self.accept_kw("COMPARE"):
            w = self.next()
            if w.kind != "word" or w.up not in ("BINARY", "TEXT", "DATABASE"):
                self.error("Expected: Binary or Text")
            m.option_compare_text = w.up == "TEXT"
        elif self.accept_kw("PRIVATE"):
            self.expect_kw("MODULE")
            m.option_private = True
        else:
            self.error("Expected: Explicit, Base, Compare or Private Module")

    def deftype(self, m: A.ModuleAst, kind: str) -> None:
        tname = {
            "DEFBOOL": "Boolean", "DEFBYTE": "Byte", "DEFINT": "Integer", "DEFLNG": "Long", "DEFLNGLNG": "LongLong",
            "DEFLNGPTR": "LongPtr", "DEFCUR": "Currency", "DEFSNG": "Single", "DEFDBL": "Double", "DEFDEC": "Decimal",
            "DEFDATE": "Date", "DEFSTR": "String", "DEFOBJ": "Object", "DEFVAR": "Variant",
        }[kind]
        while True:
            a = self.next()
            if a.kind != "word" or len(a.text) != 1:
                self.error("Expected: letter range")
            b = a
            if self.accept_op("-"):
                b = self.next()
            m.deftypes.append((a.text.upper(), b.text.upper(), tname))
            if not self.accept_op(","):
                break

    def declaration(self, m: A.ModuleAst) -> None:
        scope = "default"
        t = self.peek()
        if self.at_kw("PUBLIC", "PRIVATE", "GLOBAL", "FRIEND"):
            scope = {"GLOBAL": "public"}.get(t.up, t.up.lower())
            self.next()
        static = bool(self.accept_kw("STATIC"))
        if self.at_kw("SUB", "FUNCTION", "PROPERTY"):
            m.procs.append(self.procedure("public" if scope in ("default", "public") else scope, static))
            return
        if static:
            self.error("Invalid outside procedure")
        if self.accept_kw("CONST"):
            for c in self.const_list():
                m.constants.append(A.ModuleConst(c, "public" if scope == "public" else "private"))
            return
        if self.at_kw("TYPE"):
            line = self.next().line
            m.types.append(self.type_decl("private" if scope == "private" else "public", line))
            return
        if self.at_kw("ENUM"):
            line = self.next().line
            m.enums.append(self.enum_decl("private" if scope == "private" else "public", line))
            return
        if self.at_kw("DECLARE"):
            m.declares.append(self.declare("private" if scope == "private" else "public"))
            return
        if self.at_kw("EVENT"):
            line = self.next().line
            name = self.ident().text
            params = self.params() if self.at_op("(") else []
            m.events.append(A.EventDecl(name, params, line))
            return
        if self.at_kw("IMPLEMENTS"):
            self.next()
            m.implements.append(self.type_name())
            return
        if scope == "default":
            if not self.accept_kw("DIM"):
                self.error("Invalid outside procedure")
            scope = "private"
        else:
            self.accept_kw("DIM")
        withevents = bool(self.accept_kw("WITHEVENTS"))
        for v in self.var_list(withevents):
            m.variables.append(A.ModuleVar(v, "public" if scope == "public" else "private"))

    def type_name(self) -> str:
        t = self.peek()
        if t.kind == "bracket":
            self.next()
            return t.value
        if t.kind != "word":
            self.error("Expected: type name")
        parts = [self.next().text]
        while self.at_op(".") and self.peek(1).kind == "word":
            self.next()
            parts.append(self.next().text)
        return ".".join(parts)

    def as_type(self) -> A.TypeRef:
        if not self.accept_kw("AS"):
            return A.TypeRef()
        new = bool(self.accept_kw("NEW"))
        name = self.type_name()
        tr = A.TypeRef(name, new=new, explicit=True)
        if name.upper() == "STRING" and self.accept_op("*"):
            t = self.next()
            if t.kind == "num":
                tr.string_len = int(t.value[0])
            elif t.kind == "word":
                tr.string_len = -1  # a constant, resolved by the compiler
                tr.name = "String*" + t.text
            else:
                self.error("Expected: integer constant")
        return tr

    def bounds(self) -> List[Tuple[Optional[A.Expr], A.Expr]]:
        out: List[Tuple[Optional[A.Expr], A.Expr]] = []
        if self.at_op(")"):
            return out
        while True:
            a = self.expr()
            if self.accept_kw("TO"):
                b = self.expr()
                out.append((a, b))
            else:
                out.append((None, a))
            if not self.accept_op(","):
                break
        return out

    def var_list(self, withevents: bool = False) -> List[A.VarDecl]:
        out = []
        while True:
            t = self.ident("variable name")
            name = t.text
            suffix = t.suffix
            bounds = None
            is_array = False
            if self.accept_op("("):
                is_array = True
                b = self.bounds()
                self.expect_op(")")
                bounds = b if b else None
            tr = self.as_type()
            if suffix and not tr.explicit:
                tr = A.TypeRef(_SUFFIX_TYPE[suffix], explicit=True)
            out.append(A.VarDecl(name, tr, bounds, is_array, withevents, t.line, suffix))
            if not self.accept_op(","):
                break
            withevents = bool(self.accept_kw("WITHEVENTS")) or False
        return out

    def const_list(self) -> List[A.ConstDecl]:
        out = []
        while True:
            t = self.ident("constant name")
            tr = self.as_type()
            if t.suffix and not tr.explicit:
                tr = A.TypeRef(_SUFFIX_TYPE[t.suffix], explicit=True)
            self.expect_op("=")
            out.append(A.ConstDecl(t.text, tr, self.expr(), t.line))
            if not self.accept_op(","):
                break
        return out

    def type_decl(self, scope: str, line: int) -> A.TypeDecl:
        name = self.ident("type name").text
        self.end_stmt()
        fields: List[A.VarDecl] = []
        while True:
            self.skip_newlines()
            if self.at_kw("END") and self.at_kw_at(1, "TYPE"):
                self.next()
                self.next()
                break
            if self.peek().kind == "eof":
                self.error("Expected: End Type")
            fields.extend(self.var_list())
            self.end_stmt()
        return A.TypeDecl(name, fields, scope, line)

    def enum_decl(self, scope: str, line: int) -> A.EnumDecl:
        name = self.ident("enum name").text
        self.end_stmt()
        members: List[Tuple[str, Optional[A.Expr]]] = []
        while True:
            self.skip_newlines()
            if self.at_kw("END") and self.at_kw_at(1, "ENUM"):
                self.next()
                self.next()
                break
            if self.peek().kind == "eof":
                self.error("Expected: End Enum")
            t = self.ident("enum member")
            e = self.expr() if self.accept_op("=") else None
            members.append((t.text, e))
            self.end_stmt()
        return A.EnumDecl(name, members, scope, line)

    def declare(self, scope: str) -> A.DeclareDecl:
        line = self.next().line
        self.accept_kw("PTRSAFE")
        kind = self.next()
        if kind.kind != "word" or kind.up not in ("SUB", "FUNCTION"):
            self.error("Expected: Sub or Function")
        name = self.ident().text
        self.accept_kw("CDECL")
        if not self.accept_kw("LIB"):
            self.error("Expected: Lib")
        lib = self.next()
        alias = ""
        if self.accept_kw("ALIAS"):
            alias = self.next().value or ""
        params = self.params() if self.at_op("(") else []
        ret = self.as_type()
        return A.DeclareDecl(name, kind.up.lower(), lib.value or "", alias, params, ret, scope, line)

    def params(self) -> List[A.Param]:
        self.expect_op("(")
        out: List[A.Param] = []
        if self.accept_op(")"):
            return out
        while True:
            line = self.peek().line
            optional = bool(self.accept_kw("OPTIONAL"))
            byval = False
            if self.accept_kw("BYVAL"):
                byval = True
            elif self.accept_kw("BYREF"):
                byval = False
            paramarray = bool(self.accept_kw("PARAMARRAY"))
            t = self.ident("parameter name")
            is_array = False
            if self.accept_op("("):
                self.expect_op(")")
                is_array = True
            tr = self.as_type()
            if t.suffix and not tr.explicit:
                tr = A.TypeRef(_SUFFIX_TYPE[t.suffix], explicit=True)
            default = None
            if self.accept_op("="):
                default = self.expr()
            out.append(A.Param(t.text, tr, byval, optional or default is not None, default, paramarray, is_array, line))
            if not self.accept_op(","):
                break
        self.expect_op(")")
        return out

    def procedure(self, scope: str, static: bool) -> A.Proc:
        t = self.next()
        line = t.line
        kind = t.up.lower()
        if kind == "property":
            k = self.next()
            if k.kind != "word" or k.up not in ("GET", "LET", "SET"):
                self.error("Expected: Get, Let or Set")
            kind = k.up.lower()
        nt = self.peek()
        if nt.kind == "bracket":
            self.next()
            name, suffix = nt.value, ""
        else:
            if nt.kind != "word":
                self.error("Expected: identifier")
            self.next()
            name, suffix = nt.text, nt.suffix
        params = self.params() if self.at_op("(") else []
        ret = A.TypeRef()
        ret_is_array = False
        if self.at_kw("AS"):
            ret = self.as_type()
            if self.accept_op("("):
                self.expect_op(")")
                ret_is_array = True
        elif suffix:
            ret = A.TypeRef(_SUFFIX_TYPE[suffix], explicit=True)
        self.accept_kw("STATIC")
        self.end_stmt()
        end_word = {"sub": "SUB", "function": "FUNCTION"}.get(kind, "PROPERTY")
        body = self.block({("END", end_word)})
        end_line = self.peek().line
        self.next()
        self.next()
        return A.Proc(kind, name, params, ret, body, scope, static, line, end_line, ret_is_array)

    # ------------------------------------------------------------ blocks

    def block_end(self, enders: Set[Tuple[str, ...]]) -> Optional[Tuple[str, ...]]:
        t = self.peek()
        if self.pending_next and ("NEXT",) in enders:
            return ("NEXT",)
        if t.kind != "word" or t.suffix:
            return None
        up = t.up
        for e in enders:
            if e[0] == up:
                if len(e) == 1 or self.at_kw_at(1, e[1]):
                    return e
            if e == ("END", "IF") and up == "ENDIF":
                return e
        return None

    def block(self, enders: Set[Tuple[str, ...]]) -> List[A.Stmt]:
        body: List[A.Stmt] = []
        while True:
            self.skip_newlines()
            if self.peek().kind == "eof":
                want = sorted(enders)[0]
                self.error(f"Expected: {' '.join(w.title() for w in want)}")
            if self.block_end(enders) is not None:
                return body
            body.extend(self.statement())
            if self.block_end(enders) is not None:
                continue
            self.end_stmt()

    # ------------------------------------------------------------ statements

    def statement(self) -> List[A.Stmt]:
        t = self.peek()
        line = t.line
        # Labels and line numbers start a line.
        if self.at_line_start():
            if t.kind == "num" and not t.value[2]:
                self.next()
                self.accept_op(":")
                out: List[A.Stmt] = [A.Label(str(t.value[0]), int(t.value[0]), line)]
                if not self.at_end():
                    out.extend(self.statement())
                return out
            if t.kind == "word" and not t.suffix and self.peek(1).kind == "op" and self.peek(1).text == ":" and t.up not in RESERVED and t.up not in ("REM",):
                self.next()
                self.next()
                out = [A.Label(t.text, None, line)]
                if not self.at_end():
                    out.extend(self.statement())
                return out
        if t.kind == "word" and not t.suffix:
            up = t.up
            nxt = self.peek(1)
            soft_name = up in SOFT_STATEMENTS and nxt.kind == "op" and (nxt.text in ("=", ".") or (nxt.text == "(" and not nxt.space_before))
            if not soft_name:
                h = _STATEMENTS.get(up)
                if h is not None:
                    return h(self, line)
        if t.kind == "word" and t.up == "DEBUG" and self.peek(1).kind == "op" and self.peek(1).text == "." and self.peek(2).kind == "word":
            w = self.peek(2).up
            if w == "PRINT":
                self.i += 3
                return [self.print_items(line)]
        return [self.expr_statement(line)]

    def expr_statement(self, line: int) -> A.Stmt:
        start = self.i
        target, split = self.lexpr()
        if self.at_op("="):
            self.next()
            value = self.expr()
            return self._assign(target, value, "let", line)
        if self.at_end() or self.at_kw("ELSE"):
            if isinstance(target, A.Call):
                args = target.args
                if len(args) == 1 and args[0].expr is not None and args[0].name is None:
                    # `Foo (x)`: one parenthesized argument, passed by value.
                    args = [A.Arg(A.Paren(args[0].expr, line))]
                return A.CallStmt(target.callee, args, False, line)
            return A.CallStmt(target, [], False, line)
        # `Foo a, b` or `Foo (a) + 1, b`: arguments without parentheses.
        if split is not None:
            self.i = split[1]
            target = split[0]
        args = self.arg_list_bare()
        return A.CallStmt(target, args, False, line)

    def _assign(self, target: A.Expr, value: A.Expr, kind: str, line: int) -> A.Stmt:
        if kind == "let" and isinstance(target, A.Call) and isinstance(target.callee, A.Ident) and target.callee.name.upper() in ("MID", "MIDB") and 2 <= len(target.args) <= 3:
            a = target.args
            return A.MidAssign(a[0].expr, a[1].expr, a[2].expr if len(a) > 2 else None, value, line)
        return A.Assign(target, value, kind, line)

    def lexpr(self) -> Tuple[A.Expr, Optional[Tuple[A.Expr, int]]]:
        """A statement's target: names, members and arguments. Also returns where the target
        could end instead — before its last spaced ``(`` — for ``Foo (a), b``."""
        t = self.peek()
        split: Optional[Tuple[A.Expr, int]] = None
        if t.kind == "op" and t.text in (".", "!"):
            self.next()
            name = self.member_name()
            e: A.Expr = A.Member(None, name, t.line) if t.text == "." else A.Bang(None, name, t.line)
        elif t.kind == "word" and t.up == "ME" and not t.suffix:
            self.next()
            e = A.MeExpr(t.line)
        elif t.kind == "bracket":
            self.next()
            e = A.BracketEval(t.value, t.line)
        elif t.kind == "word" and (t.up not in RESERVED or t.suffix):
            self.next()
            e = A.Ident(t.text, t.suffix, t.line)
        else:
            self.error("Expected: statement")
        while True:
            if self.at_op("."):
                dot = self.peek()
                if dot.space_before and self.i > 0 and self.t[self.i - 1].line == dot.line:
                    # `MsgBox .Name` in a With block: a call whose argument is a With member
                    # (a member on the next line after ` _` is still a member).
                    break
                self.next()
                e = A.Member(e, self.member_name(), t.line)
            elif self.at_op("!"):
                self.next()
                e = A.Bang(e, self.member_name(), t.line)
            elif self.at_op("("):
                paren = self.peek()
                pos = self.i
                self.next()
                args = self.arg_list()
                if paren.space_before:
                    split = (e, pos)
                e = A.Call(e, args, t.line, paren.space_before)
            else:
                break
        return e, split

    def member_name(self) -> str:
        t = self.next()
        if t.kind == "word":
            return t.text
        if t.kind == "bracket":
            return t.value
        self.error("Expected: identifier", t)
        return ""

    def arg_list_bare(self) -> List[A.Arg]:
        """Arguments of a call statement without parentheses, up to the end of the statement."""
        args: List[A.Arg] = []
        while True:
            if self.at_end() or self.at_kw("ELSE"):
                if args:
                    args.append(A.Arg(None, missing=True))
                break
            if self.at_op(","):
                self.next()
                args.append(A.Arg(None, missing=True))
                continue
            args.append(self.one_arg())
            if not self.accept_op(","):
                break
            if self.at_end():
                args.append(A.Arg(None, missing=True))
                break
        return args

    def one_arg(self) -> A.Arg:
        t = self.peek()
        if t.kind == "word" and self.peek(1).kind == "op" and self.peek(1).text == ":=":
            self.next()
            self.next()
            return A.Arg(self.expr(), t.text)
        if self.at_kw("BYVAL"):
            self.next()
        return A.Arg(self.expr())

    def arg_list(self) -> List[A.Arg]:
        """Arguments after ``(``, through the closing ``)``."""
        args: List[A.Arg] = []
        if self.accept_op(")"):
            return args
        while True:
            if self.at_op(","):
                self.next()
                args.append(A.Arg(None, missing=True))
                if self.at_op(")"):
                    args.append(A.Arg(None, missing=True))
                    break
                continue
            args.append(self.one_arg())
            if self.accept_op(","):
                if self.at_op(")"):
                    args.append(A.Arg(None, missing=True))
                    break
                continue
            break
        self.expect_op(")")
        return args

    # -- individual statements

    def st_dim(self, line: int) -> List[A.Stmt]:
        static = self.next().up == "STATIC"
        self.accept_kw("WITHEVENTS")
        return [A.Dim(self.var_list(), static, line)]

    def st_const(self, line: int) -> List[A.Stmt]:
        self.next()
        return [A.Const(self.const_list(), line)]

    def st_redim(self, line: int) -> List[A.Stmt]:
        self.next()
        preserve = bool(self.accept_kw("PRESERVE"))
        items = []
        while True:
            t = self.peek()
            target, _ = self._redim_target()
            if not isinstance(target, A.Call):
                self.error("Expected: (")
            bounds = []
            for a in target.args:
                if a.expr is None:
                    self.error("Expected: expression")
                if isinstance(a.expr, A.Binary) and a.expr.op == "TO":
                    bounds.append((a.expr.left, a.expr.right))
                else:
                    bounds.append((None, a.expr))
            tr = self.as_type() if self.at_kw("AS") else None
            items.append((target.callee, bounds, tr))
            if not self.accept_op(","):
                break
        return [A.ReDim(preserve, items, line)]

    def _redim_target(self):
        t = self.ident("variable name")
        e: A.Expr = A.Ident(t.text, t.suffix, t.line)
        while self.at_op("."):
            self.next()
            e = A.Member(e, self.member_name(), t.line)
        self.expect_op("(")
        bounds: List[A.Arg] = []
        while True:
            a = self.expr()
            if self.accept_kw("TO"):
                b = self.expr()
                a = A.Binary("TO", a, b, t.line)
            bounds.append(A.Arg(a))
            if not self.accept_op(","):
                break
        self.expect_op(")")
        return A.Call(e, bounds, t.line), None

    def st_set(self, line: int) -> List[A.Stmt]:
        self.next()
        target, _ = self.lexpr()
        self.expect_op("=")
        return [A.Assign(target, self.expr(), "set", line)]

    def st_let(self, line: int) -> List[A.Stmt]:
        self.next()
        target, _ = self.lexpr()
        self.expect_op("=")
        return [self._assign(target, self.expr(), "let", line)]

    def st_lset(self, line: int) -> List[A.Stmt]:
        kind = self.next().up.lower()
        target, _ = self.lexpr()
        self.expect_op("=")
        return [A.Assign(target, self.expr(), kind, line)]

    def st_call(self, line: int) -> List[A.Stmt]:
        self.next()
        target, _ = self.lexpr()
        if isinstance(target, A.Call):
            return [A.CallStmt(target.callee, target.args, True, line)]
        return [A.CallStmt(target, [], True, line)]

    def st_if(self, line: int) -> List[A.Stmt]:
        self.next()
        cond = self.expr()
        if not self.accept_kw("THEN"):
            if self.at_kw("GOTO"):
                self.next()
                return [A.If([(cond, [A.GoTo(self.label_name(), False, line)])], None, line)]
            self.error("Expected: Then or GoTo")
        if self.peek().kind in ("nl", "eof"):
            return [self.block_if(cond, line)]
        # Single-line If: statements up to Else, then up to the end of the line.
        then = self.line_statements()
        orelse = None
        if self.accept_kw("ELSE"):
            orelse = self.line_statements()
        return [A.If([(cond, then)], orelse, line)]

    def line_statements(self) -> List[A.Stmt]:
        out: List[A.Stmt] = []
        if self.peek().kind == "num" and not self.peek().value[2]:
            t = self.next()
            return [A.GoTo(str(t.value[0]), False, t.line)]
        while True:
            if self.at_end() and not self.at_op(":"):
                break
            if self.at_kw("ELSE"):
                break
            if self.at_op(":"):
                self.next()
                continue
            out.extend(self.statement())
            if self.at_op(":"):
                self.next()
                continue
            break
        return out

    def block_if(self, cond: A.Expr, line: int) -> A.Stmt:
        branches = []
        enders = {("ELSEIF",), ("ELSE",), ("END", "IF")}
        body = self.block(enders)
        branches.append((cond, body))
        orelse = None
        while True:
            e = self.block_end(enders)
            if e == ("ELSEIF",):
                self.next()
                c = self.expr()
                self.expect_kw("THEN")
                branches.append((c, self.block(enders)))
            elif e == ("ELSE",):
                self.next()
                self.accept_op(":")
                orelse = self.block({("END", "IF")})
            else:
                if self.peek().up == "ENDIF":
                    self.next()
                else:
                    self.next()
                    self.next()
                break
        return A.If(branches, orelse, line)

    def st_select(self, line: int) -> List[A.Stmt]:
        self.next()
        self.expect_kw("CASE")
        subject = self.expr()
        self.end_stmt()
        cases = []
        orelse = None
        enders = {("CASE",), ("END", "SELECT")}
        self.skip_newlines()
        while True:
            self.skip_newlines()
            if self.at_kw("END") and self.at_kw_at(1, "SELECT"):
                self.next()
                self.next()
                break
            if not self.at_kw("CASE"):
                self.error("Expected: Case")
            cl = self.next().line
            if self.accept_kw("ELSE"):
                self.accept_op(":")
                orelse = self.block(enders)
                continue
            clauses = []
            while True:
                if self.accept_kw("IS"):
                    op = self.next()
                    if op.kind != "op" or op.text not in _COMPARE:
                        self.error("Expected: comparison operator")
                    clauses.append(A.CaseClause("is", self.expr(), None, op.text))
                else:
                    a = self.expr()
                    if self.accept_kw("TO"):
                        clauses.append(A.CaseClause("range", a, self.expr()))
                    else:
                        clauses.append(A.CaseClause("eq", a))
                if not self.accept_op(","):
                    break
            self.accept_op(":")
            cases.append((clauses, self.block(enders), cl))
        return [A.Select(subject, cases, orelse, line)]

    def st_for(self, line: int) -> List[A.Stmt]:
        self.next()
        if self.accept_kw("EACH"):
            var, _ = self._for_var()
            self.expect_kw("IN")
            coll = self.expr()
            body = self.block({("NEXT",)})
            end_line = self.peek().line
            self._take_next()
            return [A.ForEach(var, coll, body, line, end_line)]
        var, _ = self._for_var()
        self.expect_op("=")
        start = self.expr()
        self.expect_kw("TO")
        end = self.expr()
        step = self.expr() if self.accept_kw("STEP") else None
        body = self.block({("NEXT",)})
        end_line = self.peek().line
        self._take_next()
        return [A.For(var, start, end, step, body, line, end_line)]

    def _for_var(self):
        t = self.ident("loop variable")
        e: A.Expr = A.Ident(t.text, t.suffix, t.line)
        while self.at_op("."):
            self.next()
            e = A.Member(e, self.member_name(), t.line)
        return e, t

    def _take_next(self) -> None:
        if self.pending_next:
            self.pending_next -= 1
            # `Next j, i`: this loop's variable name follows the comma.
            if self.peek().kind == "word" and not self.at_end():
                self.next()
                if self.accept_op(","):
                    self.pending_next += 1
            return
        self.expect_kw("NEXT")
        if self.peek().kind == "word" and not self.at_kw("ELSE"):
            self.next()
            while self.at_op("."):
                self.next()
                self.next()
            if self.accept_op(","):
                self.pending_next += 1

    def st_do(self, line: int) -> List[A.Stmt]:
        self.next()
        pre = None
        if self.at_kw("WHILE", "UNTIL"):
            k = self.next().up.lower()
            pre = (k, self.expr())
        body = self.block({("LOOP",)})
        end_line = self.peek().line
        self.next()
        post = None
        if self.at_kw("WHILE", "UNTIL"):
            k = self.next().up.lower()
            post = (k, self.expr())
        return [A.Do(pre, post, body, line, end_line)]

    def st_while(self, line: int) -> List[A.Stmt]:
        self.next()
        cond = self.expr()
        body = self.block({("WEND",)})
        end_line = self.peek().line
        self.next()
        return [A.While(cond, body, line, end_line)]

    def st_with(self, line: int) -> List[A.Stmt]:
        self.next()
        e = self.expr()
        body = self.block({("END", "WITH")})
        self.next()
        self.next()
        return [A.With(e, body, line)]

    def st_exit(self, line: int) -> List[A.Stmt]:
        self.next()
        w = self.next()
        if w.kind != "word" or w.up not in ("SUB", "FUNCTION", "PROPERTY", "FOR", "DO"):
            self.error("Expected: Do, For, Sub, Function or Property")
        return [A.Exit(w.up.lower(), line)]

    def st_end(self, line: int) -> List[A.Stmt]:
        if self.peek(1).kind == "word" and self.peek(1).up in _BLOCK_ENDS:
            what = self.peek(1).text
            self.error(f"End {what} without block {what}" if what.upper() in ("IF", "WITH", "SELECT") else f"Expected: statement (End {what} here doesn't close anything)")
        self.next()
        return [A.EndStmt(line)]

    def st_stop(self, line: int) -> List[A.Stmt]:
        self.next()
        return [A.StopStmt(line)]

    def label_name(self) -> str:
        t = self.next()
        if t.kind == "num":
            return str(t.value[0])
        if t.kind == "word":
            return t.text
        self.error("Expected: label", t)
        return ""

    def st_goto(self, line: int) -> List[A.Stmt]:
        gosub = self.next().up == "GOSUB"
        return [A.GoTo(self.label_name(), gosub, line)]

    def st_return(self, line: int) -> List[A.Stmt]:
        self.next()
        return [A.ReturnStmt(line)]

    def st_on(self, line: int) -> List[A.Stmt]:
        self.next()
        self.accept_kw("LOCAL")
        if self.accept_kw("ERROR"):
            if self.accept_kw("RESUME"):
                self.expect_kw("NEXT")
                return [A.OnError("resume_next", "", line)]
            self.expect_kw("GOTO")
            neg = self.accept_op("-")
            t = self.next()
            if t.kind == "num" and t.value[0] == 0 and not neg:
                return [A.OnError("goto0", "", line)]
            if t.kind == "num" and t.value[0] == 1 and neg:
                return [A.OnError("goto_minus1", "", line)]
            if t.kind == "num":
                return [A.OnError("goto", str(t.value[0]), line)]
            if t.kind == "word":
                return [A.OnError("goto", t.text, line)]
            self.error("Expected: label")
        e = self.expr()
        if self.accept_kw("GOTO"):
            gosub = False
        else:
            self.expect_kw("GOSUB")
            gosub = True
        labels = [self.label_name()]
        while self.accept_op(","):
            labels.append(self.label_name())
        return [A.OnGoTo(e, labels, gosub, line)]

    def st_resume(self, line: int) -> List[A.Stmt]:
        self.next()
        if self.accept_kw("NEXT"):
            return [A.Resume("next", "", line)]
        if self.at_end() or self.at_kw("ELSE"):
            return [A.Resume("retry", "", line)]
        t = self.next()
        if t.kind == "num" and t.value[0] == 0:
            return [A.Resume("retry", "", line)]
        return [A.Resume("label", str(t.value[0]) if t.kind == "num" else t.text, line)]

    def st_erase(self, line: int) -> List[A.Stmt]:
        self.next()
        targets = []
        while True:
            e, _ = self.lexpr()
            targets.append(e)
            if not self.accept_op(","):
                break
        return [A.Erase(targets, line)]

    def st_raiseevent(self, line: int) -> List[A.Stmt]:
        self.next()
        name = self.ident("event name").text
        args = []
        if self.accept_op("("):
            args = self.arg_list()
        return [A.RaiseEvent(name, args, line)]

    def st_error(self, line: int) -> List[A.Stmt]:
        self.next()
        return [A.ErrorStmt(self.expr(), line)]

    def st_randomize(self, line: int) -> List[A.Stmt]:
        self.next()
        e = None if self.at_end() or self.at_kw("ELSE") else self.expr()
        return [A.Randomize(e, line)]

    def st_load(self, line: int) -> List[A.Stmt]:
        t = self.next()
        return [A.LoadUnload(self.expr(), t.up == "UNLOAD", line)]

    def st_attribute(self, line: int) -> List[A.Stmt]:
        self.next()
        start = self.i
        self.rest_of_statement()
        text = " ".join(t.text for t in self.t[start:self.i])
        self.attributes.append("Attribute " + text)
        return [A.NoOp(line)]

    def st_beep(self, line: int) -> List[A.Stmt]:
        self.next()
        return [A.NoOp(line)]

    def st_unsupported(self, line: int) -> List[A.Stmt]:
        t = self.next()
        what = t.text
        if t.up == "LINE" and self.at_kw("INPUT"):
            what = "Line Input"
        elif t.up in ("PRINT", "WRITE", "INPUT", "GET", "PUT", "OPEN", "CLOSE", "SEEK", "LOCK", "UNLOCK", "WIDTH", "LINE", "RESET"):
            what = f"File I/O ({t.text})"
        elif t.up in ("DATE", "TIME"):
            what = f"Setting the system {t.text.lower()}"
        self.rest_of_statement()
        return [A.Unsupported(what, line)]

    def print_items(self, line: int) -> A.PrintStmt:
        items: List[Tuple[Optional[A.Expr], str]] = []
        while not (self.at_end() or self.at_kw("ELSE")):
            if self.at_op(";") or self.at_op(","):
                items.append((None, self.next().text))
                continue
            e = self.expr()
            sep = ""
            if self.at_op(";") or self.at_op(","):
                sep = self.next().text
            items.append((e, sep))
        return A.PrintStmt(items, "debug", line)

    # ------------------------------------------------------------ expressions

    def expr(self, min_prec: int = 1) -> A.Expr:
        left = self.unary()
        while True:
            t = self.peek()
            if t.kind == "op":
                op = t.text
            elif t.kind == "word" and not t.suffix and t.up in ("AND", "OR", "XOR", "EQV", "IMP", "MOD", "LIKE", "IS"):
                op = t.up
            else:
                break
            p = _PREC.get(op)
            if p is None or p < min_prec:
                break
            self.next()
            right = self.expr(p + 1)
            left = A.Binary(op, left, right, t.line)
        return left

    def unary(self) -> A.Expr:
        t = self.peek()
        if t.kind == "word" and t.up == "NOT" and not t.suffix:
            self.next()
            return A.Unary("NOT", self.expr(7), t.line)
        if t.kind == "op" and t.text == "-":
            self.next()
            return A.Unary("-", self.expr(14), t.line)
        if t.kind == "op" and t.text == "+":
            self.next()
            return A.Unary("+", self.expr(14), t.line)
        return self.postfix(self.primary())

    def primary(self) -> A.Expr:
        t = self.next()
        k = t.kind
        if k == "num":
            return _number(t, self)
        if k == "str":
            return A.Lit(t.value, "String", t.line)
        if k == "date":
            return A.Lit(t.value, "Date", t.line)
        if k == "bracket":
            body = t.value.strip()
            if body.startswith("_") and body[1:].replace("_", "").isalnum():
                return A.Ident(body, "", t.line)
            return A.BracketEval(body, t.line)
        if k == "op":
            if t.text == "(":
                e = self.expr()
                self.expect_op(")")
                return A.Paren(e, t.line)
            if t.text == ".":
                return A.Member(None, self.member_name(), t.line)
            if t.text == "!":
                return A.Bang(None, self.member_name(), t.line)
            self.error("Expected: expression", t)
        if k == "word":
            up = t.up
            if not t.suffix:
                if up == "TRUE":
                    return A.Lit(True, "Boolean", t.line)
                if up == "FALSE":
                    return A.Lit(False, "Boolean", t.line)
                if up == "NOTHING":
                    return A.Lit(None, "Nothing", t.line)
                if up == "EMPTY":
                    return A.Lit(None, "Empty", t.line)
                if up == "NULL":
                    return A.Lit(None, "Null", t.line)
                if up == "ME":
                    return A.MeExpr(t.line)
                if up == "NEW":
                    return A.New(self.type_name(), t.line)
                if up == "TYPEOF":
                    e = self.postfix(self.primary())
                    self.expect_kw("IS")
                    return A.TypeOfIs(e, self.type_name(), t.line)
                if up == "ADDRESSOF":
                    return A.AddressOf(self.ident().text, t.line)
                if up in RESERVED:
                    self.error("Expected: expression", t)
            return A.Ident(t.text, t.suffix, t.line)
        self.error("Expected: expression", t)
        return A.Lit(None, "Empty")

    def postfix(self, e: A.Expr) -> A.Expr:
        while True:
            t = self.peek()
            if t.kind != "op":
                return e
            if t.text == ".":
                self.next()
                e = A.Member(e, self.member_name(), t.line)
            elif t.text == "!":
                self.next()
                e = A.Bang(e, self.member_name(), t.line)
            elif t.text == "(":
                self.next()
                e = A.Call(e, self.arg_list(), t.line, t.space_before)
            else:
                return e


def _number(t: Tok, p: Parser) -> A.Lit:
    val, suffix, is_float, base = t.value
    line = t.line
    if base != 10:
        # Hex and octal literals are two's complement in their type.
        if suffix == "%" or (not suffix and val <= 0xFFFF):
            if val > 0xFFFF:
                p.error("Overflow", t)
            return A.Lit(Integer(val - 0x10000 if val >= 0x8000 else val), "Integer", line)
        if suffix == "&" or (not suffix and val <= 0xFFFFFFFF):
            if val > 0xFFFFFFFF:
                p.error("Overflow", t)
            return A.Lit(val - 0x100000000 if val >= 0x80000000 else val, "Long", line)
        if val > 0xFFFFFFFFFFFFFFFF:
            p.error("Overflow", t)
        return A.Lit(LongLong(val - (1 << 64) if val >= (1 << 63) else val), "LongLong", line)
    if suffix == "&":
        if not is_float and val <= 2147483647:
            return A.Lit(int(val), "Long", line)
        p.error("Overflow", t)
    if suffix == "%":
        if is_float or val > 32767:
            p.error("Overflow", t)
        return A.Lit(Integer(val), "Integer", line)
    if suffix == "^":
        return A.Lit(LongLong(int(val)), "LongLong", line)
    if suffix == "!":
        return A.Lit(Single(float(val)), "Single", line)
    if suffix == "#":
        return A.Lit(float(val), "Double", line)
    if suffix == "@":
        return A.Lit(Currency.of(float(val)) if is_float else Currency(int(val) * 10000), "Currency", line)
    if is_float:
        return A.Lit(float(val), "Double", line)
    if val <= 32767:
        return A.Lit(Integer(val), "Integer", line)
    if val <= 2147483647:
        return A.Lit(int(val), "Long", line)
    return A.Lit(float(val), "Double", line)


_STATEMENTS: Dict[str, Callable[[Parser, int], List[A.Stmt]]] = {
    "DIM": Parser.st_dim,
    "STATIC": Parser.st_dim,
    "CONST": Parser.st_const,
    "REDIM": Parser.st_redim,
    "SET": Parser.st_set,
    "LET": Parser.st_let,
    "LSET": Parser.st_lset,
    "RSET": Parser.st_lset,
    "CALL": Parser.st_call,
    "IF": Parser.st_if,
    "SELECT": Parser.st_select,
    "FOR": Parser.st_for,
    "DO": Parser.st_do,
    "WHILE": Parser.st_while,
    "WITH": Parser.st_with,
    "EXIT": Parser.st_exit,
    "END": Parser.st_end,
    "STOP": Parser.st_stop,
    "GOTO": Parser.st_goto,
    "GOSUB": Parser.st_goto,
    "RETURN": Parser.st_return,
    "ON": Parser.st_on,
    "RESUME": Parser.st_resume,
    "ERASE": Parser.st_erase,
    "RAISEEVENT": Parser.st_raiseevent,
    "ERROR": Parser.st_error,
    "RANDOMIZE": Parser.st_randomize,
    "BEEP": Parser.st_beep,
    "ATTRIBUTE": Parser.st_attribute,
    "LOAD": Parser.st_load,
    "UNLOAD": Parser.st_load,
}
for _w in ("OPEN", "CLOSE", "PRINT", "WRITE", "INPUT", "LINE", "GET", "PUT", "SEEK", "LOCK", "UNLOCK", "WIDTH", "KILL", "MKDIR", "RMDIR", "CHDIR", "CHDRIVE", "FILECOPY", "NAME", "SAVESETTING", "DELETESETTING", "SENDKEYS", "APPACTIVATE", "RESET", "DATE", "TIME", "SETATTR"):
    _STATEMENTS[_w] = Parser.st_unsupported
