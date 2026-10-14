"""The VBA syntax tree the parser builds and the compiler turns into closures."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, List, Optional, Tuple

# ---------------------------------------------------------------- types


@dataclass
class TypeRef:
    """``As <name>``: a built-in type, a class, a user-defined type or an Excel/library type
    (``Range``, ``Worksheet``…). ``string_len`` for ``String * n``; ``new`` for ``As New``."""

    name: str = "Variant"
    string_len: Optional[int] = None
    new: bool = False
    explicit: bool = False


# ---------------------------------------------------------------- expressions


class Expr:
    __slots__ = ("line",)


class Lit(Expr):
    __slots__ = ("value", "vtype")

    def __init__(self, value: Any, vtype: str, line: int = 0):
        self.value = value
        self.vtype = vtype
        self.line = line


class Ident(Expr):
    __slots__ = ("name", "suffix")

    def __init__(self, name: str, suffix: str = "", line: int = 0):
        self.name = name
        self.suffix = suffix
        self.line = line


class Member(Expr):
    """``obj.name``; ``obj`` is None for ``.name`` inside a With block."""

    __slots__ = ("obj", "name")

    def __init__(self, obj: Optional[Expr], name: str, line: int = 0):
        self.obj = obj
        self.name = name
        self.line = line


class Bang(Expr):
    """``obj!name``: the default member called with the name as a string."""

    __slots__ = ("obj", "name")

    def __init__(self, obj: Optional[Expr], name: str, line: int = 0):
        self.obj = obj
        self.name = name
        self.line = line


class Arg:
    __slots__ = ("expr", "name", "missing")

    def __init__(self, expr: Optional[Expr], name: Optional[str] = None, missing: bool = False):
        self.expr = expr
        self.name = name
        self.missing = missing


class Call(Expr):
    """``callee(args)``: a call, an array element or a default member with arguments."""

    __slots__ = ("callee", "args", "spaced")

    def __init__(self, callee: Expr, args: List[Arg], line: int = 0, spaced: bool = False):
        self.callee = callee
        self.args = args
        self.line = line
        self.spaced = spaced


class Paren(Expr):
    __slots__ = ("expr",)

    def __init__(self, expr: Expr, line: int = 0):
        self.expr = expr
        self.line = line


class Unary(Expr):
    __slots__ = ("op", "expr")

    def __init__(self, op: str, expr: Expr, line: int = 0):
        self.op = op
        self.expr = expr
        self.line = line


class Binary(Expr):
    __slots__ = ("op", "left", "right")

    def __init__(self, op: str, left: Expr, right: Expr, line: int = 0):
        self.op = op
        self.left = left
        self.right = right
        self.line = line


class New(Expr):
    __slots__ = ("type_name",)

    def __init__(self, type_name: str, line: int = 0):
        self.type_name = type_name
        self.line = line


class TypeOfIs(Expr):
    __slots__ = ("expr", "type_name")

    def __init__(self, expr: Expr, type_name: str, line: int = 0):
        self.expr = expr
        self.type_name = type_name
        self.line = line


class AddressOf(Expr):
    __slots__ = ("name",)

    def __init__(self, name: str, line: int = 0):
        self.name = name
        self.line = line


class MeExpr(Expr):
    __slots__ = ()

    def __init__(self, line: int = 0):
        self.line = line


class BracketEval(Expr):
    """``[A1:B5]``: Excel's shorthand for ``Evaluate("A1:B5")``."""

    __slots__ = ("text",)

    def __init__(self, text: str, line: int = 0):
        self.text = text
        self.line = line


# ---------------------------------------------------------------- statements


class Stmt:
    __slots__ = ("line",)


class Assign(Stmt):
    __slots__ = ("target", "value", "kind")

    def __init__(self, target: Expr, value: Expr, kind: str = "let", line: int = 0):
        self.target = target
        self.value = value
        self.kind = kind  # let set lset rset
        self.line = line


class CallStmt(Stmt):
    __slots__ = ("callee", "args", "explicit")

    def __init__(self, callee: Expr, args: List[Arg], explicit: bool = False, line: int = 0):
        self.callee = callee
        self.args = args
        self.explicit = explicit
        self.line = line


class VarDecl:
    __slots__ = ("name", "type", "bounds", "is_array", "withevents", "line", "suffix")

    def __init__(self, name: str, type_: TypeRef, bounds: Optional[List[Tuple[Optional[Expr], Expr]]] = None, is_array: bool = False, withevents: bool = False, line: int = 0, suffix: str = ""):
        self.name = name
        self.type = type_
        self.bounds = bounds  # fixed array bounds [(lower or None, upper)]
        self.is_array = is_array  # x() or x(bounds)
        self.withevents = withevents
        self.line = line
        self.suffix = suffix


class ConstDecl:
    __slots__ = ("name", "type", "expr", "line")

    def __init__(self, name: str, type_: TypeRef, expr: Expr, line: int = 0):
        self.name = name
        self.type = type_
        self.expr = expr
        self.line = line


class Dim(Stmt):
    __slots__ = ("vars", "static")

    def __init__(self, vars_: List[VarDecl], static: bool = False, line: int = 0):
        self.vars = vars_
        self.static = static
        self.line = line


class Const(Stmt):
    __slots__ = ("consts",)

    def __init__(self, consts: List[ConstDecl], line: int = 0):
        self.consts = consts
        self.line = line


class ReDim(Stmt):
    __slots__ = ("preserve", "items")

    def __init__(self, preserve: bool, items: List[Tuple[Expr, List[Tuple[Optional[Expr], Expr]], Optional[TypeRef]]], line: int = 0):
        self.preserve = preserve
        self.items = items
        self.line = line


class Erase(Stmt):
    __slots__ = ("targets",)

    def __init__(self, targets: List[Expr], line: int = 0):
        self.targets = targets
        self.line = line


class If(Stmt):
    __slots__ = ("branches", "orelse")

    def __init__(self, branches: List[Tuple[Expr, List[Stmt]]], orelse: Optional[List[Stmt]], line: int = 0):
        self.branches = branches
        self.orelse = orelse
        self.line = line


class CaseClause:
    __slots__ = ("kind", "a", "b", "op")

    def __init__(self, kind: str, a: Expr, b: Optional[Expr] = None, op: str = "="):
        self.kind = kind  # eq range is
        self.a = a
        self.b = b
        self.op = op


class Select(Stmt):
    __slots__ = ("expr", "cases", "orelse")

    def __init__(self, expr: Expr, cases: List[Tuple[List[CaseClause], List[Stmt], int]], orelse: Optional[List[Stmt]], line: int = 0):
        self.expr = expr
        self.cases = cases
        self.orelse = orelse
        self.line = line


class For(Stmt):
    __slots__ = ("var", "start", "end", "step", "body", "end_line")

    def __init__(self, var: Expr, start: Expr, end: Expr, step: Optional[Expr], body: List[Stmt], line: int = 0, end_line: int = 0):
        self.var = var
        self.start = start
        self.end = end
        self.step = step
        self.body = body
        self.line = line
        self.end_line = end_line


class ForEach(Stmt):
    __slots__ = ("var", "coll", "body", "end_line")

    def __init__(self, var: Expr, coll: Expr, body: List[Stmt], line: int = 0, end_line: int = 0):
        self.var = var
        self.coll = coll
        self.body = body
        self.line = line
        self.end_line = end_line


class Do(Stmt):
    __slots__ = ("pre", "post", "body", "end_line")

    def __init__(self, pre: Optional[Tuple[str, Expr]], post: Optional[Tuple[str, Expr]], body: List[Stmt], line: int = 0, end_line: int = 0):
        self.pre = pre
        self.post = post
        self.body = body
        self.line = line
        self.end_line = end_line


class While(Stmt):
    __slots__ = ("cond", "body", "end_line")

    def __init__(self, cond: Expr, body: List[Stmt], line: int = 0, end_line: int = 0):
        self.cond = cond
        self.body = body
        self.line = line
        self.end_line = end_line


class With(Stmt):
    __slots__ = ("expr", "body")

    def __init__(self, expr: Expr, body: List[Stmt], line: int = 0):
        self.expr = expr
        self.body = body
        self.line = line


class Exit(Stmt):
    __slots__ = ("what",)

    def __init__(self, what: str, line: int = 0):
        self.what = what  # sub function property for do
        self.line = line


class EndStmt(Stmt):
    __slots__ = ()

    def __init__(self, line: int = 0):
        self.line = line


class StopStmt(Stmt):
    __slots__ = ()

    def __init__(self, line: int = 0):
        self.line = line


class GoTo(Stmt):
    __slots__ = ("label", "gosub")

    def __init__(self, label: str, gosub: bool = False, line: int = 0):
        self.label = label
        self.gosub = gosub
        self.line = line


class ReturnStmt(Stmt):
    __slots__ = ()

    def __init__(self, line: int = 0):
        self.line = line


class OnError(Stmt):
    __slots__ = ("mode", "label")

    def __init__(self, mode: str, label: str = "", line: int = 0):
        self.mode = mode  # goto resume_next goto0 goto_minus1
        self.label = label
        self.line = line


class OnGoTo(Stmt):
    __slots__ = ("expr", "labels", "gosub")

    def __init__(self, expr: Expr, labels: List[str], gosub: bool, line: int = 0):
        self.expr = expr
        self.labels = labels
        self.gosub = gosub
        self.line = line


class Resume(Stmt):
    __slots__ = ("kind", "label")

    def __init__(self, kind: str, label: str = "", line: int = 0):
        self.kind = kind  # retry next label
        self.label = label
        self.line = line


class Label(Stmt):
    __slots__ = ("name", "number")

    def __init__(self, name: str, number: Optional[int] = None, line: int = 0):
        self.name = name
        self.number = number
        self.line = line


class RaiseEvent(Stmt):
    __slots__ = ("name", "args")

    def __init__(self, name: str, args: List[Arg], line: int = 0):
        self.name = name
        self.args = args
        self.line = line


class MidAssign(Stmt):
    __slots__ = ("target", "start", "length", "value", "bytes_")

    def __init__(self, target: Expr, start: Expr, length: Optional[Expr], value: Expr, line: int = 0):
        self.target = target
        self.start = start
        self.length = length
        self.value = value
        self.line = line


class PrintStmt(Stmt):
    """``Debug.Print a; b, c``: items with the separator that follows each."""

    __slots__ = ("items", "target")

    def __init__(self, items: List[Tuple[Optional[Expr], str]], target: str = "debug", line: int = 0):
        self.items = items
        self.target = target
        self.line = line


class Unsupported(Stmt):
    __slots__ = ("what",)

    def __init__(self, what: str, line: int = 0):
        self.what = what
        self.line = line


class ErrorStmt(Stmt):
    __slots__ = ("expr",)

    def __init__(self, expr: Expr, line: int = 0):
        self.expr = expr
        self.line = line


class Randomize(Stmt):
    __slots__ = ("expr",)

    def __init__(self, expr: Optional[Expr], line: int = 0):
        self.expr = expr
        self.line = line


class NoOp(Stmt):
    __slots__ = ()

    def __init__(self, line: int = 0):
        self.line = line


# ---------------------------------------------------------------- module level


@dataclass
class Param:
    name: str
    type: TypeRef
    byval: bool = False
    optional: bool = False
    default: Optional[Expr] = None
    paramarray: bool = False
    is_array: bool = False
    line: int = 0


@dataclass
class Proc:
    kind: str  # sub function get let set
    name: str
    params: List[Param]
    ret: TypeRef
    body: List[Stmt]
    scope: str = "public"  # public private friend
    static: bool = False
    line: int = 0
    end_line: int = 0
    ret_is_array: bool = False


@dataclass
class TypeDecl:
    name: str
    fields: List[VarDecl]
    scope: str = "public"
    line: int = 0


@dataclass
class EnumDecl:
    name: str
    members: List[Tuple[str, Optional[Expr]]]
    scope: str = "public"
    line: int = 0


@dataclass
class DeclareDecl:
    name: str
    kind: str
    lib: str
    alias: str
    params: List[Param]
    ret: TypeRef
    scope: str = "public"
    line: int = 0


@dataclass
class EventDecl:
    name: str
    params: List[Param]
    line: int = 0


@dataclass
class ModuleVar:
    decl: VarDecl
    scope: str  # public private


@dataclass
class ModuleConst:
    decl: ConstDecl
    scope: str


@dataclass
class ModuleAst:
    name: str
    option_explicit: bool = False
    option_base: int = 0
    option_compare_text: bool = False
    option_private: bool = False
    variables: List[ModuleVar] = field(default_factory=list)
    constants: List[ModuleConst] = field(default_factory=list)
    procs: List[Proc] = field(default_factory=list)
    types: List[TypeDecl] = field(default_factory=list)
    enums: List[EnumDecl] = field(default_factory=list)
    declares: List[DeclareDecl] = field(default_factory=list)
    events: List[EventDecl] = field(default_factory=list)
    implements: List[str] = field(default_factory=list)
    deftypes: List[Tuple[str, str, str]] = field(default_factory=list)
    # `Attribute X.VB_… = …` lines found in the code itself.
    attributes: List[str] = field(default_factory=list)
