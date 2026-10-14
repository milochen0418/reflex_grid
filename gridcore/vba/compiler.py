"""Compiling VBA modules into a runnable :class:`ProjectRuntime`.

Each procedure becomes a list of instructions — closures ``f(frame) -> next instruction`` —
with expressions compiled to closures too. Names are resolved once, here: locals, then the
module, then other modules' public members, then the VBA library, then the host's globals
(Excel's ``Range``, ``Cells``…), as VBA's reference order has it. Static types travel with each
compiled expression, because VBA's operators behave differently for typed operands and Variants
(overflow, number/string comparison).
"""

from __future__ import annotations

import re

import itertools
from contextlib import contextmanager
from typing import Any, Callable, Dict, Iterator, List, Optional, Sequence, Set, Tuple

from . import ast as A
from . import values as V
from .errors import MacroStopped, NotSupported, VbaError, VbaSyntaxError
from .parser import parse_module
from .runtime import (
    AutoNewSlot,
    ElementRef,
    Execution,
    FieldRef,
    Frame,
    HostObject,
    Instance,
    ModuleRuntime,
    ParamSpec,
    Procedure,
    Slot,
    UdtDef,
    WithEventsSlot,
    call_procedure,
    enumerate_object,
    get_default,
    get_member,
    index_value,
    let_default,
    let_index,
    let_member,
    _unref,
    _unref_named,
    run_terminations,
    set_member,
    slot_value,
)
from .values import EMPTY, MISSING, NOTHING, NULL, VArray, copy_value, is_object, is_true

OBJECT_TYPES = {"OBJECT"}
VALUE_TYPES = {"INTEGER", "LONG", "LONGLONG", "LONGPTR", "BYTE", "SINGLE", "DOUBLE", "CURRENCY", "DECIMAL", "DATE", "STRING", "BOOLEAN", "VARIANT", "ANY"}
_NUM_STATIC = {"Byte": 1, "Boolean": 2, "Integer": 2, "Long": 3, "LongLong": 4, "Single": 5, "Double": 6, "Currency": 7, "Decimal": 8}


class CompileContext:
    """Where an expression is compiled: the module, the procedure (None at module level) and
    its local names."""


# ---------------------------------------------------------------- bindings


class B:
    """A resolved name."""

    __slots__ = ("kind", "a", "b", "type")

    def __init__(self, kind: str, a: Any = None, b: Any = None, type_: str = "Variant"):
        self.kind = kind
        self.a = a
        self.b = b
        self.type = type_


class LocalInfo:
    __slots__ = ("index", "type", "is_array", "is_object", "auto_new", "const", "value", "udt")

    def __init__(self, index: int, type_: str = "Variant", is_array: bool = False, is_object: bool = False, auto_new: bool = False):
        self.index = index
        self.type = type_
        self.is_array = is_array
        self.is_object = is_object
        self.auto_new = auto_new
        self.const = False
        self.value: Any = None
        self.udt = None


# The enumerations of the libraries a project references (VBA's own, and the host's: Excel's
# XlDirection…), lower case, with their members. A variable of an enum type is a Long.
_LIB_ENUMS: Dict[str, Dict[str, Any]] = {}


def register_enums(enums: Dict[str, Dict[str, Any]]) -> None:
    for name, members in enums.items():
        _LIB_ENUMS.setdefault(name.lower(), {k.lower(): v for k, v in members.items()})


def _vba_enums() -> None:
    from .stdlib import CONSTANTS, ENUMS

    register_enums({e: {m: CONSTANTS[m] for m in ms} for e, ms in ENUMS.items()})


def _norm_type(name: str) -> str:
    """``Long``, ``VBA.String``, ``Excel.Range`` → the type's own name (a library enum: Long)."""
    n = name.split(".")[-1] if "." in name else name
    m = {
        "integer": "Integer", "long": "Long", "longlong": "LongLong", "longptr": "LongPtr", "byte": "Byte",
        "single": "Single", "double": "Double", "currency": "Currency", "decimal": "Decimal", "date": "Date",
        "string": "String", "boolean": "Boolean", "variant": "Variant", "object": "Object", "any": "Variant",
    }
    low = n.lower()
    if low in m:
        return m[low]
    if low in _LIB_ENUMS:
        return "Long"
    return n


_BRACKET_IDENT = re.compile(r"^[A-Za-z][A-Za-z0-9_]*$")


class ProjectRuntime:
    """A compiled VBA project with its module-level state."""

    def __init__(self, project_name: str = "VBAProject"):
        self.project_name = project_name
        self.modules: Dict[str, ModuleRuntime] = {}
        self.order: List[ModuleRuntime] = []
        self.errors: Dict[str, VbaSyntaxError] = {}
        self.host_globals: Set[str] = set()
        self.host_constants: Dict[str, Any] = {}
        self.host_types: Set[str] = set()
        self.dying: List[Instance] = []
        self.generation = 0
        self.rnd_state = 0x50000
        self.rnd_last = 0.0
        self.timer_base = 0.0
        self._resetters: List[Callable[[], None]] = []
        self.udts: Dict[str, UdtDef] = {}
        self.data: Dict[str, Any] = {}
        # Document modules: "workbook" (ThisWorkbook) or "worksheet".
        self.doc_kinds: Dict[str, str] = {}
        # UserForms: each form module's design (gridcore.model.forms.FormDesign), by name.
        self.form_designs: Dict[str, Any] = {}

    # ------------------------------------------------------------ lookups

    def module(self, name: str) -> Optional[ModuleRuntime]:
        return self.modules.get(name.lower())

    def find(self, name: str) -> Optional[Procedure]:
        """A procedure by ``Name`` or ``Module.Name`` (``Book.xlsm!Module.Name`` too)."""
        n = name.strip()
        if "!" in n:
            n = n.split("!", 1)[1]
        n = n.strip("'")
        if "." in n:
            mod, proc = n.rsplit(".", 1)
            m = self.module(mod)
            return m.proc(proc) if m is not None else None
        for m in self.order:
            if m.kind == "standard":
                p = m.proc(n)
                if p is not None:
                    return p
        for m in self.order:
            p = m.proc(n)
            if p is not None:
                return p
        return None

    def macros(self) -> List[Procedure]:
        """What the Macros dialog lists: public Subs without parameters."""
        out = []
        for m in self.order:
            if m.kind in ("class", "form"):
                continue
            for d in m.procs.values():
                p = d.get("sub")
                if p is not None and p.scope != "private" and not p.params and not m.private_module:
                    out.append(p)
        return out

    def udf(self, name: str) -> Optional[Procedure]:
        """A worksheet function: a public Function of a standard module."""
        n = name.lower()
        if "." in n:
            mod, fn = n.rsplit(".", 1)
            m = self.module(mod)
            if m is None or m.kind != "standard":
                return None
            p = m.proc(fn, "function")
            return p if p is not None and p.scope != "private" else None
        for m in self.order:
            if m.kind == "standard" and not m.private_module:
                p = m.proc(n, "function")
                if p is not None and p.scope != "private":
                    return p
        return None

    def reset_state(self) -> None:
        """Module-level variables back to their initial values (``End``, or the code changed).
        Objects released this way don't run Class_Terminate, as in VBA."""
        self.generation += 1
        self.dying.clear()
        for f in self._resetters:
            f()

    # ------------------------------------------------------------ calling

    def call(self, ex: Execution, proc: Procedure, args: Sequence[Any] = (), named: Optional[dict] = None, me: Any = None) -> Any:
        mod = proc.module
        if mod.name.lower() in self.errors:
            raise self.errors[mod.name.lower()]
        if mod.kind == "document" and me is None:
            me = ex.host.document(ex, mod)
        r = call_procedure(ex, proc, args, named, me)
        if self.dying:
            # Objects the procedure's own variables held.
            run_terminations(ex)
        return r

    def run(self, ex: Execution, name: str, args: Sequence[Any] = ()) -> Any:
        p = self.find(name)
        if p is None:
            raise VbaError(1004, f"Cannot run the macro '{name}'. The macro may not be available in this workbook.")
        with ex:
            return self.call(ex, p, list(args))

    def handler(self, module: str, name: str) -> Optional[Procedure]:
        m = self.module(module)
        if m is None or m.name.lower() in self.errors:
            return None
        return m.proc(name, "sub")

    # ------------------------------------------------------------ classes

    def new_instance(self, ex: Execution, cls: ModuleRuntime) -> Instance:
        if cls.name.lower() in self.errors:
            raise self.errors[cls.name.lower()]
        if cls.kind == "form":
            # `New UserForm1`: a form of its own, loaded now (UserForm_Initialize runs).
            inst = self.make_form(ex, cls)
            ex.host.form_registry(ex).add(inst.form)
            inst.form.start(ex)
            return inst
        inst = Instance(cls, self)
        init = cls.proc("Class_Initialize", "sub")
        if init is not None:
            call_procedure(ex, init, [], None, inst)
        return inst

    def make_form(self, ex: Execution, cls: ModuleRuntime) -> Instance:
        """A UserForm's instance with its controls as designed (not started yet)."""
        from ..model.forms import FormDesign
        from .msforms import UserForm

        if cls.name.lower() in self.errors:
            raise self.errors[cls.name.lower()]
        inst = Instance(cls, self)
        inst.form = UserForm(self, inst, self.form_designs.get(cls.name.lower()) or FormDesign(cls.name))
        return inst

    def form_default(self, ex: Execution, cls: ModuleRuntime) -> Instance:
        """``UserForm1`` used by name: the form's default instance, loaded on first use (and
        again after it was unloaded), as in VBA."""
        return ex.host.form_registry(ex).default(ex, cls)

    def instance_get(self, ex: Execution, inst: Instance, name: str, args: list, named: Optional[dict]) -> Any:
        cls = inst.cls
        low = name.lower()
        d = cls.procs.get(low)
        if d is None and cls.implements:
            for iface in cls.implements:
                d = cls.procs.get(f"{iface.lower()}_{low}")
                if d is not None:
                    break
        if d is not None:
            p = d.get("get") or d.get("function") or d.get("sub")
            if p is not None:
                if p.scope == "private" and not p.name.lower().startswith(tuple(i.lower() + "_" for i in cls.implements)):
                    raise VbaError(438, f"Object doesn't support this property or method: {name}")
                v = call_procedure(ex, p, args if p.params or not args else [], named, inst)
                if args and not p.params:
                    return index_value(ex, v, args, named)
                return v
        i = cls.field_index.get(low)
        if i is not None and cls.var_public.get(low, False):
            v = inst.fields[i].v
            if type(inst.fields[i]) is AutoNewSlot:
                v = inst.fields[i].get()
            return index_value(ex, v, args, named) if args else v
        if inst.form is not None:
            return inst.form.member(ex).vba_get(ex, name, [_unref(a) for a in args], _unref_named(named))
        raise VbaError(438, f"Object doesn't support this property or method: {name}")

    def instance_let(self, ex: Execution, inst: Instance, name: str, args: list, named: Optional[dict], value: Any, is_set: bool) -> None:
        cls = inst.cls
        low = name.lower()
        d = cls.procs.get(low)
        if d is None and cls.implements:
            for iface in cls.implements:
                d = cls.procs.get(f"{iface.lower()}_{low}")
                if d is not None:
                    break
        if d is not None:
            p = d.get("set" if is_set else "let") or d.get("let" if is_set else "set")
            if p is not None:
                call_procedure(ex, p, list(args) + [value], named, inst)
                return
            g = d.get("get") or d.get("function")
            if g is not None:
                obj = call_procedure(ex, g, args, named, inst)
                if is_object(obj) and not is_set:
                    return let_default(ex, obj, [], None, value)
                if type(obj) is VArray and args:
                    return let_index(ex, obj, args, value)
            raise VbaError(450, f"Can't assign to read-only property {name}")
        i = cls.field_index.get(low)
        if i is not None and cls.var_public.get(low, False):
            s = inst.fields[i]
            if args:
                return let_index(ex, s.v, args, value)
            if is_set:
                s.v = value
            else:
                s.let(copy_value(value))
            return
        if inst.form is not None:
            form = inst.form.member(ex)
            a, nm = [_unref(x) for x in args], _unref_named(named)
            return form.vba_set(ex, name, a, nm, value) if is_set else form.vba_let(ex, name, a, nm, value)
        raise VbaError(438, f"Object doesn't support this property or method: {name}")

    def instance_enum(self, ex: Execution, inst: Instance) -> Iterator[Any]:
        name = inst.cls.enum_member
        if name is None:
            raise VbaError(438, f"For Each can't go through a {inst.cls.name}")
        v = self.instance_get(ex, inst, name, [], None)
        return enumerate_object(ex, v)

    def module_get(self, ex: Execution, mod: ModuleRuntime, name: str, args: list, named: Optional[dict]) -> Any:
        low = name.lower()
        p = mod.proc(low)
        if p is not None and p.scope != "private":
            return self.call(ex, p, args, named)
        if low in mod.vars and mod.var_public.get(low):
            v = slot_value(mod.vars[low])
            return index_value(ex, v, args, named) if args else v
        if low in mod.consts and mod.const_public.get(low):
            return mod.consts[low]
        raise VbaError(438, f"Method or data member not found: {name}")

    def module_let(self, ex: Execution, mod: ModuleRuntime, name: str, args: list, named: Optional[dict], value: Any, is_set: bool) -> None:
        low = name.lower()
        d = mod.procs.get(low)
        if d:
            p = d.get("set" if is_set else "let") or d.get("let")
            if p is not None:
                self.call(ex, p, list(args) + [value], named)
                return
        if low in mod.vars and mod.var_public.get(low):
            s = mod.vars[low]
            if args:
                return let_index(ex, s.v, args, value)
            if is_set:
                s.v = value
            else:
                s.let(copy_value(value))
            return
        raise VbaError(438, f"Method or data member not found: {name}")

    def document_member(self, ex: Execution, module_name: str, name: str, args: list, named: Optional[dict], me: Any) -> Tuple[bool, Any]:
        """A public member of a document module (``Sheet1.MyMacro``): (found, value)."""
        mod = self.module(module_name)
        if mod is None:
            return False, None
        low = name.lower()
        p = mod.proc(low)
        if p is not None and p.scope != "private":
            return True, call_procedure(ex, p, args, named, me)
        if low in mod.vars and mod.var_public.get(low):
            v = slot_value(mod.vars[low])
            return True, (index_value(ex, v, args, named) if args else v)
        return False, None


# ---------------------------------------------------------------- compiling a project


def compile_project(modules: Sequence[Tuple[str, str, str, Sequence[Tuple[str, str]]]], host_globals: Optional[Set[str]] = None, host_constants: Optional[Dict[str, Any]] = None, project_name: str = "VBAProject", host_types: Optional[Set[str]] = None, host_enums: Optional[Dict[str, Dict[str, Any]]] = None) -> ProjectRuntime:
    """Compiles ``(name, kind, code, member attributes)`` modules. Syntax and declaration errors
    are collected in ``ProjectRuntime.errors`` (by module); the other modules still run."""
    rt = ProjectRuntime(project_name)
    rt.host_globals = {g.lower() for g in (host_globals or ())}
    rt.host_constants = {k.lower(): v for k, v in (host_constants or {}).items()}
    rt.host_types = {t.lower() for t in (host_types or ())}
    if not _LIB_ENUMS:
        _vba_enums()
    if host_enums:
        register_enums(host_enums)
    asts: Dict[str, A.ModuleAst] = {}
    attrs: Dict[str, Sequence[Tuple[str, str]]] = {}
    for item in modules:
        name, kind, code, member_attrs = item[:4]
        if len(item) > 4 and item[4]:
            rt.doc_kinds[name.lower()] = item[4]
        if len(item) > 5 and item[5] is not None:
            rt.form_designs[name.lower()] = item[5]
        m = ModuleRuntime(name, kind)
        rt.modules[name.lower()] = m
        rt.order.append(m)
        attrs[name.lower()] = member_attrs
        try:
            asts[name.lower()] = parse_module(code, name)
        except VbaSyntaxError as e:
            rt.errors[name.lower()] = e
    # Declarations first (types, enums, constants, variables, procedure signatures), so bodies
    # can use any public name of any module.
    for m in rt.order:
        a = asts.get(m.name.lower())
        if a is None:
            continue
        try:
            _declare(rt, m, a, attrs.get(m.name.lower(), ()))
        except VbaSyntaxError as e:
            e.module = e.module or m.name
            rt.errors[m.name.lower()] = e
    for m in rt.order:
        a = asts.get(m.name.lower())
        if a is None or m.name.lower() in rt.errors:
            continue
        try:
            _module_vars(rt, m, a)
            for p in a.procs:
                ProcCompiler(rt, m, a, p).compile()
        except VbaSyntaxError as e:
            e.module = e.module or m.name
            rt.errors[m.name.lower()] = e
    for m in rt.order:
        if m.kind == "class":
            m.terminate = m.proc("Class_Terminate", "sub")
        elif m.kind == "form":
            m.terminate = m.proc("UserForm_Terminate", "sub")
    return rt


def _declare(rt: ProjectRuntime, m: ModuleRuntime, a: A.ModuleAst, member_attrs: Sequence[Tuple[str, str]]) -> None:
    m.option_base = a.option_base
    m.option_explicit = a.option_explicit
    m.compare_text = a.option_compare_text
    m.private_module = a.option_private
    m.implements = list(a.implements)
    lines = [ln for _, ln in member_attrs] + list(a.attributes)
    for ln in lines:
        import re

        mm = re.match(r"^\s*Attribute\s+([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*VB_UserMemId\s*=\s*(-\s*)?(\d+)", ln, re.I)
        if mm:
            num = int(mm.group(3)) * (-1 if mm.group(2) else 1)
            if num == 0:
                m.default_member = mm.group(1).lower()
            elif num == -4:
                m.enum_member = mm.group(1).lower()
    for t in a.types:
        ud = UdtDef(t.name)
        m.types[t.name.lower()] = ud
        if t.scope != "private":
            rt.udts.setdefault(t.name.lower(), ud)
    for e in a.enums:
        vals: Dict[str, int] = {}
        nxt = 0
        cc = _ConstCompiler(rt, m)
        for name, ex in e.members:
            if ex is not None:
                nxt = V.to_long(cc.value(ex, extra=vals))
            vals[name.lower()] = nxt
            nxt += 1
        m.enums[e.name.lower()] = vals
        m.enum_public[e.name.lower()] = e.scope != "private"
    for d in a.declares:
        m.declares[d.name.lower()] = d
    for ev in a.events:
        m.events[ev.name.lower()] = ev
    # UDT fields (after all types exist, so fields may use other types).
    for t in a.types:
        ud = m.types[t.name.lower()]
        for f in t.fields:
            ud.index[f.name.lower()] = len(ud.names)
            ud.names.append(f.name)
            co, init, tname = _var_spec(rt, m, f)
            ud.coerces.append(co)
            ud.inits.append(init)
            ud.types.append(tname)
    for p in a.procs:
        d = m.procs.setdefault(p.name.lower(), {})
        if p.kind in d:
            raise VbaSyntaxError(f"Ambiguous name detected: {p.name}", p.line, m.name)
        proc = Procedure(m, p.name, p.kind, p.scope, p.line)
        proc.member_of_class = m.kind in ("class", "form")
        d[p.kind] = proc
    for c in a.constants:
        m.consts[c.decl.name.lower()] = _Pending(c.decl)
        m.const_public[c.decl.name.lower()] = c.scope == "public"
    # Constants may refer to each other in any order.
    for c in a.constants:
        _const_value(rt, m, c.decl.name.lower())
    for v in a.variables:
        low = v.decl.name.lower()
        m.var_public[low] = v.scope == "public"
        m.var_types[low] = _decl_type(m, a, v.decl)


class _Pending:
    def __init__(self, decl: A.ConstDecl):
        self.decl = decl
        self.busy = False


def _const_value(rt: ProjectRuntime, m: ModuleRuntime, low: str) -> Any:
    v = m.consts.get(low)
    if not isinstance(v, _Pending):
        return v
    if v.busy:
        raise VbaSyntaxError(f"Circular dependency between constants: {v.decl.name}", v.decl.line, m.name)
    v.busy = True
    val = _ConstCompiler(rt, m).value(v.decl.expr)
    tname = _norm_type(v.decl.type.name) if v.decl.type.explicit else _const_type(val)
    co = V.COERCE.get(tname)
    if co is not None:
        try:
            val = co(val)
        except VbaError as e:
            raise VbaSyntaxError(f"{e.description} in constant {v.decl.name}", v.decl.line, m.name)
    m.consts[low] = val
    m.const_types[low] = tname
    return val


def _const_type(v: Any) -> str:
    n = V.type_name(v)
    return n if n in _NUM_STATIC or n in ("String", "Date", "Boolean") else "Variant"


def _decl_type(m: ModuleRuntime, a: A.ModuleAst, d: A.VarDecl) -> str:
    if d.type.explicit:
        return _norm_type(d.type.name)
    if d.suffix:
        return {"%": "Integer", "&": "Long", "^": "LongLong", "!": "Single", "#": "Double", "@": "Currency", "$": "String"}[d.suffix]
    first = d.name[:1].upper()
    for lo, hi, t in a.deftypes:
        if lo <= first <= hi:
            return t
    return "Variant"


def _var_spec(rt: ProjectRuntime, m: ModuleRuntime, d: A.VarDecl, tname: Optional[str] = None) -> Tuple[Optional[Callable], Callable[[], Any], str]:
    """(coercion, a function making the initial value, type name) of a declared variable."""
    tname = tname or (_norm_type(d.type.name) if d.type.explicit else ({"%": "Integer", "&": "Long", "^": "LongLong", "!": "Single", "#": "Double", "@": "Currency", "$": "String"}.get(d.suffix, "Variant")))
    co: Optional[Callable] = None
    fixed_len = d.type.string_len
    if tname.startswith("String*"):
        const = tname.split("*", 1)[1]
        fixed_len = V.to_long(_ConstCompiler(rt, m).value(A.Ident(const)))
        tname = "String"
    udt = _udt(rt, m, tname)
    is_obj = _is_object_type(rt, m, tname)
    if fixed_len:
        co = V.fixed_string(fixed_len)
        base: Callable[[], Any] = (lambda n=fixed_len: " " * n) if fixed_len > 0 else (lambda: "")
    elif udt is not None:
        co = None
        base = udt.new
    elif is_obj:
        co = None
        base = lambda: NOTHING
    else:
        co = V.COERCE.get(tname)
        dv = V.DEFAULTS.get(tname, EMPTY)
        base = lambda dv=dv: dv
    if d.is_array:
        etype = tname if not is_obj and udt is None else ("Object" if is_obj else tname)
        if d.bounds:
            cc = _ConstCompiler(rt, m)
            bounds = []
            for lo, hi in d.bounds:
                lb = V.to_long(cc.value(lo)) if lo is not None else m.option_base
                bounds.append((lb, V.to_long(cc.value(hi))))
            elem_init = base
            elem_co = co

            def make_fixed(bounds=bounds, elem_init=elem_init, elem_co=elem_co, etype=etype):
                return VArray(bounds, None, elem_co, etype, True, elem_init if (udt is not None or fixed_len) else elem_init())

            return None, make_fixed, tname
        elem_co2 = co

        def make_dynamic(elem_co2=elem_co2, etype=etype):
            return VArray([], [], elem_co2, etype, False)

        return None, make_dynamic, tname
    return co, base, tname


def _udt(rt: ProjectRuntime, m: ModuleRuntime, tname: str) -> Optional[UdtDef]:
    low = tname.lower()
    return m.types.get(low) or rt.udts.get(low)


def _is_object_type(rt: ProjectRuntime, m: ModuleRuntime, tname: str) -> bool:
    if tname.upper() in VALUE_TYPES or tname.startswith("String*"):
        return False
    if _udt(rt, m, tname) is not None:
        return False
    if tname.lower() in m.enums or any(tname.lower() in mm.enums for mm in rt.order):
        return False
    return True


def _module_vars(rt: ProjectRuntime, m: ModuleRuntime, a: A.ModuleAst) -> None:
    for v in a.variables:
        d = v.decl
        low = d.name.lower()
        tname = m.var_types[low]
        co, init, tname = _var_spec(rt, m, d, tname)
        is_obj = _is_object_type(rt, m, tname) and not d.is_array
        if getattr(d, "withevents", False):
            m.has_withevents = True

            def mk(name=d.name):
                return WithEventsSlot(name)
        elif d.type.new:
            cls_name = tname

            def mk(cls_name=cls_name, co=co):
                return AutoNewSlot(lambda: _new_object(rt, None, cls_name, m), None)
        else:
            def mk(init=init, co=co):
                return Slot(init(), co)

        if m.kind in ("class", "form"):
            m.field_index[low] = len(m.field_inits)
            m.field_inits.append(mk)
            m.field_types[low] = tname
        else:
            s = mk()
            if type(s) is WithEventsSlot:
                s.owner = m
            m.vars[low] = s
            m.var_inits[low] = mk

            def reset(s=s, mk=mk):
                fresh = mk()
                s.v = slot_value(fresh) if type(fresh) is not AutoNewSlot else NOTHING

            rt._resetters.append(reset)
        m.var_types[low] = tname


def _new_object(rt: ProjectRuntime, ex: Optional[Execution], tname: str, m: Optional[ModuleRuntime]) -> Any:
    from .runtime import current

    ex = ex or current()
    low = tname.lower()
    short = low.rsplit(".", 1)[-1]
    cls = rt.module(short)
    if cls is not None and cls.kind in ("class", "form"):
        return rt.new_instance(ex, cls)
    from .objects import new_library_object

    obj = new_library_object(ex, short)
    if obj is not None:
        return obj
    raise VbaSyntaxError(f"User-defined type not defined: {tname}")


# ---------------------------------------------------------------- constants


class _ConstCompiler:
    """Evaluates constant expressions (Const, Enum values, array bounds)."""

    def __init__(self, rt: ProjectRuntime, m: ModuleRuntime, locals_: Optional[Dict[str, Any]] = None):
        self.rt = rt
        self.m = m
        self.locals = locals_ or {}

    def value(self, e: A.Expr, extra: Optional[Dict[str, int]] = None) -> Any:
        try:
            return self._v(e, extra or {})
        except VbaError as x:
            raise VbaSyntaxError(f"Constant expression: {x.description}", getattr(e, "line", 0), self.m.name)

    def _v(self, e: A.Expr, extra: Dict[str, int]) -> Any:
        if isinstance(e, A.Lit):
            return _lit_value(e)
        if isinstance(e, A.Paren):
            return self._v(e.expr, extra)
        if isinstance(e, A.Ident):
            low = e.name.lower()
            if low in extra:
                return extra[low]
            if low in self.locals:
                return self.locals[low]
            if low in self.m.consts:
                return _const_value(self.rt, self.m, low)
            for mm in self.rt.order:
                if low in mm.consts and mm.const_public.get(low):
                    return _const_value(self.rt, mm, low)
            for mm in [self.m] + self.rt.order:
                for en in mm.enums.values():
                    if low in en:
                        return en[low]
            from .stdlib import CONSTANTS

            if low in CONSTANTS:
                return CONSTANTS[low]
            if low in self.rt.host_constants:
                return self.rt.host_constants[low]
            raise VbaSyntaxError(f"Constant expression required: {e.name}", e.line, self.m.name)
        if isinstance(e, A.Member) and isinstance(e.obj, A.Ident):
            low = e.obj.name.lower()
            for mm in [self.m] + self.rt.order:
                en = mm.enums.get(low)
                if en is not None and e.name.lower() in en:
                    return en[e.name.lower()]
            mod = self.rt.module(low)
            if mod is not None and e.name.lower() in mod.consts:
                return _const_value(self.rt, mod, e.name.lower())
            if low in ("vba", "excel", "office", "xlbuiltins"):
                return self._v(A.Ident(e.name, "", e.line), extra)
            raise VbaSyntaxError("Constant expression required", e.line, self.m.name)
        if isinstance(e, A.Unary):
            x = self._v(e.expr, extra)
            if e.op == "-":
                return V.neg(x, True)
            if e.op == "+":
                return x
            return V.op_not(x)
        if isinstance(e, A.Binary):
            a, b = self._v(e.left, extra), self._v(e.right, extra)
            f = _BINOPS.get(e.op)
            if f is None:
                raise VbaSyntaxError("Constant expression required", e.line, self.m.name)
            return f(a, b)
        if isinstance(e, A.Call) and isinstance(e.callee, A.Ident):
            from .stdlib import lookup

            b = lookup(e.callee.name)
            if b is not None and b.const_ok:
                args = [self._v(x.expr, extra) if x.expr is not None else MISSING for x in e.args]
                return b.fn(None, args, self.m)
        raise VbaSyntaxError("Constant expression required", getattr(e, "line", 0), self.m.name)


def _lit_value(e: A.Lit) -> Any:
    t = e.vtype
    if t == "Empty":
        return EMPTY
    if t == "Null":
        return NULL
    if t == "Nothing":
        return NOTHING
    return e.value


_BINOPS: Dict[str, Callable] = {
    "+": V.add,
    "-": V.sub,
    "*": V.mul,
    "/": V.div,
    "\\": V.intdiv,
    "MOD": V.mod,
    "^": V.power,
    "&": V.concat,
    "AND": V.op_and,
    "OR": V.op_or,
    "XOR": V.op_xor,
    "EQV": V.op_eqv,
    "IMP": V.op_imp,
    "=": lambda a, b: V.eq(a, b),
    "<>": lambda a, b: V.ne(a, b),
    "<": lambda a, b: V.lt(a, b),
    ">": lambda a, b: V.gt(a, b),
    "<=": lambda a, b: V.le(a, b),
    ">=": lambda a, b: V.ge(a, b),
}


# ---------------------------------------------------------------- procedures


class _Unit:
    __slots__ = ("start",)


class ProcCompiler:
    def __init__(self, rt: ProjectRuntime, m: ModuleRuntime, mast: A.ModuleAst, p: A.Proc):
        self.rt = rt
        self.m = m
        self.mast = mast
        self.p = p
        self.proc = m.procs[p.name.lower()][p.kind]
        self.locals: Dict[str, LocalInfo] = {}
        self.code: List[Callable] = []
        self.lines: List[int] = []
        self.next_stmt: List[int] = []
        self.erl: List[int] = []
        self.cur_erl = 0
        self.labels: Dict[str, List[Optional[int]]] = {}
        self.label_used: Dict[str, int] = {}
        self.loops: List[Tuple[str, List[Optional[int]]]] = []
        self.withs: List[int] = []
        self.nslots = 0
        self.line = p.line
        self.implicit_ok = not m.option_explicit
        self.init: List[Tuple[int, Callable[[], Slot]]] = []
        self.statics: Dict[int, Slot] = {}

    # ------------------------------------------------------------ helpers

    def error(self, msg: str, line: Optional[int] = None) -> None:
        raise VbaSyntaxError(msg, line or self.line, self.m.name)

    def new_slot(self) -> int:
        i = self.nslots
        self.nslots += 1
        return i

    def hidden(self) -> int:
        i = self.new_slot()
        self.init.append((i, lambda: Slot(EMPTY, None)))
        return i

    def emit(self, f: Callable, line: int) -> int:
        self.code.append(f)
        self.lines.append(line)
        self.next_stmt.append(-1)
        self.erl.append(self.cur_erl)
        return len(self.code) - 1

    @contextmanager
    def unit(self) -> Iterator[None]:
        start = len(self.code)
        yield
        end = len(self.code)
        for i in range(start, end):
            if self.next_stmt[i] < 0:
                self.next_stmt[i] = end

    def simple(self, action: Callable[[Frame], Any], line: int) -> None:
        with self.unit():
            nxt = len(self.code) + 1

            def ins(fr: Frame, action=action, nxt=nxt) -> int:
                action(fr)
                return nxt

            self.emit(ins, line)

    def jump(self, target: List[Optional[int]], line: int) -> int:
        def ins(fr: Frame, target=target) -> int:
            return target[0]  # type: ignore[return-value]

        return self.emit(ins, line)

    def label_cell(self, name: str) -> List[Optional[int]]:
        return self.labels.setdefault(name.lower(), [None])

    # ------------------------------------------------------------ names

    def declare_local(self, name: str, info: LocalInfo) -> None:
        low = name.lower()
        if low in self.locals:
            self.error(f"Duplicate declaration in current scope: {name}")
        self.locals[low] = info

    def resolve(self, name: str, line: int, implicit: bool = True) -> B:
        low = name.lower()
        li = self.locals.get(low)
        if li is not None:
            if li.const:
                return B("const", li.value, type_=li.type)
            return B("local", li, type_=li.type)
        m = self.m
        if low in m.consts:
            return B("const", _const_value(self.rt, m, low), type_=m.const_types.get(low, "Variant"))
        if m.kind in ("class", "form") and low in m.field_index:
            return B("field", m.field_index[low], type_=m.field_types.get(low, "Variant"))
        if low in m.vars:
            return B("mvar", m.vars[low], type_=m.var_types.get(low, "Variant"))
        if low in m.procs:
            return B("proc", m, low, self._proc_type(m, low))
        for en in m.enums.values():
            if low in en:
                return B("const", en[low], type_="Long")
        if low in m.enums:
            return B("enumtype", m.enums[low])
        if low in m.declares:
            return B("declare", m.declares[low])
        # In ThisWorkbook and sheet modules, the object's own members come first: an
        # unqualified Range in a sheet module is that sheet's.
        if m.kind == "document" and low in _DOCUMENT_MEMBERS.get(self.rt.doc_kinds.get(m.name.lower(), "worksheet"), ()):
            return B("memberofme", name)
        if m.kind == "form" and (low in _FORM_MEMBERS or self._form_control(low)):
            return B("memberofme", name)
        mod = self.rt.module(low)
        if mod is not None:
            return B("module", mod)
        for mm in self.rt.order:
            if mm is m or mm.kind != "standard":
                continue
            if low in mm.consts and mm.const_public.get(low):
                return B("const", _const_value(self.rt, mm, low), type_=mm.const_types.get(low, "Variant"))
            if low in mm.vars and mm.var_public.get(low):
                return B("mvar", mm.vars[low], type_=mm.var_types.get(low, "Variant"))
            d = mm.procs.get(low)
            if d and any(p.scope != "private" for p in d.values()):
                return B("proc", mm, low, self._proc_type(mm, low))
            for en_name, en in mm.enums.items():
                if low in en and mm.enum_public.get(en_name, True):
                    return B("const", en[low], type_="Long")
            if low in mm.enums and mm.enum_public.get(low, True):
                return B("enumtype", mm.enums[low])
            if low in mm.declares and mm.declares[low].scope != "private":
                return B("declare", mm.declares[low])
        if low == "err":
            return B("err")
        if low == "debug":
            return B("debug")
        from .stdlib import CONSTANTS, lookup

        if low in CONSTANTS:
            v = CONSTANTS[low]
            return B("const", v, type_=_const_type(v))
        bi = lookup(name)
        if bi is not None:
            return B("builtin", bi, type_=bi.ret)
        if low in ("vba", "excel", "office", "stdole", "vbaproject"):
            return B("library", low)
        if low in self.rt.host_constants:
            v = self.rt.host_constants[low]
            return B("const", v, type_=_const_type(v))
        if low in self.rt.host_globals:
            return B("host", name)
        if low in _LIB_ENUMS:
            return B("enumtype", _LIB_ENUMS[low])
        if self.implicit_ok and implicit:
            idx = self.new_slot()
            info = LocalInfo(idx)
            self.locals[low] = info
            self.init.append((idx, lambda: Slot(EMPTY, None)))
            return B("local", info)
        if not implicit:
            self.error(f"Sub or Function not defined: {name}", line)
        self.error(f"Variable not defined: {name}", line)
        return B("const", EMPTY)

    def _form_control(self, low: str) -> bool:
        design = self.rt.form_designs.get(self.m.name.lower())
        return design is not None and design.find(low) is not None

    def _proc_type(self, m: ModuleRuntime, low: str) -> str:
        d = m.procs.get(low, {})
        for k in ("function", "get"):
            p = d.get(k)
            if p is not None:
                return getattr(p, "ret_type_decl", "Variant")
        return "Variant"

    def is_own_function(self, low: str) -> bool:
        return self.p.kind in ("function", "get") and low == self.p.name.lower()

    # ------------------------------------------------------------ procedure

    def compile(self) -> None:
        p, proc = self.p, self.proc
        proc.ret_type = _norm_type(p.ret.name) if p.ret.explicit else "Variant"
        # Parameters.
        for prm in p.params:
            idx = self.new_slot()
            tname = _norm_type(prm.type.name) if prm.type.explicit else "Variant"
            is_obj = _is_object_type(self.rt, self.m, tname) and not prm.is_array
            co = None if (prm.is_array or is_obj or _udt(self.rt, self.m, tname) is not None) else V.COERCE.get(tname)
            default = None
            if prm.default is not None:
                default = _ConstCompiler(self.rt, self.m).value(prm.default)
            spec = ParamSpec(prm.name, prm.byval, prm.optional, default, prm.paramarray, co, tname, prm.is_array or prm.paramarray, is_obj)
            proc.params.append(spec)
            if prm.paramarray:
                proc.has_paramarray = True
            info = LocalInfo(idx, tname if not prm.paramarray else "Variant", prm.is_array or prm.paramarray, is_obj)
            info.udt = _udt(self.rt, self.m, tname)
            self.declare_local(prm.name, info)
        # The return value.
        if p.kind in ("function", "get"):
            idx = self.new_slot()
            proc.ret_slot = idx
            tname = proc.ret_type
            is_obj = _is_object_type(self.rt, self.m, tname) and not p.ret_is_array
            udt = _udt(self.rt, self.m, tname)
            if p.ret_is_array:
                self.init.append((idx, lambda: Slot(EMPTY, None)))
            elif udt is not None:
                self.init.append((idx, lambda udt=udt: Slot(udt.new(), None)))
            elif is_obj:
                self.init.append((idx, lambda: Slot(NOTHING, None)))
            else:
                co = V.COERCE.get(tname)
                dv = V.DEFAULTS.get(tname, EMPTY)
                self.init.append((idx, lambda co=co, dv=dv: Slot(dv, co)))
            info = LocalInfo(idx, tname, p.ret_is_array, is_obj)
            info.udt = udt
            self.locals[p.name.lower()] = info
            proc.ret_is_object = is_obj
        proc.ret_type_decl = proc.ret_type  # type: ignore[attr-defined]
        self.static_all = p.static
        # Declarations anywhere in the body are procedure-wide (VBA allocates them on entry).
        self.predeclare(p.body)
        self.block(p.body)
        end = len(self.code)
        for cell in self._ends:
            cell[0] = end
        for name, cell in self.labels.items():
            if cell[0] is None:
                self.error(f"Label not defined: {name}", self.label_used.get(name, self.line))
        proc.code = self.code
        proc.lines = self.lines
        proc.next_stmt = [x if x >= 0 else i + 1 for i, x in enumerate(self.next_stmt)]
        proc.erl = self.erl
        proc.nslots = self.nslots
        proc.init = self.init
        proc.statics = self.statics
        proc.labels = {k: v[0] for k, v in self.labels.items()}  # type: ignore[misc]
        del end

    def predeclare(self, body: List[A.Stmt]) -> None:
        for s in _walk(body):
            if isinstance(s, A.Dim):
                for d in s.vars:
                    self.dim_local(d, s.static or self.static_all)
            elif isinstance(s, A.Const):
                for c in s.consts:
                    val = _ConstCompiler(self.rt, self.m, {k: v.value for k, v in self.locals.items() if v.const}).value(c.expr)
                    tname = _norm_type(c.type.name) if c.type.explicit else _const_type(val)
                    co = V.COERCE.get(tname)
                    if co is not None:
                        val = co(val)
                    info = LocalInfo(-1, tname)
                    info.const = True
                    info.value = val
                    self.declare_local(c.name, info)

    def dim_local(self, d: A.VarDecl, static: bool) -> None:
        idx = self.new_slot()
        tname = _decl_type(self.m, self.mast, d)
        co, init, tname = _var_spec(self.rt, self.m, d, tname)
        is_obj = _is_object_type(self.rt, self.m, tname) and not d.is_array
        info = LocalInfo(idx, tname, d.is_array, is_obj, d.type.new)
        info.udt = _udt(self.rt, self.m, tname) if not d.is_array else None
        self.declare_local(d.name, info)
        rt, m = self.rt, self.m
        if d.type.new:
            cls_name = tname

            def mk(cls_name=cls_name):
                return AutoNewSlot(lambda: _new_object(rt, None, cls_name, m), None)
        else:
            def mk(init=init, co=co):
                return Slot(init(), co)
        if static:
            self.statics[idx] = mk()
        else:
            self.init.append((idx, mk))

    # ------------------------------------------------------------ statements

    def block(self, body: List[A.Stmt]) -> None:
        for s in body:
            self.line = s.line or self.line
            self.stmt(s)

    def stmt(self, s: A.Stmt) -> None:
        h = getattr(self, "s_" + type(s).__name__, None)
        if h is None:
            self.error(f"Unsupported statement {type(s).__name__}", s.line)
        h(s)

    def s_NoOp(self, s: A.NoOp) -> None:
        pass

    def s_Dim(self, s: A.Dim) -> None:
        pass  # declared on entry

    def s_Const(self, s: A.Const) -> None:
        pass

    def s_Label(self, s: A.Label) -> None:
        cell = self.label_cell(s.name)
        if cell[0] is not None:
            self.error(f"Duplicate label: {s.name}", s.line)
        cell[0] = len(self.code)
        if s.number is not None:
            # Line numbers on lines without code (Dim…) share the next line's position; Erl
            # reports the first of them, as VBA does.
            if getattr(self, "_erl_pc", -1) != len(self.code):
                self.cur_erl = s.number
                self._erl_pc = len(self.code)

    def s_Assign(self, s: A.Assign) -> None:
        f = self.assignment(s.target, s.value, s.kind, s.line)
        self.simple(f, s.line)

    def s_MidAssign(self, s: A.MidAssign) -> None:
        get_ref = self.ref(s.target)
        if get_ref is None:
            self.error("Expected: variable", s.line)
        start, _ = self.expr(s.start)
        length = self.expr(s.length)[0] if s.length is not None else None
        value, _ = self.expr(s.value)

        def act(fr: Frame) -> None:
            slot = get_ref(fr)
            cur = V.to_string(slot_value(slot))
            st = V.to_long(start(fr))
            if st < 1:
                raise VbaError(5)
            new = V.to_string(value(fr))
            n = len(new) if length is None else min(V.to_long(length(fr)), len(new))
            n = min(n, max(len(cur) - st + 1, 0))
            if n < 0:
                raise VbaError(5)
            res = cur[:st - 1] + new[:n] + cur[st - 1 + n:]
            slot.v = slot.co(res) if slot.co else res

        self.simple(act, s.line)

    def s_CallStmt(self, s: A.CallStmt) -> None:
        f = self.call(s.callee, s.args, s.line, want_value=False)
        self.simple(f, s.line)

    def s_If(self, s: A.If) -> None:
        end: List[Optional[int]] = [None]
        for cond_e, body in s.branches:
            cond, _ = self.expr(cond_e)
            skip: List[Optional[int]] = [None]
            with self.unit():
                nxt = len(self.code) + 1

                def test(fr: Frame, cond=cond, nxt=nxt, skip=skip) -> int:
                    v = cond(fr)
                    if v is True or (v is not False and is_true(v)):
                        return nxt
                    return skip[0]  # type: ignore[return-value]

                self.emit(test, cond_e.line or s.line)
            self.block(body)
            self.jump(end, s.line)
            skip[0] = len(self.code)
        if s.orelse:
            self.block(s.orelse)
        end[0] = len(self.code)

    def s_Select(self, s: A.Select) -> None:
        subj_f, subj_t = self.expr(s.expr)
        h = self.hidden()
        end: List[Optional[int]] = [None]

        def store(fr: Frame, subj_f=subj_f, h=h) -> None:
            v = subj_f(fr)
            fr.slots[h].v = V.value_of(v) if is_object(v) else v

        self.simple(store, s.line)
        text = self.m.compare_text
        sv = subj_t == "Variant"
        for clauses, body, cline in s.cases:
            tests = []
            for c in clauses:
                af, at = self.expr(c.a)
                av = at == "Variant"
                if c.kind == "eq":
                    tests.append(lambda fr, af=af, av=av: _cmp(fr.slots[h].v, af(fr), text, sv, av) == 0)
                elif c.kind == "range":
                    bf, bt = self.expr(c.b)
                    bv = bt == "Variant"
                    tests.append(lambda fr, af=af, bf=bf, av=av, bv=bv: _cmp(fr.slots[h].v, af(fr), text, sv, av) >= 0 and _cmp(fr.slots[h].v, bf(fr), text, sv, bv) <= 0)
                else:
                    op = c.op
                    tests.append(lambda fr, af=af, op=op, av=av: _relop(op, _cmp(fr.slots[h].v, af(fr), text, sv, av)))
            skip: List[Optional[int]] = [None]
            with self.unit():
                nxt = len(self.code) + 1

                def test(fr: Frame, tests=tests, nxt=nxt, skip=skip) -> int:
                    for t in tests:
                        if t(fr):
                            return nxt
                    return skip[0]  # type: ignore[return-value]

                self.emit(test, cline)
            self.block(body)
            self.jump(end, cline)
            skip[0] = len(self.code)
        if s.orelse:
            self.block(s.orelse)
        end[0] = len(self.code)

    def s_For(self, s: A.For) -> None:
        get_ref = self.ref(s.var)
        if get_ref is None:
            self.error("Invalid Next control variable reference", s.line)
        var_t = self.static_type(s.var)
        start_f, start_t = self.expr(s.start)
        end_f, end_t = self.expr(s.end)
        step_f, step_t = self.expr(s.step) if s.step is not None else ((lambda fr: V.Integer(1)), "Integer")
        h_end, h_step = self.hidden(), self.hidden()
        exit_: List[Optional[int]] = [None]
        test_at: List[Optional[int]] = [None]
        variant_var = var_t == "Variant"

        def init(fr: Frame) -> None:
            st = start_f(fr)
            en = end_f(fr)
            sp = step_f(fr)
            st, en, sp = _num(st), _num(en), _num(sp)
            slot = get_ref(fr)
            if variant_var:
                # A Variant counter takes the widest type of start, end and step.
                st = _widest(st, en, sp)
            slot.v = slot.co(st) if slot.co else st
            if slot.co is V.to_long:
                if type(en) in (V.Integer, V.Byte):
                    en = int(en)
                if type(sp) in (V.Integer, V.Byte):
                    sp = int(sp)
            fr.slots[h_end].v = en
            fr.slots[h_step].v = sp

        self.simple(init, s.line)
        with self.unit():
            test_pos = len(self.code)
            test_at[0] = test_pos
            nxt = test_pos + 1

            def test(fr: Frame) -> int:
                slot = get_ref(fr)
                v = slot.v
                en = fr.slots[h_end].v
                sp = fr.slots[h_step].v
                if type(v) is int and type(en) is int and type(sp) in (int, V.Integer):
                    if (sp >= 0 and v > en) or (sp < 0 and v < en):
                        return exit_[0]  # type: ignore[return-value]
                    return nxt
                if sp is EMPTY:
                    raise VbaError(92)
                c = V.compare(v, en)
                if (V.compare(sp, 0) >= 0 and c > 0) or (V.compare(sp, 0) < 0 and c < 0):
                    return exit_[0]  # type: ignore[return-value]
                return nxt

            self.emit(test, s.line)
        self.loops.append(("for", exit_))
        self.block(s.body)
        self.loops.pop()
        typed = not variant_var
        with self.unit():

            def nxt_(fr: Frame) -> int:
                fr.ex.tick()
                slot = get_ref(fr)
                v = slot.v
                sp = fr.slots[h_step].v
                if sp is EMPTY:
                    raise VbaError(92)
                if type(v) is int and type(sp) is int:
                    r = v + sp
                    if not -2147483648 <= r <= 2147483647:
                        raise VbaError(6)
                    slot.v = r
                else:
                    r = V.add(v, sp, not typed)
                    slot.v = slot.co(r) if slot.co else r
                return test_at[0]  # type: ignore[return-value]

            self.emit(nxt_, s.end_line or s.line)
        exit_[0] = len(self.code)

    def s_ForEach(self, s: A.ForEach) -> None:
        get_ref = self.ref(s.var)
        if get_ref is None:
            self.error("Invalid Next control variable reference", s.line)
        coll_f, _ = self.expr(s.coll)
        h_it = self.hidden()
        h_arr = self.hidden()
        exit_: List[Optional[int]] = [None]
        top: List[Optional[int]] = [None]

        def init(fr: Frame) -> None:
            c = coll_f(fr)
            if type(c) is VArray:
                c.locked += 1
                fr.slots[h_arr].v = c
            fr.slots[h_it].v = enumerate_object(fr.ex, c)

        self.simple(init, s.line)
        with self.unit():
            pos = len(self.code)
            top[0] = pos
            nxt = pos + 1

            def step(fr: Frame) -> int:
                it = fr.slots[h_it].v
                try:
                    item = next(it)
                except StopIteration:
                    arr = fr.slots[h_arr].v
                    if type(arr) is VArray:
                        arr.locked -= 1
                        fr.slots[h_arr].v = EMPTY
                    fr.slots[h_it].v = EMPTY
                    return exit_[0]  # type: ignore[return-value]
                slot = get_ref(fr)
                slot.v = slot.co(item) if slot.co and not is_object(item) else item
                return nxt

            self.emit(step, s.line)
        self.loops.append(("for", exit_))
        self.block(s.body)
        self.loops.pop()
        with self.unit():

            def back(fr: Frame) -> int:
                fr.ex.tick()
                return top[0]  # type: ignore[return-value]

            self.emit(back, s.end_line or s.line)
        exit_[0] = len(self.code)

    def s_Do(self, s: A.Do) -> None:
        exit_: List[Optional[int]] = [None]
        top = len(self.code)
        if s.pre is not None:
            kind, ce = s.pre
            cond, _ = self.expr(ce)
            until = kind == "until"
            with self.unit():
                nxt = len(self.code) + 1

                def pre(fr: Frame) -> int:
                    c = is_true(cond(fr))
                    if c == until:
                        return exit_[0]  # type: ignore[return-value]
                    return nxt

                self.emit(pre, s.line)
        self.loops.append(("do", exit_))
        self.block(s.body)
        self.loops.pop()
        with self.unit():
            if s.post is not None:
                kind, ce = s.post
                cond, _ = self.expr(ce)
                until = kind == "until"
                nxt = len(self.code) + 1

                def post(fr: Frame) -> int:
                    fr.ex.tick()
                    c = is_true(cond(fr))
                    if c == until:
                        return nxt
                    return top

                self.emit(post, s.end_line or s.line)
            else:

                def loop(fr: Frame) -> int:
                    fr.ex.tick()
                    return top

                self.emit(loop, s.end_line or s.line)
        exit_[0] = len(self.code)

    def s_While(self, s: A.While) -> None:
        exit_: List[Optional[int]] = [None]
        top = len(self.code)
        cond, _ = self.expr(s.cond)
        with self.unit():
            nxt = len(self.code) + 1

            def pre(fr: Frame) -> int:
                if is_true(cond(fr)):
                    return nxt
                return exit_[0]  # type: ignore[return-value]

            self.emit(pre, s.line)
        self.block(s.body)
        with self.unit():

            def loop(fr: Frame) -> int:
                fr.ex.tick()
                return top

            self.emit(loop, s.end_line or s.line)
        exit_[0] = len(self.code)

    def s_With(self, s: A.With) -> None:
        f, _ = self.expr(s.expr)
        h = self.hidden()

        def enter(fr: Frame) -> None:
            v = f(fr)
            fr.slots[h].v = v

        self.simple(enter, s.line)
        self.withs.append(h)
        self.block(s.body)
        self.withs.pop()

    def s_Exit(self, s: A.Exit) -> None:
        if s.what in ("sub", "function", "property"):
            want = {"sub": "sub", "function": "function"}.get(s.what)
            if s.what == "property" and self.p.kind not in ("get", "let", "set"):
                self.error("Exit Property not allowed in Sub or Function", s.line)
            if want is not None and self.p.kind != want:
                self.error(f"Exit {s.what.title()} not allowed in {'Sub' if self.p.kind == 'sub' else 'Function' if self.p.kind == 'function' else 'Property'}", s.line)
            cell: List[Optional[int]] = [None]
            self._ends.append(cell)
            with self.unit():
                self.jump(cell, s.line)
            return
        for kind, cell2 in reversed(self.loops):
            if kind == s.what:
                with self.unit():
                    self.jump(cell2, s.line)
                return
        self.error(f"Exit {s.what.title()} not within {s.what.title()}", s.line)

    @property
    def _ends(self) -> List[List[Optional[int]]]:
        if not hasattr(self, "_end_cells"):
            self._end_cells: List[List[Optional[int]]] = []
        return self._end_cells

    def s_EndStmt(self, s: A.EndStmt) -> None:
        rt = self.rt

        def act(fr: Frame) -> None:
            raise MacroStopped("end")

        self.simple(act, s.line)

    def s_StopStmt(self, s: A.StopStmt) -> None:
        def act(fr: Frame) -> None:
            raise MacroStopped("stop", f"Stop statement in {self.m.name}.{self.p.name}, line {s.line}")

        self.simple(act, s.line)

    def s_GoTo(self, s: A.GoTo) -> None:
        cell = self.label_cell(s.label)
        self.label_used.setdefault(s.label.lower(), s.line)
        if s.gosub:
            with self.unit():
                ret = len(self.code) + 1

                def gosub(fr: Frame, cell=cell, ret=ret) -> int:
                    if len(fr.gosub) > 1000:
                        raise VbaError(28)
                    fr.gosub.append(ret)
                    return cell[0]  # type: ignore[return-value]

                self.emit(gosub, s.line)
            return
        with self.unit():

            def goto(fr: Frame, cell=cell) -> int:
                fr.ex.tick()
                return cell[0]  # type: ignore[return-value]

            self.emit(goto, s.line)

    def s_ReturnStmt(self, s: A.ReturnStmt) -> None:
        with self.unit():

            def ret(fr: Frame) -> int:
                if not fr.gosub:
                    raise VbaError(3)
                return fr.gosub.pop()

            self.emit(ret, s.line)

    def s_OnGoTo(self, s: A.OnGoTo) -> None:
        f, _ = self.expr(s.expr)
        cells = [self.label_cell(l) for l in s.labels]
        for l in s.labels:
            self.label_used.setdefault(l.lower(), s.line)
        gosub = s.gosub
        with self.unit():
            nxt = len(self.code) + 1

            def on(fr: Frame) -> int:
                i = V.to_long(f(fr))
                if i < 0 or i > 255:
                    raise VbaError(5)
                if i == 0 or i > len(cells):
                    return nxt
                if gosub:
                    fr.gosub.append(nxt)
                return cells[i - 1][0]  # type: ignore[return-value]

            self.emit(on, s.line)

    def s_OnError(self, s: A.OnError) -> None:
        mode = s.mode
        cell = self.label_cell(s.label) if mode == "goto" else None
        if cell is not None:
            self.label_used.setdefault(s.label.lower(), s.line)

        def act(fr: Frame) -> None:
            fr.ex.err.clear()
            if mode == "resume_next":
                fr.on_error = 1
            elif mode == "goto":
                fr.on_error = 2
                fr.handler = cell[0]  # type: ignore[index]
            elif mode == "goto0":
                fr.on_error = 0
            else:
                fr.active = False

        self.simple(act, s.line)

    def s_Resume(self, s: A.Resume) -> None:
        kind = s.kind
        cell = self.label_cell(s.label) if kind == "label" else None
        if cell is not None:
            self.label_used.setdefault(s.label.lower(), s.line)
        proc = self.proc
        with self.unit():

            def resume(fr: Frame) -> int:
                if not fr.active:
                    raise VbaError(20)
                fr.active = False
                fr.ex.err.clear()
                if kind == "retry":
                    return fr.resume_pc
                if kind == "next":
                    return proc.next_stmt[fr.resume_pc]
                return cell[0]  # type: ignore[index]

            self.emit(resume, s.line)

    def s_ErrorStmt(self, s: A.ErrorStmt) -> None:
        f, _ = self.expr(s.expr)

        def act(fr: Frame) -> None:
            raise VbaError(V.to_long(f(fr)))

        self.simple(act, s.line)

    def s_Randomize(self, s: A.Randomize) -> None:
        f = self.expr(s.expr)[0] if s.expr is not None else None
        rt = self.rt

        def act(fr: Frame) -> None:
            from .stdlib import randomize

            randomize(rt, f(fr) if f is not None else None)

        self.simple(act, s.line)

    def s_LoadUnload(self, s: A.LoadUnload) -> None:
        """``Load UserForm1`` loads the form (referring to it does: UserForm_Initialize runs);
        ``Unload Me`` asks the form (QueryClose) and unloads it."""
        f, _ = self.expr(s.expr)
        unload = s.unload

        def act(fr: Frame) -> None:
            obj = f(fr)
            form = obj.form if type(obj) is Instance else None
            if form is None:
                raise VbaError(361, "Can't load or unload this object")
            if unload:
                form.query_unload(fr.ex, 1)

        self.simple(act, s.line)

    def s_Unsupported(self, s: A.Unsupported) -> None:
        what = s.what

        def act(fr: Frame) -> None:
            raise NotSupported(what)

        self.simple(act, s.line)

    def s_PrintStmt(self, s: A.PrintStmt) -> None:
        items = []
        for e, sep in s.items:
            if e is None:
                items.append((None, sep, None))
                continue
            if isinstance(e, A.Call) and isinstance(e.callee, A.Ident) and e.callee.name.upper() in ("SPC", "TAB"):
                arg = self.expr(e.args[0].expr)[0] if e.args and e.args[0].expr is not None else None
                items.append(("spc" if e.callee.name.upper() == "SPC" else "tab", sep, arg))
                continue
            items.append((self.expr(e)[0], sep, None))

        def act(fr: Frame) -> None:
            from .stdlib import debug_print

            debug_print(fr.ex, [(f if not callable(f) else f(fr), sep, a(fr) if a is not None else None) for f, sep, a in items])

        self.simple(act, s.line)

    def s_RaiseEvent(self, s: A.RaiseEvent) -> None:
        name = s.name
        args = [self.arg_value(a) for a in s.args]

        def act(fr: Frame) -> None:
            inst = fr.inst
            if inst is None:
                return
            vals = [f(fr) for f in args]
            for listener, var in list(inst.listeners):
                handler = None
                owner = listener
                mod = listener.cls if type(listener) is Instance else listener
                d = mod.procs.get(f"{var.lower()}_{name.lower()}") if isinstance(mod, ModuleRuntime) else None
                if d:
                    handler = d.get("sub")
                if handler is not None:
                    call_procedure(fr.ex, handler, vals, None, owner if type(owner) is Instance else None)

        self.simple(act, s.line)

    def s_ReDim(self, s: A.ReDim) -> None:
        preserve = s.preserve
        for target, bounds, tr in s.items:
            if isinstance(target, A.Ident) and target.name.lower() not in self.locals:
                try:
                    self.resolve(target.name, s.line, implicit=False)
                except VbaSyntaxError:
                    # ReDim also declares a dynamic array.
                    idx = self.new_slot()
                    tname = _norm_type(tr.name) if tr is not None else "Variant"
                    info = LocalInfo(idx, tname, True)
                    self.locals[target.name.lower()] = info
                    self.init.append((idx, lambda: Slot(EMPTY, None)))
            get_ref = self.ref(target, for_redim=True)
            if get_ref is None:
                self.error("Expected: variable", s.line)
            bfs = [((self.expr(lo)[0] if lo is not None else None), self.expr(hi)[0]) for lo, hi in bounds]
            base = self.m.option_base
            elem_t = "Variant"
            if tr is not None:
                elem_t = _norm_type(tr.name)
            elif isinstance(target, A.Ident):
                b = self.resolve(target.name, s.line)
                if b.kind == "local" and b.a.is_array:
                    elem_t = b.a.type
                elif b.kind == "mvar":
                    elem_t = self._mvar_elem_type(target.name)
                elif b.kind == "field":
                    elem_t = self.m.field_types.get(target.name.lower(), "Variant")
            udt = _udt(self.rt, self.m, elem_t)
            is_obj = elem_t != "Variant" and _is_object_type(self.rt, self.m, elem_t)
            co = None if (udt is not None or is_obj) else V.COERCE.get(elem_t)
            fill: Any = udt.new if udt is not None else (NOTHING if is_obj else V.DEFAULTS.get(elem_t, EMPTY))
            explicit_type = tr is not None

            def act(fr: Frame, get_ref=get_ref, bfs=bfs, co=co, fill=fill, elem_t=elem_t, explicit_type=explicit_type) -> None:
                slot = get_ref(fr)
                bounds_v = []
                for lo, hi in bfs:
                    lb = V.to_long(lo(fr)) if lo is not None else base
                    ub = V.to_long(hi(fr))
                    if ub < lb - 1:
                        raise VbaError(9)
                    bounds_v.append((lb, ub))
                old = slot.v
                if type(old) is VArray and (old.fixed or old.locked):
                    raise VbaError(10)
                c, et = co, elem_t
                if type(old) is VArray and not explicit_type and old.etype != "Variant":
                    c, et = old.coerce, old.etype
                new = VArray(bounds_v, None, c, et, False, fill)
                if preserve and type(old) is VArray and old.bounds:
                    if len(old.bounds) != len(bounds_v) or old.bounds[:-1] != bounds_v[:-1] or old.bounds[-1][0] != bounds_v[-1][0]:
                        raise VbaError(9)
                    keep = min(len(old.data), len(new.data))
                    new.data[:keep] = old.data[:keep]
                slot.v = new

            self.simple(act, s.line)

    def _mvar_elem_type(self, name: str) -> str:
        low = name.lower()
        for mm in [self.m] + self.rt.order:
            if low in mm.var_types:
                return mm.var_types[low]
        return "Variant"

    def s_Erase(self, s: A.Erase) -> None:
        for t in s.targets:
            get_ref = self.ref(t)
            if get_ref is None:
                self.error("Expected: array", s.line)

            def act(fr: Frame, get_ref=get_ref) -> None:
                slot = get_ref(fr)
                a = slot.v
                if type(a) is not VArray:
                    raise VbaError(13)
                if a.locked:
                    raise VbaError(10)
                if a.fixed:
                    if a.data:
                        proto = V.DEFAULTS.get(a.etype, EMPTY) if a.etype != "Object" else NOTHING
                        if a.data and type(a.data[0]) is V.Record:
                            a.data = [x.typ.new() for x in a.data]
                        else:
                            a.data = [proto] * len(a.data)
                else:
                    slot.v = VArray([], [], a.coerce, a.etype, False)

            self.simple(act, s.line)

    # ------------------------------------------------------------ assignment

    def assignment(self, target: A.Expr, value_e: A.Expr, kind: str, line: int) -> Callable[[Frame], None]:
        is_set = kind == "set"
        vf, vt = self.expr(value_e)
        lvalue_src = isinstance(value_e, (A.Ident, A.Member, A.Call))
        if kind in ("lset", "rset"):
            ref = self.ref(target)
            if ref is None:
                self.error("Expected: variable", line)

            def lset(fr: Frame) -> None:
                slot = ref(fr)
                cur = slot.v
                if type(cur) is V.Record:
                    src = vf(fr)
                    if type(src) is V.Record:
                        slot.v = src.copy()
                        return
                cur_s = V.to_string(cur)
                new = V.to_string(vf(fr))
                n = len(cur_s)
                res = (new[:n] + " " * max(0, n - len(new))) if kind == "lset" else (" " * max(0, n - len(new)) + new[:n])
                slot.v = slot.co(res) if slot.co else res

            return lset
        if isinstance(target, A.Ident):
            low = target.name.lower()
            b = self.resolve(target.name, line)
            if b.kind == "local":
                info: LocalInfo = b.a
                idx = info.index
                if is_set:
                    if not info.is_object and info.type not in ("Variant",):
                        self.error("Object required", line)

                    def set_local(fr: Frame) -> None:
                        v = vf(fr)
                        if v is not NOTHING and not is_object(v):
                            raise VbaError(424)
                        fr.slots[idx].v = v

                    return set_local
                if info.is_object:

                    def let_obj_local(fr: Frame) -> None:
                        cur = fr.slots[idx].v
                        if type(fr.slots[idx]) is AutoNewSlot:
                            cur = fr.slots[idx].get()
                        let_default(fr.ex, cur, [], None, vf(fr))

                    return let_obj_local
                if info.udt is not None or info.is_array:

                    def let_copy(fr: Frame) -> None:
                        v = vf(fr)
                        s = fr.slots[idx]
                        if info.is_array and type(v) is not VArray:
                            raise VbaError(13)
                        if type(v) is VArray and type(s.v) is VArray and s.v.fixed:
                            raise VbaError(10)
                        s.v = copy_value(v) if lvalue_src else v

                    return let_copy

                def let_local(fr: Frame) -> None:
                    v = vf(fr)
                    s = fr.slots[idx]
                    co = s.co
                    if co is not None:
                        s.v = co(v)
                    elif type(v) is VArray or type(v) is V.Record:
                        s.v = copy_value(v) if lvalue_src else v
                    elif v is NOTHING or (type(v) not in _PLAIN and is_object(v)):
                        s.v = V.value_of(v)
                    else:
                        s.v = v

                return let_local
            if b.kind == "mvar":
                slot: Slot = b.a
                if is_set:

                    def set_m(fr: Frame) -> None:
                        v = vf(fr)
                        if v is not NOTHING and not is_object(v):
                            raise VbaError(424)
                        slot.v = v

                    return set_m

                def let_m(fr: Frame) -> None:
                    v = vf(fr)
                    if is_object(v) and v is not NOTHING and not _is_object_slot(slot):
                        v = V.value_of(v)
                    elif type(v) is VArray or type(v) is V.Record:
                        v = copy_value(v) if lvalue_src else v
                    co = slot.co
                    slot.v = co(v) if co is not None else v

                return let_m
            if b.kind == "field":
                fi = b.a

                def let_f(fr: Frame) -> None:
                    v = vf(fr)
                    s = fr.inst.fields[fi]
                    if is_set:
                        if v is not NOTHING and not is_object(v):
                            raise VbaError(424)
                        s.v = v
                        return
                    if is_object(v) and v is not NOTHING and type(s.v) is not Instance and not is_object(s.v):
                        v = V.value_of(v)
                    elif type(v) is VArray or type(v) is V.Record:
                        v = copy_value(v) if lvalue_src else v
                    s.v = s.co(v) if s.co is not None else v

                return let_f
            if b.kind == "proc":
                mod, low2 = b.a, b.b
                d = mod.procs[low2]
                p = d.get("set" if is_set else "let") or d.get("let") or d.get("set")
                if p is None:
                    self.error(f"Can't assign to {target.name}", line)
                me_f = self.me_for(mod)
                rt = self.rt

                def let_prop(fr: Frame) -> None:
                    call_procedure(fr.ex, p, [vf(fr)], None, me_f(fr))

                return let_prop
            if b.kind == "host":
                name = b.a

                def let_host(fr: Frame) -> None:
                    if is_set:
                        fr.ex.host.global_set(fr.ex, name, [], None, vf(fr))
                    else:
                        fr.ex.host.global_let(fr.ex, name, [], None, vf(fr))

                return let_host
            if b.kind == "memberofme":
                name = b.a
                me_f = self.me_for(self.m)

                def let_me(fr: Frame) -> None:
                    (set_member if is_set else let_member)(fr.ex, me_f(fr), name, [], None, vf(fr))

                return let_me
            if b.kind == "const":
                self.error("Assignment to constant not permitted", line)
            self.error(f"Can't assign to {target.name}", line)
        if isinstance(target, (A.Member, A.Bang)):
            name = target.name
            of = self.object_of(target)
            if isinstance(target, A.Bang):

                def let_bang(fr: Frame) -> None:
                    (let_default)(fr.ex, of(fr), [name], None, vf(fr))

                return let_bang
            if of is None:
                mod_b = self._qualified_module(target)
                if mod_b is not None:
                    mod = mod_b

                    def let_q(fr: Frame) -> None:
                        fr.ex.rt.module_let(fr.ex, mod, name, [], None, vf(fr), is_set)

                    return let_q
                self.error("Invalid use of property", line)

            def let_mem(fr: Frame) -> None:
                obj = of(fr)
                v = vf(fr)
                if is_set:
                    set_member(fr.ex, obj, name, [], None, v)
                else:
                    let_member(fr.ex, obj, name, [], None, v)

            return let_mem
        if isinstance(target, A.Call):
            callee = target.callee
            argfs, named = self.args(target.args, refs=False)
            if isinstance(callee, A.Ident):
                b = self.resolve(callee.name, line)
                if b.kind in ("local", "mvar", "field") and not (b.kind == "local" and self.is_own_function(callee.name.lower()) and False):
                    getv = self.ident_value(b, callee.name)

                    def let_elem(fr: Frame) -> None:
                        cur = getv(fr)
                        args = [f(fr) for f in argfs]
                        v = vf(fr)
                        if type(cur) is VArray:
                            if type(v) is VArray or type(v) is V.Record:
                                v = copy_value(v) if lvalue_src else v
                            cur.set([V.to_long(a) for a in args], v)
                        elif is_set:
                            raise VbaError(450)
                        else:
                            let_default(fr.ex, cur, args, None, v)

                    return let_elem
                if b.kind == "proc":
                    mod, low2 = b.a, b.b
                    d = mod.procs[low2]
                    p = d.get("set" if is_set else "let") or d.get("let") or d.get("set")
                    me_f = self.me_for(mod)
                    if p is not None:

                        def let_prop_args(fr: Frame) -> None:
                            call_procedure(fr.ex, p, [f(fr) for f in argfs] + [vf(fr)], None, me_f(fr))

                        return let_prop_args
                    g = d.get("get") or d.get("function")

                    def let_via_get(fr: Frame) -> None:
                        obj = call_procedure(fr.ex, g, [f(fr) for f in argfs], None, me_f(fr))
                        if type(obj) is VArray:
                            raise VbaError(450)
                        let_default(fr.ex, obj, [], None, vf(fr))

                    return let_via_get
                if b.kind == "host":
                    name = b.a

                    def let_host_call(fr: Frame) -> None:
                        args = [f(fr) for f in argfs]
                        nm = {k: f(fr) for k, f in named.items()} if named else None
                        if is_set:
                            fr.ex.host.global_set(fr.ex, name, args, nm, vf(fr))
                        else:
                            fr.ex.host.global_let(fr.ex, name, args, nm, vf(fr))

                    return let_host_call
                if b.kind == "memberofme":
                    name = b.a
                    me_f = self.me_for(self.m)

                    def let_me_call(fr: Frame) -> None:
                        (set_member if is_set else let_member)(fr.ex, me_f(fr), name, [f(fr) for f in argfs], None, vf(fr))

                    return let_me_call
                self.error(f"Can't assign to {callee.name}", line)
            if isinstance(callee, A.Member):
                name = callee.name
                of = self.object_of(callee)
                if of is None:
                    mod = self._qualified_module(callee)
                    if mod is None:
                        self.error("Invalid use of property", line)

                    def let_qa(fr: Frame) -> None:
                        fr.ex.rt.module_let(fr.ex, mod, name, [f(fr) for f in argfs], None, vf(fr), is_set)

                    return let_qa

                def let_mem_args(fr: Frame) -> None:
                    obj = of(fr)
                    args = [f(fr) for f in argfs]
                    nm = {k: f(fr) for k, f in named.items()} if named else None
                    v = vf(fr)
                    if type(obj) is V.Record:
                        let_member(fr.ex, obj, name, args, nm, v)
                        return
                    if is_set:
                        set_member(fr.ex, obj, name, args, nm, v)
                    else:
                        let_member(fr.ex, obj, name, args, nm, v)

                return let_mem_args
            cf, _ = self.expr(callee)

            def let_other(fr: Frame) -> None:
                let_index(fr.ex, cf(fr), [f(fr) for f in argfs], vf(fr))

            return let_other
        if isinstance(target, A.BracketEval):
            text = target.text

            def let_bracket(fr: Frame) -> None:
                obj = fr.ex.host.evaluate(fr.ex, text)
                let_default(fr.ex, obj, [], None, vf(fr))

            return let_bracket
        if isinstance(target, A.MeExpr):
            self.error("Invalid use of Me keyword", line)
        self.error("Expected: variable or procedure, not expression", line)
        return lambda fr: None

    # ------------------------------------------------------------ references (ByRef)

    def ref(self, e: A.Expr, for_redim: bool = False) -> Optional[Callable[[Frame], Slot]]:
        """A function giving the variable ``e`` names (for ByRef arguments, loop counters,
        ReDim and Mid statements), or None when ``e`` isn't a variable."""
        if isinstance(e, A.Ident):
            low = e.name.lower()
            b = self.resolve(e.name, e.line)
            if b.kind == "local":
                idx = b.a.index
                if b.a.auto_new:
                    def r_auto(fr: Frame) -> Slot:
                        s = fr.slots[idx]
                        s.get()
                        return s

                    return r_auto
                return lambda fr: fr.slots[idx]
            if b.kind == "mvar":
                slot = b.a
                return lambda fr: slot
            if b.kind == "field":
                fi = b.a
                return lambda fr: fr.inst.fields[fi]
            return None
        if isinstance(e, A.Call) and not for_redim:
            callee = e.callee
            if isinstance(callee, A.Ident):
                b = self.resolve(callee.name, e.line)
                if b.kind in ("local", "mvar", "field") and not (b.kind == "local" and self.is_own_function(callee.name.lower())):
                    getv = self.ident_value(b, callee.name)
                    argfs, named = self.args(e.args, refs=False)
                    if named:
                        return None

                    def r_elem(fr: Frame) -> Slot:
                        arr = getv(fr)
                        if type(arr) is not VArray:
                            # An object's default member: no variable to refer to.
                            return Slot(index_value(fr.ex, arr, [f(fr) for f in argfs]), None)
                        return ElementRef(arr, [V.to_long(f(fr)) for f in argfs])

                    return r_elem
            return None
        if isinstance(e, A.Member):
            of = self.object_of(e) if e.obj is not None or self.withs else None
            name = e.name.lower()
            if of is None:
                return None

            def r_field(fr: Frame) -> Slot:
                obj = of(fr)
                if type(obj) is V.Record:
                    i = obj.typ.index.get(name)
                    if i is None:
                        raise VbaError(461)
                    return FieldRef(obj, i)
                if type(obj) is Instance and name in obj.cls.field_index and obj.cls.var_public.get(name):
                    return obj.fields[obj.cls.field_index[name]]
                return Slot(get_member(fr.ex, obj, e.name, [], None), None)

            return r_field
        return None

    # ------------------------------------------------------------ expressions

    def static_type(self, e: A.Expr) -> str:
        try:
            return self.expr(e)[1]
        except VbaSyntaxError:
            return "Variant"

    def me_for(self, mod: ModuleRuntime) -> Callable[[Frame], Any]:
        """The ``Me`` a call into ``mod`` gets: this object for this class, else none (a document
        module's procedures find their object through the host)."""
        if mod is self.m and self.m.kind in ("class", "form"):
            return lambda fr: fr.me
        if mod is self.m and self.m.kind == "document":
            return lambda fr: fr.me if fr.me is not None else fr.ex.host.document(fr.ex, mod)
        if mod.kind == "document":
            return lambda fr: fr.ex.host.document(fr.ex, mod)
        return lambda fr: None

    def ident_value(self, b: B, name: str) -> Callable[[Frame], Any]:
        k = b.kind
        if k == "local":
            info: LocalInfo = b.a
            idx = info.index
            if info.auto_new:
                return lambda fr: fr.slots[idx].get()
            return lambda fr: fr.slots[idx].v
        if k == "mvar":
            slot = b.a
            if type(slot) is AutoNewSlot:
                return lambda fr: slot.get()
            return lambda fr: slot.v
        if k == "field":
            fi = b.a
            return lambda fr: slot_value(fr.inst.fields[fi])
        raise AssertionError(k)

    def expr(self, e: Optional[A.Expr]) -> Tuple[Callable[[Frame], Any], str]:
        if e is None:
            return (lambda fr: MISSING), "Variant"
        h = getattr(self, "e_" + type(e).__name__)
        return h(e)

    def e_Lit(self, e: A.Lit) -> Tuple[Callable, str]:
        v = _lit_value(e)
        return (lambda fr: v), e.vtype if e.vtype not in ("Empty", "Null", "Nothing") else "Variant"

    def e_Paren(self, e: A.Paren) -> Tuple[Callable, str]:
        return self.expr(e.expr)

    def e_MeExpr(self, e: A.MeExpr) -> Tuple[Callable, str]:
        if self.m.kind in ("class", "form"):
            return (lambda fr: fr.me), "Object"
        if self.m.kind == "document":
            mod = self.m
            return (lambda fr: fr.me if fr.me is not None else fr.ex.host.document(fr.ex, mod)), "Object"
        self.error("Invalid use of Me keyword", e.line)
        return (lambda fr: NOTHING), "Object"

    def e_BracketEval(self, e: A.BracketEval) -> Tuple[Callable, str]:
        """``[A1:B2]``, ``[SUM(A1:A3)]``: the host evaluates it. A bracketed identifier VBA itself
        knows (``[x]`` for a variable, ``[Rate]`` the function) is that identifier, as in VBA."""
        text = e.text
        if _BRACKET_IDENT.match(text):
            saved = self.implicit_ok
            self.implicit_ok = False
            try:
                b: Optional[B] = self.resolve(text, e.line, implicit=False)
            except VbaSyntaxError:
                b = None
            finally:
                self.implicit_ok = saved
            if b is not None and b.kind != "library":
                return self.expr(A.Ident(text, "", e.line))
        return (lambda fr: fr.ex.host.evaluate(fr.ex, text)), "Variant"

    def e_AddressOf(self, e: A.AddressOf) -> Tuple[Callable, str]:
        def f(fr: Frame) -> Any:
            raise NotSupported("AddressOf (callbacks to Windows)")

        return f, "LongPtr"

    def e_New(self, e: A.New) -> Tuple[Callable, str]:
        rt, m, tname = self.rt, self.m, e.type_name
        short = tname.rsplit(".", 1)[-1].lower()
        cls = rt.module(short)
        if cls is None:
            from .objects import LIBRARY_CLASSES

            if short not in LIBRARY_CLASSES and short not in rt.host_types:
                self.error(f"User-defined type not defined: {tname}", e.line)
        return (lambda fr: _new_object(rt, fr.ex, tname, m)), "Object"

    def e_TypeOfIs(self, e: A.TypeOfIs) -> Tuple[Callable, str]:
        f, _ = self.expr(e.expr)
        want = e.type_name.rsplit(".", 1)[-1].lower()

        def typeof(fr: Frame) -> bool:
            v = f(fr)
            if v is NOTHING:
                return False
            if not is_object(v):
                raise VbaError(424)
            return _is_type(v, want)

        return typeof, "Boolean"

    def e_Ident(self, e: A.Ident) -> Tuple[Callable, str]:
        b = self.resolve(e.name, e.line)
        k = b.kind
        if k in ("local", "mvar", "field"):
            return self.ident_value(b, e.name), b.type
        if k == "const":
            v = b.a
            return (lambda fr: v), b.type
        if k == "proc":
            return self.call(e, [], e.line, want_value=True), b.type
        if k == "builtin":
            return self.call(e, [], e.line, want_value=True), ("String" if e.suffix == "$" else b.type)
        if k == "host":
            name = b.a
            return (lambda fr: fr.ex.host.global_get(fr.ex, name, [], None)), "Variant"
        if k == "memberofme":
            name = b.a
            me_f = self.me_for(self.m)
            return (lambda fr: get_member(fr.ex, me_f(fr), name, [], None)), "Variant"
        if k == "err":
            return (lambda fr: fr.ex.err), "Object"
        if k == "debug":
            from .objects import DEBUG

            return (lambda fr: DEBUG), "Object"
        if k == "module":
            mod = b.a
            if mod.kind == "document":
                return (lambda fr: fr.ex.host.document(fr.ex, mod)), "Object"
            if mod.kind == "form":
                return (lambda fr: fr.ex.rt.form_default(fr.ex, mod)), "Object"
            self.error(f"Expected: variable or procedure, not module ({e.name})", e.line)
        if k == "declare":
            return self.call(e, [], e.line, want_value=True), "Variant"
        if k == "enumtype":
            self.error(f"Expected: expression ({e.name} is an Enum)", e.line)
        if k == "library":
            self.error(f"Expected: expression ({e.name})", e.line)
        self.error(f"Expected: expression ({e.name})", e.line)
        return (lambda fr: EMPTY), "Variant"

    def object_of(self, e: A.Expr) -> Optional[Callable[[Frame], Any]]:
        """The object a member access (``obj.name``) applies to; None for ``Module.name``."""
        obj = e.obj  # type: ignore[attr-defined]
        if obj is None:
            if not self.withs:
                self.error("Invalid or unqualified reference", e.line)
            h = self.withs[-1]
            return lambda fr: fr.slots[h].v
        if isinstance(obj, A.Ident):
            b = self.resolve(obj.name, obj.line)
            if b.kind == "module" and b.a.kind not in ("document", "form"):
                return None
            if b.kind in ("enumtype", "library"):
                return None
        f, _ = self.expr(obj)
        return f

    def _qualified_module(self, e: A.Expr) -> Optional[ModuleRuntime]:
        obj = e.obj  # type: ignore[attr-defined]
        if isinstance(obj, A.Ident):
            b = self.resolve(obj.name, obj.line)
            if b.kind == "module":
                return b.a
        return None

    def e_Member(self, e: A.Member) -> Tuple[Callable, str]:
        return self.call(e, [], e.line, want_value=True), self._member_type(e)

    def _member_type(self, e: A.Expr) -> str:
        if isinstance(e, A.Member) and isinstance(e.obj, A.Ident):
            b = self.resolve(e.obj.name, e.obj.line)
            if b.kind == "enumtype":
                return "Long"
            if b.kind == "module" and b.a.kind == "standard":
                low = e.name.lower()
                mod = b.a
                if low in mod.consts:
                    return mod.const_types.get(low, "Variant")
                if low in mod.vars:
                    return mod.var_types.get(low, "Variant")
        return "Variant"

    def e_Bang(self, e: A.Bang) -> Tuple[Callable, str]:
        of = self.object_of(e)
        if of is None:
            self.error("Invalid use of !", e.line)
        name = e.name
        return (lambda fr: get_default(fr.ex, of(fr), [name], None)), "Variant"

    def e_Call(self, e: A.Call) -> Tuple[Callable, str]:
        t = "Variant"
        if isinstance(e.callee, A.Ident):
            b = self.resolve(e.callee.name, e.line)
            if b.kind in ("builtin", "proc"):
                t = b.type
            elif b.kind in ("local", "mvar", "field") and b.type != "Variant" and not self.is_own_function(e.callee.name.lower()):
                # An element of a typed array.
                info = b.a if b.kind == "local" else None
                if info is None or info.is_array:
                    t = b.type
        return self.call(e.callee, e.args, e.line, want_value=True), t

    def e_Unary(self, e: A.Unary) -> Tuple[Callable, str]:
        f, t = self.expr(e.expr)
        var = t not in _NUM_STATIC
        if e.op == "-":
            return (lambda fr: V.neg(f(fr), var)), ("Variant" if var else t)
        if e.op == "+":
            return f, t
        return (lambda fr: V.op_not(f(fr), var)), ("Boolean" if t == "Boolean" else ("Variant" if var else t))

    def e_Binary(self, e: A.Binary) -> Tuple[Callable, str]:
        op = e.op
        lf, lt_ = self.expr(e.left)
        rf, rt_ = self.expr(e.right)
        lv, rv = lt_ == "Variant" or lt_ not in _NUM_STATIC and lt_ not in ("String", "Date"), rt_ == "Variant" or rt_ not in _NUM_STATIC and rt_ not in ("String", "Date")
        variant = lv or rv
        text = self.m.compare_text
        if op in ("=", "<>", "<", ">", "<=", ">="):
            cmp = {"=": V.eq, "<>": V.ne, "<": V.lt, ">": V.gt, "<=": V.le, ">=": V.ge}[op]
            return (lambda fr: cmp(lf(fr), rf(fr), text, lv, rv)), "Boolean"
        if op == "LIKE":
            return (lambda fr: V.like(lf(fr), rf(fr), text)), "Boolean"
        if op == "IS":

            def is_(fr: Frame) -> bool:
                a, b = lf(fr), rf(fr)
                if not (a is NOTHING or is_object(a)) or not (b is NOTHING or is_object(b)):
                    raise VbaError(424)
                return a is b or _same_object(a, b)

            return is_, "Boolean"
        if op == "&":
            return (lambda fr: V.concat(lf(fr), rf(fr))), ("String" if not variant else "Variant")
        fn = _BINOPS[op]
        rtype = _static_result(op, lt_, rt_)
        if op == "+" and lt_ == "Long" and rt_ in ("Long", "Integer"):

            def add_long(fr: Frame) -> Any:
                a, b = lf(fr), rf(fr)
                if type(a) is int and type(b) is int:
                    r = a + b
                    if -2147483648 <= r <= 2147483647:
                        return r
                return V.add(a, b, False)

            return add_long, rtype
        return (lambda fr: fn(lf(fr), rf(fr), variant)), rtype

    # ------------------------------------------------------------ calls

    def arg_value(self, a: A.Arg) -> Callable[[Frame], Any]:
        if a.missing or a.expr is None:
            return lambda fr: MISSING
        return self.expr(a.expr)[0]

    def args(self, args: List[A.Arg], refs: bool) -> Tuple[List[Callable[[Frame], Any]], Dict[str, Callable[[Frame], Any]]]:
        out: List[Callable[[Frame], Any]] = []
        named: Dict[str, Callable[[Frame], Any]] = {}
        for a in args:
            if a.missing or a.expr is None:
                f: Callable[[Frame], Any] = lambda fr: MISSING
            elif refs and not isinstance(a.expr, A.Paren):
                r = self.ref(a.expr)
                f = r if r is not None else self.expr(a.expr)[0]
            else:
                f = self.expr(a.expr)[0]
            if a.name:
                named[a.name] = f
            else:
                if named:
                    self.error("Expected: named parameter")
                out.append(f)
        return out, named

    def call(self, callee: A.Expr, args: List[A.Arg], line: int, want_value: bool) -> Callable[[Frame], Any]:
        if isinstance(callee, A.Ident):
            low = callee.name.lower()
            b = self.resolve(callee.name, line, implicit=want_value and not args)
            k = b.kind
            if k in ("local", "mvar", "field") and not (args and self.is_own_function(low)):
                getv = self.ident_value(b, callee.name)
                argfs, named = self.args(args, refs=False)
                if not argfs and not named:
                    if not want_value:
                        # A statement that names a variable holding an object calls its default member.
                        def call_var(fr: Frame) -> Any:
                            v = getv(fr)
                            if is_object(v):
                                return get_default(fr.ex, v, [], None)
                            raise VbaError(450 if v is not EMPTY else 35)

                        return call_var
                    return getv

                def idx(fr: Frame) -> Any:
                    v = getv(fr)
                    if type(v) is VArray and not named:
                        return v.get([V.to_long(f(fr)) for f in argfs])
                    return index_value(fr.ex, v, [f(fr) for f in argfs], {n: f(fr) for n, f in named.items()} if named else None)

                return idx
            if k == "local" and args and self.is_own_function(low):
                k = "proc"
                b = B("proc", self.m, low)
            if k == "proc":
                mod, low2 = b.a, b.b
                d = mod.procs[low2]
                p = d.get("function") or d.get("sub") or d.get("get")
                if p is None:
                    p = next(iter(d.values()))
                argfs, named = self.args(args, refs=True)
                me_f = self.me_for(mod)
                if mod.kind == "class" and mod is not self.m:
                    self.error(f"Sub or Function not defined: {callee.name}", line)
                errors = self.rt.errors

                def call_p(fr: Frame) -> Any:
                    a = [f(fr) for f in argfs]
                    nm = {n: f(fr) for n, f in named.items()} if named else None
                    if mod.name.lower() in errors:
                        raise errors[mod.name.lower()]
                    return call_procedure(fr.ex, p, a, nm, me_f(fr))

                return call_p
            if k == "builtin":
                bi = b.a
                if callee.suffix == "$":
                    from .stdlib import lookup

                    bi = lookup(callee.name + "$") or bi
                return self.call_builtin(bi, args, line)
            if k == "host":
                name = b.a
                argfs, named = self.args(args, refs=False)

                def call_h(fr: Frame) -> Any:
                    a = [f(fr) for f in argfs]
                    nm = {n: f(fr) for n, f in named.items()} if named else None
                    return fr.ex.host.global_get(fr.ex, name, a, nm)

                return call_h
            if k == "memberofme" and args and self.m.kind == "form" and low in _FORM_SHADOWED:
                from .stdlib import lookup

                return self.call_builtin(lookup(callee.name + ("$" if callee.suffix == "$" else "")) or lookup(callee.name), args, line)
            if k == "memberofme":
                name = b.a
                argfs, named = self.args(args, refs=False)
                me_f = self.me_for(self.m)

                def call_me(fr: Frame) -> Any:
                    return get_member(fr.ex, me_f(fr), name, [f(fr) for f in argfs], {n: f(fr) for n, f in named.items()} if named else None)

                return call_me
            if k == "declare":
                return self.call_declare(b.a, args, line)
            if k == "err":
                if args:
                    self.error("Wrong number of arguments", line)
                return lambda fr: fr.ex.err
            if k == "debug":
                from .objects import DEBUG

                return lambda fr: DEBUG
            if k == "const":
                v = b.a
                if args:
                    self.error(f"Expected: array ({callee.name} is a constant)", line)
                return lambda fr: v
            if k == "module" and b.a.kind == "document":
                mod = b.a
                return lambda fr: fr.ex.host.document(fr.ex, mod)
            if k == "module" and b.a.kind == "form":
                mod = b.a
                if not args:
                    return lambda fr: fr.ex.rt.form_default(fr.ex, mod)
                argfs, named = self.args(args, refs=False)

                def form_default_member(fr: Frame) -> Any:
                    inst = fr.ex.rt.form_default(fr.ex, mod)
                    return get_default(fr.ex, inst, [f(fr) for f in argfs], {n: f(fr) for n, f in named.items()} if named else None)

                return form_default_member
            self.error(f"Sub or Function not defined: {callee.name}", line)
        if isinstance(callee, A.Member):
            name = callee.name
            # VBA.Left, Excel.Range…: the library's own names.
            # Excel.XlDirection.xlUp: a library's enum member.
            inner = callee.obj
            if isinstance(inner, A.Member) and isinstance(inner.obj, A.Ident) and inner.name.lower() in _LIB_ENUMS and self.resolve(inner.obj.name, line).kind == "library":
                members = _LIB_ENUMS[inner.name.lower()]
                if name.lower() not in members:
                    self.error(f"Method or data member not found: {name}", line)
                if args:
                    self.error(f"Expected: array ({name} is a constant)", line)
                cv = members[name.lower()]
                return lambda fr: cv
            if isinstance(callee.obj, A.Ident):
                lb = self.resolve(callee.obj.name, line)
                if lb.kind == "library":
                    lib = lb.a
                    from .stdlib import lookup

                    if lib == "vba":
                        bi = lookup(name)
                        if bi is None:
                            self.error(f"Method or data member not found: VBA.{name}", line)
                        return self.call_builtin(bi, args, line, qualified=True)
                    if lib in ("excel",):
                        return self.call(A.Ident(name, "", line), args, line, want_value)
                    if lib == "vbaproject":
                        return self.call(A.Ident(name, "", line), args, line, want_value)
                if lb.kind == "enumtype":
                    vals = lb.a
                    v = vals.get(name.lower())
                    if v is None:
                        self.error(f"Method or data member not found: {name}", line)
                    return lambda fr: v
            mod = self._qualified_module(callee)
            if mod is not None and mod.kind == "standard":
                low = name.lower()
                if low in mod.procs:
                    d = mod.procs[low]
                    p = d.get("function") or d.get("sub") or d.get("get")
                    argfs, named = self.args(args, refs=True)

                    def call_q(fr: Frame) -> Any:
                        return fr.ex.rt.call(fr.ex, p, [f(fr) for f in argfs], {n: f(fr) for n, f in named.items()} if named else None)

                    return call_q
                argfs, named = self.args(args, refs=False)

                def get_q(fr: Frame) -> Any:
                    return fr.ex.rt.module_get(fr.ex, mod, name, [f(fr) for f in argfs], None)

                return get_q
            of = self.object_of(callee)
            if of is None:
                self.error(f"Method or data member not found: {name}", line)
            argfs, named = self.args(args, refs=True)

            def call_m(fr: Frame) -> Any:
                obj = of(fr)
                a = [f(fr) for f in argfs]
                nm = {n: f(fr) for n, f in named.items()} if named else None
                return get_member(fr.ex, obj, name, a, nm)

            return call_m
        if isinstance(callee, A.Bang):
            of = self.object_of(callee)
            nm_ = callee.name
            argfs, named = self.args(args, refs=False)

            def call_bang(fr: Frame) -> Any:
                v = get_default(fr.ex, of(fr), [nm_], None)
                return index_value(fr.ex, v, [f(fr) for f in argfs]) if argfs else v

            return call_bang
        cf, _ = self.expr(callee)
        if not args:
            return cf
        argfs, named = self.args(args, refs=False)

        def call_v(fr: Frame) -> Any:
            return index_value(fr.ex, cf(fr), [f(fr) for f in argfs], {n: f(fr) for n, f in named.items()} if named else None)

        return call_v

    def call_builtin(self, bi, args: List[A.Arg], line: int, qualified: bool = False) -> Callable[[Frame], Any]:
        if bi.name in ("Len", "LenB") and len(args) == 1 and args[0].expr is not None:
            a = args[0].expr
            if isinstance(a, A.Lit) and a.vtype in _NUM_STATIC and a.vtype != "Boolean":
                self.error("Variable required - can't assign to this expression", line)
            if isinstance(a, A.Ident):
                b = self.resolve(a.name, line)
                size = _TYPE_SIZES.get(b.type) if b.kind in ("local", "mvar", "field") else None
                if size is not None and not (b.kind == "local" and b.a.is_array):
                    return lambda fr: size
        argfs, named = self.args(args, refs=bi.refs)
        m = self.m
        if named:
            names = [n.lower() for n in bi.params]
            order: List[Optional[Callable]] = list(argfs) + [None] * max(0, len(names) - len(argfs))
            for k, f in named.items():
                if k.lower() not in names:
                    self.error(f"Named argument not found: {k}", line)
                order[names.index(k.lower())] = f
            while order and order[-1] is None:
                order.pop()
            argfs = [f if f is not None else (lambda fr: MISSING) for f in order]
        n = len(argfs)
        if n < bi.min_args:
            self.error(f"Argument not optional: {bi.name}", line)
        if bi.max_args >= 0 and n > bi.max_args:
            self.error(f"Wrong number of arguments or invalid property assignment: {bi.name}", line)
        fn = bi.fn
        ctx = _CallCtx(m, qualified)
        if n == 0:
            return lambda fr: fn(fr.ex, [], ctx)
        if n == 1:
            a0 = argfs[0]
            return lambda fr: fn(fr.ex, [a0(fr)], ctx)
        if n == 2:
            a0, a1 = argfs
            return lambda fr: fn(fr.ex, [a0(fr), a1(fr)], ctx)
        return lambda fr: fn(fr.ex, [f(fr) for f in argfs], ctx)

    def call_declare(self, d: A.DeclareDecl, args: List[A.Arg], line: int) -> Callable[[Frame], Any]:
        argfs, _ = self.args(args, refs=False)
        from .stdlib import declared_function

        name = d.alias or d.name
        lib = d.lib

        def call_d(fr: Frame) -> Any:
            return declared_function(fr.ex, lib, name, [f(fr) for f in argfs])

        return call_d


class _CallCtx:
    """What a library function may need from where it is called: Option Base and Compare."""

    __slots__ = ("module", "qualified")

    def __init__(self, module: ModuleRuntime, qualified: bool = False):
        self.module = module
        self.qualified = qualified

    @property
    def option_base(self) -> int:
        return 0 if self.qualified else self.module.option_base

    @property
    def compare_text(self) -> bool:
        return self.module.compare_text


# Len of a variable of these types is its size in bytes.
_TYPE_SIZES = {"Byte": 1, "Boolean": 2, "Integer": 2, "Long": 4, "LongLong": 8, "LongPtr": 8, "Single": 4, "Double": 8, "Currency": 8, "Date": 8}

_PLAIN = frozenset({int, float, str, bool, V.Integer, V.Byte, V.LongLong, V.Single, V.Date, V.Currency})

# A UserForm's own members, reachable without "Me." in its module (its controls are too).
_FORM_MEMBERS = frozenset({
    "activecontrol", "backcolor", "bordercolor", "borderstyle", "canpaste", "canredo", "canundo", "caption",
    "controls", "copy", "cut", "cycle", "drawbuffer", "enabled", "font", "forecolor", "height", "helpcontextid",
    "hide", "insideheight", "insidewidth", "keepscrollbarsvisible", "left", "mouseicon", "mousepointer", "move",
    "paste", "picture", "picturealignment", "picturesizemode", "picturetiling", "printform", "redoaction", "repaint",
    "scroll", "scrollbars", "scrollheight", "scrollleft", "scrolltop", "scrollwidth", "setdefaulttaborder", "show",
    "specialeffect", "startupposition", "tag", "top", "undoaction", "verticalscrollbarside", "visible",
    "whatsthisbutton", "whatsthishelp", "width", "zoom",
})
# Form members that are also VBA functions: called with arguments, the function.
_FORM_SHADOWED = frozenset({"left"})

_DOCUMENT_MEMBERS = {
    "worksheet": (
        "range", "cells", "rows", "columns", "name", "usedrange", "activate", "select", "calculate", "protect", "unprotect",
        "shapes", "chartobjects", "listobjects", "names", "hyperlinks", "comments", "index", "visible", "evaluate", "paste",
        "autofiltermode", "autofilter", "sort", "parent", "next", "previous", "codename", "tab", "pagesetup", "delete", "copy", "move",
    ),
    "workbook": ("worksheets", "sheets", "name", "fullname", "path", "save", "activesheet", "names", "charts", "windows", "protect", "unprotect", "codename"),
}


def _is_object_slot(s: Slot) -> bool:
    return s.co is None and is_object(s.v) and s.v is not EMPTY


def _num(v: Any) -> Any:
    if is_object(v):
        v = V.value_of(v)
    if type(v) is str:
        x = V.parse_number(v)
        if x is None:
            raise VbaError(13)
        return float(x) if not isinstance(x, int) else x
    if v is EMPTY:
        return V.Integer(0)
    if v is NULL:
        raise VbaError(94)
    if type(v) is bool:
        return V.Integer(-1 if v else 0)
    return v


def _widest(st: Any, en: Any, sp: Any) -> Any:
    ranks = {V.Byte: 1, V.Integer: 2, int: 3, V.LongLong: 4, V.Single: 5, float: 6, V.Currency: 7, V.Date: 6}
    import decimal

    r = max(ranks.get(type(x), 6 if type(x) is not decimal.Decimal else 8) for x in (st, en, sp))
    conv = {1: V.to_byte, 2: V.to_integer, 3: V.to_long, 4: V.to_longlong, 5: V.to_single, 6: V.to_double, 7: V.to_currency, 8: V.to_decimal}[r]
    if type(st) is V.Date:
        return st
    return conv(st)


def _cmp(a: Any, b: Any, text: bool, va: bool, vb: bool) -> int:
    r = V.compare(a, b, text, va, vb)
    if r is NULL:
        return 2  # nothing matches Null
    return r


def _relop(op: str, c: int) -> bool:
    if c == 2:
        return False
    return {"=": c == 0, "<>": c != 0, "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0}[op]


def _static_result(op: str, a: str, b: str) -> str:
    if op in ("+", "-", "*") and a in _NUM_STATIC and b in _NUM_STATIC:
        r = max(_NUM_STATIC[a], _NUM_STATIC[b])
        if r == 5 and min(_NUM_STATIC[a], _NUM_STATIC[b]) in (3, 4):
            r = 6
        return {1: "Byte", 2: "Integer", 3: "Long", 4: "LongLong", 5: "Single", 6: "Double", 7: "Currency", 8: "Decimal"}[r]
    if op in ("\\", "MOD") and a in _NUM_STATIC and b in _NUM_STATIC:
        def ir(t: str) -> int:
            r = _NUM_STATIC[t]
            return r if r <= 4 else 3

        return {1: "Byte", 2: "Integer", 3: "Long", 4: "LongLong"}[max(ir(a), ir(b))]
    if op in ("/", "^") and a in _NUM_STATIC and b in _NUM_STATIC:
        return "Double"
    if op in ("AND", "OR", "XOR", "EQV", "IMP") and a == "Boolean" and b == "Boolean":
        return "Boolean"
    return "Variant"


def _same_object(a: Any, b: Any) -> bool:
    same = getattr(a, "vba_same", None)
    return bool(same(b)) if same is not None else False


def _is_type(v: Any, want: str) -> bool:
    if want == "object":
        return True
    if type(v) is Instance:
        if v.cls.name.lower() == want or (v.form is not None and want == "userform"):
            return True
        return any(i.rsplit(".", 1)[-1].lower() == want for i in v.cls.implements)
    names = getattr(v, "TYPE_NAMES", None)
    if names is not None and want in names:
        return True
    tn = V.type_name(v).lower()
    return tn == want


def _walk(body: List[A.Stmt]) -> Iterator[A.Stmt]:
    for s in body:
        yield s
        if isinstance(s, A.If):
            for _, b in s.branches:
                yield from _walk(b)
            if s.orelse:
                yield from _walk(s.orelse)
        elif isinstance(s, A.Select):
            for _, b, _ in s.cases:
                yield from _walk(b)
            if s.orelse:
                yield from _walk(s.orelse)
        elif isinstance(s, (A.For, A.ForEach, A.Do, A.While, A.With)):
            yield from _walk(s.body)



def compile_immediate(rt: ProjectRuntime, code: str) -> Procedure:
    """Statements typed in the Immediate window (``? expr`` prints it), compiled as a procedure
    of the project: they see every public name, and implicit variables are allowed."""
    lines = []
    for ln in code.replace("\r\n", "\n").split("\n"):
        t = ln.strip()
        if t.startswith("?"):
            t = "Debug.Print " + t[1:]
        lines.append(t)
    src = "Sub Immediate__()\n" + "\n".join(lines) + "\nEnd Sub\n"
    try:
        mast = parse_module(src, "Immediate")
        m = ModuleRuntime("Immediate", "standard")
        _declare(rt, m, mast, ())
        for p in mast.procs:
            ProcCompiler(rt, m, mast, p).compile()
    except VbaSyntaxError as e:
        # Line numbers as typed (the wrapping Sub line isn't the person's).
        raise VbaSyntaxError(e.msg, max(1, min(e.line - 1, len(lines))) if e.line else 0) from None
    return m.proc("Immediate__", "sub")
