"""The VBA runtime: variables, frames, procedures, class instances, calls and member access.

A compiled procedure is a list of instructions (closures over a :class:`Frame`), run by
:func:`run_frame`. ``On Error``, ``Resume`` and ``GoTo`` are jumps between instructions, which
is how VBA behaves too (``Resume Next`` after an error in an ``If`` condition enters the
``Then`` block). One :class:`Execution` is one run — a macro, an event handler or a worksheet
function — with its own ``Err``, limits and output; module-level variables live in the
:class:`ProjectRuntime` and last until the code changes or ``End`` runs, as in Excel.

Objects VBA code can use are :class:`HostObject` subclasses (members are methods named
``g_<name>`` to read or call, ``l_<name>`` to assign, ``s_<name>`` to ``Set``), class module
instances (:class:`Instance`) and the library's Collection, Dictionary…
"""

from __future__ import annotations

import threading
import time
from typing import Any, Callable, Dict, Iterator, List, Optional, Sequence, Tuple

from .errors import MacroStopped, NotSupported, VbaError
from .values import (
    EMPTY,
    MISSING,
    NOTHING,
    NULL,
    COERCE,
    DEFAULTS,
    Record,
    VArray,
    copy_value,
    is_object,
    set_default_value_hook,
    to_long,
    value_of,
)

MAX_DEPTH = 900


class Slot:
    """A variable. ``co`` converts what is assigned to the declared type (None: Variant)."""

    __slots__ = ("v", "co")

    def __init__(self, v: Any = EMPTY, co: Optional[Callable] = None):
        self.v = v
        self.co = co

    def let(self, value: Any) -> None:
        co = self.co
        self.v = co(value) if co is not None else value

    def __repr__(self) -> str:
        return f"Slot({self.v!r})"


class AutoNewSlot(Slot):
    """``Dim x As New Class``: created on first use (and again after ``Set x = Nothing``)."""

    __slots__ = ("factory",)

    def __init__(self, factory: Callable[[], Any], co: Optional[Callable] = None):
        super().__init__(NOTHING, co)
        self.factory = factory

    def get(self) -> Any:
        v = Slot.v.__get__(self)  # type: ignore[attr-defined]
        if v is NOTHING:
            v = self.factory()
            Slot.v.__set__(self, v)  # type: ignore[attr-defined]
        return v


class ElementRef(Slot):
    """A ByRef argument that is an array element."""

    __slots__ = ("arr", "idx")

    def __init__(self, arr: VArray, idx: Sequence[int]):
        self.arr = arr
        self.idx = tuple(idx)
        self.co = arr.coerce

    @property
    def v(self) -> Any:  # type: ignore[override]
        return self.arr.get(self.idx)

    @v.setter
    def v(self, value: Any) -> None:
        self.arr.data[self.arr.offset(self.idx)] = value


class FieldRef(Slot):
    """A ByRef argument that is a field of a user-defined type."""

    __slots__ = ("rec", "i")

    def __init__(self, rec: Record, i: int):
        self.rec = rec
        self.i = i
        self.co = rec.typ.coerces[i]

    @property
    def v(self) -> Any:  # type: ignore[override]
        return self.rec.fields[self.i]

    @v.setter
    def v(self, value: Any) -> None:
        self.rec.fields[self.i] = value


class WithEventsSlot(Slot):
    """A ``WithEvents`` variable: the object it holds raises its events to the variable's owner
    (``Private WithEvents src As Shape`` gets ``src_Changed``)."""

    __slots__ = ("_v", "owner", "name")

    def __init__(self, name: str):
        self._v = NOTHING
        self.co = None
        self.owner: Any = None
        self.name = name

    @property
    def v(self) -> Any:  # type: ignore[override]
        return self._v

    @v.setter
    def v(self, value: Any) -> None:
        old = self._v
        if type(old) is Instance and self.owner is not None:
            low = self.name.lower()
            old.listeners = [(o, n) for o, n in old.listeners if not (o is self.owner and n.lower() == low)]
        self._v = value
        if type(value) is Instance and self.owner is not None:
            value.listeners.append((self.owner, self.name))


def slot_value(s: Slot) -> Any:
    if type(s) is AutoNewSlot:
        return s.get()
    return s.v


# ---------------------------------------------------------------- objects


class HostObject:
    """An object VBA can use. Members are methods: ``g_name(ex, *args)`` reads a property or
    calls a method, ``l_name(ex, value, *args)`` assigns a property, ``s_name`` sets an object
    property. ``PARAMS`` names each member's parameters for named arguments; ``DEFAULT`` is the
    default member."""

    __slots__ = ()
    TYPE_NAME = "Object"
    PARAMS: Dict[str, Tuple[str, ...]] = {}
    DEFAULT: Optional[str] = None

    def vba_type_name(self) -> str:
        return self.TYPE_NAME

    def _member(self, prefix: str, name: str):
        low = name.lower()
        f = getattr(self, prefix + low, None)
        return f, low

    def vba_get(self, ex: "Execution", name: str, args: list, named: Optional[dict] = None) -> Any:
        f, low = self._member("g_", name)
        if f is None:
            return self.vba_get_missing(ex, name, args, named)
        if named:
            args = bind_named(self.PARAMS.get(low), args, named, name)
        if args and _arity(f)[1] == 0:
            # `rng.FormatConditions(1)`: arguments of a property without parameters go to
            # the default member of what it returns.
            return index_value(ex, _call_host(f, ex, [], name), list(args))
        return _call_host(f, ex, args, name)

    def vba_get_missing(self, ex: "Execution", name: str, args: list, named: Optional[dict]) -> Any:
        raise VbaError(438, f"Object doesn't support this property or method: {name}")

    def vba_let(self, ex: "Execution", name: str, args: list, named: Optional[dict], value: Any) -> None:
        f, low = self._member("l_", name)
        if f is None:
            if getattr(self, "g_" + low, None) is not None or self.has_dynamic_member(name):
                # `ws.Cells(1, 1) = 5`: the property's object takes the value through its default member.
                obj = self.vba_get(ex, name, args, named)
                if is_object(obj):
                    return let_default(ex, obj, [], None, value)
                raise VbaError(450, f"Can't assign to read-only property {name}")
            raise VbaError(438, f"Object doesn't support this property or method: {name}")
        if named:
            args = bind_named(self.PARAMS.get(low), args, named, name)
        _call_host(f, ex, [value] + list(args), name)

    def has_dynamic_member(self, name: str) -> bool:
        return False

    def vba_set(self, ex: "Execution", name: str, args: list, named: Optional[dict], value: Any) -> None:
        f, low = self._member("s_", name)
        if f is None:
            return self.vba_let(ex, name, args, named, value)
        if named:
            args = bind_named(self.PARAMS.get(low), args, named, name)
        _call_host(f, ex, [value] + list(args), name)

    def vba_default_get(self, ex: "Execution", args: list, named: Optional[dict] = None) -> Any:
        if self.DEFAULT is None:
            raise VbaError(438, f"{self.vba_type_name()} has no default property")
        return self.vba_get(ex, self.DEFAULT, args, named)

    def vba_default_let(self, ex: "Execution", args: list, named: Optional[dict], value: Any) -> None:
        if self.DEFAULT is None:
            raise VbaError(438, f"{self.vba_type_name()} has no default property")
        self.vba_let(ex, self.DEFAULT, args, named, value)

    def vba_enum(self, ex: "Execution") -> Iterator[Any]:
        raise VbaError(438, f"For Each can't go through a {self.vba_type_name()}")

    def __repr__(self) -> str:
        return f"<{self.vba_type_name()}>"


_ARITY: Dict[Any, Tuple[int, int]] = {}


def _arity(f: Callable) -> Tuple[int, int]:
    """(fewest, most) arguments a host member method takes."""
    fn = getattr(f, "__func__", f)
    ar = _ARITY.get(fn)
    if ar is None:
        code = fn.__code__
        n = code.co_argcount - 2  # self, ex
        if code.co_flags & 0x04:
            mx = 10_000
        else:
            mx = n
        mn = n - len(fn.__defaults__ or ())
        ar = _ARITY[fn] = (mn, mx)
    return ar


def _call_host(f: Callable, ex: "Execution", args: Sequence[Any], name: str) -> Any:
    ar = _arity(f)
    if not ar[0] <= len(args) <= ar[1]:
        # Trailing omitted arguments are fine; missing required ones aren't.
        while len(args) > ar[1] and args and args[-1] is MISSING:
            args = list(args)[:-1]
        if not ar[0] <= len(args) <= ar[1]:
            raise VbaError(450, f"Wrong number of arguments or invalid property assignment: {name}")
    return f(ex, *args)


def bind_named(params: Optional[Sequence[str]], args: list, named: dict, what: str) -> list:
    """Positional and named arguments as one positional list (MISSING for the gaps)."""
    if not params:
        raise VbaError(446, f"Object doesn't support named arguments: {what}")
    low = [p.lower() for p in params]
    out = list(args) + [MISSING] * max(0, len(params) - len(args))
    for k, v in named.items():
        try:
            i = low.index(k.lower())
        except ValueError:
            raise VbaError(448, f"Named argument not found: {k}")
        if i < len(args) and args[i] is not MISSING:
            raise VbaError(5, f"Argument {k} given twice")
        out[i] = v
    while out and out[-1] is MISSING:
        out.pop()
    return out


# ---------------------------------------------------------------- the current execution

_tls = threading.local()


def current() -> "Execution":
    stack = getattr(_tls, "stack", None)
    if not stack:
        raise VbaError(51, "no VBA code is running")
    return stack[-1]


def _default_value(obj: Any) -> Any:
    ex = current()
    return get_default(ex, obj, [], None)


set_default_value_hook(_default_value)


# ---------------------------------------------------------------- procedures and modules


class ParamSpec:
    __slots__ = ("name", "byval", "optional", "default", "paramarray", "co", "type_name", "is_array", "is_object")

    def __init__(self, name: str, byval: bool, optional: bool, default: Any, paramarray: bool, co: Optional[Callable], type_name: str, is_array: bool, is_object: bool):
        self.name = name
        self.byval = byval
        self.optional = optional
        self.default = default
        self.paramarray = paramarray
        self.co = co
        self.type_name = type_name
        self.is_array = is_array
        self.is_object = is_object


class Procedure:
    """A compiled Sub, Function or Property procedure."""

    def __init__(self, module: "ModuleRuntime", name: str, kind: str, scope: str, line: int):
        self.module = module
        self.name = name
        self.kind = kind  # sub function get let set
        self.scope = scope
        self.line = line
        self.params: List[ParamSpec] = []
        self.nslots = 0
        self.code: List[Callable] = []
        self.lines: List[int] = []
        self.next_stmt: List[int] = []
        self.labels: Dict[str, int] = {}
        self.ret_slot = -1
        self.ret_co: Optional[Callable] = None
        self.ret_type = "Variant"
        self.ret_is_object = False
        self.init: List[Tuple[int, Callable[[], Slot]]] = []
        self.static_all = False
        self.statics: Dict[int, Slot] = {}
        self.static_init: List[Tuple[int, Callable[[], Slot]]] = []
        self.has_paramarray = False
        self.member_of_class = False
        self.erl: List[int] = []

    @property
    def qualified(self) -> str:
        return f"{self.module.name}.{self.name}"

    def __repr__(self) -> str:
        return f"<Procedure {self.qualified}>"


class ModuleRuntime:
    def __init__(self, name: str, kind: str):
        self.name = name
        self.kind = kind  # standard class document form
        self.option_base = 0
        self.option_explicit = False
        self.compare_text = False
        self.private_module = False
        self.vars: Dict[str, Slot] = {}
        self.var_public: Dict[str, bool] = {}
        self.var_types: Dict[str, str] = {}
        self.consts: Dict[str, Any] = {}
        self.const_public: Dict[str, bool] = {}
        self.const_types: Dict[str, str] = {}
        # name -> {kind: Procedure}
        self.procs: Dict[str, Dict[str, Procedure]] = {}
        self.types: Dict[str, "UdtDef"] = {}
        self.enums: Dict[str, Dict[str, int]] = {}
        self.enum_public: Dict[str, bool] = {}
        self.declares: Dict[str, Any] = {}
        self.events: Dict[str, Any] = {}
        self.implements: List[str] = []
        self.default_member: Optional[str] = None
        self.enum_member: Optional[str] = None
        # Class modules: fields by name, in instance order.
        self.field_index: Dict[str, int] = {}
        self.field_inits: List[Callable[[], Slot]] = []
        self.field_types: Dict[str, str] = {}
        self.var_inits: Dict[str, Callable[[], Slot]] = {}
        # Class modules: Class_Terminate, and whether some field is WithEvents.
        self.terminate: Optional["Procedure"] = None
        self.has_withevents = False

    def proc(self, name: str, kind: Optional[str] = None) -> Optional[Procedure]:
        d = self.procs.get(name.lower())
        if not d:
            return None
        if kind is not None:
            return d.get(kind)
        for k in ("function", "sub", "get"):
            if k in d:
                return d[k]
        return next(iter(d.values()))

    def __repr__(self) -> str:
        return f"<Module {self.name}>"


class UdtDef:
    """A user-defined type: its fields' names, types and initial values."""

    def __init__(self, name: str):
        self.name = name
        self.names: List[str] = []
        self.index: Dict[str, int] = {}
        self.coerces: List[Optional[Callable]] = []
        self.inits: List[Callable[[], Any]] = []
        self.types: List[str] = []

    def new(self) -> Record:
        return Record(self, [f() for f in self.inits])


class Instance:
    """An object made from a class module (``New Class1``). When the last reference to it goes,
    its ``Class_Terminate`` runs (at the next statement, see :func:`run_terminations`)."""

    __slots__ = ("cls", "fields", "rt", "__weakref__", "listeners", "terminated", "gen")

    def __init__(self, cls: ModuleRuntime, rt: "ProjectRuntime"):
        self.cls = cls
        self.rt = rt
        self.fields: List[Slot] = [f() for f in cls.field_inits]
        self.listeners: List[Tuple[Any, str]] = []
        self.terminated = cls.terminate is None
        self.gen = rt.generation
        if cls.has_withevents:
            for f in self.fields:
                if type(f) is WithEventsSlot:
                    f.owner = self

    def __del__(self) -> None:
        # Kept alive until its Class_Terminate has run (VBA code doesn't run inside a
        # finalizer: the runtime picks it up between statements).
        if not self.terminated:
            self.terminated = True
            try:
                self.rt.dying.append(self)
            except Exception:
                pass

    def vba_type_name(self) -> str:
        return self.cls.name

    def __repr__(self) -> str:
        return f"<{self.cls.name} instance>"


# ---------------------------------------------------------------- frames and execution


class Frame:
    __slots__ = ("slots", "proc", "me", "ex", "on_error", "handler", "active", "resume_pc", "gosub", "inst")

    def __init__(self, proc: Procedure, ex: "Execution", me: Any = None):
        self.proc = proc
        self.ex = ex
        self.me = me
        self.inst = me if type(me) is Instance else None
        self.slots: List[Slot] = [None] * proc.nslots  # type: ignore[list-item]
        self.on_error = 0  # 0 off, 1 Resume Next, 2 GoTo handler
        self.handler = -1
        self.active = False
        self.resume_pc = -1
        self.gosub: List[int] = []


class ErrObject(HostObject):
    """VBA's ``Err``."""

    TYPE_NAME = "ErrObject"
    DEFAULT = "number"
    PARAMS = {"raise": ("Number", "Source", "Description", "HelpFile", "HelpContext")}

    def __init__(self) -> None:
        self.number = 0
        self.description = ""
        self.source = ""
        self.help_file = ""
        self.help_context = 0
        self.last_dll_error = 0
        self.erl = 0

    def clear(self) -> None:
        self.number = 0
        self.description = ""
        self.source = ""
        self.help_file = ""
        self.help_context = 0
        self.erl = 0

    def set(self, e: VbaError, project: str = "VBAProject", erl: int = 0) -> None:
        self.number = e.number
        self.description = e.description
        self.source = e.source or project
        self.help_file = e.help_file
        self.help_context = e.help_context
        self.erl = erl

    def g_number(self, ex):
        return self.number

    def l_number(self, ex, v):
        from .errors import message

        self.number = to_long(v)
        if not self.description:
            self.description = message(self.number)

    def g_description(self, ex):
        return self.description

    def l_description(self, ex, v):
        from .values import to_string

        self.description = to_string(v)

    def g_source(self, ex):
        return self.source

    def l_source(self, ex, v):
        from .values import to_string

        self.source = to_string(v)

    def g_helpfile(self, ex):
        return self.help_file

    def l_helpfile(self, ex, v):
        from .values import to_string

        self.help_file = to_string(v)

    def g_helpcontext(self, ex):
        return self.help_context

    def l_helpcontext(self, ex, v):
        self.help_context = to_long(v)

    def g_lastdllerror(self, ex):
        return 0

    def g_clear(self, ex):
        self.clear()
        return EMPTY

    def g_raise(self, ex, number=MISSING, source=MISSING, description=MISSING, helpfile=MISSING, helpcontext=MISSING):
        from .values import to_string

        if number is MISSING:
            raise VbaError(449)
        n = to_long(number)
        e = VbaError(n, to_string(description) if description is not MISSING else None)
        if description is MISSING and self.number == n and self.description:
            e.description = self.description
        e.source = to_string(source) if source is not MISSING else ex.rt.project_name
        if helpfile is not MISSING:
            e.help_file = to_string(helpfile)
        if helpcontext is not MISSING:
            e.help_context = to_long(helpcontext)
        raise e


class Host:
    """What VBA code reaches beyond the language: the workbook (Excel's object model), dialogs
    and output. This base has nothing; :mod:`gridcore.engine.macro` provides Excel."""

    # Names (lower case) the host resolves as globals: Range, Cells, ActiveSheet…
    globals: frozenset = frozenset()

    def constant(self, name: str) -> Any:
        return None

    def global_get(self, ex: "Execution", name: str, args: list, named: Optional[dict]) -> Any:
        raise VbaError(35, f"Sub or Function not defined: {name}")

    def global_let(self, ex: "Execution", name: str, args: list, named: Optional[dict], value: Any) -> None:
        obj = self.global_get(ex, name, args, named)
        let_default(ex, obj, [], None, value)

    def global_set(self, ex: "Execution", name: str, args: list, named: Optional[dict], value: Any) -> None:
        raise VbaError(450)

    def document(self, ex: "Execution", module: ModuleRuntime) -> Any:
        """The object behind a document module (``Me`` in ThisWorkbook or a sheet module)."""
        return NOTHING

    def prompt(self, ex: "Execution", kind: str, prompt: str, title: str, buttons: int, default: Any = None, **extra: Any) -> Any:
        """MsgBox (``kind`` "msgbox", returns the button's value) or InputBox (returns text,
        or "" when cancelled). Without a user the answer is the default button."""
        from .stdlib import msgbox_default

        return msgbox_default(buttons) if kind == "msgbox" else (default if default is not None else "")

    def output(self, ex: "Execution", text: str) -> None:
        ex.output.append(text)

    def checkpoint(self, ex: "Execution") -> None:
        pass

    def create_object(self, ex: "Execution", progid: str) -> Any:
        from .objects import create_object

        return create_object(ex, progid)

    def evaluate(self, ex: "Execution", text: str) -> Any:
        raise VbaError(1004, "Evaluate needs a workbook")

    def status(self, ex: "Execution", text: str) -> None:
        pass


class Execution:
    """One run of VBA code: limits, ``Err``, the call depth and what ``Debug.Print`` wrote."""

    def __init__(self, rt: "ProjectRuntime", host: Optional[Host] = None, time_limit: float = 60.0, max_steps: int = 0, udf: bool = False, max_depth: int = MAX_DEPTH):
        self.rt = rt
        self.host = host or Host()
        self.err = ErrObject()
        self.depth = 0
        self.max_depth = max_depth
        self.steps = 0
        self.max_steps = max_steps
        self.time_limit = time_limit
        self.deadline = time.monotonic() + time_limit if time_limit else 0.0
        self.output: List[str] = []
        self.print_col = 0
        self.udf = udf
        self.cancelled = False
        self.messages: List[str] = []
        # Last error that stopped the run, with where it happened.
        self.last_error: Optional[VbaError] = None
        self.started = time.monotonic()

    def __enter__(self) -> "Execution":
        st = getattr(_tls, "stack", None)
        if st is None:
            st = _tls.stack = []
        st.append(self)
        return self

    def __exit__(self, *exc) -> None:
        _tls.stack.pop()

    def tick(self) -> None:
        self.steps += 1
        if not self.steps & 1023:
            self.check()

    def check(self) -> None:
        if self.cancelled:
            raise MacroStopped("cancelled", "The macro was stopped.")
        if self.deadline and time.monotonic() > self.deadline:
            raise MacroStopped("timeout", f"The macro ran for more than {self.time_limit:g} seconds and was stopped.")
        if self.max_steps and self.steps > self.max_steps:
            raise MacroStopped("steps", "The macro ran too many steps and was stopped.")
        self.host.checkpoint(self)

    def extend_deadline(self, seconds: float) -> None:
        if self.deadline:
            self.deadline += seconds


# ---------------------------------------------------------------- running procedures


def run_terminations(ex: "Execution") -> None:
    """Runs ``Class_Terminate`` of the objects whose last reference went (objects released by
    ``End`` don't run it, as in VBA)."""
    dying = ex.rt.dying
    while dying:
        inst = dying.pop(0)
        if inst.gen != ex.rt.generation:
            continue
        call_procedure(ex, inst.cls.terminate, [], None, inst)


def run_frame(fr: Frame) -> None:
    proc = fr.proc
    code = proc.code
    n = len(code)
    pc = 0
    dying = fr.ex.rt.dying
    while True:
        try:
            while pc < n:
                pc = code[pc](fr)
                if dying:
                    run_terminations(fr.ex)
            return
        except VbaError as e:
            pc = _trap(fr, e, pc)
        except MacroStopped:
            raise
        except RecursionError:
            pc = _trap(fr, VbaError(28), pc)
        except ZeroDivisionError:
            pc = _trap(fr, VbaError(11), pc)
        except OverflowError:
            pc = _trap(fr, VbaError(6), pc)
        except MemoryError:
            pc = _trap(fr, VbaError(7), pc)
        except (KeyboardInterrupt, SystemExit):
            raise
        except Exception as e:  # a bug in the runtime: a VBA error the code can trap
            pc = _trap(fr, VbaError(51, f"Internal error ({type(e).__name__}: {e})"), pc)


def _trap(fr: Frame, e: VbaError, pc: int) -> int:
    proc = fr.proc
    if not e.module:
        e.module = proc.module.name
        e.proc = proc.name
        e.line = proc.lines[pc] if 0 <= pc < len(proc.lines) else proc.line
    ex = fr.ex
    if fr.on_error == 0 or fr.active:
        raise e
    erl = proc.erl[pc] if 0 <= pc < len(proc.erl) else 0
    ex.err.set(e, ex.rt.project_name, erl)
    if fr.on_error == 1:
        nxt = proc.next_stmt[pc] if 0 <= pc < len(proc.next_stmt) else pc + 1
        return nxt
    fr.active = True
    fr.resume_pc = pc
    return fr.handler


def call_procedure(ex: Execution, proc: Procedure, args: Sequence[Any], named: Optional[dict] = None, me: Any = None) -> Any:
    """Calls ``proc``. ``args`` are values or :class:`Slot` references (ByRef arguments);
    MISSING marks an omitted argument."""
    if ex.depth >= ex.max_depth:
        raise VbaError(28)
    fr = Frame(proc, ex, me)
    slots = fr.slots
    params = proc.params
    if named:
        args = _bind_named_proc(proc, list(args), named)
    na = len(args)
    np_ = len(params)
    if na > np_ and not proc.has_paramarray:
        raise VbaError(450, f"Wrong number of arguments: {proc.name}")
    for i, p in enumerate(params):
        if p.paramarray:
            rest = []
            for a in args[i:]:
                rest.append(a.v if isinstance(a, Slot) else a)
            slots[i] = Slot(VArray([(0, len(rest) - 1)], rest), None)
            break
        a = args[i] if i < na else MISSING
        if a is MISSING:
            if not p.optional:
                raise VbaError(449, f"Argument not optional: {p.name}")
            if p.default is not None:
                d = p.default
                slots[i] = Slot(p.co(d) if p.co is not None else d, p.co)
            elif p.co is None and not p.is_object:
                slots[i] = Slot(MISSING, None)
            else:
                slots[i] = Slot(DEFAULTS.get(p.type_name, NOTHING if p.is_object else EMPTY) if not p.is_array else EMPTY, p.co)
            continue
        if isinstance(a, Slot):
            if p.byval:
                v = slot_value(a)
                if p.is_object:
                    slots[i] = Slot(_check_object(v), None)
                else:
                    v = copy_value(v)
                    slots[i] = Slot(p.co(v) if p.co is not None else v, p.co)
            elif p.co is None or a.co is p.co or (p.is_object and a.co is None) or (p.is_array and type(a.v) is VArray):
                slots[i] = a
            else:
                v = slot_value(a)
                slots[i] = Slot(p.co(v), p.co)
        else:
            if p.is_object:
                slots[i] = Slot(_check_object(a), None)
            else:
                slots[i] = Slot(p.co(a) if p.co is not None else a, p.co)
    for idx, mk in proc.init:
        slots[idx] = mk()
    if proc.statics:
        for idx, s in proc.statics.items():
            slots[idx] = s
    ex.depth += 1
    try:
        run_frame(fr)
    finally:
        ex.depth -= 1
    if fr.active or fr.on_error:
        # Leaving a procedure clears its error state (Err stays for the caller only if the
        # error is still being raised).
        if fr.active:
            ex.err.clear()
    if proc.ret_slot >= 0:
        rs = slots[proc.ret_slot]
        return rs.v
    return EMPTY


def _check_object(v: Any) -> Any:
    if v is NOTHING or is_object(v):
        return v
    raise VbaError(13)


def _bind_named_proc(proc: Procedure, args: list, named: dict) -> list:
    names = [p.name.lower() for p in proc.params]
    out = list(args) + [MISSING] * max(0, len(proc.params) - len(args))
    for k, v in named.items():
        try:
            i = names.index(k.lower())
        except ValueError:
            raise VbaError(448, f"Named argument not found: {k}")
        out[i] = v
    while out and out[-1] is MISSING:
        out.pop()
    return out


# ---------------------------------------------------------------- member access


def get_member(ex: Execution, obj: Any, name: str, args: list, named: Optional[dict] = None) -> Any:
    """``obj.name(args)``: a property, method or field."""
    t = type(obj)
    if t is Instance:
        return ex.rt.instance_get(ex, obj, name, args, named)
    if obj is NOTHING:
        raise VbaError(91)
    if isinstance(obj, HostObject):
        return obj.vba_get(ex, name, [_unref(a) for a in args], _unref_named(named))
    if t is Record:
        i = obj.typ.index.get(name.lower())
        if i is None:
            raise VbaError(461, f"Method or data member not found: {name}")
        v = obj.fields[i]
        if args:
            return index_value(ex, v, args, named)
        return v
    if isinstance(obj, ModuleRuntime):
        return ex.rt.module_get(ex, obj, name, args, named)
    raise VbaError(424)


def let_member(ex: Execution, obj: Any, name: str, args: list, named: Optional[dict], value: Any) -> None:
    t = type(obj)
    if t is Instance:
        return ex.rt.instance_let(ex, obj, name, args, named, value, False)
    if obj is NOTHING:
        raise VbaError(91)
    if isinstance(obj, HostObject):
        return obj.vba_let(ex, name, [_unref(a) for a in args], _unref_named(named), value)
    if t is Record:
        i = obj.typ.index.get(name.lower())
        if i is None:
            raise VbaError(461, f"Method or data member not found: {name}")
        if args:
            cur = obj.fields[i]
            return let_index(ex, cur, args, value)
        co = obj.typ.coerces[i]
        obj.fields[i] = co(value) if co is not None else copy_value(value)
        return None
    if isinstance(obj, ModuleRuntime):
        return ex.rt.module_let(ex, obj, name, args, named, value, False)
    raise VbaError(424)


def set_member(ex: Execution, obj: Any, name: str, args: list, named: Optional[dict], value: Any) -> None:
    t = type(obj)
    if t is Instance:
        return ex.rt.instance_let(ex, obj, name, args, named, value, True)
    if obj is NOTHING:
        raise VbaError(91)
    if isinstance(obj, HostObject):
        return obj.vba_set(ex, name, [_unref(a) for a in args], _unref_named(named), value)
    if t is Record:
        i = obj.typ.index.get(name.lower())
        if i is None:
            raise VbaError(461)
        obj.fields[i] = value
        return None
    if isinstance(obj, ModuleRuntime):
        return ex.rt.module_let(ex, obj, name, args, named, value, True)
    raise VbaError(424)


def _unref(a: Any) -> Any:
    return slot_value(a) if isinstance(a, Slot) else a


def _unref_named(named: Optional[dict]) -> Optional[dict]:
    return {k: _unref(v) for k, v in named.items()} if named else named


def get_default(ex: Execution, obj: Any, args: list, named: Optional[dict]) -> Any:
    """An object's default member, read (``x = Range("A1")``, ``coll(1)``)."""
    t = type(obj)
    if t is Instance:
        name = obj.cls.default_member
        if name is None:
            raise VbaError(438, f"{obj.cls.name} has no default member")
        return ex.rt.instance_get(ex, obj, name, args, named)
    if obj is NOTHING:
        raise VbaError(91)
    if isinstance(obj, HostObject):
        return obj.vba_default_get(ex, [_unref(a) for a in args], _unref_named(named))
    raise VbaError(438)


def let_default(ex: Execution, obj: Any, args: list, named: Optional[dict], value: Any) -> None:
    t = type(obj)
    if t is Instance:
        name = obj.cls.default_member
        if name is None:
            raise VbaError(438, f"{obj.cls.name} has no default member")
        return ex.rt.instance_let(ex, obj, name, args, named, value, False)
    if obj is NOTHING:
        raise VbaError(91)
    if isinstance(obj, HostObject):
        return obj.vba_default_let(ex, [_unref(a) for a in args], _unref_named(named), value)
    raise VbaError(438)


def index_value(ex: Execution, v: Any, args: list, named: Optional[dict] = None) -> Any:
    """``v(args)`` where ``v`` is a value: an array element or an object's default member."""
    if type(v) is VArray:
        if named:
            raise VbaError(446)
        return v.get([to_long(_unref(a)) for a in args])
    if is_object(v):
        return get_default(ex, v, args, named)
    if not args:
        return v
    if v is EMPTY:
        raise VbaError(13)
    raise VbaError(13)


def let_index(ex: Execution, v: Any, args: list, value: Any) -> None:
    if type(v) is VArray:
        v.set([to_long(_unref(a)) for a in args], copy_value(value))
        return
    if is_object(v):
        return let_default(ex, v, args, None, value)
    raise VbaError(13)


def enumerate_object(ex: Execution, v: Any) -> Iterator[Any]:
    """What ``For Each`` goes through."""
    t = type(v)
    if t is VArray:
        return iter(list(v.data))
    if t is Instance:
        return ex.rt.instance_enum(ex, v)
    if isinstance(v, HostObject):
        return v.vba_enum(ex)
    if v is NOTHING:
        raise VbaError(91)
    raise VbaError(424 if not isinstance(v, (str, int, float)) else 92)
