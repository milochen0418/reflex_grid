"""The worksheet function registry.

Every function is described by a :class:`FnSpec` (name, arity, category, which parameters take
scalars, a signature and a one-line description) and implemented as a plain Python function over
already-evaluated arguments. :func:`call` applies Excel's automatic array lifting: passing an
array where a scalar is expected evaluates the function once per element.

Special forms that need references, laziness or hidden-row information (ROW, OFFSET, INDIRECT,
LET, LAMBDA, SUBTOTAL, AGGREGATE, TEXT…) are evaluated by :mod:`gridcore.calc.eval`.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from ..core.date import DateSystem
from ..core.value import Array, CellError, Value, XlError

# Largest array a function may build.
MAX_CELLS = 10_000_000
# Longest text a cell can hold.
MAX_TEXT = 32_767
VAR = 255

# Scalar-position lifting tables.
S: Tuple[bool, ...] = (True,)
A: Tuple[bool, ...] = (False,)


class Arg:
    """An evaluated argument. Ranges arrive as :class:`Array` (blank cells = ``None``) with
    ``from_ref`` set."""

    __slots__ = ("value", "from_ref")

    def __init__(self, value: Value, from_ref: bool = False):
        self.value = value
        self.from_ref = from_ref

    def __repr__(self) -> str:
        return f"Arg({self.value!r}{', ref' if self.from_ref else ''})"


class Ctx:
    """Services a function may need from the host."""

    date_system: DateSystem = DateSystem.D1900

    def now_serial(self) -> float:  # pragma: no cover - overridden
        from ..core.date import now_serial

        return now_serial()

    def random(self) -> float:  # pragma: no cover - overridden
        import random

        return random.random()


Impl = Callable[[List[Arg], Ctx], Value]


@dataclass
class FnSpec:
    name: str
    min_args: int
    max_args: int
    category: str
    volatile: bool
    scalar_args: Tuple[bool, ...]
    signature: str
    description: str
    impl: Impl

    def is_scalar_arg(self, i: int) -> bool:
        if i < len(self.scalar_args):
            return self.scalar_args[i]
        return self.scalar_args[-1] if self.scalar_args else False

    def accepts(self, n: int) -> bool:
        return self.min_args <= n <= self.max_args and n <= 255


REGISTRY: Dict[str, FnSpec] = {}


def fn(
    name: str,
    min_args: int,
    max_args: int,
    category: str,
    scalar_args: Sequence[bool],
    signature: str,
    description: str,
    volatile: bool = False,
):
    """Decorator registering a worksheet function."""

    def deco(impl: Impl) -> Impl:
        REGISTRY[name] = FnSpec(name, min_args, max_args, category, volatile, tuple(scalar_args), signature, description, impl)
        return impl

    return deco


def alias(name: str, target: str, description: Optional[str] = None, category: Optional[str] = None) -> None:
    """Registers ``name`` as another name of an already registered function."""
    t = REGISTRY[target]
    REGISTRY[name] = FnSpec(
        name,
        t.min_args,
        t.max_args,
        category or t.category,
        t.volatile,
        t.scalar_args,
        t.signature.replace(target, name, 1),
        description or t.description,
        t.impl,
    )


def lookup(name: str) -> Optional[FnSpec]:
    """Looks a function up by name (case-insensitive; ``_xlfn.``/``_xlws.`` prefixes ignored)."""
    n = name.strip().upper()
    while True:
        if n.startswith("_XLFN."):
            n = n[6:]
        elif n.startswith("_XLWS."):
            n = n[6:]
        else:
            break
    return REGISTRY.get(n)


def finish(v: Value) -> Value:
    """Turns an implementation result into a cell value: non-finite numbers become ``#NUM!``,
    text longer than 32,767 characters becomes ``#VALUE!``."""
    t = type(v)
    if t is float:
        if math.isfinite(v):  # type: ignore[arg-type]
            return 0.0 if v == 0.0 else v
        return CellError.NUM
    if t is int:
        return float(v)  # type: ignore[arg-type]
    if t is str and len(v) > MAX_TEXT:  # type: ignore[arg-type]
        return CellError.VALUE
    return v


def _invoke(spec: FnSpec, args: List[Arg], ctx: Ctx) -> Value:
    try:
        return finish(spec.impl(args, ctx))
    except XlError as e:
        return e.error
    except (OverflowError, ZeroDivisionError):
        return CellError.NUM
    except RecursionError:
        return CellError.CALC
    except MemoryError:
        return CellError.NUM
    except Exception:  # never crash: a bug in one function is that cell's #VALUE!
        return CellError.VALUE


def call(spec: FnSpec, args: List[Arg], ctx: Ctx) -> Value:
    """Calls a function with array lifting applied. Wrong arity gives ``#VALUE!``."""
    if not spec.accepts(len(args)):
        return CellError.VALUE
    rows = cols = 0
    lifted = False
    for i, a in enumerate(args):
        if isinstance(a.value, Array) and spec.is_scalar_arg(i):
            lifted = True
            rows = max(rows, a.value.rows)
            cols = max(cols, a.value.cols)
    if not lifted:
        return _invoke(spec, args, ctx)
    if rows * cols > MAX_CELLS:
        return CellError.NUM
    cell_args = [Arg(a.value, a.from_ref) for a in args]
    data: List[Value] = []
    for r in range(rows):
        for c in range(cols):
            for i, a in enumerate(args):
                if isinstance(a.value, Array) and spec.is_scalar_arg(i):
                    v = a.value.get_broadcast(r, c)
                    if isinstance(v, Array):
                        v = v.data[0] if v.data else None
                    cell_args[i].value = v
            out = _invoke(spec, cell_args, ctx)
            if isinstance(out, Array):
                out = out.data[0] if out.data else None
            data.append(out)
    if rows == 1 and cols == 1:
        return data[0]
    return Array(rows, cols, data)


def all_functions() -> List[FnSpec]:
    return list(REGISTRY.values())
