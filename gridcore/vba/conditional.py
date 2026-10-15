"""Conditional compilation: ``#Const``, ``#If … Then``, ``#ElseIf … Then``, ``#Else``, ``#End If`` (MS-VBAL 3.4).

The directives are applied to a module's source before it is tokenized: the directive lines and the lines
of the branches not taken become empty lines, so line numbers stay those of the source (and code for
another platform, such as ``Declare`` statements for Mac libraries, is never parsed).

ReflexGrid compiles as 64-bit VBA 7 for Windows, as it follows Excel for Windows elsewhere: ``VBA7``,
``VBA6``, ``Win64`` and ``Win32`` are True, ``Mac`` and ``Win16`` False. A project's own conditional
compilation arguments (Project Properties, ``VbaProject.constants``: ``DEBUG_MODE = 1 : LEVEL = -2``)
apply to every module; a ``#Const`` applies to the rest of its module.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Mapping, Optional, Tuple

from . import values as V
from .errors import VbaSyntaxError

TRUE = -1
BUILTIN: Dict[str, Any] = {"vba7": TRUE, "vba6": TRUE, "win64": TRUE, "win32": TRUE, "win16": 0, "mac": 0}

_DIRECTIVE = re.compile(r"^[ \t]*#[ \t]*(const|if|elseif|else|end[ \t]*if)\b(.*)$", re.I)
_ANY_DIRECTIVE = re.compile(r"^[ \t]*#[ \t]*(?:const|if|elseif|else|end[ \t]*if)\b", re.I | re.M)
_CONST = re.compile(r"^\s*([A-Za-z][A-Za-z0-9_]*)\s*=(.*)$")
_PROJECT_CONST = re.compile(r"^\s*([A-Za-z][A-Za-z0-9_]*)\s*=\s*(-?\d+)\s*$")


def project_constants(text: str) -> Dict[str, Any]:
    """A project's conditional compilation arguments: ``Name = integer`` pairs separated by ``:``."""
    out: Dict[str, Any] = {}
    for part in (text or "").split(":"):
        m = _PROJECT_CONST.match(part)
        if m:
            out[m.group(1).lower()] = int(m.group(2))
    return out


def has_directives(src: str) -> bool:
    return "#" in src and _ANY_DIRECTIVE.search(src) is not None


def _strip_comment(text: str) -> str:
    """The directive's text without a trailing ``'`` comment (quotes respected)."""
    in_str = False
    for i, ch in enumerate(text):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return text[:i]
    return text


def evaluate(expr: str, constants: Mapping[str, Any], line: int = 0, module: str = "Module") -> Any:
    """The value of a conditional compilation expression: literals, compilation constants (an
    unknown one is Empty) and operators."""
    from . import ast as A
    from .compiler import _BINOPS, _lit_value
    from .errors import VbaError
    from .parser import parse_expression

    def value(e: Any) -> Any:
        if isinstance(e, A.Lit):
            return _lit_value(e)
        if isinstance(e, A.Paren):
            return value(e.expr)
        if isinstance(e, A.Ident):
            return constants.get(e.name.lower(), V.EMPTY)
        if isinstance(e, A.Unary):
            x = value(e.expr)
            if e.op == "-":
                return V.neg(x, True)
            if e.op == "+":
                return x
            return V.op_not(x)
        if isinstance(e, A.Binary) and e.op in _BINOPS:
            return _BINOPS[e.op](value(e.left), value(e.right))
        raise VbaSyntaxError("Constant expression required", line, module)

    try:
        return value(parse_expression(expr))
    except VbaSyntaxError as e:
        raise VbaSyntaxError(e.msg, line, module)
    except VbaError as e:
        raise VbaSyntaxError(f"Constant expression: {e.description}", line, module)


def _truth(v: Any, line: int, module: str) -> bool:
    from .errors import VbaError

    try:
        return V.to_boolean(v)
    except VbaError:
        raise VbaSyntaxError("Type mismatch", line, module)


def preprocess(src: str, constants: Optional[Mapping[str, Any]] = None, module: str = "Module") -> str:
    """The module's source with its directives applied (unchanged when it has none)."""
    if not has_directives(src):
        return src
    lines = src.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    consts: Dict[str, Any] = dict(BUILTIN)
    consts.update({k.lower(): v for k, v in (constants or {}).items()})
    out: List[str] = []
    # One frame per open #If: (whether its enclosing code is active, whether a branch was taken, its line).
    stack: List[Tuple[bool, bool, int]] = []
    active = True
    for no, text in enumerate(lines, 1):
        m = _DIRECTIVE.match(text)
        if not m:
            out.append(text if active else "")
            continue
        kind = re.sub(r"[ \t]+", "", m.group(1).lower())
        rest = _strip_comment(m.group(2)).strip()
        if kind == "const":
            if active:
                c = _CONST.match(rest)
                if not c or not c.group(2).strip():
                    raise VbaSyntaxError("Expected: identifier = expression", no, module)
                consts[c.group(1).lower()] = evaluate(c.group(2), consts, no, module)
        elif kind in ("if", "elseif"):
            if not re.search(r"\bthen\s*$", rest, re.I):
                raise VbaSyntaxError("Expected: Then", no, module)
            cond_text = re.sub(r"\bthen\s*$", "", rest, flags=re.I).strip()
            if kind == "if":
                taken = active and _truth(evaluate(cond_text, consts, no, module), no, module)
                stack.append((active, taken, no))
                active = taken
            else:
                if not stack:
                    raise VbaSyntaxError("#ElseIf without #If", no, module)
                outer, done, start = stack[-1]
                taken = outer and not done and _truth(evaluate(cond_text, consts, no, module), no, module)
                stack[-1] = (outer, done or taken, start)
                active = taken
        elif kind == "else":
            if not stack:
                raise VbaSyntaxError("#Else without #If", no, module)
            outer, done, start = stack[-1]
            active = outer and not done
            stack[-1] = (outer, True, start)
        else:  # end if
            if not stack:
                raise VbaSyntaxError("#End If without #If block", no, module)
            active = stack.pop()[0]
        out.append("")
    if stack:
        raise VbaSyntaxError("#If block without #End If", stack[-1][2], module)
    return "\n".join(out)
