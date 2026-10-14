"""VBA errors: compile errors (syntax, undeclared names…) and run-time errors with VBA's numbers
and messages, which ``Err.Number`` / ``Err.Description`` report and ``On Error`` handles."""

from __future__ import annotations

from typing import Optional

MESSAGES = {
    3: "Return without GoSub",
    5: "Invalid procedure call or argument",
    6: "Overflow",
    7: "Out of memory",
    9: "Subscript out of range",
    10: "This array is fixed or temporarily locked",
    11: "Division by zero",
    13: "Type mismatch",
    14: "Out of string space",
    16: "Expression too complex",
    17: "Can't perform requested operation",
    18: "User interrupt occurred",
    20: "Resume without error",
    28: "Out of stack space",
    35: "Sub or Function not defined",
    51: "Internal error",
    52: "Bad file name or number",
    53: "File not found",
    54: "Bad file mode",
    55: "File already open",
    57: "Device I/O error",
    58: "File already exists",
    62: "Input past end of file",
    67: "Too many files",
    70: "Permission denied",
    75: "Path/File access error",
    76: "Path not found",
    91: "Object variable or With block variable not set",
    92: "For loop not initialized",
    93: "Invalid pattern string",
    94: "Invalid use of Null",
    380: "Invalid property value",
    381: "Invalid property-array index",
    382: "Set not supported at runtime",
    383: "Set not supported (read-only property)",
    385: "Need property-array index",
    387: "Set not permitted",
    393: "Get not supported at runtime",
    394: "Get not supported (write-only property)",
    422: "Property not found",
    423: "Property or method not found",
    424: "Object required",
    425: "Invalid object use",
    429: "ActiveX component can't create object",
    430: "Class does not support Automation or does not support expected interface",
    438: "Object doesn't support this property or method",
    440: "Automation error",
    443: "Automation object does not have a default value",
    445: "Object doesn't support this action",
    446: "Object doesn't support named arguments",
    448: "Named argument not found",
    449: "Argument not optional",
    450: "Wrong number of arguments or invalid property assignment",
    451: "Property let procedure not defined and property get procedure did not return an object",
    452: "Invalid ordinal",
    453: "Specified DLL function not found",
    457: "This key is already associated with an element of this collection",
    458: "Variable uses an Automation type not supported in Visual Basic",
    459: "Object or class does not support the set of events",
    461: "Method or data member not found",
    744: "Search text not found",
    1004: "Application-defined or object-defined error",
    5017: "Application-defined or object-defined error",
}


def message(number: int) -> str:
    return MESSAGES.get(number, "Application-defined or object-defined error")


class VbaSyntaxError(Exception):
    """A compile error: the code can't run until it is fixed. ``line`` is the physical line
    in its module (1-based)."""

    def __init__(self, msg: str, line: int = 0, module: str = ""):
        super().__init__(msg)
        self.msg = msg
        self.line = line
        self.module = module

    def __str__(self) -> str:
        where = f"{self.module}, line {self.line}: " if self.module else (f"line {self.line}: " if self.line else "")
        return f"Compile error: {where}{self.msg}"


class VbaError(Exception):
    """A run-time error. ``number`` and ``description`` are what ``Err`` reports; ``module``,
    ``proc`` and ``line`` say where it happened (filled in as it leaves the procedure)."""

    def __init__(self, number: int, description: Optional[str] = None, source: str = "", help_file: str = "", help_context: int = 0):
        self.number = int(number)
        self.description = description if description is not None else message(self.number)
        self.source = source
        self.help_file = help_file
        self.help_context = help_context
        self.module = ""
        self.proc = ""
        self.line = 0
        super().__init__(self.description)

    def __str__(self) -> str:
        return f"Run-time error '{self.number}': {self.description}"

    def where(self) -> str:
        if not self.module:
            return ""
        return f"{self.module}.{self.proc}, line {self.line}" if self.proc else f"{self.module}, line {self.line}"


class MacroStopped(Exception):
    """The macro ended early: ``End`` or ``Stop``, the user cancelled, or it exceeded a limit
    (``reason``: end, stop, cancelled, timeout, steps)."""

    def __init__(self, reason: str, msg: str = ""):
        super().__init__(msg or reason)
        self.reason = reason
        self.msg = msg


class NotSupported(VbaError):
    """Something ReflexGrid's sandbox doesn't provide (files, shell, Windows API, COM…)."""

    def __init__(self, what: str, number: int = 445):
        super().__init__(number, f"{what} isn't available in ReflexGrid macros")
        self.what = what
