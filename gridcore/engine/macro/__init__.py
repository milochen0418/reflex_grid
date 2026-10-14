"""Excel macros (VBA) on the engine: the Excel object model VBA code uses (Application, Range,
Worksheet…), running macros, event macros and worksheet functions written in VBA.

Importing this package connects VBA to the engine: formulas can call the workbook's public
VBA Functions, and commands run its event macros (when enabled)."""

from ...calc import eval as _eval
from .. import session as _session
from .events import after_command, udf
from .project import compile_vba, diagnostics, procedures, runtime_for
from .runner import MacroRun, call_macro_job

_eval.UDF_HOOK = udf
_session.EVENT_HOOK = after_command

__all__ = ["MacroRun", "after_command", "call_macro_job", "compile_vba", "diagnostics", "procedures", "runtime_for", "udf"]
