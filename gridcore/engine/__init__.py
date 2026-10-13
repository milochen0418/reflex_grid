"""The engine: documents, every user-visible action as a command (``Session.execute``),
undo/redo, the clipboard, and what the grid draws (:mod:`.viewport`)."""

from . import (  # noqa: F401  (imported for command registration)
    cmd_cf,
    cmd_data,
    cmd_edit,
    cmd_file,
    cmd_format,
    cmd_formulas,
    cmd_insert,
    cmd_sheet,
    cmd_view,
    cmd_whatif,
)
from .cmd_file import open_bytes
from .selection import Selection
from .session import COMMANDS, DocState, EngineError, Session
from .viewport import geometry, objects, ui_state, viewport

__all__ = ["COMMANDS", "DocState", "EngineError", "Selection", "Session", "geometry", "objects", "open_bytes", "ui_state", "viewport"]
