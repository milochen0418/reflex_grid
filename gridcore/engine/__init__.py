"""The engine: documents, every user-visible action as a command (``Session.execute``),
undo/redo, the clipboard, what the grid draws (:mod:`.viewport`) and several people editing
one workbook (:mod:`.collab`)."""

from . import (  # noqa: F401  (imported for command registration)
    cmd_cf,
    cmd_data,
    cmd_edit,
    cmd_file,
    cmd_format,
    cmd_formulas,
    cmd_insert,
    cmd_macro,
    cmd_sheet,
    cmd_view,
    cmd_whatif,
)
from . import macro as _macro  # noqa: F401  (VBA: worksheet functions and event macros)
from .cmd_file import open_bytes
from .collab import CollabSession, SharedBook, View
from .selection import Selection
from .session import COMMANDS, DocState, EngineError, Session
from .viewport import geometry, objects, ui_state, viewport

__all__ = ["COMMANDS", "CollabSession", "DocState", "EngineError", "Selection", "Session", "SharedBook", "View", "geometry", "objects", "open_bytes", "ui_state", "viewport"]
