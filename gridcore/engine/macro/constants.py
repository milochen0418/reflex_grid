"""Excel's and Office's enumeration constants as VBA code uses them (``xlUp``, ``xlPasteValues``,
``msoTrue``…) and their enumeration types (``Dim d As XlDirection`` declares a Long, and
``XlDirection.xlUp`` names a member). The table is generated from Excel's type libraries
(:mod:`.constants_data`)."""

from __future__ import annotations

from typing import Dict

from .constants_data import DATA

CONSTANTS: Dict[str, int] = {}
ENUMS: Dict[str, Dict[str, int]] = {}


def _load() -> None:
    members: Dict[str, int] = {}
    for line in DATA.splitlines():
        if not line:
            continue
        if line.startswith("# "):
            members = ENUMS.setdefault(line[2:].split(".", 1)[1].lower(), {})
            continue
        name, value = line.split()
        low = name.lower()
        members[low] = int(value)
        CONSTANTS.setdefault(low, members[low])


_load()
