"""A workbook's VBA project: its modules' source code and what the file needs to write the project
back (references, project properties, the original ``vbaProject.bin``).

Projects are immutable: a change makes a new :class:`VbaProject`, so workbook snapshots (undo)
can share them. ``base`` is the ``vbaProject.bin`` read from the file; as long as ``modified``
is false it is written back byte for byte (with Excel's compiled code and any signature), and
once the project changes it is rebuilt from the source, which Excel recompiles when it opens the
file. Reading and writing the binary format is :mod:`gridcore.xlsx.ovba`.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field, replace
from typing import Dict, List, Optional, Tuple

# Module kinds as the VBA editor groups them.
STANDARD = "standard"   # Modules
CLASS = "class"         # Class Modules
DOCUMENT = "document"   # Microsoft Excel Objects: ThisWorkbook and one per sheet
FORM = "form"           # Forms (the designer is kept, the form itself doesn't run)

KINDS = (STANDARD, CLASS, DOCUMENT, FORM)

# The type libraries behind ThisWorkbook and the worksheet modules (their ``VB_Base`` attribute),
# and the base Office documents give class modules (Excel won't compile a class without it).
WORKBOOK_BASE = "0{00020819-0000-0000-C000-000000000046}"
WORKSHEET_BASE = "0{00020820-0000-0000-C000-000000000046}"
CLASS_BASE = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"

_NAME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,30}$")
_ATTR_RE = re.compile(r'^\s*Attribute\s+([A-Za-z_][A-Za-z0-9_]*)\.(VB_[A-Za-z_]+)', re.I)


def valid_module_name(name: str) -> Optional[str]:
    """An error message when ``name`` can't name a module, else ``None``."""
    if not name:
        return "A module needs a name."
    if not _NAME_RE.match(name):
        return "Module names start with a letter and contain only letters, digits and underscores (at most 31)."
    return None


@dataclass(frozen=True)
class VbaReference:
    """A project reference (``Tools › References`` in the VBA editor): a registered type library
    or another project. ``raw`` is its records in the ``dir`` stream, written back unchanged."""

    name: str
    kind: str  # registered | project | control
    libid: str
    raw: bytes = b""

    @property
    def description(self) -> str:
        """The library's display name (the last ``#`` field of a registered libid)."""
        parts = self.libid.split("#")
        return parts[-1] if len(parts) >= 4 and parts[-1] else self.name


@dataclass(frozen=True)
class VbaModule:
    """One module. ``code`` is the source as the VBA editor shows it (``\\n`` line ends, without
    the hidden ``Attribute VB_…`` lines); ``attributes`` are the module's leading attribute lines
    and ``member_attributes`` the hidden lines after procedure headers (descriptions, shortcut
    keys), as ``(procedure, line)``."""

    name: str
    kind: str
    code: str = ""
    attributes: Tuple[str, ...] = ()
    member_attributes: Tuple[Tuple[str, str], ...] = ()
    stream: str = ""
    doc_string: str = ""
    help_context: int = 0
    read_only: bool = False
    private: bool = False

    @property
    def stream_name(self) -> str:
        return self.stream or self.name

    def source(self) -> str:
        """The full module text as stored in the file (attributes included, ``\\n`` line ends)."""
        pending: Dict[str, List[str]] = {}
        for member, line in self.member_attributes:
            pending.setdefault(member.lower(), []).append(line)
        body = self.code.split("\n") if self.code else []
        out: List[str] = list(self.attributes)
        i = 0
        while i < len(body):
            # A declaration may continue over several lines (` _`); attributes follow its last line.
            j = i
            while body[j].rstrip().endswith(" _") and j + 1 < len(body):
                j += 1
            out.extend(body[i:j + 1])
            name = member_decl_name(" ".join(body[i:j + 1]))
            if name is not None and name.lower() in pending:
                out.extend(pending.pop(name.lower()))
            i = j + 1
        return "\n".join(out)


_DECL_RE = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)"
    r"|^\s*(?:Public|Private|Dim|Global)\s+(?:WithEvents\s+)?([A-Za-z_][A-Za-z0-9_]*)",
    re.I,
)


def member_decl_name(line: str) -> Optional[str]:
    """The member a procedure header or module-level variable declaration declares, else ``None``."""
    m = _DECL_RE.match(line)
    if not m:
        return None
    name = m.group(1) or m.group(2)
    if name and name.lower() in ("sub", "function", "property", "const", "type", "enum", "declare", "event"):
        return None
    return name


def split_source(text: str) -> Tuple[str, Tuple[str, ...], Tuple[Tuple[str, str], ...]]:
    """Splits a module's stored text into the code the editor shows, the leading attribute
    lines and the attribute lines inside the code (kept with the procedure they follow)."""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    head: List[str] = []
    i = 0
    while i < len(lines) and re.match(r"^\s*Attribute\s+VB_", lines[i], re.I):
        head.append(lines[i].rstrip())
        i += 1
    body: List[str] = []
    members: List[Tuple[str, str]] = []
    for line in lines[i:]:
        m = _ATTR_RE.match(line)
        if m:
            members.append((m.group(1), line.rstrip()))
            continue
        body.append(line)
    while body and body[-1] == "":
        body.pop()
    return "\n".join(body), tuple(head), tuple(members)


def default_attributes(name: str, kind: str, base: str = "") -> Tuple[str, ...]:
    """The attribute lines Excel gives a new module of this kind."""
    lines = [f'Attribute VB_Name = "{name}"']
    if kind == STANDARD:
        return tuple(lines)
    if kind == DOCUMENT:
        lines.append(f'Attribute VB_Base = "{base or WORKSHEET_BASE}"')
    elif kind == CLASS:
        lines.append(f'Attribute VB_Base = "{base or CLASS_BASE}"')
    lines += [
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = False",
        f"Attribute VB_PredeclaredId = {'True' if kind in (DOCUMENT, FORM) else 'False'}",
        f"Attribute VB_Exposed = {'True' if kind == DOCUMENT else 'False'}",
    ]
    if kind in (DOCUMENT, CLASS):
        lines += ["Attribute VB_TemplateDerived = False", f"Attribute VB_Customizable = {'True' if kind == DOCUMENT else 'False'}"]
    return tuple(lines)


def renamed_attributes(attrs: Tuple[str, ...], name: str) -> Tuple[str, ...]:
    out = []
    for a in attrs:
        if re.match(r"^\s*Attribute\s+VB_Name\s*=", a, re.I):
            out.append(f'Attribute VB_Name = "{name}"')
        else:
            out.append(a)
    return tuple(out)


@dataclass(frozen=True)
class VbaProject:
    modules: Tuple[VbaModule, ...] = ()
    name: str = "VBAProject"
    code_page: int = 1252
    references: Tuple[VbaReference, ...] = ()
    sys_kind: int = 3  # 0 Win16, 1 Win32, 2 Macintosh, 3 Win64
    lcid: int = 0x409
    lcid_invoke: int = 0x409
    version_major: int = 1
    version_minor: int = 0
    compat_version: Optional[int] = None
    doc_string: str = ""
    help_file: str = ""
    help_context: int = 0
    lib_flags: int = 0
    constants: str = ""
    project_id: str = ""
    # Encrypted protection state, password and visibility (CMG, DPB, GC) as the PROJECT stream had them.
    protection: Tuple[str, str, str] = ("", "", "")
    # Locked for viewing in the VBA editor (the code still runs).
    locked: bool = False
    # The PROJECT stream's lines that aren't about modules (kept in order when rewriting).
    project_lines: Tuple[str, ...] = ()
    # The file's own vbaProject.bin and the parts signing it (written back while unmodified).
    base: Optional[bytes] = field(default=None, repr=False, compare=False)
    signatures: Tuple[Tuple[str, str, bytes], ...] = field(default=(), repr=False, compare=False)
    modified: bool = False
    # A project whose binary couldn't be read: kept byte for byte, but not editable or runnable.
    unreadable: str = ""
    warnings: Tuple[str, ...] = ()

    # ------------------------------------------------------------ lookups

    def module(self, name: str) -> Optional[VbaModule]:
        n = name.lower()
        for m in self.modules:
            if m.name.lower() == n:
                return m
        return None

    def index(self, name: str) -> Optional[int]:
        n = name.lower()
        for i, m in enumerate(self.modules):
            if m.name.lower() == n:
                return i
        return None

    def document_names(self) -> List[str]:
        return [m.name for m in self.modules if m.kind == DOCUMENT]

    def new_module_name(self, kind: str) -> str:
        stem = {STANDARD: "Module", CLASS: "Class", FORM: "UserForm"}.get(kind, "Sheet")
        n = 1
        while self.module(f"{stem}{n}") is not None:
            n += 1
        return f"{stem}{n}"

    # ------------------------------------------------------------ changes (new projects)

    def _changed(self, modules: Tuple[VbaModule, ...]) -> "VbaProject":
        return replace(self, modules=modules, modified=True, signatures=())

    def with_module(self, module: VbaModule) -> "VbaProject":
        """Adds ``module`` or replaces the module of the same name."""
        i = self.index(module.name)
        mods = list(self.modules)
        if i is None:
            mods.append(module)
        else:
            mods[i] = module
        return self._changed(tuple(mods))

    def without_module(self, name: str) -> "VbaProject":
        return self._changed(tuple(m for m in self.modules if m.name.lower() != name.lower()))

    def renamed_module(self, old: str, new: str) -> "VbaProject":
        mods = []
        for m in self.modules:
            if m.name.lower() == old.lower():
                stream = m.stream if m.stream and m.stream.lower() != m.name.lower() else new
                m = replace(m, name=new, stream=stream, attributes=renamed_attributes(m.attributes, new))
            mods.append(m)
        return self._changed(tuple(mods))

    def with_code(self, name: str, code: str) -> "VbaProject":
        m = self.module(name)
        if m is None:
            raise KeyError(name)
        return self.with_module(replace(m, code=normalize_code(code)))


def normalize_code(code: str) -> str:
    text = code.replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.rstrip() for ln in text.split("\n")]
    while lines and lines[-1] == "":
        lines.pop()
    return "\n".join(lines)


def new_project(sheet_code_names: List[str], workbook_code_name: str = "ThisWorkbook") -> VbaProject:
    """An empty project as Excel starts one: ThisWorkbook and a module per sheet."""
    mods = [VbaModule(workbook_code_name, DOCUMENT, "", default_attributes(workbook_code_name, DOCUMENT, WORKBOOK_BASE))]
    for n in sheet_code_names:
        mods.append(VbaModule(n, DOCUMENT, "", default_attributes(n, DOCUMENT, WORKSHEET_BASE)))
    return VbaProject(modules=tuple(mods), modified=True)


def reconcile(project: Optional[VbaProject], workbook_code_name: str, sheet_code_names: List[str]) -> Tuple[Optional[VbaProject], str, List[str]]:
    """Matches a project's document modules to the workbook's sheets, as Excel requires: one for
    the workbook (``ThisWorkbook``) and one per sheet, named by the sheet's code name. Sheets
    without a code name get ``SheetN``; modules of sheets that are gone are dropped. Returns the
    project (the same object when nothing had to change) and the code names to use."""
    if project is None:
        return None, workbook_code_name, list(sheet_code_names)
    taken = {m.name.lower() for m in project.modules if m.kind != DOCUMENT}
    wb_name = workbook_code_name
    if not wb_name:
        wb_name = next((m.name for m in project.modules if m.kind == DOCUMENT and any(WORKBOOK_BASE[2:10] in a for a in m.attributes)), "ThisWorkbook")
    names: List[str] = []
    seen = {wb_name.lower()} | taken
    for cn in sheet_code_names:
        if cn and cn.lower() not in seen:
            names.append(cn)
            seen.add(cn.lower())
        else:
            names.append("")
    for i, cn in enumerate(names):
        if not cn:
            n = 1
            while f"sheet{n}" in seen:
                n += 1
            names[i] = f"Sheet{n}"
            seen.add(f"sheet{n}")
    if project.unreadable:
        return project, wb_name, names
    wanted = {wb_name.lower(): WORKBOOK_BASE}
    for cn in names:
        wanted[cn.lower()] = WORKSHEET_BASE
    mods: List[VbaModule] = []
    present = set()
    for m in project.modules:
        if m.kind == DOCUMENT:
            if m.name.lower() not in wanted or m.name.lower() in present:
                continue
            present.add(m.name.lower())
        mods.append(m)
    for name in [wb_name] + names:
        if name.lower() not in present:
            mods.append(VbaModule(name, DOCUMENT, "", default_attributes(name, DOCUMENT, wanted[name.lower()])))
            present.add(name.lower())
    if tuple(mods) == project.modules:
        return project, wb_name, names
    return project._changed(tuple(mods)), wb_name, names
