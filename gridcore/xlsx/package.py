"""The OPC zip package: bounded part reads and relationships."""

from __future__ import annotations

import io
import posixpath
import zipfile
from dataclasses import dataclass
from typing import Dict, List, Optional
from xml.etree import ElementTree as ET

from .xml import attr, kids, local

MAX_PART = 512 * 1024 * 1024
MAX_PACKAGE = 1024 * 1024 * 1024


class XlsxError(Exception):
    """The file is not a readable spreadsheet package."""


@dataclass
class Rel:
    id: str
    kind: str  # last segment of the relationship type (worksheet, styles, …)
    type: str
    target: str  # resolved part name (no leading slash) or the external URL
    external: bool


class Package:
    def __init__(self, data: bytes):
        try:
            self.zf = zipfile.ZipFile(io.BytesIO(data))
        except (zipfile.BadZipFile, OSError, ValueError) as e:
            raise XlsxError(f"the file is not a valid zip package: {e}") from None
        self.names: Dict[str, str] = {}
        total = 0
        for info in self.zf.infolist():
            self.names[info.filename.lower().lstrip("/")] = info.filename
            total += info.file_size
        if total > MAX_PACKAGE:
            raise XlsxError("the file is too large")
        self.read_bytes = 0

    def has(self, name: str) -> bool:
        return name.lower().lstrip("/") in self.names

    def read(self, name: str) -> Optional[bytes]:
        real = self.names.get(name.lower().lstrip("/"))
        if real is None:
            return None
        with self.zf.open(real) as f:
            data = f.read(MAX_PART + 1)
        if len(data) > MAX_PART:
            raise XlsxError(f"part {name} is too large")
        self.read_bytes += len(data)
        if self.read_bytes > MAX_PACKAGE:
            raise XlsxError("the file is too large")
        return data

    def read_xml(self, name: str) -> Optional[ET.Element]:
        data = self.read(name)
        if data is None:
            return None
        try:
            return ET.fromstring(data)
        except ET.ParseError as e:
            raise XlsxError(f"{name} contains invalid XML: {e}") from None

    def rels(self, part: str) -> List[Rel]:
        """Relationships of ``part`` ("" for the package root)."""
        if part:
            d, b = posixpath.split(part)
            rels_name = posixpath.join(d, "_rels", b + ".rels")
        else:
            d, rels_name = "", "_rels/.rels"
        try:
            root = self.read_xml(rels_name)
        except XlsxError:
            return []
        if root is None:
            return []
        out = []
        for r in kids(root, "Relationship"):
            rid = attr(r, "Id") or ""
            ty = attr(r, "Type") or ""
            target = attr(r, "Target") or ""
            external = (attr(r, "TargetMode") or "") == "External"
            if not external:
                if target.startswith("/"):
                    target = target.lstrip("/")
                else:
                    target = posixpath.normpath(posixpath.join(d, target)) if target else target
            out.append(Rel(rid, ty.rsplit("/", 1)[-1], ty, target, external))
        return out
