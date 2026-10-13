"""File formats: XLSX (transitional and strict) read/write and CSV/TSV import/export.

Implemented from ECMA-376 and observed Excel behaviour. Reading is lenient: unsupported or broken optional parts are skipped and
reported as warnings; only a missing/invalid package or workbook part is an error."""

from __future__ import annotations

from .csv_io import read_csv, write_csv
from .package import XlsxError
from .read import read_xlsx
from .write import write_xlsx


def sniff(data: bytes) -> str:
    """Guesses the format from content: ``xlsx``, ``ods``, ``xls``, ``csv`` or ``unknown``."""
    if data[:4] in (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08"):
        if b"mimetypeapplication/vnd.oasis.opendocument.spreadsheet" in data[:200]:
            return "ods"
        return "xlsx"
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "xls"
    if not data:
        return "unknown"
    head = data[:4096]
    if data[:3] == b"\xef\xbb\xbf" or data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return "csv"
    control = sum(1 for b in head if b < 9 or 14 <= b < 32 or b == 127)
    if b"\x00" in head or control * 20 > len(head):
        return "unknown"
    return "csv"


__all__ = ["XlsxError", "read_csv", "read_xlsx", "sniff", "write_csv", "write_xlsx"]
