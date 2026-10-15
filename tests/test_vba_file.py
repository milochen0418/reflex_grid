"""VBA projects in files: the compound file container, MS-OVBA compression and obfuscation,
vbaProject.bin read/write and .xlsm packages (checked with openpyxl)."""

import base64
import io
import os
import random
import struct
import zipfile
from dataclasses import replace

import pytest

from gridcore.engine import Session
from gridcore.model.vba import CLASS, DOCUMENT, STANDARD, VbaModule, default_attributes, new_project, reconcile, split_source
from gridcore.xlsx import read_xlsx, write_xlsx
from gridcore.xlsx.cfb import NOSTREAM, CfbError, Storage, read_cfb, write_cfb
from gridcore.xlsx.ovba import _rec, _records, compress, decompress, decrypt, encrypt, project_key, read_project, write_project

# [MS-OVBA] 3.2: the compression examples, byte for byte.
VECTORS = {
    b"abcdefghijklmnopqrstuv.": "01 19 B0 00 61 62 63 64 65 66 67 68 00 69 6A 6B 6C 6D 6E 6F 70 00 71 72 73 74 75 76 2E",
    b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa": "01 2F B0 00 23 61 61 61 62 63 64 65 82 66 00 70 61 67 68 69 6A 01 38 08 61 6B 6C 00 20 6D 6E 6F 70 06 71 02 70 04 00 72 73 74 75 76 10 77 78 79 7A 00 2C",
    b"a" * 73: "01 03 B0 02 61 45 00",
}


@pytest.mark.parametrize("text", list(VECTORS))
def test_compression_matches_specification(text):
    want = bytes.fromhex(VECTORS[text])
    assert compress(text) == want
    assert decompress(want) == text


def test_compression_round_trips():
    rnd = random.Random(7)
    for n in (0, 1, 3, 4095, 4096, 4097, 9000, 70000):
        text = bytes(rnd.choice(b"Sub Foo()\r\n    Range(\"A1\").Value = 1\r\nEnd Sub\r\n") for _ in range(n))
        assert decompress(compress(text)) == text
        noise = os.urandom(n)
        back = decompress(compress(noise))
        assert back[:n] == noise and set(back[n:]) <= {0}


def test_obfuscation_matches_specification():
    # [MS-OVBA] 3.1: the PROJECT stream example's CMG, DPB and GC.
    pid = "{917DED54-440B-4FD1-A5C1-74ACF261E600}"
    assert project_key(pid) == 0xDF
    assert encrypt(b"\0\0\0\0", pid, 0x07, b"\x07\x07\x07") == "0705D8E3D8EDDBF1DBF1DBF1DBF1"
    assert encrypt(b"\0", pid, 0x0E, b"\x07\x07\x07") == "0E0CD1ECDFF4E7F5E7F5E7"
    assert encrypt(b"\xff", pid, 0x15, b"\x07\x07") == "1517CAF1D6F9D7F9D706"
    assert decrypt("0705D8E3D8EDDBF1DBF1DBF1DBF1") == b"\0\0\0\0"
    assert decrypt("1517CAF1D6F9D7F9D706") == b"\xff"
    assert decrypt(encrypt(b"hello", "{X}")) == b"hello"
    assert decrypt("zz") is None


# ---------------------------------------------------------------- compound files


def _check_tree(data: bytes) -> None:
    """Every storage's children form a binary search tree in [MS-CFB] order with equal black
    heights and no red node with a red child."""
    import struct as st

    dir_start = st.unpack_from("<I", data, 48)[0]
    fat = []
    nfat = st.unpack_from("<I", data, 44)[0]
    for s in st.unpack_from("<109I", data, 76)[:nfat]:
        fat.extend(st.unpack_from("<128I", data, 512 + s * 512))
    raw = b""
    s = dir_start
    while s != 0xFFFFFFFE:
        raw += data[512 + s * 512:512 + (s + 1) * 512]
        s = fat[s]
    entries = []
    for off in range(0, len(raw), 128):
        e = raw[off:off + 128]
        nl = st.unpack_from("<H", e, 64)[0]
        entries.append((e[:max(nl - 2, 0)].decode("utf-16-le"), e[66], e[67], *st.unpack_from("<III", e, 68)))

    def key(i):
        n = entries[i][0]
        return (len(n), n.upper())

    def black_height(i, lo, hi):
        if i == NOSTREAM:
            return 1
        name, kind, color, left, right, child = entries[i]
        assert (lo is None or key(i) > lo) and (hi is None or key(i) < hi)
        if color == 0:
            for c in (left, right):
                assert c == NOSTREAM or entries[c][2] == 1
        a, b = black_height(left, lo, key(i)), black_height(right, key(i), hi)
        assert a == b
        if kind in (1, 5):
            black_height(child, None, None)
        return a + (1 if color == 1 else 0)

    black_height(entries[0][5], None, None)


def test_cfb_round_trip_and_tree():
    root = Storage()
    for i in range(37):
        root.set_stream(f"stream{i}", bytes([i]) * (i * 211))
    vba = root.add_storage("VBA")
    vba.set_stream("dir", b"d" * 5000)
    vba.set_stream("Module1", b"m")
    vba.add_storage("Form").set_stream("\x03VBFrame", b"frame")
    data = write_cfb(root)
    _check_tree(data)
    back = read_cfb(data)
    assert all(back.stream(f"STREAM{i}") == bytes([i]) * (i * 211) for i in range(37))
    assert back.storage("vba").stream("dir") == b"d" * 5000
    assert back.storage("VBA").storage("Form").stream("\x03VBFrame") == b"frame"
    assert sorted(n for n, _ in back.walk()) == sorted(n for n, _ in root.walk())


def test_cfb_large_file_uses_difat():
    root = Storage()
    root.set_stream("big", b"\x5a" * (9 * 1024 * 1024))
    data = write_cfb(root)
    assert struct.unpack_from("<I", data, 72)[0] > 0  # DIFAT sectors
    assert read_cfb(data).stream("big") == b"\x5a" * (9 * 1024 * 1024)


def test_cfb_damage_is_reported():
    with pytest.raises(CfbError):
        read_cfb(b"not a compound file" * 40)
    root = Storage()
    root.set_stream("x", b"y" * 10000)
    data = bytearray(write_cfb(root))
    # Make the big stream's chain loop on itself.
    fat_sector = struct.unpack_from("<I", data, 76)[0]
    start = None
    for i in range(128):
        v = struct.unpack_from("<I", data, 512 + fat_sector * 512 + i * 4)[0]
        if v == 0xFFFFFFFE and i > 2:
            start = i
    struct.pack_into("<I", data, 512 + fat_sector * 512 + start * 4, start)
    with pytest.raises(CfbError):
        read_cfb(bytes(data))


# ---------------------------------------------------------------- projects


def _project():
    code = 'Option Explicit\n\n\' 你好\nSub Hello()\n    Range("A1").Value = "Hi"\nEnd Sub'
    p = new_project(["Sheet1"])
    p = p.with_module(VbaModule("Module1", STANDARD, code, default_attributes("Module1", STANDARD)))
    p = p.with_module(VbaModule("Class1", CLASS, "Public Name As String", default_attributes("Class1", CLASS)))
    return replace(p, code_page=950)


def test_project_round_trip():
    q = read_project(write_project(_project()))
    assert [(m.name, m.kind) for m in q.modules] == [("ThisWorkbook", DOCUMENT), ("Sheet1", DOCUMENT), ("Module1", STANDARD), ("Class1", CLASS)]
    assert q.code_page == 950
    assert "你好" in q.module("Module1").code
    assert q.module("Module1").attributes == ('Attribute VB_Name = "Module1"',)
    assert any("VB_PredeclaredId = True" in a for a in q.module("Sheet1").attributes)
    assert [r.name for r in q.references] == ["VBA", "Excel", "stdole", "Office"]
    assert q.references[0].libid.startswith(r"*\G{000204EF")
    assert decrypt(q.protection[2]) == b"\xff" and not q.locked
    assert q.project_id.startswith("{")
    # Unmodified: written back byte for byte; modified: rebuilt with the references kept.
    assert write_project(q) is q.base
    r = read_project(write_project(q.with_code("Module1", "Sub X()\nEnd Sub")))
    assert r.module("Module1").code == "Sub X()\nEnd Sub"
    assert [a.raw for a in r.references] == [a.raw for a in q.references]
    assert r.project_id == q.project_id


def test_member_attributes_survive_editing():
    text = (
        'Attribute VB_Name = "Module1"\r\nSub Macro1()\r\nAttribute Macro1.VB_Description = "Recorded"\r\n'
        'Attribute Macro1.VB_ProcData.VB_Invoke_Func = "q\\n14"\r\n    Range("A1").Select\r\nEnd Sub\r\n'
    )
    code, attrs, members = split_source(text)
    assert code == 'Sub Macro1()\n    Range("A1").Select\nEnd Sub'
    m = VbaModule("Module1", STANDARD, code, attrs, members)
    m2 = replace(m, code="' new comment\n" + code)
    src = m2.source().split("\n")
    assert src[0] == 'Attribute VB_Name = "Module1"'
    assert src[2] == "Sub Macro1()" and src[3].startswith("Attribute Macro1.VB_Description")
    # A deleted procedure takes its attributes with it.
    assert "Macro1.VB" not in replace(m, code="Sub Other()\nEnd Sub").source()


def test_reads_excel_style_project():
    """A project as Excel saves it: compiled code before each module's source (MODULEOFFSET),
    __SRP_ caches, the PROJECTCOMPATVERSION record, the Mac code page and Mac references."""
    p = _project()
    p = replace(p, code_page=10000, compat_version=0x20000)
    p = p.with_module(replace(p.module("Module1"), code='Sub Hello()\n    MsgBox "Café"\nEnd Sub'))
    root = read_cfb(write_project(p))
    vba = root.storage("VBA")
    pcode = os.urandom(333)
    vba.set_stream("Module1", pcode + vba.stream("Module1"))
    vba.set_stream("__SRP_0", os.urandom(900))
    vba.set_stream("_VBA_PROJECT", os.urandom(3000))
    # Module1's MODULEOFFSET points past the compiled code.
    out, current = b"", None
    for rid, data in _records(decompress(vba.stream("dir"))):
        if rid == 0x0019:
            current = data
        if rid == 0x0031 and current == b"Module1":
            data = struct.pack("<I", len(pcode))
        out += struct.pack("<HIIH", 9, 4, *struct.unpack("<IH", data)) if rid == 0x0009 else _rec(rid, data)
    vba.set_stream("dir", compress(out))
    q = read_project(write_cfb(root))
    assert q.code_page == 10000 and q.compat_version == 0x20000
    assert q.module("Module1").code == 'Sub Hello()\n    MsgBox "Café"\nEnd Sub'
    # Rebuilt after an edit: no compiled code, no caches, Excel recompiles from source.
    r = read_cfb(write_project(q.with_code("Module1", "Sub Y()\nEnd Sub")))
    assert r.storage("VBA").stream("_VBA_PROJECT") == b"\xCC\x61\xFF\xFF\x00\x00\x00"
    assert r.storage("VBA").stream("__SRP_0") is None


def test_locked_project_is_detected():
    p = replace(_project(), project_id="{11111111-2222-3333-4444-555555555555}")
    pid = p.project_id
    p = replace(p, protection=(encrypt(b"\x01\x00\x00\x00", pid), encrypt(b"\x00", pid), encrypt(b"\x00", pid)))
    q = read_project(write_project(p))
    assert q.locked


def test_reconcile_document_modules():
    p = new_project(["Sheet1", "Old"])
    q, wb_name, names = reconcile(p, "ThisWorkbook", ["Sheet1", "", "Sheet1"])
    assert wb_name == "ThisWorkbook" and names == ["Sheet1", "Sheet2", "Sheet3"]
    assert [m.name for m in q.modules] == ["ThisWorkbook", "Sheet1", "Sheet2", "Sheet3"]
    assert reconcile(q, wb_name, names)[0] is q


# ---------------------------------------------------------------- packages


def _book_with_macros():
    s = Session()
    s.new_workbook(2)
    p = new_project([]).with_module(VbaModule("Module1", STANDARD, 'Sub Hello()\n    Range("A1").Value = 1\nEnd Sub', default_attributes("Module1", STANDARD)))
    s.doc().wb.vba = p
    return s


def test_xlsm_package():
    s = _book_with_macros()
    r = s.run("file.saveBytes")
    assert r["name"].endswith(".xlsm") and r["mime"] == "application/vnd.ms-excel.sheet.macroEnabled.12"
    data = base64.b64decode(r["data"])
    z = zipfile.ZipFile(io.BytesIO(data))
    ct = z.read("[Content_Types].xml").decode()
    assert "application/vnd.ms-excel.sheet.macroEnabled.main+xml" in ct and "application/vnd.ms-office.vbaProject" in ct
    assert "relationships/vbaProject" in z.read("xl/_rels/workbook.xml.rels").decode()
    assert 'codeName="ThisWorkbook"' in z.read("xl/workbook.xml").decode()
    assert '<sheetPr codeName="Sheet2">' in z.read("xl/worksheets/sheet2.xml").decode()
    openpyxl = pytest.importorskip("openpyxl")
    book = openpyxl.load_workbook(io.BytesIO(data), keep_vba=True)
    try:
        assert book.sheetnames == ["Sheet1", "Sheet2"] and book.vba_archive is not None
    finally:
        book.vba_archive.close()
    # Read back: the project, the code names and every document module.
    wb, warnings = read_xlsx(data)
    assert not warnings
    assert wb.code_name == "ThisWorkbook" and [x.code_name for x in wb.sheets] == ["Sheet1", "Sheet2"]
    assert [m.name for m in wb.vba.modules] == ["ThisWorkbook", "Module1", "Sheet1", "Sheet2"]
    # Saving as .xlsx leaves the macros out (and says so).
    plain = zipfile.ZipFile(io.BytesIO(write_xlsx(wb, macros=False)))
    assert "xl/vbaProject.bin" not in plain.namelist()
    assert "macroEnabled" not in plain.read("[Content_Types].xml").decode()


def test_signatures_kept_only_while_unmodified():
    wb, _ = read_xlsx(write_xlsx(_book_with_macros().doc().wb))
    sig = ("http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature", "vbaProjectSignature.bin", b"SIGNATURE")
    wb.vba = replace(wb.vba, signatures=(sig,))
    data = write_xlsx(wb)
    z = zipfile.ZipFile(io.BytesIO(data))
    assert z.read("xl/vbaProjectSignature.bin") == b"SIGNATURE"
    assert "vbaProjectSignature" in z.read("xl/_rels/vbaProject.bin.rels").decode()
    back, _ = read_xlsx(data)
    assert back.vba.signatures and back.vba.signatures[0][2] == b"SIGNATURE"
    wb.vba = wb.vba.with_code("Module1", "Sub Changed()\nEnd Sub")
    z2 = zipfile.ZipFile(io.BytesIO(write_xlsx(wb)))
    assert "xl/vbaProjectSignature.bin" not in z2.namelist()


def test_unreadable_project_is_kept():
    wb, _ = read_xlsx(write_xlsx(_book_with_macros().doc().wb))
    junk = b"\xd0\xcf\x11\xe0 definitely not a VBA project"
    buf = io.BytesIO()
    src = zipfile.ZipFile(io.BytesIO(write_xlsx(wb)))
    with zipfile.ZipFile(buf, "w") as z:
        for n in src.namelist():
            z.writestr(n, junk if n == "xl/vbaProject.bin" else src.read(n))
    back, warnings = read_xlsx(buf.getvalue())
    assert back.vba is not None and back.vba.unreadable
    assert any("VBA project can't be read" in w for w in warnings)
    again = zipfile.ZipFile(io.BytesIO(write_xlsx(back)))
    assert again.read("xl/vbaProject.bin") == junk


def test_rebuilt_project_text_has_one_id():
    # A project read from a file, then edited: the PROJECT stream is rebuilt with its ID once.
    from gridcore.xlsx.ovba import read_project, write_project
    from gridcore.model.vba import new_project
    from dataclasses import replace

    p = new_project(["Sheet1"]).with_module(VbaModule("Module1", STANDARD, "Sub A()\nEnd Sub\n", default_attributes("Module1", STANDARD)))
    q = read_project(write_project(p))
    r = replace(q.with_module(replace(q.module("Module1"), code="Sub B()\nEnd Sub\n")), modified=True)
    text = read_cfb(write_project(r)).stream("PROJECT").decode("cp1252")
    assert [ln for ln in text.split("\r\n") if ln.startswith("ID=")] == [text.split("\r\n")[0]]
