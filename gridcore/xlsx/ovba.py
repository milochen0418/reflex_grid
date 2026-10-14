"""VBA project storage ([MS-OVBA]): reading a ``vbaProject.bin`` into a
:class:`~gridcore.model.vba.VbaProject` and writing one back.

A project is a compound file (:mod:`.cfb`) holding the ``PROJECT`` stream (text: modules and
project properties), ``PROJECTwm`` (module names), and the ``VBA`` storage: ``dir`` (compressed
records describing the project, its references and modules), ``_VBA_PROJECT`` and one stream per
module whose compressed source follows a version-specific cache of compiled code.

Writing: an unmodified project is written back exactly as read. A modified (or new) project is
rebuilt from source, keeping everything else the original had (references, designer storages of
UserForms, ``PROJECTlk``, project properties); its ``_VBA_PROJECT`` stream carries version
0xFFFF and no compiled code, which tells Excel to compile the modules from their source when it
opens the file ([MS-OVBA] 2.3.4.1).
"""

from __future__ import annotations

import codecs
import os
import re
import struct
import uuid
from dataclasses import replace
from typing import Dict, List, Optional, Tuple

from ..model.vba import CLASS, DOCUMENT, FORM, STANDARD, VbaModule, VbaProject, VbaReference, split_source
from .cfb import CfbError, Storage, read_cfb, write_cfb

MAX_SOURCE = 16 * 1024 * 1024


class OvbaError(Exception):
    """The VBA project can't be read."""


# ---------------------------------------------------------------- code pages


_CODECS = {
    65001: "utf-8",
    1200: "utf-16-le",
    10000: "mac_roman",
    10001: "shift_jis",
    10002: "big5",
    10003: "euc_kr",
    10008: "gb2312",
    20127: "ascii",
    28591: "latin-1",
    936: "gbk",
}


def codec(code_page: int) -> str:
    name = _CODECS.get(code_page) or f"cp{code_page}"
    try:
        codecs.lookup(name)
        return name
    except LookupError:
        return "cp1252"


def _dec(b: bytes, cp: int) -> str:
    return b.decode(codec(cp), "replace")


def _enc(s: str, cp: int) -> bytes:
    try:
        return s.encode(codec(cp))
    except UnicodeEncodeError as e:
        raise OvbaError(f"“{s[e.start:e.end]}” can't be stored in this VBA project's code page ({code_page_name(cp)})") from None


def code_page_name(cp: int) -> str:
    return {1252: "Western European (Windows-1252)", 950: "Traditional Chinese (Big5)", 936: "Simplified Chinese (GBK)", 932: "Japanese (Shift-JIS)", 949: "Korean", 10000: "Mac Roman", 65001: "UTF-8"}.get(cp, f"code page {cp}")


def pick_code_page(texts: List[str]) -> int:
    """A code page that can hold all of ``texts`` (Windows-1252 when it can, as Excel's default)."""
    for cp in (1252, 950, 936, 932, 949, 1250, 1251, 1253, 1254, 1255, 1256, 1257, 1258, 874):
        try:
            for t in texts:
                t.encode(codec(cp))
            return cp
        except UnicodeEncodeError:
            continue
    return 1252


# ---------------------------------------------------------------- compression ([MS-OVBA] 2.4.1)


def _copy_token_help(diff: int) -> Tuple[int, int, int]:
    """(length mask, bit count, maximum length) for a token ``diff`` bytes into its chunk."""
    bits = max((diff - 1).bit_length(), 4)
    mask = 0xFFFF >> bits
    return mask, bits, mask + 3


def decompress(data: bytes, limit: int = MAX_SOURCE) -> bytes:
    if not data or data[0] != 0x01:
        raise OvbaError("compressed container has no signature")
    out = bytearray()
    pos = 1
    n = len(data)
    while pos + 2 <= n:
        header = data[pos] | (data[pos + 1] << 8)
        size = (header & 0x0FFF) + 3
        compressed = header & 0x8000
        if (header >> 12) & 0x7 != 0b011:
            raise OvbaError("bad compressed chunk signature")
        end = min(n, pos + size)
        pos += 2
        start = len(out)
        if not compressed:
            out += data[pos:pos + 4096]
            pos += 4096
        else:
            while pos < end:
                flags = data[pos]
                pos += 1
                for bit in range(8):
                    if pos >= end:
                        break
                    if not flags & (1 << bit):
                        out.append(data[pos])
                        pos += 1
                    else:
                        if pos + 1 >= end:
                            raise OvbaError("truncated copy token")
                        token = data[pos] | (data[pos + 1] << 8)
                        pos += 2
                        mask, bits, _ = _copy_token_help(len(out) - start)
                        length = (token & mask) + 3
                        offset = (token >> (16 - bits)) + 1
                        src = len(out) - offset
                        if src < start:
                            raise OvbaError("copy token points before its chunk")
                        if offset >= length:
                            out += out[src:src + length]
                        else:
                            for k in range(length):
                                out.append(out[src + k])
            pos = end
        if len(out) > limit:
            raise OvbaError("decompressed data is too large")
    return bytes(out)


def compress(data: bytes) -> bytes:
    """Compresses as [MS-OVBA] 2.4.1.3 does: the longest match, the nearest one when several are
    as long (so the output matches the specification's examples byte for byte)."""
    out = bytearray(b"\x01")
    n = len(data)
    pos = 0
    while pos < n:
        chunk_start = pos
        chunk_end = min(n, pos + 4096)
        body = bytearray()
        # Positions of each 3-byte prefix inside the chunk, nearest last.
        index: Dict[bytes, List[int]] = {}
        p = chunk_start

        def note(i: int) -> None:
            if i + 3 <= chunk_end:
                index.setdefault(data[i:i + 3], []).append(i)

        while p < chunk_end:
            flag_at = len(body)
            body.append(0)
            flags = 0
            for bit in range(8):
                if p >= chunk_end:
                    break
                best_len = 0
                best = -1
                cands = index.get(data[p:p + 3]) if p + 3 <= chunk_end else None
                if cands:
                    _, _, max_len = _copy_token_help(p - chunk_start)
                    for c in reversed(cands):
                        ln = 3
                        lim = chunk_end - p
                        while ln < lim and data[c + ln] == data[p + ln]:
                            ln += 1
                        if ln > best_len:
                            best_len, best = ln, c
                            if ln >= lim:
                                break
                if best_len >= 3:
                    mask, bits, max_len = _copy_token_help(p - chunk_start)
                    length = min(best_len, max_len)
                    offset = p - best
                    token = ((offset - 1) << (16 - bits)) | (length - 3)
                    body += struct.pack("<H", token)
                    flags |= 1 << bit
                    for i in range(p, p + length):
                        note(i)
                    p += length
                else:
                    body.append(data[p])
                    note(p)
                    p += 1
            body[flag_at] = flags
        if len(body) > 4096:
            # Incompressible: a raw chunk (the specification pads a short last chunk with zeros).
            raw = data[chunk_start:chunk_start + 4096]
            raw = raw + b"\0" * (4096 - len(raw))
            out += struct.pack("<H", 0x3000 | 0x0FFF)
            out += raw
        else:
            out += struct.pack("<H", 0xB000 | (len(body) + 2 - 3))
            out += body
        pos = chunk_end
    return bytes(out)


# ---------------------------------------------------------------- data encryption ([MS-OVBA] 2.4.3)


def project_key(project_id: str) -> int:
    return sum(project_id.encode("latin-1", "replace")) & 0xFF


def encrypt(data: bytes, project_id: str, seed: Optional[int] = None, ignored: Optional[bytes] = None) -> str:
    """Obfuscates ``data`` as the CMG, DPB and GC properties are; returns the hex text."""
    if seed is None:
        seed = os.urandom(1)[0]
    version = 2
    key = project_key(project_id)
    v_enc, k_enc = seed ^ version, seed ^ key
    out = [seed, v_enc, k_enc]
    plain1, enc1, enc2 = key, k_enc, v_enc
    count = (seed & 6) // 2
    junk = ignored if ignored is not None else os.urandom(count)

    def put(b: int) -> None:
        nonlocal plain1, enc1, enc2
        e = (b ^ ((enc2 + plain1) & 0xFF)) & 0xFF
        out.append(e)
        enc2, enc1, plain1 = enc1, e, b

    for k in range(count):
        put(junk[k])
    for b in struct.pack("<I", len(data)) + data:
        put(b)
    return bytes(out).hex().upper()


def decrypt(text: str) -> Optional[bytes]:
    """The data behind an obfuscated CMG/DPB/GC value, or ``None`` when it doesn't decode."""
    try:
        b = bytes.fromhex(text.strip().strip('"'))
    except ValueError:
        return None
    if len(b) < 7:
        return None
    seed, v_enc, k_enc = b[0], b[1], b[2]
    plain1, enc1, enc2 = seed ^ k_enc, k_enc, v_enc
    pos = 3

    def get() -> int:
        nonlocal plain1, enc1, enc2, pos
        e = b[pos]
        pos += 1
        p = (e ^ ((enc2 + plain1) & 0xFF)) & 0xFF
        enc2, enc1, plain1 = enc1, e, p
        return p

    try:
        for _ in range((seed & 6) // 2):
            get()
        length = struct.unpack("<I", bytes(get() for _ in range(4)))[0]
        if length > len(b):
            return None
        return bytes(get() for _ in range(length))
    except IndexError:
        return None


# ---------------------------------------------------------------- dir stream records


def _records(data: bytes) -> List[Tuple[int, bytes]]:
    """The ``dir`` stream as (id, data) records. Every record is id, size, data, except
    PROJECTVERSION whose 4-byte "size" is followed by 6 bytes."""
    out: List[Tuple[int, bytes]] = []
    i = 0
    n = len(data)
    while i + 6 <= n:
        rid, size = struct.unpack_from("<HI", data, i)
        i += 6
        if rid == 0x0009:
            size = 6
        if size > n - i:
            raise OvbaError(f"dir record 0x{rid:04X} is truncated")
        out.append((rid, data[i:i + size]))
        i += size
        if rid == 0x0010:
            break
    return out


def _rec(rid: int, data: bytes) -> bytes:
    return struct.pack("<HI", rid, len(data)) + data


def _libid(rid: int, data: bytes) -> str:
    try:
        if rid in (0x000D, 0x002F, 0x0030, 0x000E):
            n = struct.unpack_from("<I", data, 0)[0]
            return data[4:4 + n].decode("latin-1")
        if rid == 0x0033:
            return data.decode("latin-1")
    except struct.error:
        pass
    return ""


def _parse_references(recs: List[Tuple[int, bytes]], cp: int) -> List[VbaReference]:
    """Groups the reference records: an optional REFERENCENAME, then a registered, project or
    control reference (a control may follow a REFERENCEORIGINAL and carry a name of its own)."""
    refs: List[Dict[str, object]] = []
    cur: Optional[Dict[str, object]] = None
    in_control = False

    def start() -> Dict[str, object]:
        r: Dict[str, object] = {"name": "", "uname": "", "kind": None, "libid": "", "raw": b"", "done": False}
        refs.append(r)
        return r

    for rid, data in recs:
        if in_control and cur is not None and rid in (0x0016, 0x003E):
            cur["raw"] += _rec(rid, data)  # type: ignore[operator]
            continue
        if rid == 0x003E:
            if cur is not None:
                cur["raw"] += _rec(rid, data)  # type: ignore[operator]
                cur["uname"] = data.decode("utf-16-le", "replace")
            continue
        continues = cur is not None and not cur["done"] and (rid != 0x0016 or not cur["name"]) and not (rid == 0x0016 and cur["kind"] is not None)
        if rid == 0x002F and cur is not None and cur["kind"] == "original" and not cur["done"]:
            continues = True
        if not continues:
            cur = start()
        assert cur is not None
        cur["raw"] += _rec(rid, data)  # type: ignore[operator]
        if rid == 0x0016:
            cur["name"] = _dec(data, cp)
        elif rid in (0x000D, 0x000E):
            cur["kind"] = "registered" if rid == 0x000D else "project"
            cur["libid"] = _libid(rid, data)
            cur["done"] = True
        elif rid == 0x0033:
            cur["kind"], cur["libid"] = "original", _libid(rid, data)
        elif rid == 0x002F:
            if cur["kind"] != "original":
                cur["libid"] = _libid(rid, data)
            cur["kind"] = "control"
            in_control = True
        elif rid == 0x0030:
            in_control = False
            cur["done"] = True
        else:
            cur["done"] = True
    out = []
    for r in refs:
        name = str(r["uname"] or r["name"])
        out.append(VbaReference(name, str(r["kind"] or "registered"), str(r["libid"]), bytes(r["raw"])))  # type: ignore[arg-type]
    return out


def registered_reference(name: str, libid: str, cp: int = 1252) -> VbaReference:
    lib = libid.encode("latin-1")
    raw = _rec(0x0016, _enc(name, cp)) + _rec(0x003E, name.encode("utf-16-le"))
    raw += _rec(0x000D, struct.pack("<I", len(lib)) + lib + struct.pack("<IH", 0, 0))
    return VbaReference(name, "registered", libid, raw)


# The references Excel gives a new project.
DEFAULT_REFERENCES = (
    ("VBA", r"*\G{000204EF-0000-0000-C000-000000000046}#4.2#9#C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL#Visual Basic For Applications"),
    ("Excel", r"*\G{00020813-0000-0000-C000-000000000046}#1.9#0#C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE#Microsoft Excel 16.0 Object Library"),
    ("stdole", r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation"),
    ("Office", r"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library"),
)


def default_references(cp: int = 1252) -> Tuple[VbaReference, ...]:
    return tuple(registered_reference(n, lib, cp) for n, lib in DEFAULT_REFERENCES)


# ---------------------------------------------------------------- PROJECT stream


_MODULE_KEYS = ("document", "module", "class", "baseclass", "package")


def _project_props(text: str) -> Tuple[List[Tuple[str, str]], Dict[str, str], List[str], List[str]]:
    """(property lines before the first section as (key, value), module kinds by name,
    host extender lines, workspace lines)."""
    props: List[Tuple[str, str]] = []
    kinds: Dict[str, str] = {}
    host: List[str] = []
    workspace: List[str] = []
    section = ""
    for line in text.replace("\r\n", "\n").split("\n"):
        s = line.strip()
        if not s:
            continue
        if s.startswith("[") and s.endswith("]"):
            section = s[1:-1].lower()
            continue
        if section == "host extender info":
            host.append(s)
            continue
        if section == "workspace":
            workspace.append(s)
            continue
        if section:
            continue
        if "=" not in s:
            continue
        k, v = s.split("=", 1)
        kl = k.strip().lower()
        if kl == "document":
            kinds[v.split("/", 1)[0].strip().lower()] = DOCUMENT
        elif kl == "module":
            kinds[v.strip().lower()] = STANDARD
        elif kl == "class":
            kinds[v.strip().lower()] = CLASS
        elif kl == "baseclass":
            kinds[v.strip().lower()] = FORM
        props.append((k.strip(), v))
    return props, kinds, host, workspace


def build_project_text(p: VbaProject) -> str:
    props: Dict[str, str] = {}
    packages: List[str] = []
    host: List[str] = []
    workspace: Dict[str, str] = {}
    if p.project_lines:
        pl, _, host, ws = _project_props("\n".join(p.project_lines))
        for k, v in pl:
            kl = k.lower()
            if kl == "package":
                packages.append(v)
            elif kl not in _MODULE_KEYS:
                props[kl] = f"{k}={v}"
        for line in ws:
            if "=" in line:
                k, v = line.split("=", 1)
                workspace[k.strip().lower()] = v
    pid = p.project_id or "{00000000-0000-0000-0000-000000000000}"
    lines = [f'ID="{pid}"']
    forms = [m for m in p.modules if m.kind == FORM]
    for m in p.modules:
        if m.kind == DOCUMENT:
            lines.append(f"Document={m.name}/&H00000000")
        elif m.kind == STANDARD:
            lines.append(f"Module={m.name}")
        elif m.kind == CLASS:
            lines.append(f"Class={m.name}")
    if forms:
        for pk in packages or ["{AC9F2F90-E877-11CE-9F68-00AA00574A4F}"]:
            lines.append(f"Package={pk}")
        for m in forms:
            lines.append(f"BaseClass={m.name}")
    for key in ("helpfile", "exename32"):
        if key in props:
            lines.append(props.pop(key))
    lines.append(props.pop("name", f'Name="{p.name}"'))
    lines.append(props.pop("helpcontextid", f'HelpContextID="{p.help_context}"'))
    if "description" in props:
        lines.append(props.pop("description"))
    lines.append(props.pop("versioncompatible32", 'VersionCompatible32="393222000"'))
    cmg, dpb, gc = p.protection
    cmg = cmg or encrypt(b"\x00\x00\x00\x00", pid)
    dpb = dpb or encrypt(b"\x00", pid)
    gc = gc or encrypt(b"\xff", pid)
    lines.append(props.pop("cmg", f'CMG="{cmg}"'))
    lines.append(props.pop("dpb", f'DPB="{dpb}"'))
    lines.append(props.pop("gc", f'GC="{gc}"'))
    for k in ("cmg", "dpb", "gc"):
        props.pop(k, None)
    lines.extend(props.values())
    lines.append("")
    lines.append("[Host Extender Info]")
    lines.extend(host or ["&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000"])
    lines.append("")
    lines.append("[Workspace]")
    for m in p.modules:
        lines.append(f"{m.name}={workspace.get(m.name.lower(), '0, 0, 0, 0, C')}")
    return "\r\n".join(lines) + "\r\n"


# ---------------------------------------------------------------- reading


def read_project(data: bytes) -> VbaProject:
    """Reads a ``vbaProject.bin``. Raises :class:`OvbaError` when it isn't a VBA project."""
    try:
        root = read_cfb(data)
    except CfbError as e:
        raise OvbaError(f"the VBA project container is damaged: {e}") from None
    vba = root.storage("VBA")
    if vba is None:
        raise OvbaError("the file has no VBA storage")
    dir_raw = vba.stream("dir")
    if dir_raw is None:
        raise OvbaError("the VBA project has no dir stream")
    recs = _records(decompress(dir_raw))
    warnings: List[str] = []
    info: Dict[int, bytes] = {}
    i = 0
    while i < len(recs) and recs[i][0] not in (0x0016, 0x0033, 0x002F, 0x000D, 0x000E, 0x000F):
        info.setdefault(recs[i][0], recs[i][1])
        i += 1
    cp = struct.unpack("<H", info[0x0003][:2])[0] if len(info.get(0x0003, b"")) >= 2 else 1252
    j = i
    while j < len(recs) and recs[j][0] != 0x000F:
        j += 1
    references = _parse_references(recs[i:j], cp)

    def u32(rid: int, default: int = 0) -> int:
        b = info.get(rid, b"")
        return struct.unpack("<I", b[:4])[0] if len(b) >= 4 else default

    ver = info.get(0x0009, b"")
    major, minor = struct.unpack("<IH", ver[:6]) if len(ver) >= 6 else (1, 0)
    name = _dec(info.get(0x0004, b"VBAProject"), cp)
    doc = info.get(0x0040)
    project_text = _dec(root.stream("PROJECT") or b"", cp)
    props, kinds, _, _ = _project_props(project_text)
    pmap = {k.lower(): v for k, v in props}

    def unq(v: str) -> str:
        v = v.strip()
        return v[1:-1] if len(v) >= 2 and v[0] == '"' and v[-1] == '"' else v

    project_id = unq(pmap.get("id", ""))
    protection = (unq(pmap.get("cmg", "")), unq(pmap.get("dpb", "")), unq(pmap.get("gc", "")))
    visibility = decrypt(protection[2]) if protection[2] else None
    locked = visibility is not None and len(visibility) >= 1 and visibility[0] == 0x00

    modules: List[VbaModule] = []
    k = j + 1
    while k < len(recs) and recs[k][0] in (0x0013,):
        k += 1
    cur: Optional[Dict[str, object]] = None
    for rid, rdata in recs[k:]:
        if rid == 0x0019:
            cur = {"name": _dec(rdata, cp), "stream": "", "offset": 0, "type": 0x0021, "doc": "", "help": 0, "ro": False, "priv": False}
        elif cur is None:
            continue
        elif rid == 0x0047:
            cur["name"] = rdata.decode("utf-16-le", "replace") or cur["name"]
        elif rid == 0x001A:
            cur["stream"] = _dec(rdata, cp)
        elif rid == 0x0032:
            cur["stream"] = rdata.decode("utf-16-le", "replace") or cur["stream"]
        elif rid == 0x001C:
            cur["doc"] = _dec(rdata, cp)
        elif rid == 0x0048:
            cur["doc"] = rdata.decode("utf-16-le", "replace") or cur["doc"]
        elif rid == 0x0031 and len(rdata) >= 4:
            cur["offset"] = struct.unpack("<I", rdata[:4])[0]
        elif rid == 0x001E and len(rdata) >= 4:
            cur["help"] = struct.unpack("<I", rdata[:4])[0]
        elif rid in (0x0021, 0x0022):
            cur["type"] = rid
        elif rid == 0x0025:
            cur["ro"] = True
        elif rid == 0x0028:
            cur["priv"] = True
        elif rid == 0x002B:
            mname = str(cur["name"])
            stream_name = str(cur["stream"] or mname)
            raw = vba.stream(stream_name)
            text = ""
            if raw is None:
                warnings.append(f"module {mname} has no code stream")
            else:
                off = int(cur["offset"])  # type: ignore[arg-type]
                try:
                    text = _dec(decompress(raw[off:]), cp) if off < len(raw) else ""
                except OvbaError as e:
                    warnings.append(f"module {mname}: {e}")
            code, attrs, members = split_source(text)
            kind = kinds.get(mname.lower())
            if kind is None:
                if cur["type"] == 0x0021:
                    kind = STANDARD
                elif any("VB_Base" in a for a in attrs):
                    kind = DOCUMENT if any(("00020819" in a or "00020820" in a) for a in attrs) else FORM
                else:
                    kind = CLASS
            if cur["type"] == 0x0021:
                kind = STANDARD
            modules.append(
                VbaModule(
                    mname,
                    kind,
                    code,
                    attrs,
                    members,
                    stream_name if stream_name != mname else "",
                    str(cur["doc"]),
                    int(cur["help"]),  # type: ignore[arg-type]
                    bool(cur["ro"]),
                    bool(cur["priv"]),
                )
            )
            cur = None
    keep = tuple(line for line in project_text.replace("\r\n", "\n").split("\n"))
    return VbaProject(
        modules=tuple(modules),
        name=name,
        code_page=cp,
        references=tuple(references),
        sys_kind=u32(0x0001, 1),
        lcid=u32(0x0002, 0x409),
        lcid_invoke=u32(0x0014, 0x409),
        version_major=major,
        version_minor=minor,
        compat_version=u32(0x004A) if 0x004A in info else None,
        doc_string=doc.decode("utf-16-le", "replace") if doc is not None else _dec(info.get(0x0005, b""), cp),
        help_file=_dec(info.get(0x0006, b""), cp),
        help_context=u32(0x0007),
        lib_flags=u32(0x0008),
        constants=(info.get(0x003C) or b"").decode("utf-16-le", "replace") or _dec(info.get(0x000C, b""), cp),
        project_id=project_id,
        protection=protection,
        locked=locked,
        project_lines=keep,
        base=bytes(data),
        modified=False,
        warnings=tuple(warnings),
    )


# ---------------------------------------------------------------- writing


def _dir_stream(p: VbaProject) -> bytes:
    cp = p.code_page
    out = bytearray()
    out += _rec(0x0001, struct.pack("<I", p.sys_kind))
    if p.compat_version is not None:
        out += _rec(0x004A, struct.pack("<I", p.compat_version))
    out += _rec(0x0002, struct.pack("<I", p.lcid))
    out += _rec(0x0014, struct.pack("<I", p.lcid_invoke))
    out += _rec(0x0003, struct.pack("<H", cp))
    out += _rec(0x0004, _enc(p.name, cp))
    out += _rec(0x0005, _enc(p.doc_string, cp)) + _rec(0x0040, p.doc_string.encode("utf-16-le"))
    hf = _enc(p.help_file, cp)
    out += _rec(0x0006, hf) + _rec(0x003D, hf)
    out += _rec(0x0007, struct.pack("<I", p.help_context))
    out += _rec(0x0008, struct.pack("<I", p.lib_flags))
    out += struct.pack("<HIIH", 0x0009, 4, p.version_major & 0xFFFFFFFF, p.version_minor & 0xFFFF)
    out += _rec(0x000C, _enc(p.constants, cp)) + _rec(0x003C, p.constants.encode("utf-16-le"))
    for r in p.references:
        out += r.raw
    out += _rec(0x000F, struct.pack("<H", len(p.modules)))
    out += _rec(0x0013, struct.pack("<H", 0xFFFF))
    for m in p.modules:
        out += _rec(0x0019, _enc(m.name, cp)) + _rec(0x0047, m.name.encode("utf-16-le"))
        out += _rec(0x001A, _enc(m.stream_name, cp)) + _rec(0x0032, m.stream_name.encode("utf-16-le"))
        out += _rec(0x001C, _enc(m.doc_string, cp)) + _rec(0x0048, m.doc_string.encode("utf-16-le"))
        out += _rec(0x0031, struct.pack("<I", 0))
        out += _rec(0x001E, struct.pack("<I", m.help_context))
        out += _rec(0x002C, struct.pack("<H", 0xFFFF))
        out += struct.pack("<HI", 0x0021 if m.kind == STANDARD else 0x0022, 0)
        if m.read_only:
            out += struct.pack("<HI", 0x0025, 0)
        if m.private:
            out += struct.pack("<HI", 0x0028, 0)
        out += struct.pack("<HI", 0x002B, 0)
    out += struct.pack("<HI", 0x0010, 0)
    return bytes(out)


def _projectwm(p: VbaProject) -> bytes:
    out = bytearray()
    for m in p.modules:
        out += _enc(m.name, p.code_page) + b"\x00" + m.name.encode("utf-16-le") + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)


def prepare(p: VbaProject) -> VbaProject:
    """Fills in what a new project needs before it is written (an ID, references)."""
    if not p.project_id:
        p = replace(p, project_id="{" + str(uuid.uuid4()).upper() + "}")
    if not p.references:
        p = replace(p, references=default_references(p.code_page))
    return p


def write_project(p: VbaProject) -> bytes:
    """The project as a ``vbaProject.bin``: the original bytes while unmodified, else rebuilt."""
    if p.base is not None and not p.modified:
        return p.base
    if p.unreadable:
        if p.base is None:
            raise OvbaError("an unreadable VBA project can't be written")
        return p.base
    p = prepare(p)
    root = Storage()
    if p.base is not None:
        try:
            root = read_cfb(p.base)
        except CfbError:
            root = Storage()
    vba = root.storage("VBA")
    if vba is not None:
        root.remove("VBA")
    vba = root.add_storage("VBA")
    # No compiled code: version 0xFFFF makes Excel compile the modules from source.
    vba.set_stream("_VBA_PROJECT", b"\xCC\x61\xFF\xFF\x00\x00\x00")
    vba.set_stream("dir", compress(_dir_stream(p)))
    seen = set()
    for m in p.modules:
        key = m.stream_name.upper()
        if key in seen:
            raise OvbaError(f"two modules share the stream name {m.stream_name}")
        seen.add(key)
        text = m.source().replace("\n", "\r\n") + "\r\n"
        vba.set_stream(m.stream_name, compress(_enc(text, p.code_page)))
    root.set_stream("PROJECT", _enc(build_project_text(p), p.code_page))
    root.set_stream("PROJECTwm", _projectwm(p))
    # Designer storages (UserForm layouts) stay for the forms that still exist.
    forms = {m.name.upper() for m in p.modules if m.kind == FORM}
    for name in list(root.storages):
        if name.upper() != "VBA" and name.upper() not in forms:
            root.remove(name)
    return write_cfb(root)


def describe(p: VbaProject) -> List[str]:
    """One line per module (for logs and the CLI)."""
    return [f"{m.kind:9} {m.name} ({len(m.code.splitlines())} lines)" for m in p.modules]

