"""Compound File Binary ([MS-CFB]): the OLE container a VBA project is stored in
(``xl/vbaProject.bin``).

:func:`read_cfb` turns a file into a tree of :class:`Storage` objects holding stream bytes;
:func:`write_cfb` writes such a tree as a version 3 file (512-byte sectors). Reading is bounded:
sector chains are checked for loops and out-of-range sectors, so a damaged file raises
:class:`CfbError` instead of hanging or exhausting memory. Writing produces a directory whose
red-black trees are balanced and correctly coloured.
"""

from __future__ import annotations

import struct
from typing import Dict, Iterator, List, Optional, Tuple

SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
MAXREGSECT = 0xFFFFFFFA
NOSTREAM = 0xFFFFFFFF

MINI_CUTOFF = 4096
MAX_FILE = 256 * 1024 * 1024
MAX_ENTRIES = 100_000

_STORAGE, _STREAM, _ROOT = 1, 2, 5


class CfbError(Exception):
    """The data is not a readable compound file."""


class Storage:
    """A storage: named streams (bytes) and named sub-storages. Names keep their case; lookups
    ignore it, as in the file format."""

    def __init__(self, name: str = "Root Entry", clsid: bytes = b"\0" * 16):
        self.name = name
        self.clsid = clsid
        self.state = 0
        self.streams: Dict[str, bytes] = {}
        self.storages: Dict[str, "Storage"] = {}

    def _key(self, d: Dict[str, object], name: str) -> Optional[str]:
        if name in d:
            return name
        up = name.upper()
        for k in d:
            if k.upper() == up:
                return k
        return None

    def stream(self, name: str) -> Optional[bytes]:
        k = self._key(self.streams, name)  # type: ignore[arg-type]
        return self.streams[k] if k is not None else None

    def storage(self, name: str) -> Optional["Storage"]:
        k = self._key(self.storages, name)  # type: ignore[arg-type]
        return self.storages[k] if k is not None else None

    def set_stream(self, name: str, data: bytes) -> None:
        k = self._key(self.streams, name)  # type: ignore[arg-type]
        self.streams[k if k is not None else name] = bytes(data)

    def add_storage(self, name: str) -> "Storage":
        s = self.storage(name)
        if s is None:
            s = self.storages[name] = Storage(name)
        return s

    def remove(self, name: str) -> None:
        k = self._key(self.streams, name)  # type: ignore[arg-type]
        if k is not None:
            del self.streams[k]
        k = self._key(self.storages, name)  # type: ignore[arg-type]
        if k is not None:
            del self.storages[k]

    def walk(self, prefix: str = "") -> Iterator[Tuple[str, bytes]]:
        """Every stream below this storage as ``("VBA/dir", data)`` pairs."""
        for n, d in self.streams.items():
            yield prefix + n, d
        for n, s in self.storages.items():
            yield from s.walk(prefix + n + "/")

    def __repr__(self) -> str:
        return f"Storage({self.name!r}, streams={list(self.streams)}, storages={list(self.storages)})"


# ---------------------------------------------------------------- reading


def read_cfb(data: bytes) -> Storage:
    if len(data) > MAX_FILE:
        raise CfbError("the compound file is too large")
    if len(data) < 512 or data[:8] != SIGNATURE:
        raise CfbError("not a compound file")
    (major, byte_order, sector_shift, mini_shift) = struct.unpack_from("<HHHH", data, 26)
    if byte_order != 0xFFFE:
        raise CfbError("bad byte order mark")
    if sector_shift not in (9, 12) or mini_shift != 6:
        raise CfbError("unsupported sector size")
    sector_size = 1 << sector_shift
    n_fat, first_dir = struct.unpack_from("<II", data, 44)
    cutoff, first_minifat, n_minifat, first_difat, n_difat = struct.unpack_from("<IIIII", data, 56)
    if cutoff != MINI_CUTOFF:
        cutoff = MINI_CUTOFF
    n_sectors = max((len(data) - sector_size) // sector_size + (1 if (len(data) - sector_size) % sector_size else 0), 0)

    def sector(i: int) -> bytes:
        if i >= n_sectors or i > MAXREGSECT:
            raise CfbError(f"sector {i} is outside the file")
        off = (i + 1) * sector_size
        chunk = data[off:off + sector_size]
        if len(chunk) < sector_size:
            chunk = chunk + b"\0" * (sector_size - len(chunk))
        return chunk

    # The FAT sectors: 109 in the header, the rest in a chain of DIFAT sectors.
    fat_sectors: List[int] = [s for s in struct.unpack_from("<109I", data, 76) if s <= MAXREGSECT]
    per = sector_size // 4
    d, seen = first_difat, set()
    for _ in range(n_difat):
        if d > MAXREGSECT:
            break
        if d in seen:
            raise CfbError("the DIFAT chain loops")
        seen.add(d)
        vals = struct.unpack("<%dI" % per, sector(d))
        fat_sectors.extend(s for s in vals[:-1] if s <= MAXREGSECT)
        d = vals[-1]
    if len(fat_sectors) > n_sectors:
        raise CfbError("the FAT is larger than the file")
    fat: List[int] = []
    for s in fat_sectors[: max(n_fat, 0) or len(fat_sectors)]:
        fat.extend(struct.unpack("<%dI" % per, sector(s)))

    def chain(start: int, table: List[int], limit: int) -> List[int]:
        out: List[int] = []
        seen_: set = set()
        i = start
        while i != ENDOFCHAIN:
            if i > MAXREGSECT or i >= len(table) or i in seen_:
                if i == FREESECT and out:
                    break
                raise CfbError("a sector chain is broken or loops")
            seen_.add(i)
            out.append(i)
            if len(out) > limit:
                raise CfbError("a sector chain is too long")
            i = table[i]
        return out

    dir_bytes = b"".join(sector(i) for i in chain(first_dir, fat, n_sectors))
    entries: List[Tuple[str, int, int, int, int, bytes, int, int, int]] = []
    for off in range(0, len(dir_bytes) - 127, 128):
        if len(entries) >= MAX_ENTRIES:
            raise CfbError("too many directory entries")
        raw = dir_bytes[off:off + 128]
        name_len, kind = struct.unpack_from("<HB", raw, 64)
        left, right, child = struct.unpack_from("<III", raw, 68)
        clsid = raw[80:96]
        state = struct.unpack_from("<I", raw, 96)[0]
        start, size_lo, size_hi = struct.unpack_from("<III", raw, 116)
        size = size_lo if major == 3 else size_lo | (size_hi << 32)
        nl = min(max(name_len - 2, 0), 62) & ~1
        name = raw[:nl].decode("utf-16-le", "replace")
        entries.append((name, kind, left, right, child, clsid, state, start, size))
    if not entries or entries[0][1] != _ROOT:
        raise CfbError("no root entry")

    root_start, root_size = entries[0][7], entries[0][8]
    mini_stream = b""
    if root_size and root_start != ENDOFCHAIN:
        mini_stream = b"".join(sector(i) for i in chain(root_start, fat, n_sectors))[:root_size]
    minifat: List[int] = []
    if n_minifat and first_minifat != ENDOFCHAIN:
        for s in chain(first_minifat, fat, n_sectors):
            minifat.extend(struct.unpack("<%dI" % per, sector(s)))

    def stream_data(start: int, size: int) -> bytes:
        if size == 0:
            return b""
        if size > len(data) * 2 + 4096:
            raise CfbError("a stream is larger than the file")
        if size < cutoff:
            parts = []
            for i in chain(start, minifat, len(mini_stream) // 64 + 1):
                parts.append(mini_stream[i * 64:(i + 1) * 64])
            out = b"".join(parts)
        else:
            out = b"".join(sector(i) for i in chain(start, fat, n_sectors))
        if len(out) < size:
            raise CfbError("a stream is truncated")
        return out[:size]

    root = Storage(entries[0][0] or "Root Entry", entries[0][5])
    root.state = entries[0][6]
    visited = {0}

    def fill(storage: Storage, child: int, depth: int) -> None:
        if depth > 64:
            raise CfbError("storages are nested too deeply")
        stack = [child]
        while stack:
            i = stack.pop()
            if i == NOSTREAM:
                continue
            if i >= len(entries) or i in visited:
                raise CfbError("the directory tree is broken")
            visited.add(i)
            name, kind, left, right, ch, clsid, state, start, size = entries[i]
            stack.append(left)
            stack.append(right)
            if kind == _STREAM:
                storage.streams[name] = stream_data(start, size)
            elif kind == _STORAGE:
                sub = Storage(name, clsid)
                sub.state = state
                storage.storages[name] = sub
                fill(sub, ch, depth + 1)

    fill(root, entries[0][4], 0)
    return root


# ---------------------------------------------------------------- writing


def _sort_key(name: str) -> Tuple[int, str]:
    # Shorter names first, then by upper-cased UTF-16 code units ([MS-CFB] 2.6.4).
    return (len(name.encode("utf-16-le")), name.upper())


class _Node:
    __slots__ = ("name", "kind", "clsid", "state", "data", "children", "left", "right", "child", "color", "start", "size", "index")

    def __init__(self, name: str, kind: int, clsid: bytes = b"\0" * 16, state: int = 0, data: bytes = b""):
        self.name = name
        self.kind = kind
        self.clsid = clsid
        self.state = state
        self.data = data
        self.children: List["_Node"] = []
        self.left = self.right = self.child = NOSTREAM
        self.color = 1
        self.start = ENDOFCHAIN
        self.size = 0
        self.index = 0


def _tree(nodes: List["_Node"]) -> int:
    """Links ``nodes`` (siblings) as a balanced binary search tree; returns the root's index.
    Nodes on the deepest level are red when that level isn't full, which keeps every path's
    black height equal."""
    nodes = sorted(nodes, key=lambda n: _sort_key(n.name))
    if not nodes:
        return NOSTREAM
    depth: Dict[int, int] = {}

    def build(lo: int, hi: int, d: int) -> int:
        if lo > hi:
            return NOSTREAM
        mid = (lo + hi) // 2
        n = nodes[mid]
        depth[mid] = d
        n.left = build(lo, mid - 1, d + 1)
        n.right = build(mid + 1, hi, d + 1)
        return n.index

    root = build(0, len(nodes) - 1, 0)
    deepest = max(depth.values())
    perfect = len(nodes) == (1 << (deepest + 1)) - 1
    for i, n in enumerate(nodes):
        n.color = 0 if (depth[i] == deepest and not perfect) else 1
    return root


def write_cfb(root: Storage) -> bytes:
    SS = 512
    nodes: List[_Node] = []

    def add(n: _Node) -> _Node:
        n.index = len(nodes)
        nodes.append(n)
        return n

    def collect(st: Storage, node: _Node) -> None:
        for name, data in st.streams.items():
            node.children.append(add(_Node(_check_name(name), _STREAM, data=data)))
        for name, sub in st.storages.items():
            child = add(_Node(_check_name(name), _STORAGE, sub.clsid, sub.state))
            node.children.append(child)
            collect(sub, child)

    rnode = add(_Node("Root Entry", _ROOT, root.clsid, root.state))
    collect(root, rnode)
    for n in nodes:
        if n.kind in (_ROOT, _STORAGE):
            n.child = _tree(n.children)

    # Small streams live in the mini stream (64-byte mini sectors chained by the mini FAT).
    mini = bytearray()
    minifat: List[int] = []
    big: List[_Node] = []
    for n in nodes:
        if n.kind != _STREAM:
            continue
        n.size = len(n.data)
        if n.size == 0:
            n.start = ENDOFCHAIN
        elif n.size < MINI_CUTOFF:
            first = len(mini) // 64
            count = (n.size + 63) // 64
            n.start = first
            minifat.extend(first + k + 1 for k in range(count - 1))
            minifat.append(ENDOFCHAIN)
            mini += n.data + b"\0" * (count * 64 - n.size)
        else:
            big.append(n)

    def sectors_for(nbytes: int) -> int:
        return (nbytes + SS - 1) // SS

    n_dir = sectors_for(len(nodes) * 128)
    n_minifat = sectors_for(len(minifat) * 4)
    n_mini = sectors_for(len(mini))
    n_big = sum(sectors_for(len(n.data)) for n in big)
    content = n_dir + n_minifat + n_mini + n_big
    n_fat = n_difat = 0
    while True:
        need_fat = sectors_for((content + n_fat + n_difat) * 4)
        need_difat = 0 if need_fat <= 109 else (need_fat - 109 + 126) // 127
        if need_fat == n_fat and need_difat == n_difat:
            break
        n_fat, n_difat = need_fat, need_difat
    total = n_fat + n_difat + content
    fat = [FREESECT] * (n_fat * (SS // 4))

    pos = 0
    fat_secs = list(range(pos, pos + n_fat))
    for s in fat_secs:
        fat[s] = FATSECT
    pos += n_fat
    difat_secs = list(range(pos, pos + n_difat))
    for s in difat_secs:
        fat[s] = DIFSECT
    pos += n_difat

    def run(count: int) -> int:
        nonlocal pos
        if count == 0:
            return ENDOFCHAIN
        first = pos
        for k in range(count - 1):
            fat[pos + k] = pos + k + 1
        fat[pos + count - 1] = ENDOFCHAIN
        pos += count
        return first

    dir_start = run(n_dir)
    minifat_start = run(n_minifat)
    mini_start = run(n_mini)
    for n in big:
        n.start = run(sectors_for(len(n.data)))
    assert pos == total
    rnode.start = mini_start if mini else ENDOFCHAIN
    rnode.size = len(mini)

    out = bytearray()
    header = bytearray(512)
    header[0:8] = SIGNATURE
    struct.pack_into("<HHHHH", header, 24, 0x003E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIIIII", header, 40, 0, n_fat, dir_start, 0, MINI_CUTOFF, minifat_start if n_minifat else ENDOFCHAIN, n_minifat, difat_secs[0] if difat_secs else ENDOFCHAIN, n_difat, 0)
    difat_head = fat_secs[:109] + [FREESECT] * (109 - min(len(fat_secs), 109))
    struct.pack_into("<109I", header, 76, *difat_head)
    out += header

    body = bytearray(total * SS)

    def put(sec: int, payload: bytes) -> None:
        body[sec * SS:sec * SS + len(payload)] = payload

    for i, s in enumerate(fat_secs):
        put(s, struct.pack("<%dI" % (SS // 4), *fat[i * (SS // 4):(i + 1) * (SS // 4)]))
    rest = fat_secs[109:]
    for i, s in enumerate(difat_secs):
        vals = rest[i * 127:(i + 1) * 127]
        vals = vals + [FREESECT] * (127 - len(vals))
        nxt = difat_secs[i + 1] if i + 1 < len(difat_secs) else ENDOFCHAIN
        put(s, struct.pack("<128I", *(vals + [nxt])))

    dir_bytes = bytearray()
    for n in nodes:
        e = bytearray(128)
        nm = n.name.encode("utf-16-le")
        e[0:len(nm)] = nm
        struct.pack_into("<HBB", e, 64, len(nm) + 2, n.kind, n.color)
        struct.pack_into("<III", e, 68, n.left, n.right, n.child)
        e[80:96] = n.clsid if n.kind != _STREAM else b"\0" * 16
        struct.pack_into("<I", e, 96, n.state if n.kind != _STREAM else 0)
        start = n.start if n.kind != _STORAGE else 0
        struct.pack_into("<III", e, 116, start, n.size & 0xFFFFFFFF, 0)
        dir_bytes += e
    while len(dir_bytes) % SS:
        e = bytearray(128)
        struct.pack_into("<III", e, 68, NOSTREAM, NOSTREAM, NOSTREAM)
        dir_bytes += e
    for k in range(n_dir):
        put(dir_start + k, bytes(dir_bytes[k * SS:(k + 1) * SS]))
    if n_minifat:
        mf = minifat + [FREESECT] * (n_minifat * (SS // 4) - len(minifat))
        for k in range(n_minifat):
            put(minifat_start + k, struct.pack("<%dI" % (SS // 4), *mf[k * (SS // 4):(k + 1) * (SS // 4)]))
    for k in range(n_mini):
        put(mini_start + k, bytes(mini[k * SS:(k + 1) * SS]))
    for n in big:
        for k in range(sectors_for(len(n.data))):
            put(n.start + k, n.data[k * SS:(k + 1) * SS])
    out += body
    return bytes(out)


def _check_name(name: str) -> str:
    if not name or len(name) > 31 or any(c in name for c in "/\\:!"):
        raise ValueError(f"invalid compound file entry name: {name!r}")
    return name
