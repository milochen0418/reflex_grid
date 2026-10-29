"""Changing a UserForm's design in place ([MS-OFORMS]), for ReflexGrid's form designer: moving,
resizing and changing controls, adding and removing them, and the form's own caption and size.

What an edit doesn't touch stays byte for byte, so a form made in Excel keeps everything ReflexGrid
doesn't model (mouse icons, IME modes, column widths, ActiveX classes…). A changed record is merged
rather than rewritten: the writer turns the control as it was and as it is now into records, and only
the fields where those two differ go into the original record (bit by bit for VariousPropertyBits and
a site's flags). Controls added here are written as the writer — checked by Excel — writes them.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, fields, replace
from typing import Any, Dict, List, Optional, Sequence, Tuple

from ..model.forms import ControlDesign, FontDesign, FormDesign, new_control
from .oforms import (
    CLASSES,
    CONTAINERS,
    FORM,
    FORM_FLAG_NOCLASSTABLE,
    LABEL,
    MORPH_DATA,
    TEXT_PROPS,
    VBFRAME,
    VPB_DEFAULT_LABEL,
    VPB_DEFAULT_MORPH,
    VPB_DEFAULT_PLAIN,
    DesignWriter,
    FormError,
    Record,
    Site,
    _read_font,
    _read_picture,
    _Reader,
    _Writer,
    _write_std_font,
    control_record,
    design_site,
    design_streams,
    form_record,
    frame_record,
    multipage_record,
    page_record,
    read_design,
    read_record,
    scan_sites,
    sites_block,
    strip_record,
    write_record,
    write_site,
)

# The controls the designer's toolbox adds.
ADDABLE = ("Label", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton", "Frame",
           "CommandButton", "TabStrip", "MultiPage", "ScrollBar", "SpinButton", "Image")
_IDENT = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,39}$")


# ---------------------------------------------------------------- merging


def _vpb_default(spec) -> int:
    return VPB_DEFAULT_MORPH if spec is MORPH_DATA else (VPB_DEFAULT_LABEL if spec is LABEL else VPB_DEFAULT_PLAIN)


def merge_record(orig: Record, base: Record, mine: Record) -> Record:
    """``orig`` with the changes from ``base`` to ``mine`` (the writer's records of a control
    before and after an edit): fields where those two agree keep what ``orig`` has."""
    spec = orig.spec
    out = Record(spec, orig.mask, dict(orig.values), dict(orig.pictures), orig.text_props, orig.tail)
    seen = set()
    entries = [(f.bit, f.name, f.kind) for f in spec.data] + [(x.bit, x.name, x.kind) for x in spec.extra] + [(bit, name, "flag") for bit, name in spec.flags]
    for bit, name, kind in entries:
        if (bit, name) in seen or kind == "pic":
            continue
        seen.add((bit, name))
        flag = 1 << bit
        b_has, m_has = bool(base.mask & flag), bool(mine.mask & flag)
        b_val, m_val = base.values.get(name), mine.values.get(name)
        if b_has == m_has and b_val == m_val:
            continue
        if name == "VariousPropertyBits":
            default = _vpb_default(spec)
            b_eff = b_val if b_has else default
            m_eff = m_val if m_has else default
            o_eff = orig.values.get(name, default) if orig.mask & flag else default
            new = o_eff ^ (b_eff ^ m_eff)
            if new == default:
                out.mask &= ~flag
                out.values.pop(name, None)
            else:
                out.mask |= flag
                out.values[name] = new
            continue
        if m_has:
            out.mask |= flag
            if kind != "flag":
                out.values[name] = m_val
        else:
            out.mask &= ~flag
            out.values.pop(name, None)
    if spec.text_props:
        empty = Record(TEXT_PROPS, 0)
        out.text_props = merge_record(orig.text_props or empty, base.text_props or empty, mine.text_props or empty)
    return out


def merge_site(orig: Site, base: Site, mine: Site) -> Site:
    out = replace(orig)
    for f in fields(Site):
        b, m = getattr(base, f.name), getattr(mine, f.name)
        if b == m:
            continue
        setattr(out, f.name, getattr(orig, f.name) ^ (b ^ m) if f.name == "bit_flags" else m)
    return out


# ---------------------------------------------------------------- one container's f stream


@dataclass
class _SiteEntry:
    site: Site
    raw: bytes
    dirty: bool = False


class Level:
    """A container's ``f`` stream, parsed so that what isn't changed is written back as it was."""

    def __init__(self, data: bytes, encoding: str) -> None:
        self.encoding = encoding
        rec, pos = read_record(data, FORM, encoding, 0, stop_after_cb=True)
        self.rec, self._rec_raw, self.rec_dirty = rec, data[:pos], False
        r = _Reader(data, pos)

        def part(bit: int, read) -> bytes:
            if not rec.mask & (1 << bit):
                return b""
            start = r.pos
            read()
            return data[start:r.pos]

        self.mouse = part(15, lambda: _read_picture(r))
        self.font_raw = part(20, lambda: _read_font(r, encoding))
        self.picture_raw = part(21, lambda: _read_picture(r))
        start = r.pos
        no_table = bool(rec.get("BooleanProperties", 4) & FORM_FLAG_NOCLASSTABLE)
        found = scan_sites(data, start, encoding, not no_table) or scan_sites(data, start, encoding, no_table)
        if found is None:
            raise FormError("the form's list of controls doesn't add up")
        class_end, spans, end = found
        self.class_table = data[start:class_end]
        self.sites = [_SiteEntry(site, data[a:b]) for site, a, b in spans]
        self._sites_raw = data[class_end:end]
        self.sites_changed = False
        self.tail = data[end:]

    def set_record(self, rec: Record) -> None:
        if rec != self.rec:
            self.rec = rec
            self.rec_dirty = True

    def set_font(self, font: FontDesign) -> None:
        w = _Writer()
        _write_std_font(w, font, self.encoding)
        self.font_raw = bytes(w.b)
        if not self.rec.mask & (1 << 20):
            self.rec.mask |= 1 << 20
        self.rec_dirty = True

    def to_bytes(self) -> bytes:
        head = write_record(self.rec) if self.rec_dirty else self._rec_raw
        if self.sites_changed or any(se.dirty for se in self.sites):
            body = sites_block([write_site(se.site) if se.dirty or not se.raw else se.raw for se in self.sites])
        else:
            body = self._sites_raw
        return head + self.mouse + self.font_raw + self.picture_raw + self.class_table + body + self.tail


# ---------------------------------------------------------------- the form


@dataclass
class _Loc:
    prefix: str
    index: int
    child: Optional[str]


class FormEditor:
    """Edits the designer storage of the form ``name`` (items as
    :attr:`gridcore.model.vba.VbaModule.designer` keeps them)."""

    def __init__(self, name: str, items: Sequence[Tuple[str, bytes]], encoding: str = "cp1252") -> None:
        self.name = name
        self.encoding = encoding
        self.files: Dict[str, bytes] = dict(items)
        self.levels: Dict[str, Level] = {}
        self.design = read_design(name, design_streams(items), encoding)
        self._index()

    # -- reading

    def level(self, prefix: str) -> Level:
        lv = self.levels.get(prefix)
        if lv is None:
            data = self.files.get(prefix + "f")
            if data is None:
                raise FormError(f"{prefix or 'the form'}: no f stream")
            lv = self.levels[prefix] = Level(data, self.encoding)
        return lv

    def _index(self) -> None:
        self.locs: Dict[str, _Loc] = {}

        def walk(prefix: str) -> None:
            for i, se in enumerate(self.level(prefix).sites):
                child = None
                if se.site.cache_index in CONTAINERS:
                    child = self._child(prefix, se.site.id)
                if se.site.name:
                    self.locs[se.site.name.lower()] = _Loc(prefix, i, child)
                if child is not None:
                    walk(child)

        walk("")

    def _child(self, prefix: str, site_id: int) -> Optional[str]:
        for path in self.files:
            if path.startswith(prefix) and path.endswith("/") and path != "/":
                part = path[len(prefix):-1]
                if "/" not in part and part[:1].lower() == "i" and part[1:].isdigit() and int(part[1:]) == site_id:
                    return path
        return None

    def _loc(self, name: str) -> _Loc:
        loc = self.locs.get(name.lower())
        if loc is None:
            raise FormError(f"There is no control named {name}.")
        return loc

    def _record_span(self, loc: _Loc) -> Tuple[int, int]:
        """Where the record of the control at ``loc`` is in its container's ``o`` stream."""
        off = 0
        for i, se in enumerate(self.level(loc.prefix).sites):
            if se.site.cache_index in CONTAINERS:
                continue
            if i == loc.index:
                return off, off + se.site.object_stream_size
            off += se.site.object_stream_size
        raise FormError("a container has no record of its own")

    def _max_id(self) -> int:
        top = 0
        for path in list(self.files):
            if path.endswith("f") and (path == "f" or path.endswith("/f")):
                lv = self.level(path[:-1])
                top = max([top, lv.rec.get("NextAvailableID", 0)] + [se.site.id for se in lv.sites])
        return top

    def _reload(self) -> None:
        self.design = read_design(self.name, design_streams(self.items()), self.encoding)
        self._index()

    # -- changes

    def set(self, name: str, changes: Dict[str, Any]) -> None:
        """Changes properties of a control (``ControlDesign`` fields)."""
        old = self.design.find(name)
        if old is None:
            raise FormError(f"There is no control named {name}.")
        new = replace(old, **changes)
        if new == old:
            return
        if new.name != old.name:
            if not _IDENT.match(new.name):
                raise FormError(f"{new.name!r} isn't a valid control name.")
            if self.design.find(new.name) is not None and new.name.lower() != old.name.lower():
                raise FormError(f"There is already a control named {new.name}.")
        loc = self._loc(name)
        lv = self.level(loc.prefix)
        se = lv.sites[loc.index]
        if old.kind == "Page":
            self._set_page(old, new, loc)
        else:
            site = merge_site(se.site, design_site(old, se.site.id, loc.index, se.site.object_stream_size), design_site(new, se.site.id, loc.index, se.site.object_stream_size))
            if old.kind not in ("Frame", "MultiPage"):
                a, b = self._record_span(loc)
                o = self.files.get(loc.prefix + "o", b"")
                spec = CLASSES[se.site.cache_index][1]
                orig, _ = read_record(o[a:b], spec, self.encoding, 0, b - a)
                rec = merge_record(orig, control_record(old), control_record(new))
                if rec != orig:
                    data = write_record(rec)
                    self.files[loc.prefix + "o"] = o[:a] + data + o[b:]
                    site.object_stream_size = len(data)
            elif old.kind == "Frame" and loc.child:
                inner = self.level(loc.child)
                inner.set_record(merge_record(inner.rec, frame_record(old), frame_record(new)))
                if old.font != new.font:
                    inner.set_font(new.font)
            elif old.kind == "MultiPage" and loc.child:
                self._set_multipage(old, new, loc)
            if site != se.site:
                se.site = site
                se.dirty = True
        self._reload()

    def _set_multipage(self, old: ControlDesign, new: ControlDesign, loc: _Loc) -> None:
        inner = self.level(loc.child)
        inner.set_record(merge_record(inner.rec, multipage_record(old), multipage_record(new)))
        # Its tabs (the TabStrip first in its storage), and the size of its pages.
        strip_loc = next((_Loc(loc.child, i, None) for i, se in enumerate(inner.sites) if se.site.cache_index == 18), None)
        if strip_loc is not None:
            self._merge_strip(strip_loc, old, new)
        for i, se in enumerate(inner.sites):
            if se.site.cache_index != 7:
                continue
            page_old = next((p for p in old.children if p.name == se.site.name), None)
            child = self._child(loc.child, se.site.id)
            if page_old is None or child is None:
                continue
            pl = self.level(child)
            pl.set_record(merge_record(pl.rec, page_record(page_old, old), page_record(page_old, new)))
        # The page on top is the one whose site is visible.
        pages = [se for se in inner.sites if se.site.cache_index == 7]
        if old.position != new.position and 0 <= new.position < len(pages):
            for j, se in enumerate(pages):
                want = (se.site.bit_flags | 2) if j == new.position else (se.site.bit_flags & ~2)
                if want != se.site.bit_flags:
                    se.site = replace(se.site, bit_flags=want)
                    se.dirty = True

    def _merge_strip(self, strip_loc: _Loc, old_mp: ControlDesign, new_mp: ControlDesign) -> None:
        a, b = self._record_span(strip_loc)
        o = self.files.get(strip_loc.prefix + "o", b"")
        orig, _ = read_record(o[a:b], CLASSES[18][1], self.encoding, 0, b - a)
        rec = merge_record(orig, strip_record(old_mp), strip_record(new_mp))
        if rec != orig:
            data = write_record(rec)
            self.files[strip_loc.prefix + "o"] = o[:a] + data + o[b:]
            se = self.level(strip_loc.prefix).sites[strip_loc.index]
            se.site = replace(se.site, object_stream_size=len(data))
            se.dirty = True

    def _set_page(self, old: ControlDesign, new: ControlDesign, loc: _Loc) -> None:
        """A page's caption and name live in its MultiPage's tabs; its tag and tip in its site."""
        mp = next((c for c in self.design.walk() if c.kind == "MultiPage" and any(p is old for p in c.children)), None)
        if mp is None:
            return
        new_mp = replace(mp, children=tuple(new if p is old else p for p in mp.children))
        strip_loc = next((_Loc(loc.prefix, i, None) for i, se in enumerate(self.level(loc.prefix).sites) if se.site.cache_index == 18), None)
        if strip_loc is not None:
            self._merge_strip(strip_loc, mp, new_mp)
        se = self.level(loc.prefix).sites[loc.index]
        site = replace(se.site, name=new.name, tag=new.tag, tip=new.tip)
        if site != se.site:
            se.site = site
            se.dirty = True

    def add(self, kind: str, parent: str = "", left: float = 6.0, top: float = 6.0, width: Optional[float] = None, height: Optional[float] = None) -> str:
        """Adds a control as the designer's toolbox does (in the form, a Frame or a page) and
        returns its name."""
        if kind not in ADDABLE:
            raise FormError(f"{kind} isn't a control the toolbox has.")
        font = self.design.font
        if parent:
            host = self.design.find(parent)
            loc = self._loc(parent)
            if host is None or host.kind not in ("Frame", "Page", "MultiPage") or not loc.child:
                raise FormError(f"Controls go on the form, in a Frame or on a page, not in {parent}.")
            if host.kind == "MultiPage":
                pages = [p for p in host.children if p.kind == "Page"]
                if not pages:
                    raise FormError(f"{parent} has no page to put a control on.")
                host = pages[host.position if 0 <= host.position < len(pages) else 0]
                loc = self._loc(host.name)
            prefix = loc.child
            font = host.font
        else:
            prefix = ""
        n = 1
        while self.design.find(f"{kind}{n}") is not None:
            n += 1
        name = f"{kind}{n}"
        lv = self.level(prefix)
        tab = max([se.site.tab_index for se in lv.sites if se.site.cache_index != 18] + [-1]) + 1
        extra: Dict[str, Any] = {"left": max(0.0, float(left)), "top": max(0.0, float(top)), "tab_index": tab, "font": font}
        if width:
            extra["width"] = max(1.0, float(width))
        if height:
            extra["height"] = max(1.0, float(height))
        c = new_control(name, kind, **extra)
        if kind == "MultiPage":
            c = replace(c, children=tuple(replace(p, font=font) for p in c.children))
        w = DesignWriter(self.encoding, iter(range(self._max_id() + 1, 1 << 30)))
        site, data, top_id = w.control(prefix, c, len(lv.sites))
        lv.sites.append(_SiteEntry(site, b"", dirty=True))
        lv.sites_changed = True
        if data:
            self.files[prefix + "o"] = self.files.get(prefix + "o", b"") + data
        for path, blob in w.out:
            self.files[path] = blob
        # The largest ID used, up the containers to the form.
        parts = prefix.split("/")[:-1]
        for k in range(len(parts), -1, -1):
            up = "".join(p + "/" for p in parts[:k])
            ul = self.level(up)
            if top_id > ul.rec.get("NextAvailableID", 0):
                rec = replace(ul.rec, values=dict(ul.rec.values), mask=ul.rec.mask | (1 << 3))
                rec.values["NextAvailableID"] = top_id
                ul.set_record(rec)
        self._reload()
        return name

    def delete(self, name: str) -> None:
        old = self.design.find(name)
        if old is None:
            raise FormError(f"There is no control named {name}.")
        if old.kind == "Page":
            raise FormError("A MultiPage's pages can't be removed here.")
        loc = self._loc(name)
        lv = self.level(loc.prefix)
        if loc.child:
            for path in [p for p in self.files if p.startswith(loc.child)]:
                del self.files[path]
            for prefix in [p for p in self.levels if p.startswith(loc.child)]:
                del self.levels[prefix]
        else:
            a, b = self._record_span(loc)
            o = self.files.get(loc.prefix + "o", b"")
            self.files[loc.prefix + "o"] = o[:a] + o[b:]
        del lv.sites[loc.index]
        lv.sites_changed = True
        self._reload()

    def set_form(self, changes: Dict[str, Any]) -> None:
        """Changes the form's own properties (``FormDesign`` fields: caption, width and height of
        its inside, start-up position, ShowModal, colours, font…)."""
        old = self.design
        new = replace(old, **changes)
        if new == old:
            return
        root = self.level("")
        root.set_record(merge_record(root.rec, form_record(old), form_record(new)))
        if old.font != new.font:
            root.set_font(new.font)
        text = self.files.get(VBFRAME, b"").decode(self.encoding, "replace")
        text = _frame_set(text, "Caption", '"' + new.caption.replace('"', '""') + '"', old.caption != new.caption)
        text = _frame_set(text, "ClientHeight", str(int(round(new.height * 20))), old.height != new.height)
        text = _frame_set(text, "ClientLeft", str(int(round(new.left * 20))), old.left != new.left)
        text = _frame_set(text, "ClientTop", str(int(round(new.top * 20))), old.top != new.top)
        text = _frame_set(text, "ClientWidth", str(int(round(new.width * 20))), old.width != new.width)
        if old.show_modal != new.show_modal:
            text = _frame_set(text, "ShowModal", None if new.show_modal else "0   'False", True)
        names = {0: "Manual", 1: "CenterOwner", 2: "CenterScreen", 3: "Windows Default"}
        text = _frame_set(text, "StartUpPosition", f"{new.start_up_position}  '{names.get(new.start_up_position, 'Manual')}", old.start_up_position != new.start_up_position)
        self.files[VBFRAME] = text.encode(self.encoding, "replace")
        self._reload()

    # -- writing

    def items(self) -> Tuple[Tuple[str, bytes], ...]:
        for prefix, lv in self.levels.items():
            if (prefix + "f") in self.files or prefix == "":
                self.files[prefix + "f"] = lv.to_bytes()
        return tuple(sorted(self.files.items()))


def renamed_designer(items: Sequence[Tuple[str, bytes]], name: str, encoding: str = "cp1252") -> Tuple[Tuple[str, bytes], ...]:
    """A form's designer items with the form named ``name`` in its VBFrame header (the
    ``Begin {…} Name`` line), as the VBA editor rewrites it when a form's (Name) changes.
    Everything else is kept byte for byte."""
    out = []
    for path, data in items:
        if path == VBFRAME:
            text = data.decode(encoding, "replace")
            sep = "\r\n" if "\r\n" in text else "\n"
            lines = text.split(sep)
            for i, line in enumerate(lines):
                parts = line.split()
                if len(parts) >= 2 and parts[0].upper() == "BEGIN":
                    lines[i] = f"Begin {parts[1]} {name} "
                    break
            data = sep.join(lines).encode(encoding, "replace")
        out.append((path, data))
    return tuple(out)


_FRAME_ORDER = ("Caption", "ClientHeight", "ClientLeft", "ClientTop", "ClientWidth", "OleObjectBlob", "ShowModal", "StartUpPosition", "TypeInfoVer")


def _frame_set(text: str, key: str, value: Optional[str], changed: bool) -> str:
    """The VBFrame header with ``key`` set to ``value`` (None: removed), in the editor's order."""
    if not changed:
        return text
    lines = text.replace("\r\n", "\n").split("\n")
    end = next((i for i, ln in enumerate(lines) if ln.strip().upper() == "END"), len(lines))
    props = [(i, ln.strip().split("=", 1)[0].strip()) for i, ln in enumerate(lines[:end]) if "=" in ln]
    line = None if value is None else f"   {key:<16}=   {value}"
    for i, k in props:
        if k.lower() == key.lower():
            if line is None:
                del lines[i]
            else:
                lines[i] = line
            return "\r\n".join(lines)
    if line is None:
        return text
    order = {k.lower(): n for n, k in enumerate(_FRAME_ORDER)}
    at = next((i for i, k in props if order.get(k.lower(), 99) > order.get(key.lower(), 99)), end)
    lines.insert(at, line)
    return "\r\n".join(lines)
