"""UserForm designs in a VBA project ([MS-OFORMS]): reading them for ReflexGrid to show and run,
and writing new ones.

A form's design lives in a storage of ``vbaProject.bin`` named after the form, beside ``VBA/``::

    UserForm1/\\x03VBFrame     the form's caption, size, start-up position (text)
    UserForm1/\\x01CompObj     the class ("Microsoft Forms 2.0 Form")
    UserForm1/f               the form's own properties, then one site per control
    UserForm1/o               each control's property record, back to back
    UserForm1/i06/f, i06/o    a container (Frame, MultiPage, Page) keeps its children in a
                              storage of its own, named by its site's ID

Every record starts with a property mask: a property is stored only when it differs from its
default, in a DataBlock (fields aligned to their own size, counted from the record's start), an
ExtraDataBlock (strings, sizes) and, after the record, pictures and fonts. The tables below
follow the specification's diagrams field by field. The reader cross-checks the lengths at
every level (each record's ``cb``, the sites' ``CountOfBytes``, their ``ObjectStreamSize``
against the length of ``o``, every child storage claimed by a container) and reports what it
can't account for instead of guessing.

The design stays in the workbook exactly as it was read (the project writer copies its
streams); this module only interprets it, or builds the streams of a new form.
"""

from __future__ import annotations

import struct
import uuid
from dataclasses import dataclass, field, replace
from typing import Any, Dict, List, Optional, Sequence, Tuple

from ..model.forms import ControlDesign, FontDesign, FormDesign

HIMETRIC_PER_POINT = 2540 / 72


class FormError(Exception):
    """A designer storage that doesn't read as [MS-OFORMS] describes it."""


# ---------------------------------------------------------------- record tables


@dataclass(frozen=True)
class F:
    """A DataBlock field: its mask bit, name, size in bytes and kind (``u``/``i`` integers,
    ``str`` a CountOfBytesWithCompressionFlag for a string in the ExtraDataBlock, ``pic`` the
    0xFFFF marker of a picture kept in the StreamData)."""

    bit: int
    name: str
    size: int
    kind: str = "u"


@dataclass(frozen=True)
class X:
    """An ExtraDataBlock field: ``size``/``pos`` (two 32-bit integers), ``str``, ``arr`` (raw
    bytes, as long as the DataBlock field ``size_from`` says)."""

    bit: int
    name: str
    kind: str
    size_from: str = ""


@dataclass(frozen=True)
class Spec:
    name: str
    data: Tuple[F, ...]
    extra: Tuple[X, ...] = ()
    stream: Tuple[Tuple[int, str], ...] = ()  # StreamData pictures, in stored order
    major: int = 2
    mask64: bool = False
    text_props: bool = False
    flags: Tuple[Tuple[int, str], ...] = ()  # mask bits that are properties without data


TEXT_PROPS = Spec(
    "TextProps",
    (F(0, "FontName", 4, "str"), F(1, "FontEffects", 4), F(2, "FontHeight", 4), F(4, "FontCharSet", 1),
     F(5, "FontPitchAndFamily", 1), F(6, "ParagraphAlign", 1), F(7, "FontWeight", 2)),
    (X(0, "FontName", "str"),),
)

COMMAND_BUTTON = Spec(
    "CommandButton",
    (F(0, "ForeColor", 4), F(1, "BackColor", 4), F(2, "VariousPropertyBits", 4), F(3, "Caption", 4, "str"),
     F(4, "PicturePosition", 4), F(6, "MousePointer", 1), F(7, "Picture", 2, "pic"), F(8, "Accelerator", 2),
     F(10, "MouseIcon", 2, "pic")),
    (X(3, "Caption", "str"), X(5, "Size", "size")),
    ((7, "Picture"), (10, "MouseIcon")),
    text_props=True,
    flags=((9, "TakeFocusOnClick"),),
)

LABEL = Spec(
    "Label",
    (F(0, "ForeColor", 4), F(1, "BackColor", 4), F(2, "VariousPropertyBits", 4), F(3, "Caption", 4, "str"),
     F(4, "PicturePosition", 4), F(6, "MousePointer", 1), F(7, "BorderColor", 4), F(8, "BorderStyle", 2),
     F(9, "SpecialEffect", 2), F(10, "Picture", 2, "pic"), F(11, "Accelerator", 2), F(12, "MouseIcon", 2, "pic")),
    (X(3, "Caption", "str"), X(5, "Size", "size")),
    ((10, "Picture"), (12, "MouseIcon")),
    text_props=True,
)

IMAGE = Spec(
    "Image",
    (F(3, "BorderColor", 4), F(4, "BackColor", 4), F(5, "BorderStyle", 1), F(6, "MousePointer", 1),
     F(7, "PictureSizeMode", 1), F(8, "SpecialEffect", 1), F(10, "Picture", 2, "pic"), F(11, "PictureAlignment", 1),
     F(13, "VariousPropertyBits", 4), F(14, "MouseIcon", 2, "pic")),
    (X(9, "Size", "size"),),
    ((10, "Picture"), (14, "MouseIcon")),
    flags=((2, "AutoSize"), (12, "PictureTiling")),
)

MORPH_DATA = Spec(
    "MorphData",
    (F(0, "VariousPropertyBits", 4), F(1, "BackColor", 4), F(2, "ForeColor", 4), F(3, "MaxLength", 4),
     F(4, "BorderStyle", 1), F(5, "ScrollBars", 1), F(6, "DisplayStyle", 1), F(7, "MousePointer", 1),
     F(9, "PasswordChar", 2), F(10, "ListWidth", 4), F(11, "BoundColumn", 2), F(12, "TextColumn", 2, "i"),
     F(13, "ColumnCount", 2, "i"), F(14, "ListRows", 2), F(15, "cColumnInfo", 2), F(16, "MatchEntry", 1),
     F(17, "ListStyle", 1), F(18, "ShowDropButtonWhen", 1), F(20, "DropButtonStyle", 1), F(21, "MultiSelect", 1),
     F(22, "Value", 4, "str"), F(23, "Caption", 4, "str"), F(24, "PicturePosition", 4), F(25, "BorderColor", 4),
     F(26, "SpecialEffect", 4), F(27, "MouseIcon", 2, "pic"), F(28, "Picture", 2, "pic"), F(29, "Accelerator", 2),
     F(32, "GroupName", 4, "str")),
    (X(8, "Size", "size"), X(22, "Value", "str"), X(23, "Caption", "str"), X(32, "GroupName", "str")),
    ((27, "MouseIcon"), (28, "Picture")),
    mask64=True,
    text_props=True,
)

SPIN_BUTTON = Spec(
    "SpinButton",
    (F(0, "ForeColor", 4), F(1, "BackColor", 4), F(2, "VariousPropertyBits", 4), F(5, "Min", 4, "i"),
     F(6, "Max", 4, "i"), F(7, "Position", 4, "i"), F(8, "PrevEnabled", 4, "i"), F(9, "NextEnabled", 4, "i"),
     F(10, "SmallChange", 4, "i"), F(11, "Orientation", 4, "i"), F(12, "Delay", 4), F(13, "MouseIcon", 2, "pic"),
     F(14, "MousePointer", 1)),
    (X(3, "Size", "size"),),
    ((13, "MouseIcon"),),
)

SCROLL_BAR = Spec(
    "ScrollBar",
    (F(0, "ForeColor", 4), F(1, "BackColor", 4), F(2, "VariousPropertyBits", 4), F(4, "MousePointer", 1),
     F(5, "Min", 4, "i"), F(6, "Max", 4, "i"), F(7, "Position", 4, "i"), F(9, "PrevEnabled", 4, "i"),
     F(10, "NextEnabled", 4, "i"), F(11, "SmallChange", 4, "i"), F(12, "LargeChange", 4, "i"),
     F(13, "Orientation", 4, "i"), F(14, "ProportionalThumb", 2, "i"), F(15, "Delay", 4), F(16, "MouseIcon", 2, "pic")),
    (X(3, "Size", "size"),),
    ((16, "MouseIcon"),),
)

TAB_STRIP = Spec(
    "TabStrip",
    (F(0, "ListIndex", 4, "i"), F(1, "BackColor", 4), F(2, "ForeColor", 4), F(5, "ItemsSize", 4),
     F(6, "MousePointer", 1), F(8, "TabOrientation", 4), F(9, "TabStyle", 4), F(11, "TabFixedWidth", 4),
     F(12, "TabFixedHeight", 4), F(15, "TipStringsSize", 4), F(17, "NamesSize", 4), F(18, "VariousPropertyBits", 4),
     F(20, "TabsAllocated", 4), F(21, "TagsSize", 4), F(22, "TabData", 4), F(23, "AcceleratorsSize", 4),
     F(24, "MouseIcon", 2, "pic")),
    (X(4, "Size", "size"), X(5, "Items", "arr", "ItemsSize"), X(15, "TipStrings", "arr", "TipStringsSize"),
     X(17, "TabNames", "arr", "NamesSize"), X(21, "Tags", "arr", "TagsSize"),
     X(23, "Accelerators", "arr", "AcceleratorsSize")),
    ((24, "MouseIcon"),),
    text_props=True,
    flags=((10, "MultiRow"), (13, "Tooltips"), (19, "NewVersion")),
)

FORM = Spec(
    "Form",
    (F(1, "BackColor", 4), F(2, "ForeColor", 4), F(3, "NextAvailableID", 4), F(6, "BooleanProperties", 4),
     F(7, "BorderStyle", 1), F(8, "MousePointer", 1), F(9, "ScrollBars", 1), F(13, "GroupCnt", 4, "i"),
     F(15, "MouseIcon", 2, "pic"), F(16, "Cycle", 1), F(17, "SpecialEffect", 1), F(18, "BorderColor", 4),
     F(19, "Caption", 4, "str"), F(20, "Font", 2, "pic"), F(21, "Picture", 2, "pic"), F(22, "Zoom", 4),
     F(23, "PictureAlignment", 1), F(25, "PictureSizeMode", 1), F(26, "ShapeCookie", 4), F(27, "DrawBuffer", 4)),
    (X(10, "DisplayedSize", "size"), X(11, "LogicalSize", "size"), X(12, "ScrollPosition", "pos"),
     X(19, "Caption", "str")),
    major=4,
    flags=((24, "PictureTiling"),),
)

# ClsidCacheIndex values ([MS-OFORMS] 2.4.5) and the control each names.
CLASSES: Dict[int, Tuple[str, Optional[Spec]]] = {
    7: ("Page", None),        # a container: its record is the f stream of its own storage
    12: ("Image", IMAGE),
    14: ("Frame", None),
    15: ("MorphData", MORPH_DATA),
    16: ("SpinButton", SPIN_BUTTON),
    17: ("CommandButton", COMMAND_BUTTON),
    18: ("TabStrip", TAB_STRIP),
    21: ("Label", LABEL),
    23: ("TextBox", MORPH_DATA),
    24: ("ListBox", MORPH_DATA),
    25: ("ComboBox", MORPH_DATA),
    26: ("CheckBox", MORPH_DATA),
    27: ("OptionButton", MORPH_DATA),
    28: ("ToggleButton", MORPH_DATA),
    47: ("ScrollBar", SCROLL_BAR),
    57: ("MultiPage", None),
}
CONTAINERS = {7, 14, 57}
CACHE_INDEX = {name: idx for idx, (name, _) in CLASSES.items() if name != "MorphData"}
# A MorphData's DisplayStyle says which control it is.
MORPH_KINDS = {1: "TextBox", 2: "ListBox", 3: "ComboBox", 4: "CheckBox", 5: "OptionButton", 6: "ToggleButton", 7: "ComboBox"}
MORPH_STYLES = {"TextBox": 1, "ListBox": 2, "ComboBox": 3, "CheckBox": 4, "OptionButton": 5, "ToggleButton": 6}

# VariousPropertyBits defaults by control ([MS-OFORMS] 2.6 VariousPropertyBits).
VPB_DEFAULT_MORPH = 0x2C80081B
VPB_DEFAULT_PLAIN = 0x0000001B
VPB_DEFAULT_LABEL = 0x0080001B
VPB_ENABLED = 1 << 1
VPB_LOCKED = 1 << 2
VPB_BACKSTYLE = 1 << 3
VPB_COLUMNHEADS = 1 << 10
VPB_MATCHREQUIRED = 1 << 12
VPB_ALIGNMENT = 1 << 13
VPB_EDITABLE = 1 << 14
VPB_ENTERKEY = 1 << 20
VPB_WORDWRAP = 1 << 23
VPB_AUTOSIZE = 1 << 28
VPB_MULTILINE = 1 << 31

SITE_FLAG_DEFAULT = 0x33
SITE_TABSTOP = 1 << 0
SITE_VISIBLE = 1 << 1
SITE_DEFAULT = 1 << 2
SITE_CANCEL = 1 << 3
SITE_STREAMED = 1 << 4
SITE_AUTOSIZE = 1 << 5
SITE_PROMOTE = 1 << 19

FORM_FLAG_DESINK = 1 << 14
FORM_FLAG_NOCLASSTABLE = 1 << 15

FONT_EFFECT_BOLD, FONT_EFFECT_ITALIC, FONT_EFFECT_UNDERLINE, FONT_EFFECT_STRIKE = 1, 2, 4, 8

CLSID_STDPICTURE = uuid.UUID("0BE35204-8F91-11CE-9DE3-00AA004BB851").bytes_le
CLSID_STDFONT = uuid.UUID("0BE35203-8F91-11CE-9DE3-00AA004BB851").bytes_le
CLSID_TEXTPROPS = uuid.UUID("AFC20920-DA4E-11CE-B943-00AA006887B4").bytes_le
# The class of a UserForm ("Microsoft Forms 2.0 Form") and its \x03VBFrame header.
CLSID_FORM = "{C62A69F0-16DC-11CE-9E98-00AA00574A4F}"
VBFRAME = "\x03VBFrame"
COMPOBJ = "\x01CompObj"

# Site DataBlock fields (in mask-bit order) and the strings of its ExtraDataBlock.
SITE_FIELDS: Tuple[Tuple[int, str, int], ...] = (
    (0, "NameData", 4), (1, "TagData", 4), (2, "ID", 4), (3, "HelpContextID", 4), (4, "BitFlags", 4),
    (5, "ObjectStreamSize", 4), (6, "TabIndex", 2), (7, "ClsidCacheIndex", 2), (9, "GroupID", 2),
    (11, "ControlTipTextData", 4), (12, "RuntimeLicKeyData", 4), (13, "ControlSourceData", 4),
    (14, "RowSourceData", 4),
)
SITE_SIGNED = {"ID", "HelpContextID", "TabIndex"}
SITE_STRINGS_BEFORE = (("NameData", "Name"), ("TagData", "Tag"))
SITE_STRINGS_AFTER = (("ControlTipTextData", "ControlTipText"), ("RuntimeLicKeyData", "RuntimeLicKey"),
                      ("ControlSourceData", "ControlSource"), ("RowSourceData", "RowSource"))


# ---------------------------------------------------------------- reading records


class _Reader:
    def __init__(self, data: bytes, pos: int = 0, end: Optional[int] = None):
        self.data = data
        self.pos = pos
        self.end = len(data) if end is None else end

    def need(self, n: int) -> None:
        if n < 0 or self.pos + n > self.end:
            raise FormError(f"read past the end at {self.pos}+{n} of {self.end}")

    def int(self, size: int, signed: bool = False) -> int:
        self.need(size)
        fmt = {1: "b", 2: "h", 4: "i"}[size] if signed else {1: "B", 2: "H", 4: "I"}[size]
        v = struct.unpack_from("<" + fmt, self.data, self.pos)[0]
        self.pos += size
        return int(v)

    def take(self, n: int) -> bytes:
        self.need(n)
        b = self.data[self.pos:self.pos + n]
        self.pos += n
        return b

    def align(self, base: int, size: int) -> None:
        over = (self.pos - base) % size
        if over:
            self.take(size - over)


@dataclass
class Record:
    """A control's properties as stored: ``values`` (integers, decoded strings, ``(w, h)`` sizes,
    raw arrays) for the mask bits that are set, its pictures, its TextProps and anything after
    them (a ListBox's column widths, a TabStrip's tab flags)."""

    spec: Spec
    mask: int
    values: Dict[str, Any] = field(default_factory=dict)
    pictures: Dict[str, bytes] = field(default_factory=dict)
    text_props: Optional["Record"] = None
    tail: bytes = b""

    def has(self, name: str) -> bool:
        return name in self.values

    def get(self, name: str, default: Any = None) -> Any:
        return self.values.get(name, default)

    def flag(self, name: str) -> bool:
        for bit, n in self.spec.flags:
            if n == name:
                return bool(self.mask & (1 << bit))
        return False


def _decode(raw: bytes, compressed: bool, encoding: str = "") -> str:
    """A stored string: "compressed" ones are UTF-16 with each (zero) high byte dropped
    ([MS-OFORMS] 3.1), so one byte per character, Latin-1."""
    return raw.decode("latin-1" if compressed else "utf-16-le", "replace")


def read_record(data: bytes, spec: Spec, encoding: str, pos: int = 0, end: Optional[int] = None, stop_after_cb: bool = False) -> Tuple[Record, int]:
    """One record at ``pos``; returns it and where it ends (with its StreamData, TextProps and,
    up to ``end``, its tail)."""
    limit = len(data) if end is None else end
    r = _Reader(data, pos, limit)
    base = r.pos
    minor, major = r.int(1), r.int(1)
    if (minor, major) != (0, spec.major):
        raise FormError(f"{spec.name}: not a record of this kind (version {major}.{minor})")
    cb = r.int(2)
    mask = r.int(4)
    if spec.mask64:
        mask |= r.int(4) << 32
    rec = Record(spec, mask)
    after_cb = base + 4 + cb
    if after_cb > limit:
        raise FormError(f"{spec.name}: its size runs past the end")
    lengths: Dict[str, int] = {}
    for f in spec.data:
        if not mask & (1 << f.bit):
            continue
        r.align(base, f.size)
        v = r.int(f.size, f.kind == "i")
        if f.kind == "str":
            lengths[f.name] = v
        elif f.kind != "pic":
            rec.values[f.name] = v
    r.align(base, 4)
    for x in spec.extra:
        if not mask & (1 << x.bit):
            continue
        if x.kind in ("size", "pos"):
            rec.values[x.name] = (r.int(4, True), r.int(4, True))
        elif x.kind == "str":
            packed = lengths.get(x.name, 0)
            raw = r.take(packed & 0x7FFFFFFF)
            rec.values[x.name] = _decode(raw, bool(packed & 0x80000000), encoding)
            r.align(base, 4)
        else:
            rec.values[x.name] = r.take(rec.values.get(x.size_from, 0))
    if r.pos != after_cb:
        if r.pos > after_cb:
            raise FormError(f"{spec.name}: its fields run past its size")
        raise FormError(f"{spec.name}: {after_cb - r.pos} bytes inside it aren't accounted for")
    if stop_after_cb:
        return rec, r.pos
    for bit, name in spec.stream:
        if mask & (1 << bit):
            rec.pictures[name] = _read_picture(r)
    if spec.text_props:
        r.need(4)
        tp_end = r.pos + 4 + struct.unpack_from("<H", data, r.pos + 2)[0]
        rec.text_props, r.pos = read_record(data, TEXT_PROPS, encoding, r.pos, tp_end, stop_after_cb=True)
    if end is not None and r.pos < end:
        rec.tail = r.take(end - r.pos)
    return rec, r.pos


def _read_picture(r: _Reader) -> bytes:
    """A GuidAndPicture: the class, a preamble, the size, then the picture file itself."""
    r.take(16)
    r.int(4)
    return r.take(r.int(4))


def _read_font(r: _Reader, encoding: str) -> FontDesign:
    """The form's or a Frame's GuidAndFont: a StdFont, or TextProps."""
    guid = r.take(16)
    if guid == CLSID_STDFONT:
        if r.int(1) != 1:
            raise FormError("unknown StdFont version")
        r.int(2)  # charset
        flags = r.int(1)
        weight = r.int(2)
        height = r.int(4)
        face = r.take(r.int(1)).decode(encoding, "replace")
        return FontDesign(face or "Tahoma", round(height / 10000.0, 4), weight >= 600, bool(flags & 2), bool(flags & 4), bool(flags & 8))
    if guid == CLSID_TEXTPROPS:
        start = r.pos
        rec, r.pos = read_record(r.data, TEXT_PROPS, encoding, start, stop_after_cb=True)
        return _font_of(rec)
    raise FormError("unknown font class in a form's StreamData")


def _font_of(tp: Optional[Record]) -> FontDesign:
    if tp is None:
        return FontDesign()
    eff = tp.get("FontEffects", 0)
    weight = tp.get("FontWeight", 0)
    return FontDesign(
        tp.get("FontName") or "Tahoma",
        tp.get("FontHeight", 160) / 20.0,
        bool(eff & FONT_EFFECT_BOLD) or weight >= 600,
        bool(eff & FONT_EFFECT_ITALIC),
        bool(eff & FONT_EFFECT_UNDERLINE),
        bool(eff & FONT_EFFECT_STRIKE),
    )


# ---------------------------------------------------------------- reading sites and form streams


@dataclass
class Site:
    """An embedded control as its container's ``f`` stream lists it (OleSiteConcreteControl)."""

    name: str = ""
    tag: str = ""
    id: int = 0
    bit_flags: int = SITE_FLAG_DEFAULT
    object_stream_size: int = 0
    tab_index: int = -1
    cache_index: int = 0x7FFF
    group_id: int = 0
    position: Tuple[int, int] = (0, 0)
    tip: str = ""
    control_source: str = ""
    row_source: str = ""
    help_context: int = 0
    lic_key: str = ""


def _read_site(r: _Reader, encoding: str) -> Site:
    start = r.pos
    if r.int(2) != 0:
        raise FormError("unknown site version")
    cb = r.int(2)
    mask = r.int(4)
    vals: Dict[str, int] = {}
    for bit, name, size in SITE_FIELDS:
        if mask & (1 << bit):
            r.align(start, size)
            vals[name] = r.int(size, name in SITE_SIGNED)
    r.align(start, 4)
    strings: Dict[str, str] = {}

    def string(len_field: str, name: str) -> None:
        if len_field in vals:
            packed = vals[len_field]
            strings[name] = _decode(r.take(packed & 0x7FFFFFFF), bool(packed & 0x80000000), encoding)
            r.align(start, 4)

    for lf, n in SITE_STRINGS_BEFORE:
        string(lf, n)
    pos = (0, 0)
    if mask & (1 << 8):
        pos = (r.int(4, True), r.int(4, True))
    for lf, n in SITE_STRINGS_AFTER:
        string(lf, n)
    if r.pos != start + 4 + cb:
        raise FormError(f"site {strings.get('Name', '?')!r}: {start + 4 + cb - r.pos} bytes aren't accounted for")
    return Site(
        name=strings.get("Name", ""),
        tag=strings.get("Tag", ""),
        id=vals.get("ID", 0),
        bit_flags=vals.get("BitFlags", SITE_FLAG_DEFAULT),
        object_stream_size=vals.get("ObjectStreamSize", 0),
        tab_index=vals.get("TabIndex", -1),
        cache_index=vals.get("ClsidCacheIndex", 0x7FFF),
        group_id=vals.get("GroupID", 0),
        position=pos,
        tip=strings.get("ControlTipText", ""),
        control_source=strings.get("ControlSource", ""),
        row_source=strings.get("RowSource", ""),
        help_context=vals.get("HelpContextID", 0),
        lic_key=strings.get("RuntimeLicKey", ""),
    )


@dataclass
class FormStream:
    record: Record
    font: Optional[FontDesign]
    picture: bytes
    sites: List[Site]


def _try_sites(data: bytes, pos: int, encoding: str, class_table: bool) -> Optional[List[Site]]:
    found = scan_sites(data, pos, encoding, class_table)
    return [site for site, _, _ in found[1]] if found is not None else None


def scan_sites(data: bytes, pos: int, encoding: str, class_table: bool) -> Optional[Tuple[int, List[Tuple[Site, int, int]], int]]:
    """The class table (when ``class_table``) and the sites at ``pos``: where the class table
    ends, each site with where it starts and ends, and where the sites end; None when they don't
    read that way."""
    try:
        r = _Reader(data, pos)
        if class_table:
            for _ in range(r.int(2)):
                # SiteClassInfo: version, cb, then its fields (ActiveX controls this form uses).
                if r.int(2) != 0:
                    return None
                r.take(r.int(2))
        class_end = r.pos
        count = r.int(4)
        nbytes = r.int(4)
        if count > 100_000:
            return None
        end = r.pos + nbytes
        if end > len(data):
            return None
        depths_start = r.pos
        seen = 0
        while seen < count:
            r.int(1)  # depth
            toc = r.int(1)
            if toc & 0x80:
                seen += toc & 0x7F
                r.int(1)
            else:
                seen += 1
        if seen != count:
            return None
        r.align(depths_start, 4)
        sites = []
        for _ in range(count):
            a = r.pos
            site = _read_site(r, encoding)
            sites.append((site, a, r.pos))
        if r.pos != end:
            return None
        return class_end, sites, end
    except FormError:
        return None


def read_form_stream(data: bytes, encoding: str) -> FormStream:
    """A container's ``f`` stream: its FormControl record, font and picture, and its sites."""
    rec, pos = read_record(data, FORM, encoding, 0, stop_after_cb=True)
    r = _Reader(data, pos)
    font = None
    picture = b""
    if rec.mask & (1 << 15):
        _read_picture(r)  # MouseIcon
    if rec.mask & (1 << 20):
        font = _read_font(r, encoding)
    if rec.mask & (1 << 21):
        picture = _read_picture(r)
    no_table = bool(rec.get("BooleanProperties", 4) & FORM_FLAG_NOCLASSTABLE)
    sites = _try_sites(data, r.pos, encoding, not no_table)
    if sites is None:
        sites = _try_sites(data, r.pos, encoding, no_table)
    if sites is None:
        raise FormError("the form's list of controls doesn't add up")
    return FormStream(rec, font, picture, sites)


# ---------------------------------------------------------------- the design


def himetric_to_points(v: int) -> float:
    return round(v / HIMETRIC_PER_POINT, 2)


def points_to_himetric(v: float) -> int:
    return int(round(v * HIMETRIC_PER_POINT))


def _strings(blob: bytes, count: int, encoding: str) -> List[str]:
    """An array of CountOfBytesWithCompressionFlag + string, each padded to 4 bytes."""
    out: List[str] = []
    r = _Reader(blob)
    try:
        while r.pos < len(blob) and len(out) < count:
            packed = r.int(4)
            out.append(_decode(r.take(packed & 0x7FFFFFFF), bool(packed & 0x80000000), encoding))
            r.align(0, 4)
    except FormError:
        pass
    return out


def _column_widths(rec: Record) -> str:
    """ColumnWidths from the MorphDataColumnInfo records after a list's TextProps (points)."""
    n = rec.get("cColumnInfo", 0)
    if not n or not rec.tail:
        return ""
    widths: List[str] = []
    r = _Reader(rec.tail)
    try:
        for _ in range(n):
            base = r.pos
            r.int(2)
            cb = r.int(2)
            mask = r.int(4)
            w = r.int(4, True) if mask & 1 else -1
            r.pos = base + 4 + cb
            widths.append("" if w < 0 else f"{himetric_to_points(w):g} pt")
    except FormError:
        return ""
    # Stored last column first.
    return ";".join(reversed(widths))


def _control(site: Site, rec: Optional[Record], children: Tuple[ControlDesign, ...], encoding: str, warnings: List[str], inherited: FontDesign) -> ControlDesign:
    kind, _ = CLASSES.get(site.cache_index, ("", None))
    if site.cache_index >= 0x8000:
        kind = "ActiveX"
    if kind == "MorphData":
        kind = MORPH_KINDS.get(rec.get("DisplayStyle", 1) if rec else 1, "TextBox")
    elif kind in ("TextBox", "ListBox", "ComboBox", "CheckBox", "OptionButton", "ToggleButton") and rec is not None and rec.has("DisplayStyle"):
        kind = MORPH_KINDS.get(rec.get("DisplayStyle"), kind)
    if not kind:
        kind = f"Control{site.cache_index}"
        warnings.append(f"{site.name}: a control ReflexGrid doesn't know (class {site.cache_index})")
    flags = site.bit_flags
    common: Dict[str, Any] = dict(
        name=site.name,
        kind=kind,
        left=himetric_to_points(site.position[0]),
        top=himetric_to_points(site.position[1]),
        id=site.id,
        tab_index=site.tab_index,
        tab_stop=bool(flags & SITE_TABSTOP),
        visible=bool(flags & SITE_VISIBLE),
        default=bool(flags & SITE_DEFAULT),
        cancel=bool(flags & SITE_CANCEL),
        tag=site.tag,
        tip=site.tip,
        control_source=site.control_source,
        row_source=site.row_source,
        children=children,
        font=inherited,
    )
    if rec is None:
        return ControlDesign(**common)
    if "Size" in rec.values:
        w, h = rec.values["Size"]
        common["width"] = himetric_to_points(w)
        common["height"] = himetric_to_points(h)
    morph = rec.spec is MORPH_DATA
    vpb_default = VPB_DEFAULT_MORPH if morph else (VPB_DEFAULT_LABEL if rec.spec is LABEL else VPB_DEFAULT_PLAIN)
    vpb = rec.get("VariousPropertyBits", vpb_default)
    common.update(
        enabled=bool(vpb & VPB_ENABLED),
        locked=bool(vpb & VPB_LOCKED),
        back_style_opaque=bool(vpb & VPB_BACKSTYLE),
        word_wrap=bool(vpb & VPB_WORDWRAP),
        auto_size=bool(vpb & VPB_AUTOSIZE),
        multi_line=bool(vpb & VPB_MULTILINE),
        column_heads=bool(vpb & VPB_COLUMNHEADS),
        match_required=bool(vpb & VPB_MATCHREQUIRED),
        enter_key_behavior=bool(vpb & VPB_ENTERKEY),
        alignment=0 if vpb & VPB_ALIGNMENT and kind in ("CheckBox", "OptionButton") else 1,
    )
    # The file format's defaults where they differ from the model's.
    if morph:
        common.update(back_color=0x80000005, fore_color=0x80000008, special_effect=2)
    if kind in ("ListBox", "ComboBox"):
        common["match_entry"] = 2
    if kind == "SpinButton":
        common["max"] = 100
    if kind == "Image":
        common["border_style"] = 1
    for key, attr in (("ForeColor", "fore_color"), ("BackColor", "back_color"), ("BorderColor", "border_color"),
                      ("BorderStyle", "border_style"), ("SpecialEffect", "special_effect"),
                      ("ScrollBars", "scroll_bars"), ("MaxLength", "max_length"), ("ListRows", "list_rows"),
                      ("ColumnCount", "column_count"), ("BoundColumn", "bound_column"), ("TextColumn", "text_column"),
                      ("ListStyle", "list_style"), ("MatchEntry", "match_entry"), ("MultiSelect", "multi_select"),
                      ("Min", "min"), ("Max", "max"), ("Position", "position"), ("SmallChange", "small_change"),
                      ("LargeChange", "large_change"), ("Orientation", "orientation"), ("Delay", "delay"),
                      ("PictureSizeMode", "picture_size_mode"), ("PictureAlignment", "picture_alignment"),
                      ("PicturePosition", "picture_position")):
        if rec.has(key):
            common[attr] = rec.get(key)
    if kind in ("CheckBox", "OptionButton", "ToggleButton") and rec.has("MultiSelect"):
        # TripleState is stored as MultiSelect on these.
        common["triple_state"] = rec.get("MultiSelect") == 1
        common["multi_select"] = 0
    if kind == "ComboBox":
        common["style"] = 2 if rec.get("DisplayStyle", 3) == 7 else 0
    if rec.has("Caption"):
        common["caption"] = rec.get("Caption")
    elif kind in ("CommandButton", "Label", "CheckBox", "OptionButton", "ToggleButton"):
        common["caption"] = ""
    if rec.has("Value"):
        common["value"] = rec.get("Value")
    if rec.has("GroupName"):
        common["group_name"] = rec.get("GroupName")
    if rec.has("PasswordChar"):
        pc = rec.get("PasswordChar")
        common["password_char"] = chr(pc) if pc else ""
    if rec.has("Accelerator"):
        acc = rec.get("Accelerator")
        common["accelerator"] = chr(acc) if acc else ""
    if "Picture" in rec.pictures:
        common["picture"] = rec.pictures["Picture"]
    if rec.text_props is not None:
        common["font"] = _font_of(rec.text_props)
        para = rec.text_props.get("ParagraphAlign", 1)
        common["text_align"] = {1: 1, 2: 3, 3: 2}.get(para, 1)
    if morph:
        common["column_widths"] = _column_widths(rec)
    if rec.spec is TAB_STRIP:
        n = rec.get("TabsAllocated", 0) or 0
        items = _strings(rec.get("Items", b""), max(n, 64), encoding)
        common["tabs"] = tuple(items)
        common["position"] = rec.get("ListIndex", 0)
    return ControlDesign(**common)


def _container_props(fs: FormStream) -> Dict[str, Any]:
    rec = fs.record
    out: Dict[str, Any] = {}
    if rec.has("DisplayedSize"):
        w, h = rec.get("DisplayedSize")
        out["width"], out["height"] = himetric_to_points(w), himetric_to_points(h)
    for key, attr in (("BackColor", "back_color"), ("ForeColor", "fore_color"), ("BorderColor", "border_color"),
                      ("BorderStyle", "border_style"), ("SpecialEffect", "special_effect"), ("ScrollBars", "scroll_bars")):
        if rec.has(key):
            out[attr] = rec.get(key)
    if rec.has("Caption"):
        out["caption"] = rec.get("Caption")
    if fs.font is not None:
        out["font"] = fs.font
    if fs.picture:
        out["picture"] = fs.picture
    out["enabled"] = bool(rec.get("BooleanProperties", 4) & 4)
    return out


def _streams_of(streams: Dict[str, bytes], prefix: str) -> Dict[str, bytes]:
    return {k[len(prefix):]: v for k, v in streams.items() if k.startswith(prefix)}


def _read_level(streams: Dict[str, bytes], prefix: str, encoding: str, warnings: List[str], inherited: FontDesign) -> Tuple[FormStream, Tuple[ControlDesign, ...]]:
    f = streams.get(prefix + "f")
    if f is None:
        raise FormError(f"{prefix or 'the form'}: no f stream")
    o = streams.get(prefix + "o", b"")
    fs = read_form_stream(f, encoding)
    total = sum(s.object_stream_size for s in fs.sites if s.cache_index not in CONTAINERS)
    if total != len(o):
        raise FormError(f"{prefix or 'the form'}: its controls account for {total} bytes of o, which holds {len(o)}")
    font = fs.font or inherited
    children: List[ControlDesign] = []
    off = 0
    claimed = set()
    for place, site in enumerate(fs.sites):
        if site.tab_index < 0:
            site = replace(site, tab_index=place)
        if site.cache_index in CONTAINERS:
            sub = _child_storage(streams, prefix, site.id)
            if sub is None:
                warnings.append(f"{site.name}: its storage is missing")
                children.append(ControlDesign(name=site.name, kind=CLASSES[site.cache_index][0]))
                continue
            claimed.add(sub)
            child_fs, grand = _read_level(streams, sub, encoding, warnings, font)
            props = _container_props(child_fs)
            kind = CLASSES[site.cache_index][0]
            pages = grand
            if kind == "MultiPage":
                # A MultiPage lists its TabStrip (the pages' captions) and then its pages.
                strip = next((g for g in grand if g.kind == "TabStrip"), None)
                tabs = strip.tabs if strip is not None else ()
                only_pages = [g for g in grand if g.kind == "Page"]
                # The page on top is the one whose site is visible (the others' aren't, though
                # as properties every page is visible), else the TabStrip's selected tab.
                on_top = [i for i, p in enumerate(only_pages) if p.visible]
                if len(on_top) == 1:
                    selected = on_top[0]
                elif strip is not None and 0 <= strip.position < len(only_pages):
                    selected = strip.position
                else:
                    selected = 0
                pages = tuple(replace(p, caption=tabs[i] if i < len(tabs) else p.name, visible=True, tab_index=i) for i, p in enumerate(only_pages))
                props["position"] = selected
                if strip is not None:
                    props.setdefault("font", strip.font)
            design = _control(site, None, pages, encoding, warnings, props.get("font", font))
            extra = {k: v for k, v in props.items() if k in ("width", "height", "back_color", "fore_color", "border_color", "border_style", "special_effect", "scroll_bars", "caption", "font", "picture", "enabled", "position")}
            design = replace(design, **extra)
            children.append(design)
            continue
        size = site.object_stream_size
        chunk = o[off:off + size]
        off += size
        _, spec = CLASSES.get(site.cache_index, ("", None))
        rec = None
        if spec is not None and chunk:
            try:
                rec, _ = read_record(chunk, spec, encoding, 0, len(chunk))
            except FormError as e:
                warnings.append(f"{site.name}: {e}")
        children.append(_control(site, rec, (), encoding, warnings, font))
    for name in _child_storage_names(streams, prefix):
        if name not in claimed:
            warnings.append(f"{name}: a storage no control claims")
    return fs, tuple(children)


def _child_storage_names(streams: Dict[str, bytes], prefix: str) -> List[str]:
    names = set()
    for k in streams:
        if k.startswith(prefix):
            rest = k[len(prefix):]
            if "/" in rest:
                names.add(prefix + rest.split("/", 1)[0] + "/")
    return sorted(names)


def _child_storage(streams: Dict[str, bytes], prefix: str, site_id: int) -> Optional[str]:
    for name in _child_storage_names(streams, prefix):
        part = name[len(prefix):-1]
        if part[:1].lower() == "i" and part[1:].isdigit() and int(part[1:]) == site_id:
            return name
    return None


def parse_vbframe(text: str) -> Dict[str, str]:
    """The ``\\x03VBFrame`` header's ``Name = value`` lines (comments dropped, quotes kept)."""
    out: Dict[str, str] = {}
    for line in text.replace("\r\n", "\n").split("\n"):
        s = line.strip()
        if "=" not in s or s.upper().startswith(("VERSION", "BEGIN")):
            continue
        k, v = s.split("=", 1)
        v = v.strip()
        if not v.startswith('"') and "'" in v:
            v = v.split("'", 1)[0].strip()
        out[k.strip()] = v
    return out


def _unquote(v: str) -> str:
    if len(v) >= 2 and v[0] == '"' and v[-1] == '"':
        return v[1:-1].replace('""', '"')
    return v


def design_streams(items: Sequence[Tuple[str, bytes]]) -> Dict[str, bytes]:
    """The streams of a module's designer items (storages, which carry CLSIDs, left out)."""
    return {path: data for path, data in items if not path.endswith("/")}


def read_design(name: str, streams: Dict[str, bytes], encoding: str = "cp1252") -> FormDesign:
    """A form's design from its designer storage (``{relative path: bytes}``, paths like ``f``,
    ``o``, ``\\x03VBFrame``, ``i06/f``). Raises :class:`FormError` when the form itself can't be
    read; problems with single controls become ``warnings``."""
    warnings: List[str] = []
    header = parse_vbframe(streams.get(VBFRAME, b"").decode(encoding, "replace"))
    fs, controls = _read_level(streams, "", encoding, warnings, FontDesign())
    props = _container_props(fs)
    caption = _unquote(header["Caption"]) if "Caption" in header else props.get("caption", name)
    width, height = props.get("width", 240.0), props.get("height", 180.0)
    try:
        if "ClientWidth" in header:
            width = int(header["ClientWidth"]) / 20.0
        if "ClientHeight" in header:
            height = int(header["ClientHeight"]) / 20.0
    except ValueError:
        pass

    def num(key: str, default: int) -> int:
        try:
            return int(header.get(key, default))
        except ValueError:
            return default

    show_modal = header.get("ShowModal", "-1").strip().lower() not in ("0", "false")
    return FormDesign(
        name=name,
        caption=caption,
        width=width,
        height=height,
        back_color=props.get("back_color", 0x8000000F),
        fore_color=props.get("fore_color", 0x80000012),
        border_color=props.get("border_color", 0x80000012),
        border_style=props.get("border_style", 0),
        special_effect=props.get("special_effect", 0),
        scroll_bars=props.get("scroll_bars", 0),
        font=props.get("font", FontDesign()),
        start_up_position=num("StartUpPosition", 1),
        left=num("ClientLeft", 120) / 20.0,
        top=num("ClientTop", 465) / 20.0,
        show_modal=show_modal,
        enabled=props.get("enabled", True),
        picture=props.get("picture", b""),
        controls=controls,
        warnings=tuple(warnings),
    )


# ---------------------------------------------------------------- .frm / .frx


_CFB_SIGNATURE = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"


def _frx_root(frx: bytes):
    from .cfb import CfbError, read_cfb

    at = frx.find(_CFB_SIGNATURE)
    if at < 0:
        raise FormError("the .frx file has no form in it")
    try:
        return read_cfb(frx[at:])
    except CfbError as e:
        raise FormError(f"the .frx file is damaged ({e})")


def frx_streams(frx: bytes) -> Dict[str, bytes]:
    """The designer storage inside an exported form's ``.frx`` (a short header, then a compound
    file holding the same streams as the workbook's storage)."""
    return dict(_frx_root(frx).walk())


def frx_items(frx: bytes) -> Tuple[Tuple[str, bytes], ...]:
    """The designer storage inside a ``.frx`` as designer items (storages with their CLSIDs)."""
    from .ovba import designer_items

    return designer_items(_frx_root(frx))


def write_frx(items: Sequence[Tuple[str, bytes]], design: FormDesign) -> bytes:
    """A ``.frx`` file for a form's designer items (the VBFrame header goes into the ``.frm``):
    "LB", its version, the compound file's length, then the form's outer size in twips."""
    from .cfb import Storage, write_cfb
    from .ovba import put_designer

    st = Storage()
    put_designer(st, [(path, data) for path, data in items if path != VBFRAME])
    cfb = write_cfb(st)
    left, top = int(round(design.left * 20)), int(round(design.top * 20))
    width = int(round(design.width * 20)) + 2 * left
    height = int(round(design.height * 20)) + top + left
    return b"LB" + struct.pack("<HI", 8, len(cfb)) + b"\0" * 8 + struct.pack("<II", width, height) + cfb


# ---------------------------------------------------------------- writing
#
# The inverse of the reader, laying records out as Excel does: what Excel's designer stores for
# each kind of control (a TextBox's VariousPropertyBits, a button's centred TextProps, the
# MultiPage's TabStrip, "x" stream and DesignExtender…) was checked against forms Excel wrote.


class _Writer:
    def __init__(self) -> None:
        self.b = bytearray()

    def int(self, v: int, size: int, signed: bool = False) -> None:
        if signed:
            self.b += struct.pack("<" + {1: "b", 2: "h", 4: "i"}[size], int(v))
        else:
            self.b += struct.pack("<" + {1: "B", 2: "H", 4: "I"}[size], int(v) & ((1 << (8 * size)) - 1))

    def align(self, base: int, size: int) -> None:
        while (len(self.b) - base) % size:
            self.b.append(0)

    def raw(self, data: bytes) -> None:
        self.b += data


def _encode(text: str) -> Tuple[bytes, int]:
    """A string as stored, with its CountOfBytesWithCompressionFlag."""
    if all(ord(ch) < 0x100 for ch in text):
        raw = text.encode("latin-1")
        return raw, len(raw) | 0x80000000
    raw = text.encode("utf-16-le")
    return raw, len(raw)


def write_record(rec: Record) -> bytes:
    """The bytes :func:`read_record` reads as ``rec`` (its mask says what is stored), with its
    pictures, TextProps and tail after it."""
    spec, mask = rec.spec, rec.mask
    w = _Writer()
    w.int(0, 1)
    w.int(spec.major, 1)
    w.int(0, 2)
    w.int(mask & 0xFFFFFFFF, 4)
    if spec.mask64:
        w.int(mask >> 32, 4)
    strings: Dict[str, bytes] = {}
    for f in spec.data:
        if not mask & (1 << f.bit):
            continue
        w.align(0, f.size)
        if f.kind == "str":
            raw, packed = _encode(rec.values.get(f.name, ""))
            strings[f.name] = raw
            w.int(packed, 4)
        elif f.kind == "pic":
            w.int(0xFFFF, 2)
        else:
            w.int(rec.values.get(f.name, 0), f.size, f.kind == "i")
    w.align(0, 4)
    for x in spec.extra:
        if not mask & (1 << x.bit):
            continue
        if x.kind in ("size", "pos"):
            a, b = rec.values.get(x.name, (0, 0))
            w.int(a, 4, True)
            w.int(b, 4, True)
        elif x.kind == "str":
            w.raw(strings.get(x.name, b""))
            w.align(0, 4)
        else:
            w.raw(rec.values.get(x.name, b""))
    if len(w.b) - 4 > 0xFFFF:
        raise FormError(f"{spec.name}: too large to store")
    struct.pack_into("<H", w.b, 2, len(w.b) - 4)
    for bit, name in spec.stream:
        if mask & (1 << bit):
            _write_picture(w, rec.pictures.get(name, b""))
    if spec.text_props:
        w.raw(write_record(rec.text_props if rec.text_props is not None else Record(TEXT_PROPS, 0)))
    w.raw(rec.tail)
    return bytes(w.b)


def _write_picture(w: _Writer, data: bytes) -> None:
    w.raw(CLSID_STDPICTURE)
    w.int(0x0000746C, 4)
    w.int(len(data), 4)
    w.raw(data)


def _write_std_font(w: _Writer, f: FontDesign, encoding: str) -> None:
    w.raw(CLSID_STDFONT)
    w.int(1, 1)
    w.int(0, 2)
    w.int((2 if f.italic else 0) | (4 if f.underline else 0) | (8 if f.strike else 0), 1)
    w.int(700 if f.bold else 400, 2)
    w.int(int(round(f.size * 10000)), 4)
    face = f.name.encode(encoding, "replace")[:255]
    w.int(len(face), 1)
    w.raw(face)


class _RecordBuilder:
    """A record's values, setting each one's mask bit."""

    def __init__(self, spec: Spec, mask: int = 0) -> None:
        self.rec = Record(spec, mask)

    def put(self, name: str, value: Any, default: Any = object()) -> None:
        """Stores ``value`` under ``name`` unless it is the file format's default."""
        if value == default:
            return
        for f in self.rec.spec.data:
            if f.name == name:
                self.rec.mask |= 1 << f.bit
                if f.kind == "pic":
                    self.rec.pictures[name] = value
                else:
                    self.rec.values[name] = value
                break
        for x in self.rec.spec.extra:
            if x.name == name:
                self.rec.mask |= 1 << x.bit
                self.rec.values[name] = value
                break

    def flag(self, bit: int) -> None:
        self.rec.mask |= 1 << bit


def _text_props(font: FontDesign, align: int = 1) -> Record:
    """A control's TextProps, as Excel writes them (name, height, charset, pitch and family)."""
    b = _RecordBuilder(TEXT_PROPS)
    b.put("FontName", font.name)
    effects = (FONT_EFFECT_BOLD if font.bold else 0) | (FONT_EFFECT_ITALIC if font.italic else 0) | (FONT_EFFECT_UNDERLINE if font.underline else 0) | (FONT_EFFECT_STRIKE if font.strike else 0)
    b.put("FontEffects", effects, 0)
    b.put("FontHeight", int(round(font.size * 20)))
    b.put("FontCharSet", 0)
    b.put("FontPitchAndFamily", 2)
    # fmTextAlign 1 left, 2 centre, 3 right; stored 1 left, 2 right, 3 centre.
    b.put("ParagraphAlign", {1: 1, 2: 3, 3: 2}.get(align, 1), 1)
    if font.bold:
        b.put("FontWeight", 700)
    return b.rec


def _vpb(d: ControlDesign, base: int) -> int:
    v = base
    for on, bit in ((d.enabled, VPB_ENABLED), (d.locked, VPB_LOCKED), (d.back_style_opaque, VPB_BACKSTYLE),
                    (d.word_wrap, VPB_WORDWRAP), (d.auto_size, VPB_AUTOSIZE)):
        v = (v | bit) if on else (v & ~bit)
    return v


_MORPH_DISPLAY = {"TextBox": 1, "ListBox": 2, "ComboBox": 3, "CheckBox": 4, "OptionButton": 5, "ToggleButton": 6}


def control_record(d: ControlDesign) -> Record:
    """The record of a control that lives in its form's ``o`` stream."""
    k = d.kind
    size = (points_to_himetric(d.width), points_to_himetric(d.height))
    if k in _MORPH_DISPLAY:
        b = _RecordBuilder(MORPH_DATA, 1 << 31)
        vpb = _vpb(d, VPB_DEFAULT_MORPH)
        for on, bit in ((d.column_heads, VPB_COLUMNHEADS), (d.match_required, VPB_MATCHREQUIRED), (d.multi_line, VPB_MULTILINE),
                        (d.enter_key_behavior, VPB_ENTERKEY), (k in ("CheckBox", "OptionButton") and d.alignment == 0, VPB_ALIGNMENT),
                        (k == "TextBox" or (k == "ComboBox" and d.style != 2), VPB_EDITABLE)):
            vpb = (vpb | bit) if on else (vpb & ~bit)
        if k not in ("TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton"):
            vpb &= ~VPB_EDITABLE
        b.put("VariousPropertyBits", vpb, VPB_DEFAULT_MORPH)
        b.put("BackColor", d.back_color, 0x80000005)
        b.put("ForeColor", d.fore_color, 0x80000008)
        b.put("MaxLength", d.max_length, 0)
        b.put("BorderStyle", d.border_style, 0)
        b.put("ScrollBars", d.scroll_bars, 0)
        display = _MORPH_DISPLAY[k] if not (k == "ComboBox" and d.style == 2) else 7
        b.put("DisplayStyle", display, 1)
        b.put("Size", size)
        b.put("PasswordChar", ord(d.password_char[0]) if d.password_char else 0, 0)
        if k in ("ListBox", "ComboBox"):
            b.put("BoundColumn", d.bound_column, 1)
            b.put("TextColumn", d.text_column, -1)
            b.put("ColumnCount", d.column_count, 1)
            b.put("ListRows", d.list_rows, 8)
            b.put("MatchEntry", d.match_entry, 2)
            b.put("ListStyle", d.list_style, 0)
        if k == "ComboBox":
            b.put("ShowDropButtonWhen", 2)
        if k == "ListBox":
            b.put("MultiSelect", d.multi_select, 0)
        if k in ("CheckBox", "OptionButton", "ToggleButton"):
            b.put("MultiSelect", 1 if d.triple_state else 0, 0)
            b.put("Value", d.value if d.value in ("0", "1") else ("" if d.triple_state else "0"), "")
            b.put("Caption", d.caption, "")
        else:
            b.put("Value", d.value, "")
        b.put("BorderColor", d.border_color, 0x80000006)
        b.put("SpecialEffect", d.special_effect, 2)
        if d.picture:
            b.put("Picture", d.picture)
        b.put("Accelerator", ord(d.accelerator[0]) if d.accelerator else 0, 0)
        b.put("GroupName", d.group_name, "")
        b.rec.text_props = _text_props(d.font, d.text_align)
        return b.rec
    if k == "CommandButton":
        b = _RecordBuilder(COMMAND_BUTTON)
        b.put("ForeColor", d.fore_color, 0x80000012)
        b.put("BackColor", d.back_color, 0x8000000F)
        b.put("VariousPropertyBits", _vpb(d, VPB_DEFAULT_PLAIN), VPB_DEFAULT_PLAIN)
        b.put("Caption", d.caption, "")
        b.put("Size", size)
        if d.picture:
            b.put("Picture", d.picture)
            b.put("PicturePosition", d.picture_position, 0x00070001)
        b.put("Accelerator", ord(d.accelerator[0]) if d.accelerator else 0, 0)
        b.rec.text_props = _text_props(d.font, 2)  # a button's caption is centred
        return b.rec
    if k == "Label":
        b = _RecordBuilder(LABEL)
        b.put("ForeColor", d.fore_color, 0x80000012)
        b.put("BackColor", d.back_color, 0x8000000F)
        b.put("VariousPropertyBits", _vpb(d, VPB_DEFAULT_LABEL), VPB_DEFAULT_LABEL)
        b.put("Caption", d.caption, "")
        b.put("Size", size)
        b.put("BorderColor", d.border_color, 0x80000006)
        b.put("BorderStyle", d.border_style, 0)
        b.put("SpecialEffect", d.special_effect, 0)
        if d.picture:
            b.put("Picture", d.picture)
            b.put("PicturePosition", d.picture_position, 0x00070001)
        b.put("Accelerator", ord(d.accelerator[0]) if d.accelerator else 0, 0)
        b.rec.text_props = _text_props(d.font, d.text_align)
        return b.rec
    if k == "Image":
        b = _RecordBuilder(IMAGE)
        b.put("BorderColor", d.border_color, 0x80000006)
        b.put("BackColor", d.back_color, 0x8000000F)
        b.put("BorderStyle", d.border_style, 1)
        b.put("PictureSizeMode", d.picture_size_mode, 0)
        b.put("SpecialEffect", d.special_effect, 0)
        b.put("Size", size)
        if d.picture:
            b.put("Picture", d.picture)
        b.put("PictureAlignment", d.picture_alignment, 2)
        b.put("VariousPropertyBits", _vpb(replace(d, word_wrap=False, auto_size=False), VPB_DEFAULT_PLAIN), VPB_DEFAULT_PLAIN)
        if d.auto_size:
            b.flag(2)
        return b.rec
    if k in ("ScrollBar", "SpinButton"):
        spin = k == "SpinButton"
        b = _RecordBuilder(SPIN_BUTTON if spin else SCROLL_BAR)
        b.put("ForeColor", d.fore_color, 0x80000012)
        b.put("BackColor", d.back_color, 0x8000000F)
        b.put("VariousPropertyBits", _vpb(replace(d, word_wrap=False, auto_size=False), VPB_DEFAULT_PLAIN), VPB_DEFAULT_PLAIN)
        b.put("Size", size)
        b.put("Min", d.min, 0)
        b.put("Max", d.max, 100 if spin else 32767)
        b.put("Position", d.position, 0)
        b.put("SmallChange", d.small_change, 1)
        if not spin:
            b.put("LargeChange", d.large_change, 1)
        b.put("Orientation", d.orientation)
        b.put("Delay", d.delay, 50)
        return b.rec
    if k == "TabStrip":
        return _tab_strip_record(list(d.tabs), [""] * len(d.tabs), d.position, size, d.font)
    raise FormError(f"{d.name}: a {k} can't be written")


def _string_array(items: Sequence[str]) -> bytes:
    w = _Writer()
    for t in items:
        raw, packed = _encode(t)
        w.int(packed, 4)
        w.raw(raw)
        w.align(0, 4)
    return bytes(w.b)


def _tab_strip_record(captions: List[str], names: List[str], selected: int, size: Tuple[int, int], font: FontDesign) -> Record:
    """A TabStrip (also the tabs of a MultiPage): captions, names, empty tips, tags and
    accelerators, and each tab's flags (enabled, visible) after its TextProps."""
    n = len(captions)
    b = _RecordBuilder(TAB_STRIP)
    b.put("ListIndex", selected, -1)
    b.put("Size", size)
    items, tips, tab_names, tags, accels = _string_array(captions), _string_array([""] * n), _string_array(names), _string_array([""] * n), b"\x00\x00\x00\x00" * n
    b.put("ItemsSize", len(items), 0)
    b.put("Items", items)
    b.put("TipStringsSize", len(tips), 0)
    b.put("TipStrings", tips)
    b.put("NamesSize", len(tab_names), 0)
    b.put("TabNames", tab_names)
    b.flag(19)
    b.put("TabsAllocated", n, 0)
    b.put("TagsSize", len(tags), 0)
    b.put("Tags", tags)
    b.put("TabData", n, 0)
    b.put("AcceleratorsSize", len(accels), 0)
    b.put("Accelerators", accels)
    b.rec.text_props = _text_props(font)
    b.rec.tail = struct.pack("<" + "I" * n, *([3] * n))
    return b.rec


def write_site(s: Site) -> bytes:
    """An OleSiteConcreteControl: what is not the file format's default."""
    vals: Dict[str, int] = {}
    raws: Dict[str, bytes] = {}
    mask = 0

    def string(bit: int, field_name: str, name: str, text: str) -> None:
        nonlocal mask
        if text:
            raw, packed = _encode(text)
            mask |= 1 << bit
            vals[field_name] = packed
            raws[name] = raw

    string(0, "NameData", "Name", s.name)
    string(1, "TagData", "Tag", s.tag)
    for bit, name, value, default in ((2, "ID", s.id, 0), (3, "HelpContextID", s.help_context, 0), (4, "BitFlags", s.bit_flags, SITE_FLAG_DEFAULT),
                                      (5, "ObjectStreamSize", s.object_stream_size, 0), (6, "TabIndex", s.tab_index, -1),
                                      (7, "ClsidCacheIndex", s.cache_index, 0x7FFF), (9, "GroupID", s.group_id, 0)):
        if value != default:
            mask |= 1 << bit
            vals[name] = value
    if s.position != (0, 0):
        mask |= 1 << 8
    string(11, "ControlTipTextData", "ControlTipText", s.tip)
    string(12, "RuntimeLicKeyData", "RuntimeLicKey", s.lic_key)
    string(13, "ControlSourceData", "ControlSource", s.control_source)
    string(14, "RowSourceData", "RowSource", s.row_source)
    w = _Writer()
    w.int(0, 2)
    w.int(0, 2)
    w.int(mask, 4)
    for bit, name, size in SITE_FIELDS:
        if mask & (1 << bit):
            w.align(0, size)
            w.int(vals[name], size, name in SITE_SIGNED)
    w.align(0, 4)
    for _, name in SITE_STRINGS_BEFORE:
        if name in raws:
            w.raw(raws[name])
            w.align(0, 4)
    if mask & (1 << 8):
        w.int(s.position[0], 4, True)
        w.int(s.position[1], 4, True)
    for _, name in SITE_STRINGS_AFTER:
        if name in raws:
            w.raw(raws[name])
            w.align(0, 4)
    struct.pack_into("<H", w.b, 2, len(w.b) - 4)
    return bytes(w.b)


def write_form_stream(rec: Record, font: Optional[FontDesign], picture: bytes, sites: Sequence[Site], encoding: str = "cp1252", design_ex: bytes = b"") -> bytes:
    """A container's ``f`` stream: its FormControl, font and picture, and its sites."""
    w = _Writer()
    w.raw(write_record(rec))
    if rec.mask & (1 << 20):
        _write_std_font(w, font or FontDesign(), encoding)
    if rec.mask & (1 << 21):
        _write_picture(w, picture)
    if not rec.get("BooleanProperties", 4) & FORM_FLAG_NOCLASSTABLE:
        w.int(0, 2)  # no classes of ActiveX controls
    w.raw(sites_block([write_site(site) for site in sites]))
    w.raw(design_ex)
    return bytes(w.b)


def sites_block(sites: Sequence[bytes]) -> bytes:
    """CountOfSites, CountOfBytes, the sites' depths and types, and the sites (written)."""
    body = _Writer()
    left = len(sites)
    while left:
        run = min(left, 127)
        body.int(0, 1)  # depth
        if run == 1:
            body.int(1, 1)  # one OleSiteConcreteControl
        else:
            body.int(0x80 | run, 1)
            body.int(1, 1)
        left -= run
    body.align(0, 4)
    for site in sites:
        body.raw(site)
    w = _Writer()
    w.int(len(sites), 4)
    w.int(len(body.b), 4)
    w.raw(body.b)
    return bytes(w.b)


CLSID_FRAME = "{6E182020-F460-11CE-9BCD-00AA00608E01}"
CLSID_MULTIPAGE = "{46E31370-3F7A-11CE-BED6-00AA00611080}"
_COMPOBJ_NAMES = {
    CLSID_FORM: ("Microsoft Forms 2.0 Form", "Forms.Form.1"),
    CLSID_FRAME: ("Microsoft Forms 2.0 Frame", "Forms.Frame.1"),
    CLSID_MULTIPAGE: ("Microsoft Forms 2.0 Form", "Forms.MultiPage.1"),
}
# What follows a MultiPage's and a Page's sites ([MS-OFORMS] DesignExtender), as Excel writes it.
_DESIGN_EX_MULTIPAGE = bytes.fromhex("00020c0019000000f08f0000ff010000")
_DESIGN_EX_PAGE = bytes.fromhex("00020c0019000000f3ff0100ff010000")
_BOOL_FRAME = 0x8004
_BOOL_PAGES = 0xC004
# A container's site: not streamed (it has a storage), auto-sized, its controls its own.
_SITE_CONTAINER = 0x40020


def _container_flags(tab_stop: bool, visible: bool) -> int:
    return _SITE_CONTAINER | (SITE_TABSTOP if tab_stop else 0) | (SITE_VISIBLE if visible else 0)


def compobj(clsid: str) -> bytes:
    """A storage's ``\\x01CompObj`` stream ([MS-OLEDS] CompObjStream) for a Forms class."""
    user, progid = _COMPOBJ_NAMES[clsid]
    w = _Writer()
    w.raw(bytes.fromhex("0100feff030a0000ffffffff"))
    w.raw(uuid.UUID(clsid).bytes_le)
    for text in (user, "Embedded Object", progid):
        raw = text.encode("ascii") + b"\x00"
        w.int(len(raw), 4)
        w.raw(raw)
    w.int(0x71B239F4, 4)
    w.raw(b"\x00" * 12)
    return bytes(w.b)


def _form_record(width: float, height: float, caption: Optional[str], back: int, fore: int, bool_props: Optional[int],
                 next_id: int, cookie: int, special_effect: int = 0, font: bool = False, picture: bool = False) -> Record:
    b = _RecordBuilder(FORM)
    b.put("BackColor", back, 0x8000000F)
    b.put("ForeColor", fore, 0x80000012)
    b.put("NextAvailableID", next_id, 0)
    if bool_props is not None:
        b.put("BooleanProperties", bool_props, 4)
    b.put("SpecialEffect", special_effect, 0)
    if caption:
        b.put("Caption", caption)
    if font:
        b.flag(20)
    if picture:
        b.flag(21)
    b.put("ShapeCookie", cookie, 0)
    b.put("DrawBuffer", 32000)
    b.put("DisplayedSize", (points_to_himetric(width), points_to_himetric(height)))
    b.put("LogicalSize", (0, 0))
    return b.rec


# A page is this much smaller than its MultiPage (HIMETRIC: borders and tabs), as Excel lays it out.
PAGE_INSET = (106, 609)


def design_site(c: ControlDesign, cid: int, place: int, object_stream_size: int = 0) -> Site:
    """The site a control of the design has in its container's ``f`` stream."""
    common = dict(name=c.name, tag=c.tag, id=cid, tab_index=c.tab_index if c.tab_index >= 0 else place,
                  position=(points_to_himetric(c.left), points_to_himetric(c.top)), tip=c.tip,
                  control_source=c.control_source, row_source=c.row_source)
    if c.kind in ("Frame", "MultiPage"):
        return Site(**common, bit_flags=_container_flags(c.tab_stop, c.visible), cache_index=CACHE_INDEX[c.kind])
    flags = (SITE_TABSTOP if c.tab_stop else 0) | (SITE_VISIBLE if c.visible else 0) | (SITE_DEFAULT if c.default else 0) | (SITE_CANCEL if c.cancel else 0)
    return Site(**common, bit_flags=flags | SITE_STREAMED | SITE_AUTOSIZE, object_stream_size=object_stream_size, cache_index=CACHE_INDEX[c.kind])


def frame_record(c: ControlDesign, next_id: int = 0, nsites: int = 0) -> Record:
    """A Frame's own FormControl (the ``f`` stream of its storage)."""
    return _form_record(c.width, c.height, c.caption or None, c.back_color, c.fore_color, _BOOL_FRAME, next_id, nsites, special_effect=c.special_effect, font=True)


def multipage_record(c: ControlDesign, next_id: int = 0, nsites: int = 0) -> Record:
    return _form_record(c.width, c.height, None, c.back_color, c.fore_color, _BOOL_PAGES, next_id, nsites)


def page_record(pg: ControlDesign, mp: ControlDesign, next_id: int = 0, nsites: int = 0) -> Record:
    """A page's FormControl: the MultiPage's size less its tabs and borders."""
    w = max(0, points_to_himetric(mp.width) - PAGE_INSET[0])
    h = max(0, points_to_himetric(mp.height) - PAGE_INSET[1])
    return _form_record(himetric_to_points(w), himetric_to_points(h), None, pg.back_color, pg.fore_color, _BOOL_PAGES, next_id, nsites)


def strip_record(c: ControlDesign) -> Record:
    """The tabs of a MultiPage: an unnamed TabStrip as large as the MultiPage."""
    pages = [p for p in c.children if p.kind == "Page"]
    selected = c.position if 0 <= c.position < len(pages) else 0
    return _tab_strip_record([p.caption for p in pages], [p.name for p in pages], selected, (points_to_himetric(c.width), points_to_himetric(c.height)), c.font)


class DesignWriter:
    """Writes controls of a design, with IDs from ``ids``; storages and streams of containers go
    to ``out``."""

    def __init__(self, encoding: str, ids) -> None:
        self.encoding = encoding
        self.ids = ids
        self.out: List[Tuple[str, bytes]] = []

    def level(self, prefix: str, controls: Sequence[ControlDesign]) -> Tuple[List[Site], bytes, int]:
        """The sites and ``o`` stream of a container's controls, and the largest ID beneath."""
        sites: List[Site] = []
        o = bytearray()
        top = 0
        for i, c in enumerate(controls):
            site, data, t = self.control(prefix, c, i)
            sites.append(site)
            o += data
            top = max(top, t)
        return sites, bytes(o), top

    def control(self, prefix: str, c: ControlDesign, place: int) -> Tuple[Site, bytes, int]:
        """One control: its site, its record (for the ``o`` stream) and the largest ID it took."""
        cid = next(self.ids)
        if c.kind in ("Frame", "MultiPage"):
            top = self.container(f"{prefix}i{cid:02d}/", c)
            return design_site(c, cid, place), b"", max(cid, top)
        data = write_record(control_record(c))
        return design_site(c, cid, place, len(data)), data, cid

    def container(self, sub: str, c: ControlDesign) -> int:
        out, encoding = self.out, self.encoding
        if c.kind == "Frame":
            out.append((sub, uuid.UUID(CLSID_FRAME).bytes_le))
            sites, o, top = self.level(sub, c.children)
            out.append((sub + COMPOBJ, compobj(CLSID_FRAME)))
            out.append((sub + "f", write_form_stream(frame_record(c, top, len(sites)), c.font, b"", sites, encoding)))
            out.append((sub + "o", o))
            return top
        # A MultiPage: its tabs (an unnamed TabStrip), then one storage per page.
        out.append((sub, uuid.UUID(CLSID_MULTIPAGE).bytes_le))
        pages = [p for p in c.children if p.kind == "Page"]
        strip_id = next(self.ids)
        top = strip_id
        selected = c.position if 0 <= c.position < len(pages) else 0
        strip_data = write_record(strip_record(c))
        sites = [Site(id=strip_id, bit_flags=SITE_FLAG_DEFAULT, object_stream_size=len(strip_data), tab_index=0, cache_index=CACHE_INDEX["TabStrip"])]
        page_ids: List[int] = []
        for i, pg in enumerate(pages):
            pid = next(self.ids)
            top = max(top, pid)
            page_ids.append(pid)
            psub = f"{sub}i{pid:02d}/"
            out.append((psub, uuid.UUID(CLSID_FORM).bytes_le))
            psites, po, ptop = self.level(psub, pg.children)
            top = max(top, ptop)
            out.append((psub + COMPOBJ, compobj(CLSID_FORM)))
            out.append((psub + "f", write_form_stream(page_record(pg, c, max(ptop, pid), len(psites)), None, b"", psites, encoding, _DESIGN_EX_PAGE)))
            out.append((psub + "o", po))
            # Only the page on top has its site visible; pages sit below the tabs (where Excel
            # puts them, unless the design says).
            at = (points_to_himetric(pg.left), points_to_himetric(pg.top)) if (pg.left or pg.top) else (53, 556)
            sites.append(Site(name=pg.name, tag=pg.tag, id=pid, bit_flags=_container_flags(True, i == selected), tab_index=i + 1, cache_index=CACHE_INDEX["Page"], position=at, tip=pg.tip))
        out.append((sub + COMPOBJ, compobj(CLSID_MULTIPAGE)))
        out.append((sub + "f", write_form_stream(multipage_record(c, top, len(sites)), None, b"", sites, encoding, _DESIGN_EX_MULTIPAGE)))
        out.append((sub + "o", strip_data))
        x = _Writer()
        for _ in range(len(pages) + 1):
            x.raw(bytes.fromhex("0002040000000000"))
        x.int(0, 1)
        x.int(2, 1)
        x.int(12, 2)
        x.int(0x6, 4)  # PageCount, ID
        x.int(len(pages), 4)
        x.int(strip_id, 4)
        for pid in page_ids:
            x.int(pid, 4)
        out.append((sub + "x", bytes(x.b)))
        return top


def form_record(design: FormDesign, next_id: int = 0, nsites: int = 0) -> Record:
    """The form's own FormControl (its ``f`` stream)."""
    return _form_record(design.width, design.height, design.caption or None, design.back_color, design.fore_color, None if design.enabled else 0, next_id, nsites,
                        special_effect=design.special_effect, font=design.font != FontDesign(), picture=bool(design.picture))


def write_design(design: FormDesign, encoding: str = "cp1252") -> List[Tuple[str, bytes]]:
    """The designer storage of ``design``, as :attr:`gridcore.model.vba.VbaModule.designer`
    keeps it (``("/", CLSID)`` and ``("i06/", CLSID)`` entries for the storages)."""
    w = DesignWriter(encoding, iter(range(1, 1 << 30)))
    w.out.append(("/", uuid.UUID(CLSID_FORM).bytes_le))
    sites, o, top = w.level("", design.controls)
    w.out.append((COMPOBJ, compobj(CLSID_FORM)))
    w.out.append(("f", write_form_stream(form_record(design, top, len(sites)), design.font, design.picture, sites, encoding)))
    w.out.append(("o", o))
    w.out.append((VBFRAME, vbframe(design).encode(encoding, "replace")))
    return sorted(w.out, key=lambda kv: kv[0])


def vbframe(design: FormDesign) -> str:
    """The ``\\x03VBFrame`` header: the form's caption, client size (twips) and position."""
    caption = design.caption.replace('"', '""')
    lines = [
        "VERSION 5.00",
        f"Begin {CLSID_FORM} {design.name} ",
        f'   Caption         =   "{caption}"',
        f"   ClientHeight    =   {int(round(design.height * 20))}",
        f"   ClientLeft      =   {int(round(design.left * 20))}",
        f"   ClientTop       =   {int(round(design.top * 20))}",
        f"   ClientWidth     =   {int(round(design.width * 20))}",
    ]
    if not design.show_modal:
        lines.append("   ShowModal       =   0   'False")
    lines.append(f"   StartUpPosition =   {design.start_up_position}  '{_STARTUP.get(design.start_up_position, 'Manual')}")
    lines += ["End", ""]
    return "\r\n".join(lines)


_STARTUP = {0: "Manual", 1: "CenterOwner", 2: "CenterScreen", 3: "Windows Default"}
