"""A UserForm's design: the form and the controls on it, as the VBA editor's form designer lays
them out. Read from a workbook's VBA project (:mod:`gridcore.xlsx.oforms`, [MS-OFORMS]); the
running form (:mod:`gridcore.vba.msforms`) starts from it.

Sizes and positions are in points, colours are OLE_COLORs (``0x00BBGGRR``, or ``0x800000nn`` for a
system colour), as VBA reports them.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional, Tuple

# OLE_COLORs of the system colours MSForms uses by default.
SYS_BUTTON_FACE = 0x8000000F
SYS_BUTTON_TEXT = 0x80000012
SYS_WINDOW = 0x80000005
SYS_WINDOW_TEXT = 0x80000008
SYS_WINDOW_FRAME = 0x80000006

# The kinds of control the designer offers (TypeName of the running control).
KINDS = (
    "Label", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton", "Frame",
    "CommandButton", "TabStrip", "MultiPage", "Page", "ScrollBar", "SpinButton", "Image",
)


@dataclass(frozen=True)
class FontDesign:
    name: str = "Tahoma"
    size: float = 8.25
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False


@dataclass(frozen=True)
class ControlDesign:
    """One control. ``kind`` is a :data:`KINDS` name (or the class name of a control ReflexGrid
    can't show). Containers (Frame, MultiPage and its Pages) hold ``children``."""

    name: str
    kind: str
    left: float = 0.0
    top: float = 0.0
    width: float = 72.0
    height: float = 18.0
    id: int = 0
    tab_index: int = 0
    tab_stop: bool = True
    visible: bool = True
    default: bool = False
    cancel: bool = False
    tag: str = ""
    tip: str = ""
    control_source: str = ""
    row_source: str = ""
    group_name: str = ""
    caption: str = ""
    # The stored Value: text for a TextBox or ComboBox, "1"/"0"/"" for a CheckBox, an
    # OptionButton or a ToggleButton (Null when neither).
    value: str = ""
    enabled: bool = True
    locked: bool = False
    fore_color: int = SYS_BUTTON_TEXT
    back_color: int = SYS_BUTTON_FACE
    back_style_opaque: bool = True
    border_style: int = 0
    border_color: int = SYS_WINDOW_FRAME
    special_effect: int = 0
    word_wrap: bool = False
    auto_size: bool = False
    multi_line: bool = False
    password_char: str = ""
    max_length: int = 0
    # fmTextAlign: 1 left, 2 center, 3 right.
    text_align: int = 1
    font: FontDesign = field(default_factory=FontDesign)
    scroll_bars: int = 0
    # Lists (ListBox, ComboBox).
    list_rows: int = 8
    column_count: int = 1
    column_widths: str = ""
    bound_column: int = 1
    text_column: int = -1
    list_style: int = 0
    match_entry: int = 1
    multi_select: int = 0
    # ComboBox: 0 fmStyleDropDownCombo, 2 fmStyleDropDownList.
    style: int = 0
    column_heads: bool = False
    # Numbers (ScrollBar, SpinButton).
    min: int = 0
    max: int = 32767
    position: int = 0
    small_change: int = 1
    large_change: int = 1
    orientation: int = -1
    delay: int = 50
    # Pictures (Image, CommandButton, Label, Frame…): the stored image data.
    picture: bytes = b""
    picture_size_mode: int = 0
    picture_alignment: int = 2
    picture_position: int = 0x00070001
    accelerator: str = ""
    triple_state: bool = False
    alignment: int = 1
    # ComboBox: only items of the list may be typed. TextBox: Enter starts a new line.
    match_required: bool = False
    enter_key_behavior: bool = False
    # TabStrip tabs (captions); a MultiPage's pages are its children.
    tabs: Tuple[str, ...] = ()
    children: Tuple["ControlDesign", ...] = ()

    def walk(self):
        yield self
        for c in self.children:
            yield from c.walk()


@dataclass(frozen=True)
class FormDesign:
    """A UserForm: its own properties and its controls (in tab order of the designer)."""

    name: str
    caption: str = ""
    width: float = 240.0  # the client area (InsideWidth), points
    height: float = 180.0
    back_color: int = SYS_BUTTON_FACE
    fore_color: int = SYS_BUTTON_TEXT
    border_color: int = SYS_BUTTON_TEXT
    border_style: int = 0
    special_effect: int = 0
    scroll_bars: int = 0
    font: FontDesign = field(default_factory=FontDesign)
    # 0 manual, 1 center on owner, 2 center on screen, 3 Windows default.
    start_up_position: int = 1
    # Where the inside of the form starts (ClientLeft, ClientTop: the designer's 120 and 465
    # twips for a new form).
    left: float = 6.0
    top: float = 23.25
    show_modal: bool = True
    enabled: bool = True
    picture: bytes = b""
    controls: Tuple[ControlDesign, ...] = ()
    # Problems met while reading (a control kind ReflexGrid doesn't know, a stream it couldn't
    # read): the form still shows what could be read.
    warnings: Tuple[str, ...] = ()

    def walk(self):
        for c in self.controls:
            yield from c.walk()

    def find(self, name: str) -> Optional[ControlDesign]:
        low = name.lower()
        return next((c for c in self.walk() if c.name.lower() == low), None)


# What the form designer gives a new control of each kind: its size and its look.
_NEW: dict = {
    "TextBox": dict(width=72, height=18, back_color=SYS_WINDOW, fore_color=SYS_WINDOW_TEXT, special_effect=2, word_wrap=True),
    "ComboBox": dict(width=72, height=18, back_color=SYS_WINDOW, fore_color=SYS_WINDOW_TEXT, special_effect=2, word_wrap=True, match_entry=1),
    "ListBox": dict(width=72, height=72, back_color=SYS_WINDOW, fore_color=SYS_WINDOW_TEXT, special_effect=2, word_wrap=True, match_entry=0, scroll_bars=3),
    "CheckBox": dict(width=108, height=18, special_effect=2, word_wrap=True, value="0"),
    "OptionButton": dict(width=108, height=18, special_effect=2, word_wrap=True, value="0"),
    "ToggleButton": dict(width=24, height=24, special_effect=2, word_wrap=True, value="0", text_align=2),
    "CommandButton": dict(width=72, height=24, text_align=2),
    "Label": dict(width=72, height=18, word_wrap=True, tab_stop=False),
    "Frame": dict(width=144, height=72, special_effect=3),
    "MultiPage": dict(width=144, height=108),
    "TabStrip": dict(width=144, height=108),
    "ScrollBar": dict(width=12.75, height=63.75),
    "SpinButton": dict(width=12.75, height=25.5, max=100),
    "Image": dict(width=72, height=72, border_style=1, tab_stop=False),
}
_CAPTIONED = ("CheckBox", "OptionButton", "ToggleButton", "CommandButton", "Label", "Frame")


def new_control(name: str, kind: str, **props) -> ControlDesign:
    """A control as the designer adds it (a TextBox white and sunken, a button's caption
    centred…), with ``props`` set on top."""
    base = dict(_NEW.get(kind, {}))
    if kind in _CAPTIONED:
        base["caption"] = name
    if kind == "MultiPage" and "children" not in props:
        base["children"] = tuple(ControlDesign(f"Page{i}", "Page", caption=f"Page{i}", tab_index=i - 1) for i in (1, 2))
    if kind == "TabStrip":
        base["tabs"] = ("Tab1", "Tab2")
    base.update(props)
    return ControlDesign(name, kind, **base)


# ---------------------------------------------------------------- the Properties window

_TEXT = ("TextBox", "ComboBox")
_LISTS = ("ListBox", "ComboBox")
_TOGGLES = ("CheckBox", "OptionButton", "ToggleButton")
_NUMBERS = ("ScrollBar", "SpinButton")
_CAPTIONS = ("Label", "CommandButton", "CheckBox", "OptionButton", "ToggleButton", "Frame", "Page")
_ALL = None

# What the form designer's Properties window lists for a control: (property as VBA names it,
# ControlDesign field — "font.name" for the font's —, type, the kinds that have it). Types: text,
# number, int, bool, color, or a list of (value, label) choices.
PROPERTY_SHEET: Tuple[Tuple[str, str, Any, Optional[Tuple[str, ...]]], ...] = (
    ("(Name)", "name", "text", _ALL),
    ("Caption", "caption", "text", _CAPTIONS),
    ("Text", "value", "text", _TEXT),
    ("Value", "value", (("1", "True"), ("0", "False"), ("", "Null")), _TOGGLES),
    ("Value", "position", "int", _NUMBERS + ("MultiPage", "TabStrip")),
    ("Left", "left", "number", _ALL),
    ("Top", "top", "number", _ALL),
    ("Width", "width", "number", _ALL),
    ("Height", "height", "number", _ALL),
    ("Enabled", "enabled", "bool", _ALL),
    ("Visible", "visible", "bool", _ALL),
    ("Locked", "locked", "bool", _TEXT + ("ListBox",) + _TOGGLES),
    ("TabIndex", "tab_index", "int", _ALL),
    ("TabStop", "tab_stop", "bool", _ALL),
    ("Default", "default", "bool", ("CommandButton",)),
    ("Cancel", "cancel", "bool", ("CommandButton",)),
    ("ControlTipText", "tip", "text", _ALL),
    ("Tag", "tag", "text", _ALL),
    ("BackColor", "back_color", "color", _ALL),
    ("ForeColor", "fore_color", "color", _ALL),
    ("BackStyle", "back_style_opaque", (("1", "1 - fmBackStyleOpaque"), ("0", "0 - fmBackStyleTransparent")), ("Label", "CommandButton", "CheckBox", "OptionButton", "ToggleButton", "TextBox", "ComboBox", "Image")),
    ("Font", "font.name", "text", _ALL),
    ("FontSize", "font.size", "number", _ALL),
    ("FontBold", "font.bold", "bool", _ALL),
    ("FontItalic", "font.italic", "bool", _ALL),
    ("TextAlign", "text_align", (("1", "1 - fmTextAlignLeft"), ("2", "2 - fmTextAlignCenter"), ("3", "3 - fmTextAlignRight")), ("Label", "TextBox", "ComboBox", "CheckBox", "OptionButton", "ToggleButton")),
    ("WordWrap", "word_wrap", "bool", ("Label", "CommandButton", "TextBox", "CheckBox", "OptionButton", "ToggleButton")),
    ("AutoSize", "auto_size", "bool", ("Label", "CommandButton", "TextBox", "ComboBox", "CheckBox", "OptionButton", "ToggleButton", "Image")),
    ("MultiLine", "multi_line", "bool", ("TextBox",)),
    ("EnterKeyBehavior", "enter_key_behavior", "bool", ("TextBox",)),
    ("PasswordChar", "password_char", "text", ("TextBox",)),
    ("MaxLength", "max_length", "int", _TEXT),
    ("Style", "style", (("0", "0 - fmStyleDropDownCombo"), ("2", "2 - fmStyleDropDownList")), ("ComboBox",)),
    ("MatchRequired", "match_required", "bool", ("ComboBox",)),
    ("MultiSelect", "multi_select", (("0", "0 - fmMultiSelectSingle"), ("1", "1 - fmMultiSelectMulti"), ("2", "2 - fmMultiSelectExtended")), ("ListBox",)),
    ("ColumnCount", "column_count", "int", _LISTS),
    ("BoundColumn", "bound_column", "int", _LISTS),
    ("ColumnHeads", "column_heads", "bool", _LISTS),
    ("ListRows", "list_rows", "int", ("ComboBox",)),
    ("MatchEntry", "match_entry", (("0", "0 - fmMatchEntryFirstLetter"), ("1", "1 - fmMatchEntryComplete"), ("2", "2 - fmMatchEntryNone")), _LISTS),
    ("RowSource", "row_source", "text", _LISTS),
    ("ControlSource", "control_source", "text", _TEXT + ("ListBox",) + _TOGGLES + _NUMBERS),
    ("GroupName", "group_name", "text", ("OptionButton",)),
    ("TripleState", "triple_state", "bool", _TOGGLES),
    ("Alignment", "alignment", (("1", "1 - fmAlignmentRight"), ("0", "0 - fmAlignmentLeft")), ("CheckBox", "OptionButton")),
    ("Min", "min", "int", _NUMBERS),
    ("Max", "max", "int", _NUMBERS),
    ("SmallChange", "small_change", "int", _NUMBERS),
    ("LargeChange", "large_change", "int", ("ScrollBar",)),
    ("Orientation", "orientation", (("-1", "-1 - fmOrientationAuto"), ("0", "0 - fmOrientationVertical"), ("1", "1 - fmOrientationHorizontal")), _NUMBERS),
    ("SpecialEffect", "special_effect", (("0", "0 - fmSpecialEffectFlat"), ("1", "1 - fmSpecialEffectRaised"), ("2", "2 - fmSpecialEffectSunken"), ("3", "3 - fmSpecialEffectEtched"), ("6", "6 - fmSpecialEffectBump")), ("Label", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "Frame", "Image")),
    ("BorderStyle", "border_style", (("0", "0 - fmBorderStyleNone"), ("1", "1 - fmBorderStyleSingle")), ("Label", "TextBox", "ComboBox", "ListBox", "Image", "Frame")),
    ("BorderColor", "border_color", "color", ("Label", "TextBox", "ComboBox", "ListBox", "Image", "Frame")),
    ("PictureSizeMode", "picture_size_mode", (("0", "0 - fmPictureSizeModeClip"), ("1", "1 - fmPictureSizeModeStretch"), ("3", "3 - fmPictureSizeModeZoom")), ("Image",)),
    ("Accelerator", "accelerator", "text", ("CommandButton", "Label", "CheckBox", "OptionButton", "ToggleButton")),
)

FORM_PROPERTY_SHEET: Tuple[Tuple[str, str, Any], ...] = (
    ("(Name)", "name", "readonly"),
    ("Caption", "caption", "text"),
    ("Width", "width", "number"),
    ("Height", "height", "number"),
    ("BackColor", "back_color", "color"),
    ("ForeColor", "fore_color", "color"),
    ("Font", "font.name", "text"),
    ("FontSize", "font.size", "number"),
    ("FontBold", "font.bold", "bool"),
    ("StartUpPosition", "start_up_position", (("0", "0 - Manual"), ("1", "1 - CenterOwner"), ("2", "2 - CenterScreen"), ("3", "3 - WindowsDefault"))),
    ("ShowModal", "show_modal", "bool"),
    ("Left", "left", "number"),
    ("Top", "top", "number"),
)


def ole_color_text(c: int) -> str:
    """An OLE_COLOR as the Properties window writes it: ``&H8000000F&``."""
    return f"&H{int(c) & 0xFFFFFFFF:08X}&"


def parse_ole_color(text: Any) -> int:
    """``&H8000000F&``, ``#FF0000`` (red, as RGB) or a number."""
    t = str(text).strip()
    if t.startswith("#") and len(t) == 7:
        rgb = int(t[1:], 16)
        return ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | (rgb >> 16)
    if t.upper().startswith("&H"):
        return int(t[2:].rstrip("&"), 16) & 0xFFFFFFFF
    return int(float(t)) & 0xFFFFFFFF


def _get(obj: Any, key: str) -> Any:
    if key.startswith("font."):
        return getattr(obj.font, key[5:])
    return getattr(obj, key)


def _show(kind: Any, value: Any) -> Any:
    if kind == "color":
        return ole_color_text(value)
    if isinstance(kind, tuple):
        return str(int(value)) if isinstance(value, bool) else str(value)
    return value


def property_rows(design: FormDesign, control: str = "") -> list:
    """The Properties window for ``control`` (the form itself when empty): ``{key, label, type,
    value, choices}`` rows, in the order the VBA editor lists them."""
    rows = []
    if not control:
        for label, key, kind in FORM_PROPERTY_SHEET:
            rows.append(_row(label, key, kind, _get(design, key)))
        return rows
    c = design.find(control)
    if c is None:
        return rows
    for label, key, kind, kinds in PROPERTY_SHEET:
        if kinds is not None and c.kind not in kinds:
            continue
        if c.kind == "Page" and key not in ("name", "caption", "tip", "tag", "enabled", "back_color", "fore_color"):
            continue
        rows.append(_row(label, key, kind, _get(c, key)))
    return rows


def _row(label: str, key: str, kind: Any, value: Any) -> dict:
    return {
        "key": key,
        "label": label,
        "type": kind if isinstance(kind, str) else "choice",
        "value": _show(kind, value),
        "choices": [{"value": v, "label": t} for v, t in kind] if isinstance(kind, tuple) else [],
    }


def property_change(target: Any, key: str, value: Any, form: bool = False) -> dict:
    """The design change (``replace`` arguments) that sets property ``key`` of ``target`` (a
    ControlDesign, or the FormDesign when ``form``) from what the Properties window gives."""
    if form:
        kind = next((kd for _, k, kd in FORM_PROPERTY_SHEET if k == key), None)
    else:
        kind = next((kd for _, k, kd, kinds in PROPERTY_SHEET if k == key and (kinds is None or target.kind in kinds)), None)
    if kind is None or kind == "readonly":
        raise ValueError(f"{key} can't be set here")
    current = _get(target, key)
    if kind == "color":
        new: Any = parse_ole_color(value)
    elif kind == "bool":
        new = value if isinstance(value, bool) else str(value).strip().lower() in ("true", "1", "yes", "on")
    elif kind == "int":
        new = int(float(value))
    elif kind == "number":
        new = round(float(value), 2)
        if key in ("width", "height", "font.size") and new <= 0:
            raise ValueError(f"{key} must be more than 0")
    elif isinstance(kind, tuple):
        allowed = [v for v, _ in kind]
        if str(value) not in allowed:
            raise ValueError(f"{value!r} isn't one of {', '.join(allowed)}")
        new = str(value) if key == "value" else (str(value) == "1" if isinstance(current, bool) else int(value))
    else:
        new = str(value)
        if key in ("password_char", "accelerator"):
            new = new[:1]
    if key.startswith("font."):
        from dataclasses import replace as _replace

        return {"font": _replace(target.font, **{key[5:]: new})}
    return {key: new}
