"""UserForms at run time: the Microsoft Forms object model a form's code uses (the form, its
controls, ``Controls``, ``Font``, MultiPage ``Pages``…), started from the form's design
(:class:`gridcore.model.forms.FormDesign`).

A UserForm in VBA is a class module with a predeclared instance; its object adds the form's own
members (``Caption``, ``Show``, ``Hide``, ``Controls``…) and one member per control
(``TextBox1``). :class:`UserForm` is that part of a form instance (``Instance.form``). Setting a
property fires the events Excel fires (setting ``TextBox1.Text`` runs ``TextBox1_Change``,
``ListBox1.ListIndex`` runs ``ListBox1_Click``); what a person does on the form in the browser
comes in as the same property changes and events (:meth:`UserForm.dispatch`).

Showing a form is the host's business (``host.show_form``): a modal form keeps its ``Show`` call
waiting while the person uses it, as a MsgBox does; a modeless one stays open after the macro.
"""

from __future__ import annotations

import base64
import dataclasses
import itertools
import weakref
from typing import Any, Dict, Iterator, List, Optional, Sequence, Tuple

from ..model.forms import ControlDesign, FontDesign, FormDesign, new_control
from . import values as V
from .errors import NotSupported, VbaError
from .runtime import HostObject, Instance, Slot, index_value, slot_value
from .values import EMPTY, MISSING, NOTHING, NULL, VArray, is_object

_ids = itertools.count(1)

# Width and height a form adds around its inside (borders, title bar), in points.
FRAME_WIDTH = 4.0
FRAME_HEIGHT = 24.0

# System colours (OLE_COLOR 0x800000nn) as the browser draws them.
SYSTEM_COLORS = {
    0x00: "#C8C8C8", 0x01: "#000000", 0x02: "#99B4D1", 0x03: "#BFCDDB", 0x04: "#F0F0F0", 0x05: "#FFFFFF",
    0x06: "#646464", 0x07: "#000000", 0x08: "#000000", 0x09: "#000000", 0x0A: "#B4B4B4", 0x0B: "#F4F7FC",
    0x0C: "#ABABAB", 0x0D: "#0078D7", 0x0E: "#FFFFFF", 0x0F: "#F0F0F0", 0x10: "#A0A0A0", 0x11: "#6D6D6D",
    0x12: "#000000", 0x13: "#000000", 0x14: "#FFFFFF", 0x15: "#696969", 0x16: "#E3E3E3", 0x17: "#000000",
    0x18: "#FFFFE1",
}

# The control a "Forms.X.1" ProgID creates (Controls.Add).
PROGIDS = {
    "forms.label.1": "Label", "forms.textbox.1": "TextBox", "forms.combobox.1": "ComboBox", "forms.listbox.1": "ListBox",
    "forms.checkbox.1": "CheckBox", "forms.optionbutton.1": "OptionButton", "forms.togglebutton.1": "ToggleButton",
    "forms.frame.1": "Frame", "forms.commandbutton.1": "CommandButton", "forms.tabstrip.1": "TabStrip",
    "forms.multipage.1": "MultiPage", "forms.scrollbar.1": "ScrollBar", "forms.spinbutton.1": "SpinButton",
    "forms.image.1": "Image",
}
_VALUE_KINDS = {"TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton", "ScrollBar", "SpinButton", "MultiPage", "TabStrip", "CommandButton"}
_LIST_KINDS = {"ListBox", "ComboBox"}
_TOGGLE_KINDS = {"CheckBox", "OptionButton", "ToggleButton"}
_CAPTION_KINDS = {"Label", "CommandButton", "CheckBox", "OptionButton", "ToggleButton", "Frame", "Page"}
_CONTAINER_KINDS = {"Frame", "Page", "MultiPage"}


def css_color(c: Any) -> str:
    try:
        n = int(c) & 0xFFFFFFFF
    except (TypeError, ValueError):
        return "#000000"
    if n & 0x80000000:
        return SYSTEM_COLORS.get(n & 0xFF, "#000000")
    return f"#{n & 0xFF:02X}{(n >> 8) & 0xFF:02X}{(n >> 16) & 0xFF:02X}"


def picture_url(data: bytes) -> str:
    """A picture as the browser can show it ("" for formats it can't: WMF, EMF)."""
    if not data:
        return ""
    mime = None
    if data[:2] == b"BM":
        mime = "image/bmp"
    elif data[:6] in (b"GIF87a", b"GIF89a"):
        mime = "image/gif"
    elif data[:3] == b"\xFF\xD8\xFF":
        mime = "image/jpeg"
    elif data[:8] == b"\x89PNG\r\n\x1a\n":
        mime = "image/png"
    elif data[:4] == b"\x00\x00\x01\x00":
        mime = "image/x-icon"
    if mime is None:
        return ""
    return f"data:{mime};base64," + base64.b64encode(data).decode("ascii")


def _long(v: Any) -> int:
    return V.to_long(_plain(v))


def _dbl(v: Any) -> float:
    return V.to_double(_plain(v))


def _bool(v: Any) -> bool:
    return V.to_boolean(_plain(v))


def _str(v: Any) -> str:
    v = _plain(v)
    return "" if v is EMPTY or v is NULL else V.text(v)


def _plain(v: Any) -> Any:
    if isinstance(v, Slot):
        v = slot_value(v)
    if is_object(v) and v is not NOTHING:
        v = V.value_of(v)
    return v


def _m(v: Any) -> bool:
    return v is MISSING


# ---------------------------------------------------------------- event arguments


class ReturnBoolean(HostObject):
    """``Cancel As MSForms.ReturnBoolean``: the handler sets it to stop what is happening."""

    TYPE_NAME = "ReturnBoolean"
    DEFAULT = "Value"

    def __init__(self, v: bool = False):
        self.v = v

    def g_value(self, ex):
        return self.v

    def l_value(self, ex, v):
        self.v = _bool(v)


class ReturnInteger(HostObject):
    """``KeyAscii As MSForms.ReturnInteger``: the handler may change or zero it."""

    TYPE_NAME = "ReturnInteger"
    DEFAULT = "Value"

    def __init__(self, v: int = 0):
        self.v = v

    def g_value(self, ex):
        return V.Integer(self.v)

    def l_value(self, ex, v):
        self.v = _long(v)


class ReturnString(HostObject):
    TYPE_NAME = "ReturnString"
    DEFAULT = "Value"

    def __init__(self, v: str = ""):
        self.v = v

    def g_value(self, ex):
        return self.v

    def l_value(self, ex, v):
        self.v = _str(v)


# ---------------------------------------------------------------- fonts


class Font(HostObject):
    """A control's (or the form's) ``Font``: changing it redraws the control."""

    TYPE_NAME = "Font"
    DEFAULT = "Name"

    def __init__(self, f: FontDesign):
        self.name, self.size, self.bold, self.italic, self.underline, self.strike = f.name, f.size, f.bold, f.italic, f.underline, f.strike

    def snapshot(self) -> Dict[str, Any]:
        return {"name": self.name, "size": self.size, "bold": self.bold, "italic": self.italic, "underline": self.underline, "strike": self.strike}

    def g_name(self, ex):
        return self.name

    def l_name(self, ex, v):
        self.name = _str(v)

    def g_size(self, ex):
        return V.Currency.of(self.size) if hasattr(V, "Currency") else self.size

    def l_size(self, ex, v):
        n = _dbl(v)
        if not 1 <= n <= 1638:
            raise VbaError(380)
        self.size = n

    def g_bold(self, ex):
        return self.bold

    def l_bold(self, ex, v):
        self.bold = _bool(v)

    def g_italic(self, ex):
        return self.italic

    def l_italic(self, ex, v):
        self.italic = _bool(v)

    def g_underline(self, ex):
        return self.underline

    def l_underline(self, ex, v):
        self.underline = _bool(v)

    def g_strikethrough(self, ex):
        return self.strike

    def l_strikethrough(self, ex, v):
        self.strike = _bool(v)

    def g_weight(self, ex):
        return 700 if self.bold else 400

    def l_weight(self, ex, v):
        self.bold = _long(v) >= 600

    def g_charset(self, ex):
        return 0

    def l_charset(self, ex, v):
        pass


# ---------------------------------------------------------------- controls


class Control(HostObject):
    """One control on a form (or in a Frame or a Page). ``kind`` is its TypeName."""

    PARAMS = {
        "move": ("Left", "Top", "Width", "Height", "Layout"),
        "additem": ("pvargItem", "pvargIndex"),
        "removeitem": ("pvargIndex",),
        "list": ("pvargIndex", "pvargColumn"),
        "column": ("pvargColumn", "pvargIndex"),
        "selected": ("pvargIndex",),
    }

    def __init__(self, form: "UserForm", d: ControlDesign, parent: Any):
        self.form = form
        self.parent = parent
        self.kind = d.kind
        self.design = d
        self.name = d.name
        self.left, self.top, self.width, self.height = d.left, d.top, d.width, d.height
        self.visible = d.visible
        self.enabled = d.enabled
        self.locked = d.locked
        self.tab_index = d.tab_index
        self.tab_stop = d.tab_stop
        self.tag = d.tag
        self.tip = d.tip
        self.caption = d.caption
        self.fore_color = d.fore_color
        self.back_color = d.back_color
        self.back_style = 1 if d.back_style_opaque else 0
        self.border_style = d.border_style
        self.border_color = d.border_color
        self.special_effect = d.special_effect
        self.font = Font(d.font)
        self.text_align = d.text_align
        self.word_wrap = d.word_wrap
        self.auto_size = d.auto_size
        self.multi_line = d.multi_line
        self.password_char = d.password_char
        self.max_length = d.max_length
        self.default = d.default
        self.cancel = d.cancel
        self.group_name = d.group_name
        self.triple_state = d.triple_state
        self.picture = d.picture
        self.control_source = d.control_source
        self.row_source = d.row_source
        self.listeners: List[Tuple[Any, str]] = []  # WithEvents variables holding this control
        # Changed by the person since it got the focus (BeforeUpdate/AfterUpdate when it loses it).
        self.dirty = False
        # Values.
        self.text = ""
        self.value: Any = EMPTY
        self.list: List[List[Any]] = []
        self.list_index = -1
        self.selected: List[bool] = []
        self.column_count = max(1, d.column_count)
        self.column_widths = d.column_widths
        self.bound_column = d.bound_column
        self.text_column = d.text_column
        self.multi_select = d.multi_select
        self.style = d.style
        self.match_required = False
        self.list_rows = d.list_rows
        self.column_heads = d.column_heads
        self.min, self.max = d.min, d.max
        self.small_change, self.large_change = d.small_change, d.large_change
        self.orientation = d.orientation
        self.sel_start = 0
        self.sel_length = 0
        self.top_index = 0
        self.tabs: List[str] = list(d.tabs)
        self.children: List[Control] = []
        if self.kind == "TextBox":
            self.text = d.value
            self.value = d.value
        elif self.kind == "ComboBox":
            self.text = d.value
            self.value = d.value
        elif self.kind in _TOGGLE_KINDS:
            self.value = {"1": True, "0": False}.get(d.value, NULL if d.triple_state and d.value == "" else False)
        elif self.kind in ("ScrollBar", "SpinButton"):
            self.value = d.position
        elif self.kind == "TabStrip":
            self.value = d.position
        elif self.kind == "CommandButton":
            self.value = False
        for c in d.children:
            self.children.append(make_control(form, c, self))
        if self.kind == "MultiPage":
            self.value = d.position if 0 <= d.position < len(self.children) else 0

    # -- identity and position

    def vba_type_name(self) -> str:
        return self.kind

    @property
    def TYPE_NAMES(self) -> Tuple[str, ...]:
        return (self.kind.lower(), "control")

    @property
    def DEFAULT(self) -> Optional[str]:  # type: ignore[override]
        if self.kind in _VALUE_KINDS:
            return "Value"
        if self.kind in ("Label", "Page"):
            return "Caption"
        if self.kind == "Frame":
            return "Controls"
        return None

    def __repr__(self) -> str:
        return f"<{self.kind} {self.name}>"

    def _need(self, *kinds: str) -> None:
        if self.kind not in kinds:
            raise VbaError(438, f"Object doesn't support this property or method ({self.kind})")

    def fire(self, ex, event: str, args: Sequence[Any] = ()) -> bool:
        return self.form.fire(ex, self, event, args)

    def g_name(self, ex):
        return self.name

    def l_name(self, ex, v):
        if self.name.lower() in self.form.designed:
            raise VbaError(382, "Could not set the Name property. Set not supported at runtime.")
        new = _str(v)
        if not new or (self.form.find(new) not in (None, self)):
            raise VbaError(380, "Could not set the Name property. Invalid property value.")
        self.name = new

    def g_left(self, ex):
        return self.left

    def l_left(self, ex, v):
        self.left = _dbl(v)

    def g_top(self, ex):
        return self.top

    def l_top(self, ex, v):
        self.top = _dbl(v)

    def g_width(self, ex):
        return self.width

    def l_width(self, ex, v):
        n = _dbl(v)
        if n < 0:
            raise VbaError(380)
        self.width = n

    def g_height(self, ex):
        return self.height

    def l_height(self, ex, v):
        n = _dbl(v)
        if n < 0:
            raise VbaError(380)
        self.height = n

    def g_move(self, ex, left=MISSING, top=MISSING, width=MISSING, height=MISSING, layout=MISSING):
        if not _m(left):
            self.left = _dbl(left)
        if not _m(top):
            self.top = _dbl(top)
        if not _m(width):
            self.width = max(0.0, _dbl(width))
        if not _m(height):
            self.height = max(0.0, _dbl(height))
        return EMPTY

    def g_zorder(self, ex, position=MISSING):
        siblings = self.parent.children if isinstance(self.parent, Control) else self.form.controls
        if self in siblings:
            siblings.remove(self)
            if _m(position) or _long(position) == 0:
                siblings.append(self)
            else:
                siblings.insert(0, self)
        return EMPTY

    def g_visible(self, ex):
        return self.visible

    def l_visible(self, ex, v):
        self.visible = _bool(v)

    def g_enabled(self, ex):
        return self.enabled

    def l_enabled(self, ex, v):
        self.enabled = _bool(v)

    def g_locked(self, ex):
        return self.locked

    def l_locked(self, ex, v):
        self.locked = _bool(v)

    def g_tag(self, ex):
        return self.tag

    def l_tag(self, ex, v):
        self.tag = _str(v)

    def g_controltiptext(self, ex):
        return self.tip

    def l_controltiptext(self, ex, v):
        self.tip = _str(v)

    def g_tabindex(self, ex):
        return V.Integer(self.tab_index)

    def l_tabindex(self, ex, v):
        self.tab_index = _long(v)

    def g_tabstop(self, ex):
        return self.tab_stop

    def l_tabstop(self, ex, v):
        self.tab_stop = _bool(v)

    def g_parent(self, ex):
        if isinstance(self.parent, Control):
            return self.parent
        owner = self.form.owner
        return owner if owner is not None else NOTHING

    def g_object(self, ex):
        return self

    def g_setfocus(self, ex):
        if not self.visible or not self.enabled:
            raise VbaError(2110, "Can't move focus to the control because it is invisible, not enabled, or of a type that does not accept the focus.")
        self.form.focus(ex, self)
        return EMPTY

    def g_font(self, ex):
        return self.font

    def g_forecolor(self, ex):
        return self.fore_color

    def l_forecolor(self, ex, v):
        self.fore_color = _long(v) & 0xFFFFFFFF

    def g_backcolor(self, ex):
        return self.back_color

    def l_backcolor(self, ex, v):
        self.back_color = _long(v) & 0xFFFFFFFF

    def g_backstyle(self, ex):
        return self.back_style

    def l_backstyle(self, ex, v):
        self.back_style = 1 if _long(v) else 0

    def g_borderstyle(self, ex):
        return self.border_style

    def l_borderstyle(self, ex, v):
        self.border_style = _long(v)

    def g_bordercolor(self, ex):
        return self.border_color

    def l_bordercolor(self, ex, v):
        self.border_color = _long(v) & 0xFFFFFFFF

    def g_specialeffect(self, ex):
        return self.special_effect

    def l_specialeffect(self, ex, v):
        self.special_effect = _long(v)

    def g_textalign(self, ex):
        return self.text_align

    def l_textalign(self, ex, v):
        self.text_align = _long(v)

    def g_wordwrap(self, ex):
        return self.word_wrap

    def l_wordwrap(self, ex, v):
        self.word_wrap = _bool(v)

    def g_autosize(self, ex):
        return self.auto_size

    def l_autosize(self, ex, v):
        self.auto_size = _bool(v)

    def g_mousepointer(self, ex):
        return 0

    def l_mousepointer(self, ex, v):
        pass

    def g_picture(self, ex):
        return NOTHING

    def l_picture(self, ex, v):
        pass

    def s_picture(self, ex, v):
        pass

    def g_accelerator(self, ex):
        return self.design.accelerator

    def l_accelerator(self, ex, v):
        pass

    def g_helpcontextid(self, ex):
        return 0

    def g_controlsource(self, ex):
        return self.control_source

    def l_controlsource(self, ex, v):
        self.control_source = _str(v)
        self._from_control_source(ex)

    def g_rowsource(self, ex):
        self._need("ListBox", "ComboBox")
        return self.row_source

    def l_rowsource(self, ex, v):
        self._need("ListBox", "ComboBox")
        self.row_source = _str(v)
        self.load_row_source(ex)

    # -- captions and values

    def g_caption(self, ex):
        if self.kind not in _CAPTION_KINDS:
            raise VbaError(438)
        return self.caption

    def l_caption(self, ex, v):
        if self.kind not in _CAPTION_KINDS:
            raise VbaError(438)
        self.caption = _str(v)
        if self.kind == "Page" and isinstance(self.parent, Control):
            mp = self.parent
            if self in mp.children:
                i = mp.children.index(self)
                while len(mp.tabs) <= i:
                    mp.tabs.append("")
                mp.tabs[i] = self.caption

    def g_value(self, ex):
        k = self.kind
        if k == "TextBox":
            return self.value if self.value is not EMPTY else ""
        if k in _LIST_KINDS:
            return self.bound_value()
        if k in _VALUE_KINDS:
            return self.value
        raise VbaError(438)

    def bound_value(self) -> Any:
        """A list's Value: the BoundColumn of the chosen row (0: the row's index); a ComboBox
        whose text matches no row has its text, a ListBox with nothing chosen Null."""
        if self.kind == "ComboBox" and (self.list_index < 0 or self.list_index >= len(self.list)):
            return self.text
        return self._list_value()

    def l_value(self, ex, v):
        self.set_value(ex, _plain(v))

    def g_text(self, ex):
        k = self.kind
        if k in ("TextBox", "ComboBox"):
            return self.text
        if k == "ListBox":
            if self.list_index < 0 or self.list_index >= len(self.list):
                return ""
            col = (self.text_column if self.text_column > 0 else 1) - 1
            row = self.list[self.list_index]
            return _str(row[col]) if col < len(row) else ""
        raise VbaError(438)

    def l_text(self, ex, v):
        k = self.kind
        if k in ("TextBox", "ComboBox"):
            self.set_text(ex, _str(v))
            return
        if k == "ListBox":
            t = _str(v)
            col = (self.text_column if self.text_column > 0 else 1) - 1
            idx = next((i for i, row in enumerate(self.list) if col < len(row) and _str(row[col]) == t), -1)
            if idx < 0:
                raise VbaError(380, "Could not set the Text property. Invalid property value.")
            self.set_list_index(ex, idx)
            return
        raise VbaError(438)

    def set_text(self, ex, t: str, user: bool = False) -> None:
        """A TextBox's or ComboBox's text (typed, or set by code): Change when it differs. Code
        writes its ControlSource cell at once; what the person types, when they leave the box."""
        if self.kind == "TextBox" and self.max_length and len(t) > self.max_length and user:
            t = t[: self.max_length]
        if t == self.text and self.value == t:
            return
        self.text = t
        self.value = t
        self.sel_start = len(t)
        self.sel_length = 0
        if self.kind == "ComboBox":
            col = (self.text_column if self.text_column > 0 else 1) - 1
            idx = next((i for i, row in enumerate(self.list) if col < len(row) and _str(row[col]) == t), -1)
            self.list_index = idx
        self.fire(ex, "Change")
        if not user:
            self._to_control_source(ex)

    def set_value(self, ex, v: Any, user: bool = False) -> None:
        k = self.kind
        if k in ("TextBox", "ComboBox"):
            self.set_text(ex, "" if v is NULL or v is EMPTY else V.text(v), user)
            return
        if k == "ListBox":
            if v is NULL or v is EMPTY:
                self.set_list_index(ex, -1)
                return
            col = (self.bound_column if self.bound_column > 0 else 1) - 1
            idx = next((i for i, row in enumerate(self.list) if col < len(row) and _str(row[col]) == _str(v)), -1)
            if idx < 0 and self.multi_select == 0:
                raise VbaError(380, "Could not set the Value property. Invalid property value.")
            self.set_list_index(ex, idx)
            return
        if k in _TOGGLE_KINDS:
            if v is NULL or v is EMPTY:
                nv: Any = NULL if self.triple_state else False
            else:
                nv = V.to_boolean(v)
            if nv is self.value or nv == self.value and type(nv) is type(self.value):
                return
            self.value = nv
            if k == "OptionButton" and nv is True:
                for other in self.form.option_group(self):
                    if other is not self and other.value is True:
                        other.value = False
                        other.fire(ex, "Change")
            self.fire(ex, "Change")
            self.fire(ex, "Click")
            self._to_control_source(ex)
            return
        if k in ("ScrollBar", "SpinButton"):
            n = V.to_long(v)
            lo, hi = min(self.min, self.max), max(self.min, self.max)
            if not lo <= n <= hi:
                raise VbaError(380, "Could not set the Value property. Invalid property value.")
            if n == self.value:
                return
            self.value = n
            self.fire(ex, "Change")
            self._to_control_source(ex)
            return
        if k == "MultiPage":
            n = V.to_long(v)
            if not 0 <= n < len(self.children):
                raise VbaError(380, "Could not set the Value property. Invalid property value.")
            if n == self.value:
                return
            self.value = n
            self.fire(ex, "Change")
            return
        if k == "TabStrip":
            n = V.to_long(v)
            if not -1 <= n < len(self.tabs):
                raise VbaError(380)
            if n == self.value:
                return
            self.value = n
            self.fire(ex, "Change")
            return
        if k == "CommandButton":
            if V.to_boolean(v):
                self.fire(ex, "Click")
            return
        raise VbaError(438)

    def _list_value(self) -> Any:
        if self.multi_select != 0 or self.list_index < 0 or self.list_index >= len(self.list):
            return NULL
        if self.bound_column == 0:
            return self.list_index
        col = self.bound_column - 1
        row = self.list[self.list_index]
        return row[col] if col < len(row) else NULL

    # -- ControlSource / RowSource: cells of the workbook

    def _from_control_source(self, ex) -> None:
        if self.control_source:
            v = ex.host.form_cell(ex, self.control_source)
            if self.kind in ("TextBox", "ComboBox"):
                self.text = self.value = "" if v is EMPTY else V.text(v)
            elif self.kind in _TOGGLE_KINDS:
                self.value = False if v is EMPTY else (V.to_boolean(v) if not isinstance(v, str) else v.upper() == "TRUE")
            elif self.kind in ("ScrollBar", "SpinButton") and v is not EMPTY:
                try:
                    self.value = V.to_long(v)
                except VbaError:
                    pass

    def _to_control_source(self, ex) -> None:
        if self.control_source and not self.form.loading:
            ex.host.form_set_cell(ex, self.control_source, self.bound_value() if self.kind in _LIST_KINDS else self.value)

    def load_row_source(self, ex) -> None:
        if not self.row_source:
            return
        rows = ex.host.form_range(ex, self.row_source)
        self.list = [[x for x in row] for row in rows]
        self.selected = [False] * len(self.list)
        if rows and self.column_count < len(rows[0]) and self.design.column_count <= 1:
            pass
        self.list_index = -1

    # -- lists

    def _list_need(self) -> None:
        if self.kind not in _LIST_KINDS:
            raise VbaError(438)

    def g_additem(self, ex, item=MISSING, index=MISSING):
        self._list_need()
        if self.row_source:
            raise VbaError(70, "Permission denied (the list comes from RowSource)")
        row = [EMPTY if _m(item) else _plain(item)] + [NULL] * (self.column_count - 1)
        if _m(index):
            self.list.append(row)
            self.selected.append(False)
        else:
            i = _long(index)
            if not 0 <= i <= len(self.list):
                raise VbaError(-2147024809, "Invalid argument.")
            self.list.insert(i, row)
            self.selected.insert(i, False)
            if self.list_index >= i:
                self.list_index += 1
        return EMPTY

    def g_removeitem(self, ex, index=MISSING):
        self._list_need()
        i = _long(index)
        if not 0 <= i < len(self.list):
            raise VbaError(-2147024809, "Invalid argument.")
        del self.list[i]
        del self.selected[i]
        if self.list_index == i:
            self.list_index = -1
            if self.kind == "ComboBox":
                self.text = self.value = ""
        elif self.list_index > i:
            self.list_index -= 1
        return EMPTY

    def g_clear(self, ex):
        self._list_need()
        if self.row_source:
            raise VbaError(70, "Permission denied (the list comes from RowSource)")
        self.list = []
        self.selected = []
        changed = self.list_index != -1
        self.list_index = -1
        if self.kind == "ComboBox" and self.text:
            self.text = self.value = ""
            changed = True
        if changed:
            self.fire(ex, "Change")
        return EMPTY

    def g_listcount(self, ex):
        self._list_need()
        return len(self.list)

    def g_list(self, ex, index=MISSING, column=MISSING):
        self._list_need()
        if _m(index):
            if not self.list:
                return NULL
            width = max(self.column_count, max(len(r) for r in self.list))
            return VArray.from_rows([[r[c] if c < len(r) else NULL for c in range(width)] for r in self.list], 0)
        i = _long(index)
        c = 0 if _m(column) else _long(column)
        if not 0 <= i < len(self.list) or c < 0:
            raise VbaError(381, "Could not get the List property. Invalid property array index.")
        row = self.list[i]
        return row[c] if c < len(row) else NULL

    def l_list(self, ex, v, index=MISSING, column=MISSING):
        self._list_need()
        if _m(index):
            v = _plain(v)
            if type(v) is not VArray:
                raise VbaError(380)
            if v.dims == 1:
                self.list = [[x] for x in v.data]
            else:
                self.list = [list(r) for r in v.rows()]
                self.column_count = max(self.column_count, len(self.list[0]) if self.list else 1)
            self.selected = [False] * len(self.list)
            self.list_index = -1
            return
        i = _long(index)
        c = 0 if _m(column) else _long(column)
        if not 0 <= i < len(self.list) or c < 0:
            raise VbaError(381, "Could not set the List property. Invalid property array index.")
        row = self.list[i]
        while len(row) <= c:
            row.append(NULL)
        row[c] = _plain(v)

    def g_column(self, ex, column=MISSING, index=MISSING):
        self._list_need()
        if _m(column):
            if not self.list:
                return NULL
            width = max(self.column_count, max(len(r) for r in self.list))
            return VArray.from_rows([[r[c] if c < len(r) else NULL for r in self.list] for c in range(width)], 0)
        c = _long(column)
        i = self.list_index if _m(index) else _long(index)
        if i < 0 or i >= len(self.list):
            return NULL
        row = self.list[i]
        return row[c] if 0 <= c < len(row) else NULL

    def l_column(self, ex, v, column=MISSING, index=MISSING):
        self._list_need()
        c = _long(column)
        i = self.list_index if _m(index) else _long(index)
        if not 0 <= i < len(self.list):
            raise VbaError(381)
        row = self.list[i]
        while len(row) <= c:
            row.append(NULL)
        row[c] = _plain(v)

    def g_listindex(self, ex):
        if self.kind in _LIST_KINDS:
            return self.list_index
        if self.kind == "TabStrip":
            return self.value
        raise VbaError(438)

    def l_listindex(self, ex, v):
        if self.kind == "TabStrip":
            self.set_value(ex, v)
            return
        self._list_need()
        n = _long(v)
        if not -1 <= n < len(self.list):
            raise VbaError(380, "Could not set the ListIndex property. Invalid property value.")
        self.set_list_index(ex, n)

    def set_list_index(self, ex, n: int) -> None:
        if n == self.list_index and self.kind != "ComboBox":
            return
        if self.kind == "ComboBox":
            if n == self.list_index:
                return
            self.list_index = n
            col = (self.text_column if self.text_column > 0 else 1) - 1
            t = "" if n < 0 else (_str(self.list[n][col]) if col < len(self.list[n]) else "")
            self.text = self.value = t
            self.fire(ex, "Change")
            if n >= 0:
                self.fire(ex, "Click")
            self._to_control_source(ex)
            return
        self.list_index = n
        if self.multi_select == 0:
            self.selected = [i == n for i in range(len(self.list))]
        self.fire(ex, "Change")
        self.fire(ex, "Click")
        self._to_control_source(ex)

    def g_selected(self, ex, index=MISSING):
        self._need("ListBox")
        i = _long(index)
        if not 0 <= i < len(self.list):
            raise VbaError(381, "Could not get the Selected property. Invalid property array index.")
        if self.multi_select == 0:
            return i == self.list_index
        return bool(self.selected[i]) if i < len(self.selected) else False

    def l_selected(self, ex, v, index=MISSING):
        self._need("ListBox")
        i = _long(index)
        if not 0 <= i < len(self.list):
            raise VbaError(381, "Could not set the Selected property. Invalid property array index.")
        on = _bool(v)
        if self.multi_select == 0:
            if on:
                self.set_list_index(ex, i)
            elif self.list_index == i:
                self.set_list_index(ex, -1)
            return
        while len(self.selected) < len(self.list):
            self.selected.append(False)
        if self.selected[i] != on:
            self.selected[i] = on
            self.list_index = i
            self.fire(ex, "Change")

    def g_multiselect(self, ex):
        self._need("ListBox")
        return self.multi_select

    def l_multiselect(self, ex, v):
        self._need("ListBox")
        self.multi_select = _long(v)
        self.selected = [False] * len(self.list)

    def g_columncount(self, ex):
        self._list_need()
        return self.column_count

    def l_columncount(self, ex, v):
        self._list_need()
        n = _long(v)
        self.column_count = max(1, n) if n != -1 else max(1, max((len(r) for r in self.list), default=1))
        for r in self.list:
            while len(r) < self.column_count:
                r.append(NULL)

    def g_columnwidths(self, ex):
        self._list_need()
        return self.column_widths

    def l_columnwidths(self, ex, v):
        self._list_need()
        self.column_widths = _str(v)

    def g_boundcolumn(self, ex):
        self._list_need()
        return self.bound_column

    def l_boundcolumn(self, ex, v):
        self._list_need()
        self.bound_column = _long(v)

    def g_textcolumn(self, ex):
        self._list_need()
        return self.text_column

    def l_textcolumn(self, ex, v):
        self._list_need()
        self.text_column = _long(v)

    def g_columnheads(self, ex):
        self._list_need()
        return self.column_heads

    def l_columnheads(self, ex, v):
        self._list_need()
        self.column_heads = _bool(v)

    def g_listrows(self, ex):
        self._need("ComboBox")
        return self.list_rows

    def l_listrows(self, ex, v):
        self._need("ComboBox")
        self.list_rows = _long(v)

    def g_style(self, ex):
        if self.kind == "ComboBox":
            return self.style
        if self.kind in ("MultiPage", "TabStrip"):
            return 0
        raise VbaError(438)

    def l_style(self, ex, v):
        if self.kind == "ComboBox":
            self.style = _long(v)
            return
        if self.kind in ("MultiPage", "TabStrip"):
            return
        raise VbaError(438)

    def g_matchrequired(self, ex):
        self._need("ComboBox")
        return self.match_required

    def l_matchrequired(self, ex, v):
        self._need("ComboBox")
        self.match_required = _bool(v)

    def g_matchentry(self, ex):
        self._list_need()
        return self.design.match_entry

    def l_matchentry(self, ex, v):
        self._list_need()

    def g_dropdown(self, ex):
        self._need("ComboBox")
        return EMPTY

    def g_topindex(self, ex):
        self._list_need()
        return self.top_index

    def l_topindex(self, ex, v):
        self._list_need()
        self.top_index = max(0, _long(v))

    def g_liststyle(self, ex):
        self._list_need()
        return self.design.list_style

    def l_liststyle(self, ex, v):
        self._list_need()

    # -- text boxes

    def g_maxlength(self, ex):
        self._need("TextBox", "ComboBox")
        return self.max_length

    def l_maxlength(self, ex, v):
        self._need("TextBox", "ComboBox")
        self.max_length = max(0, _long(v))

    def g_multiline(self, ex):
        self._need("TextBox")
        return self.multi_line

    def l_multiline(self, ex, v):
        self._need("TextBox")
        self.multi_line = _bool(v)

    def g_passwordchar(self, ex):
        self._need("TextBox")
        return self.password_char

    def l_passwordchar(self, ex, v):
        self._need("TextBox")
        self.password_char = _str(v)[:1]

    def g_textlength(self, ex):
        self._need("TextBox", "ComboBox")
        return len(self.text)

    def g_linecount(self, ex):
        self._need("TextBox")
        return max(1, self.text.count("\n") + 1)

    def g_selstart(self, ex):
        self._need("TextBox", "ComboBox")
        return self.sel_start

    def l_selstart(self, ex, v):
        self._need("TextBox", "ComboBox")
        n = _long(v)
        if n < 0:
            raise VbaError(380)
        self.sel_start = min(n, len(self.text))

    def g_sellength(self, ex):
        self._need("TextBox", "ComboBox")
        return self.sel_length

    def l_sellength(self, ex, v):
        self._need("TextBox", "ComboBox")
        n = _long(v)
        if n < 0:
            raise VbaError(380)
        self.sel_length = min(n, len(self.text) - self.sel_start)

    def g_seltext(self, ex):
        self._need("TextBox", "ComboBox")
        return self.text[self.sel_start:self.sel_start + self.sel_length]

    def l_seltext(self, ex, v):
        self._need("TextBox", "ComboBox")
        t = _str(v)
        new = self.text[: self.sel_start] + t + self.text[self.sel_start + self.sel_length:]
        start = self.sel_start
        self.set_text(ex, new)
        self.sel_start = start + len(t)
        self.sel_length = 0

    def g_scrollbars(self, ex):
        return self.design.scroll_bars

    def l_scrollbars(self, ex, v):
        pass

    def g_enterkeybehavior(self, ex):
        return False

    def l_enterkeybehavior(self, ex, v):
        pass

    def g_integralheight(self, ex):
        return True

    def l_integralheight(self, ex, v):
        pass

    def g_cutcopy(self, ex):
        return EMPTY

    def g_copy(self, ex):
        return EMPTY

    def g_paste(self, ex):
        return EMPTY

    # -- buttons and toggles

    def g_default(self, ex):
        self._need("CommandButton")
        return self.default

    def l_default(self, ex, v):
        self._need("CommandButton")
        on = _bool(v)
        if on:
            for c in self.form.all_controls():
                if c is not self and c.kind == "CommandButton":
                    c.default = False
        self.default = on

    def g_cancel(self, ex):
        self._need("CommandButton")
        return self.cancel

    def l_cancel(self, ex, v):
        self._need("CommandButton")
        on = _bool(v)
        if on:
            for c in self.form.all_controls():
                if c is not self and c.kind == "CommandButton":
                    c.cancel = False
        self.cancel = on

    def g_takefocusonclick(self, ex):
        return True

    def l_takefocusonclick(self, ex, v):
        pass

    def g_groupname(self, ex):
        self._need("OptionButton")
        return self.group_name

    def l_groupname(self, ex, v):
        self._need("OptionButton")
        self.group_name = _str(v)

    def g_triplestate(self, ex):
        self._need(*_TOGGLE_KINDS)
        return self.triple_state

    def l_triplestate(self, ex, v):
        self._need(*_TOGGLE_KINDS)
        self.triple_state = _bool(v)

    def g_alignment(self, ex):
        self._need("CheckBox", "OptionButton")
        return self.design.alignment

    def l_alignment(self, ex, v):
        pass

    # -- numbers

    def g_min(self, ex):
        self._need("ScrollBar", "SpinButton")
        return self.min

    def l_min(self, ex, v):
        self._need("ScrollBar", "SpinButton")
        self.min = _long(v)
        self._clamp(ex)

    def g_max(self, ex):
        self._need("ScrollBar", "SpinButton")
        return self.max

    def l_max(self, ex, v):
        self._need("ScrollBar", "SpinButton")
        self.max = _long(v)
        self._clamp(ex)

    def _clamp(self, ex) -> None:
        lo, hi = min(self.min, self.max), max(self.min, self.max)
        v = self.value if isinstance(self.value, int) else 0
        if not lo <= v <= hi:
            self.value = max(lo, min(hi, v))
            self.fire(ex, "Change")

    def g_smallchange(self, ex):
        self._need("ScrollBar", "SpinButton")
        return self.small_change

    def l_smallchange(self, ex, v):
        self._need("ScrollBar", "SpinButton")
        self.small_change = _long(v)

    def g_largechange(self, ex):
        self._need("ScrollBar")
        return self.large_change

    def l_largechange(self, ex, v):
        self._need("ScrollBar")
        self.large_change = _long(v)

    def g_orientation(self, ex):
        self._need("ScrollBar", "SpinButton")
        return self.orientation

    def l_orientation(self, ex, v):
        self._need("ScrollBar", "SpinButton")
        self.orientation = _long(v)

    def g_delay(self, ex):
        return self.design.delay

    def l_delay(self, ex, v):
        pass

    # -- containers

    def g_controls(self, ex, index=MISSING):
        if self.kind not in ("Frame", "Page"):
            raise VbaError(438)
        coll = Controls(self.form, self)
        return coll if _m(index) else coll.g_item(ex, index)

    def g_pages(self, ex, index=MISSING):
        self._need("MultiPage")
        coll = Pages(self)
        return coll if _m(index) else coll.g_item(ex, index)

    def g_tabs(self, ex, index=MISSING):
        self._need("TabStrip")
        coll = Tabs(self)
        return coll if _m(index) else coll.g_item(ex, index)

    def g_selecteditem(self, ex):
        if self.kind == "MultiPage":
            return self.children[self.value] if 0 <= self.value < len(self.children) else NOTHING
        if self.kind == "TabStrip":
            return Tab(self, self.value) if 0 <= self.value < len(self.tabs) else NOTHING
        raise VbaError(438)

    def g_index(self, ex):
        if self.kind == "Page" and isinstance(self.parent, Control):
            return self.parent.children.index(self)
        raise VbaError(438)

    def l_index(self, ex, v):
        if self.kind == "Page" and isinstance(self.parent, Control):
            sib = self.parent.children
            i = max(0, min(_long(v), len(sib) - 1))
            sib.remove(self)
            sib.insert(i, self)
            return
        raise VbaError(438)

    def g_activecontrol(self, ex):
        if self.kind not in _CONTAINER_KINDS:
            raise VbaError(438)
        a = self.form.active
        return a if a is not None and self._contains(a) else NOTHING

    def _contains(self, c: "Control") -> bool:
        p = c.parent
        while isinstance(p, Control):
            if p is self:
                return True
            p = p.parent
        return False

    def g_scrolltop(self, ex):
        return 0

    def l_scrolltop(self, ex, v):
        pass

    def g_scrollleft(self, ex):
        return 0

    def l_scrollleft(self, ex, v):
        pass

    def g_scrollheight(self, ex):
        return self.height

    def l_scrollheight(self, ex, v):
        pass

    def g_scrollwidth(self, ex):
        return self.width

    def l_scrollwidth(self, ex, v):
        pass

    def g_insidewidth(self, ex):
        return self.width

    def g_insideheight(self, ex):
        return self.height

    def g_repaint(self, ex):
        return EMPTY

    def g_picturesizemode(self, ex):
        return self.design.picture_size_mode

    def l_picturesizemode(self, ex, v):
        pass

    def g_pictureposition(self, ex):
        return self.design.picture_position

    def g_picturealignment(self, ex):
        return self.design.picture_alignment

    def l_picturealignment(self, ex, v):
        pass

    def g_picturetiling(self, ex):
        return False

    def l_picturetiling(self, ex, v):
        pass

    # -- snapshot

    def snapshot(self) -> Dict[str, Any]:
        k = self.kind
        out: Dict[str, Any] = {
            "name": self.name,
            "kind": k,
            "left": self.left,
            "top": self.top,
            "width": self.width,
            "height": self.height,
            "visible": self.visible,
            "enabled": self.enabled,
            "locked": self.locked,
            "tabIndex": self.tab_index,
            "tabStop": self.tab_stop,
            "tip": self.tip,
            "font": self.font.snapshot(),
            "foreColor": css_color(self.fore_color),
            "backColor": css_color(self.back_color),
            "opaque": self.back_style == 1,
            "borderStyle": self.border_style,
            "borderColor": css_color(self.border_color),
            "specialEffect": self.special_effect,
            "textAlign": self.text_align,
            "wordWrap": self.word_wrap,
            "events": sorted(self.form.handlers(self.name)),
        }
        if k in _CAPTION_KINDS:
            out["caption"] = self.caption
        if k in ("TextBox", "ComboBox"):
            out["text"] = self.text
            out["multiLine"] = self.multi_line
            out["passwordChar"] = self.password_char
            out["maxLength"] = self.max_length
        if k in _LIST_KINDS:
            width = max([self.column_count] + [len(r) for r in self.list]) if self.list else self.column_count
            out["list"] = [[("" if (x is EMPTY or x is NULL) else _safe_text(x)) for x in (r + [NULL] * (width - len(r)))][: max(1, self.column_count)] for r in self.list[:5000]]
            out["listIndex"] = self.list_index
            out["selected"] = [i for i, s in enumerate(self.selected) if s] if self.multi_select else ([self.list_index] if self.list_index >= 0 else [])
            out["multiSelect"] = self.multi_select
            out["columnCount"] = self.column_count
            out["columnWidths"] = self.column_widths
            out["style"] = self.style
        if k in _TOGGLE_KINDS:
            out["value"] = None if self.value is NULL else bool(self.value)
            out["groupName"] = self.group_name
            out["alignment"] = self.design.alignment
        if k in ("ScrollBar", "SpinButton"):
            out.update(value=self.value, min=self.min, max=self.max, smallChange=self.small_change, largeChange=self.large_change, orientation=self.orientation)
        if k == "CommandButton":
            out["default"] = self.default
            out["cancel"] = self.cancel
        if k in ("MultiPage", "TabStrip"):
            out["value"] = self.value
            out["tabs"] = [p.caption for p in self.children] if k == "MultiPage" else list(self.tabs)
        if k == "Image" or (self.picture and k in ("CommandButton", "Label", "Frame", "Page", "ToggleButton")):
            out["picture"] = picture_url(self.picture)
            out["pictureSizeMode"] = self.design.picture_size_mode
        if self.children:
            out["children"] = [c.snapshot() for c in self.children]
        return out


def _safe_text(x: Any) -> str:
    try:
        return V.text(x)
    except VbaError:
        return ""


def make_control(form: "UserForm", d: ControlDesign, parent: Any) -> Control:
    return Control(form, d, parent)


# ---------------------------------------------------------------- collections


class Controls(HostObject):
    """``Me.Controls`` / ``Frame1.Controls``: every control inside (also those in Frames and
    Pages, for the form's collection, as in Excel), by index or name."""

    TYPE_NAME = "Controls"
    DEFAULT = "Item"
    PARAMS = {"add": ("bstrProgID", "Name", "Visible")}

    def __init__(self, form: "UserForm", owner: Any = None):
        self.form = form
        self.owner = owner

    def _items(self) -> List[Control]:
        """Every control inside, those in Frames and on a MultiPage's pages too (the pages
        themselves are the MultiPage's ``Pages``, not controls)."""
        out: List[Control] = []

        def walk(cs: List[Control]) -> None:
            for c in cs:
                out.append(c)
                if c.kind != "MultiPage":
                    walk(c.children)
                else:
                    for p in c.children:
                        walk(p.children)

        walk(self.form.controls if self.owner is None else self.owner.children)
        return out

    def g_count(self, ex):
        return len(self._items())

    def g_item(self, ex, index=MISSING):
        items = self._items()
        key = _plain(index)
        if isinstance(key, str):
            low = key.lower()
            for c in items:
                if c.name.lower() == low:
                    return c
            raise VbaError(-2147024809, "Could not find the specified object.")
        i = V.to_long(key)
        if not 0 <= i < len(items):
            raise VbaError(-2147024809, "Could not find the specified object.")
        return items[i]

    def vba_enum(self, ex) -> Iterator[Any]:
        return iter(self._items())

    def g_add(self, ex, progid=MISSING, name=MISSING, visible=MISSING):
        kind = PROGIDS.get(_str(progid).lower())
        if kind is None:
            raise NotSupported(f"Adding a {_str(progid)} control")
        container = self.owner
        if container is not None and container.kind == "MultiPage":
            raise VbaError(438)
        siblings = self.form.controls if container is None else container.children
        nm = _str(name) if not _m(name) else self.form.new_name(kind)
        if self.form.find(nm) is not None:
            raise VbaError(-2147024809, f"There is already a control named {nm}.")
        # As the designer adds it, in the form's font.
        d = new_control(nm, kind, tab_index=len(siblings), visible=True if _m(visible) else _bool(visible), font=self.form.font_design)
        if kind == "MultiPage":
            d = dataclasses.replace(d, children=tuple(dataclasses.replace(p, font=self.form.font_design, width=d.width, height=d.height) for p in d.children))
        c = make_control(self.form, d, container if container is not None else self.form)
        siblings.append(c)
        return c

    def g_remove(self, ex, index=MISSING):
        c = self.g_item(ex, index)
        parent = c.parent
        siblings = parent.children if isinstance(parent, Control) else self.form.controls
        if c in siblings:
            siblings.remove(c)
        if self.form.active is c:
            self.form.active = None
        return EMPTY

    def g_clear(self, ex):
        siblings = self.form.controls if self.owner is None else self.owner.children
        siblings.clear()
        return EMPTY


class Pages(HostObject):
    TYPE_NAME = "Pages"
    DEFAULT = "Item"
    PARAMS = {"add": ("bstrName", "bstrCaption", "lIndex")}

    def __init__(self, mp: Control):
        self.mp = mp

    def g_count(self, ex):
        return len(self.mp.children)

    def g_item(self, ex, index=MISSING):
        key = _plain(index)
        if isinstance(key, str):
            for p in self.mp.children:
                if p.name.lower() == key.lower():
                    return p
            raise VbaError(-2147024809, "Could not find the specified object.")
        i = V.to_long(key)
        if not 0 <= i < len(self.mp.children):
            raise VbaError(-2147024809, "Invalid argument.")
        return self.mp.children[i]

    def vba_enum(self, ex):
        return iter(list(self.mp.children))

    def g_add(self, ex, name=MISSING, caption=MISSING, index=MISSING):
        n = len(self.mp.children)
        nm = _str(name) if not _m(name) else self.mp.form.new_name("Page")
        cap = _str(caption) if not _m(caption) else f"Page{n + 1}"
        d = ControlDesign(name=nm, kind="Page", caption=cap, width=self.mp.width, height=self.mp.height, font=self.mp.design.font)
        p = make_control(self.mp.form, d, self.mp)
        i = n if _m(index) else max(0, min(_long(index), n))
        self.mp.children.insert(i, p)
        return p

    def g_remove(self, ex, index=MISSING):
        p = self.g_item(ex, index)
        i = self.mp.children.index(p)
        self.mp.children.remove(p)
        if self.mp.value >= len(self.mp.children):
            self.mp.value = max(0, len(self.mp.children) - 1)
        elif self.mp.value > i:
            self.mp.value -= 1
        return EMPTY

    def g_clear(self, ex):
        self.mp.children.clear()
        self.mp.value = 0
        return EMPTY


class Tab(HostObject):
    TYPE_NAME = "Tab"
    DEFAULT = "Caption"

    def __init__(self, strip: Control, i: int):
        self.strip, self.i = strip, i

    def g_caption(self, ex):
        return self.strip.tabs[self.i] if self.i < len(self.strip.tabs) else ""

    def l_caption(self, ex, v):
        if self.i < len(self.strip.tabs):
            self.strip.tabs[self.i] = _str(v)

    def g_index(self, ex):
        return self.i

    def g_name(self, ex):
        return f"Tab{self.i + 1}"

    def g_visible(self, ex):
        return True

    def l_visible(self, ex, v):
        pass

    def g_enabled(self, ex):
        return True

    def l_enabled(self, ex, v):
        pass


class Tabs(HostObject):
    TYPE_NAME = "Tabs"
    DEFAULT = "Item"
    PARAMS = {"add": ("bstrName", "bstrCaption", "lIndex")}

    def __init__(self, strip: Control):
        self.strip = strip

    def g_count(self, ex):
        return len(self.strip.tabs)

    def g_item(self, ex, index=MISSING):
        i = _long(index)
        if not 0 <= i < len(self.strip.tabs):
            raise VbaError(-2147024809, "Invalid argument.")
        return Tab(self.strip, i)

    def vba_enum(self, ex):
        return iter([Tab(self.strip, i) for i in range(len(self.strip.tabs))])

    def g_add(self, ex, name=MISSING, caption=MISSING, index=MISSING):
        cap = _str(caption) if not _m(caption) else f"Tab{len(self.strip.tabs) + 1}"
        i = len(self.strip.tabs) if _m(index) else max(0, min(_long(index), len(self.strip.tabs)))
        self.strip.tabs.insert(i, cap)
        return Tab(self.strip, i)

    def g_remove(self, ex, index=MISSING):
        i = _long(index)
        if 0 <= i < len(self.strip.tabs):
            del self.strip.tabs[i]
        return EMPTY

    def g_clear(self, ex):
        self.strip.tabs.clear()
        return EMPTY


# ---------------------------------------------------------------- the form


class UserForm(HostObject):
    """The Forms part of a UserForm instance: its own properties, its controls, and showing,
    hiding and unloading it. ``owner`` is the VBA instance (``Me``), held weakly: a form nobody
    refers to goes away (and runs UserForm_Terminate) unless it is on screen."""

    PARAMS = {"show": ("Modal",), "move": ("Left", "Top", "Width", "Height", "Layout")}
    DEFAULT = "Controls"

    def __init__(self, rt: Any, owner: Any, design: FormDesign):
        self.rt = rt
        self._owner = weakref.ref(owner)
        self.module = getattr(owner, "cls", None)
        self.id = next(_ids)
        self.design = design
        self.name = design.name
        self.registry: Optional["FormRegistry"] = None
        self.loading = True
        self.loaded = True
        self.shown = False
        self.modal = False
        self.caption = design.caption
        self.width = design.width + FRAME_WIDTH
        self.height = design.height + FRAME_HEIGHT
        self.left = design.left
        self.top = design.top
        self.start_up = design.start_up_position
        self.back_color = design.back_color
        self.fore_color = design.fore_color
        self.enabled = design.enabled
        self.tag = ""
        self.font_design = design.font
        self.font = Font(design.font)
        self.picture = design.picture
        self.active: Optional[Control] = None
        self.controls: List[Control] = [make_control(self, d, self) for d in design.controls]
        # Controls laid out by the designer (their events are compiled against these names).
        self.designed = {c.name.lower() for c in self.all_controls()}

    @property
    def owner(self) -> Any:
        return self._owner()

    def start(self, ex) -> None:
        """Fills lists and values from the workbook (RowSource, ControlSource) and runs
        UserForm_Initialize."""
        for c in self.all_controls():
            if c.row_source:
                try:
                    c.load_row_source(ex)
                except VbaError:
                    pass
            if c.control_source:
                try:
                    c._from_control_source(ex)
                except VbaError:
                    pass
        first = self._tab_order()
        self.active = first[0] if first else None
        self.loading = False
        self.fire(ex, None, "Initialize")

    def member(self, ex) -> "UserForm":
        """The form, for code using it: a form that was unloaded is gone, as in VBA."""
        if not self.loaded:
            raise VbaError(-2147418105, "Automation error: the form was unloaded (Unload) and can't be used any more.")
        return self

    # -- lookups

    def vba_type_name(self) -> str:
        return self.name

    def vba_get(self, ex, name: str, args: list, named: Optional[dict] = None) -> Any:
        c = self.find(name)
        if c is not None:
            return index_value(ex, c, list(args), named) if args else c
        return HostObject.vba_get(self, ex, name, args, named)

    def vba_let(self, ex, name: str, args: list, named: Optional[dict], value: Any) -> None:
        c = self.find(name)
        if c is not None:
            from .runtime import let_default

            return let_default(ex, c, list(args), named, value)
        return HostObject.vba_let(self, ex, name, args, named, value)

    def has_dynamic_member(self, name: str) -> bool:
        return self.find(name) is not None

    def all_controls(self) -> Iterator[Control]:
        def walk(cs: Sequence[Control]) -> Iterator[Control]:
            for c in cs:
                yield c
                yield from walk(c.children)

        return walk(list(self.controls))

    def find(self, name: str) -> Optional[Control]:
        low = name.lower()
        return next((c for c in self.all_controls() if c.name.lower() == low), None)

    def new_name(self, kind: str) -> str:
        n = 1
        while self.find(f"{kind}{n}") is not None:
            n += 1
        return f"{kind}{n}"

    def option_group(self, opt: Control) -> List[Control]:
        """The OptionButtons that turn off when ``opt`` turns on: same GroupName, or (without one)
        the same container."""
        out = []
        for c in self.all_controls():
            if c.kind != "OptionButton":
                continue
            if opt.group_name:
                if c.group_name == opt.group_name:
                    out.append(c)
            elif not c.group_name and c.parent is opt.parent:
                out.append(c)
        return out

    def _tab_order(self) -> List[Control]:
        """Controls that take the focus, in tab order (containers' controls in their place)."""
        out: List[Control] = []

        def walk(cs: Sequence[Control]) -> None:
            for c in sorted(cs, key=lambda c: c.tab_index):
                if not c.visible or not c.enabled:
                    continue
                if c.kind == "MultiPage":
                    if 0 <= c.value < len(c.children):
                        walk(c.children[c.value].children)
                elif c.kind in ("Frame", "Page"):
                    walk(c.children)
                elif c.tab_stop and c.kind not in ("Label", "Image"):
                    out.append(c)

        walk(self.controls)
        return out

    def handlers(self, control: Optional[str]) -> List[str]:
        """The events of ``control`` (or of the form, for None) that code handles: the form's
        module, and WithEvents variables holding the control."""
        out: List[str] = []
        prefix = ("userform" if control is None else control.lower()) + "_"
        if control is None or control.lower() in self.designed:
            procs = getattr(self.module, "procs", {}) if self.module is not None else {}
            out = [n[len(prefix):] for n, d in procs.items() if n.startswith(prefix) and "sub" in d]
        c = self.find(control) if control else None
        if c is not None:
            for listener, var in c.listeners:
                m = listener.cls if type(listener) is Instance else listener
                vp = var.lower() + "_"
                out += [n[len(vp):] for n in getattr(m, "procs", {}) if n.startswith(vp)]
        return sorted(set(out))

    # -- events

    def fire(self, ex, control: Optional[Control], event: str, args: Sequence[Any] = ()) -> bool:
        """Runs ``<Control>_<Event>`` (``UserForm_<Event>``) if the form's module has it, and the
        handlers of WithEvents variables holding the control. Returns whether one ran."""
        if self.loading and event != "Initialize":
            return False
        owner = self.owner
        ran = False
        if owner is not None and (control is None or control.name.lower() in self.designed):
            proc = self._proc(f"{'UserForm' if control is None else control.name}_{event}")
            if proc is not None:
                self._call(ex, proc, args, owner)
                ran = True
        if control is not None:
            for listener, var in list(control.listeners):
                mod = listener.cls if type(listener) is Instance else listener
                d = getattr(mod, "procs", {}).get(f"{var}_{event}".lower())
                p = d.get("sub") if d else None
                if p is not None:
                    self._call(ex, p, args, listener if type(listener) is Instance else None)
                    ran = True
        return ran

    def _proc(self, name: str):
        mod = self.module
        if mod is None:
            return None
        d = mod.procs.get(name.lower())
        return d.get("sub") if d else None

    @staticmethod
    def _call(ex, proc, args: Sequence[Any], me: Any) -> None:
        """Calls an event procedure with the event's arguments, as many as it declares
        (``ByRef Cancel`` arguments are shared, so the handler's answer comes back)."""
        params = proc.params
        a = list(args)[: len(params)]
        for p in params[len(a):]:
            a.append(MISSING if p.optional else (NOTHING if p.is_object else V.DEFAULTS.get(p.type_name, EMPTY)))
        for x, p in zip(a, params):
            if isinstance(x, Slot) and not p.byval:
                x.co = p.co
                if p.co is not None:
                    x.v = p.co(x.v)
        ex.rt.call(ex, proc, a, None, me)

    def commit(self, ex, c: Control) -> bool:
        """The person's change to ``c`` takes effect: BeforeUpdate (which may refuse it), the
        ControlSource cell, AfterUpdate."""
        if not c.dirty:
            return True
        cancel = ReturnBoolean()
        c.fire(ex, "BeforeUpdate", [cancel])
        if cancel.v:
            return False
        c.dirty = False
        if c.kind in ("TextBox", "ComboBox"):
            c._to_control_source(ex)
        c.fire(ex, "AfterUpdate")
        return True

    def focus(self, ex, c: Optional[Control]) -> bool:
        """Moves the focus to ``c`` (None: nowhere): the control losing it gets BeforeUpdate and
        AfterUpdate (when the person changed it) and Exit, any of which may keep the focus
        where it is; then ``c`` gets Enter. Returns whether the focus moved."""
        old = self.active
        if old is c:
            return True
        if old is not None and self.find(old.name) is old:
            if not self.commit(ex, old):
                return False
            cancel = ReturnBoolean()
            old.fire(ex, "Exit", [cancel])
            if cancel.v:
                return False
        self.active = c
        if c is not None:
            c.fire(ex, "Enter")
        return True

    # -- what the person does (from the browser)

    def dispatch(self, ex, ev: Dict[str, Any]) -> None:
        """One thing the person did on the form::

            {"event": "click" | "dblclick" | "change" | "focus" | "keypress" | "keydown" |
                      "keyup" | "spin" | "scroll" | "enter" | "escape" | "close",
             "control": name, "value": …, "key": …,
             "values": {name: value}, "focus": name}

        ``values`` are what the person typed or picked since the last event (applied first,
        each firing Change as the edits would have), ``focus`` the control that has the focus."""
        values = ev.get("values") or {}
        if isinstance(values, dict):
            for name, v in values.items():
                c = self.find(str(name))
                if c is not None and c.enabled and not c.locked and self.loaded:
                    self.apply_user_value(ex, c, v)
        if not self.loaded:
            return
        kind = str(ev.get("event") or "")
        name = ev.get("control")
        c = self.find(str(name)) if name else None
        if kind == "close":
            self.query_unload(ex, 0)
            return
        if kind == "enter":
            # The Enter key outside a multi-line text box: the Default button's Click.
            btn = next((b for b in self.all_controls() if b.kind == "CommandButton" and b.default and b.enabled and b.visible), None)
            if btn is not None and self.focus(ex, btn):
                btn.fire(ex, "Click")
            return
        if kind == "escape":
            btn = next((b for b in self.all_controls() if b.kind == "CommandButton" and b.cancel and b.enabled and b.visible), None)
            if btn is not None:
                btn.fire(ex, "Click")
            return
        focus = ev.get("focus")
        if kind == "change" and c is not None and c is not self.active:
            focus = c.name  # one types into the control that has the focus
        if focus and (c is None or kind not in ("click", "dblclick", "focus")):
            f = self.find(str(focus))
            if f is not None and f.enabled and f.visible:
                self.focus(ex, f)
        # Controls the person changed and has left since (no code was watching them leave):
        # their updates now (BeforeUpdate, the ControlSource cell, AfterUpdate).
        for other in list(self.all_controls()):
            if other.dirty and other is not self.active and self.loaded:
                self.commit(ex, other)
        if not self.loaded:
            return
        if c is None:
            if kind == "click":
                self.fire(ex, None, "Click")
            elif kind == "dblclick":
                self.fire(ex, None, "DblClick", [ReturnBoolean()])
            return
        if not c.enabled or not self.loaded:
            return
        takes_focus = c.kind not in ("Label", "Image", "Frame", "Page", "MultiPage", "TabStrip") and c.tab_stop
        if kind == "focus":
            self.focus(ex, c)
        elif kind == "change":
            if not c.locked:
                self.apply_user_value(ex, c, ev.get("value"))
        elif kind == "click":
            if takes_focus and not self.focus(ex, c):
                return
            if c.kind == "CommandButton":
                c.fire(ex, "Click")
            elif c.kind in _TOGGLE_KINDS:
                if not c.locked:
                    if "value" in ev:
                        self.apply_user_value(ex, c, ev.get("value"))
                    elif c.kind == "OptionButton":
                        c.set_value(ex, True, user=True)
                    else:
                        c.set_value(ex, not bool(c.value) if c.value is not NULL else True, user=True)
            elif c.kind in _LIST_KINDS:
                if "value" in ev and not c.locked:
                    self.apply_user_value(ex, c, ev.get("value"))
            elif c.kind in ("MultiPage", "TabStrip"):
                if "value" in ev:
                    n = _int(ev.get("value"), c.value)
                    if c.kind == "MultiPage" and not 0 <= n < len(c.children):
                        return
                    c.set_value(ex, n, user=True)
                    c.fire(ex, "Click", [n])
            else:
                c.fire(ex, "Click")
        elif kind == "dblclick":
            c.fire(ex, "DblClick", [ReturnBoolean()])
        elif kind == "spin":
            up = bool(ev.get("up"))
            c.fire(ex, "SpinUp" if up else "SpinDown")
            step = c.small_change or 1
            cur = c.value if isinstance(c.value, int) else 0
            lo, hi = min(c.min, c.max), max(c.min, c.max)
            n = cur + (step if up else -step) * (1 if c.max >= c.min else -1)
            c.set_value(ex, max(lo, min(hi, n)), user=True)
        elif kind == "scroll":
            lo, hi = min(c.min, c.max), max(c.min, c.max)
            c.set_value(ex, max(lo, min(hi, _int(ev.get("value"), lo))), user=True)
            c.fire(ex, "Scroll")
        elif kind == "keypress":
            self.key_press(ex, c, str(ev.get("key") or ""), ev)
        elif kind in ("keydown", "keyup"):
            code = ReturnInteger(_int(ev.get("keyCode"), 0))
            c.fire(ex, "KeyDown" if kind == "keydown" else "KeyUp", [code, V.Integer(_int(ev.get("shift"), 0))])
        elif kind == "dropbuttonclick":
            c.fire(ex, "DropButtonClick")

    def apply_user_value(self, ex, c: Control, v: Any) -> None:
        """A value the person gave a control (typed text, a pick from a list, a ticked box)."""
        k = c.kind
        before = (c.text, c.list_index, list(c.selected), c.value if not isinstance(c.value, list) else None)
        if k in ("TextBox", "ComboBox"):
            t = "" if v is None else str(v)
            if k == "ComboBox":
                col = (c.text_column if c.text_column > 0 else 1) - 1
                idx = next((i for i, row in enumerate(c.list) if col < len(row) and _str(row[col]) == t), -1)
                if c.style == 2 or (idx >= 0 and t != c.text):
                    if c.style == 2 and idx < 0:
                        return
                    c.set_list_index(ex, idx)
                else:
                    c.set_text(ex, t, user=True)
            else:
                c.set_text(ex, t, user=True)
        elif k == "ListBox":
            if c.multi_select:
                sel = {int(i) for i in (v or []) if isinstance(i, (int, float))} if isinstance(v, list) else set()
                changed = False
                while len(c.selected) < len(c.list):
                    c.selected.append(False)
                for i in range(len(c.list)):
                    on = i in sel
                    if c.selected[i] != on:
                        c.selected[i] = on
                        changed = True
                if changed:
                    if sel:
                        c.list_index = max(sel)
                    c.fire(ex, "Change")
            else:
                n = _int(v[0] if isinstance(v, list) and v else v, -1) if not (isinstance(v, list) and not v) else -1
                if -1 <= n < len(c.list):
                    c.set_list_index(ex, n)
        elif k in _TOGGLE_KINDS:
            c.set_value(ex, NULL if v is None else bool(v), user=True)
        elif k in ("ScrollBar", "SpinButton"):
            lo, hi = min(c.min, c.max), max(c.min, c.max)
            c.set_value(ex, max(lo, min(hi, _int(v, c.value if isinstance(c.value, int) else lo))), user=True)
        elif k in ("MultiPage", "TabStrip"):
            try:
                c.set_value(ex, _int(v, c.value), user=True)
            except VbaError:
                pass
        after = (c.text, c.list_index, list(c.selected), c.value if not isinstance(c.value, list) else None)
        if after != before:
            c.dirty = True

    def key_press(self, ex, c: Control, key: str, ev: Dict[str, Any]) -> None:
        """A character typed into a TextBox or ComboBox whose KeyPress code handles: the
        handler sees it first and may change it (or set it to 0, dropping it)."""
        if not key or c.kind not in ("TextBox", "ComboBox") or c.locked:
            return
        code = ReturnInteger(ord(key[0]))
        c.fire(ex, "KeyPress", [code])
        if code.v <= 0 or not self.loaded:
            return
        ch = chr(code.v)
        start = max(0, min(_int(ev.get("selStart"), len(c.text)), len(c.text)))
        length = max(0, _int(ev.get("selLength"), 0))
        new = c.text[:start] + ch + c.text[start + length:]
        if c.max_length and len(new) > c.max_length:
            return
        c.set_text(ex, new, user=True)
        c.dirty = True
        c.sel_start = start + 1
        c.sel_length = 0

    # -- the form's members

    def g_caption(self, ex):
        return self.caption

    def l_caption(self, ex, v):
        self.caption = _str(v)

    def g_name(self, ex):
        return self.name

    def g_width(self, ex):
        return self.width

    def l_width(self, ex, v):
        self.width = max(0.0, _dbl(v))
        self.fire(ex, None, "Resize")

    def g_height(self, ex):
        return self.height

    def l_height(self, ex, v):
        self.height = max(0.0, _dbl(v))
        self.fire(ex, None, "Resize")

    def g_insidewidth(self, ex):
        return max(0.0, self.width - FRAME_WIDTH)

    def g_insideheight(self, ex):
        return max(0.0, self.height - FRAME_HEIGHT)

    def g_left(self, ex):
        return self.left

    def l_left(self, ex, v):
        self.left = _dbl(v)
        self.start_up = 0

    def g_top(self, ex):
        return self.top

    def l_top(self, ex, v):
        self.top = _dbl(v)
        self.start_up = 0

    def g_startupposition(self, ex):
        return self.start_up

    def l_startupposition(self, ex, v):
        self.start_up = _long(v)

    def g_move(self, ex, left=MISSING, top=MISSING, width=MISSING, height=MISSING, layout=MISSING):
        if not _m(left):
            self.left = _dbl(left)
            self.start_up = 0
        if not _m(top):
            self.top = _dbl(top)
            self.start_up = 0
        if not _m(width):
            self.width = max(0.0, _dbl(width))
        if not _m(height):
            self.height = max(0.0, _dbl(height))
        return EMPTY

    def g_backcolor(self, ex):
        return self.back_color

    def l_backcolor(self, ex, v):
        self.back_color = _long(v) & 0xFFFFFFFF

    def g_forecolor(self, ex):
        return self.fore_color

    def l_forecolor(self, ex, v):
        self.fore_color = _long(v) & 0xFFFFFFFF

    def g_bordercolor(self, ex):
        return self.design.border_color

    def l_bordercolor(self, ex, v):
        pass

    def g_borderstyle(self, ex):
        return self.design.border_style

    def l_borderstyle(self, ex, v):
        pass

    def g_specialeffect(self, ex):
        return self.design.special_effect

    def l_specialeffect(self, ex, v):
        pass

    def g_enabled(self, ex):
        return self.enabled

    def l_enabled(self, ex, v):
        self.enabled = _bool(v)

    def g_visible(self, ex):
        return self.shown

    def l_visible(self, ex, v):
        if _bool(v):
            self.g_show(ex)
        else:
            self.g_hide(ex)

    def g_tag(self, ex):
        return self.tag

    def l_tag(self, ex, v):
        self.tag = _str(v)

    def g_font(self, ex):
        return self.font

    def g_controls(self, ex, index=MISSING):
        coll = Controls(self)
        return coll if _m(index) else coll.g_item(ex, index)

    def g_activecontrol(self, ex):
        return self.active if self.active is not None else NOTHING

    def g_repaint(self, ex):
        return EMPTY

    def g_printform(self, ex):
        return EMPTY

    def g_setdefaulttaborder(self, ex):
        return EMPTY

    def g_whatsthisbutton(self, ex):
        return False

    def g_whatsthishelp(self, ex):
        return False

    def g_helpcontextid(self, ex):
        return 0

    def g_mousepointer(self, ex):
        return 0

    def l_mousepointer(self, ex, v):
        pass

    def g_picture(self, ex):
        return NOTHING

    def l_picture(self, ex, v):
        pass

    def s_picture(self, ex, v):
        pass

    def g_scrollbars(self, ex):
        return self.design.scroll_bars

    def l_scrollbars(self, ex, v):
        pass

    def g_scrolltop(self, ex):
        return 0

    def l_scrolltop(self, ex, v):
        pass

    def g_scrollleft(self, ex):
        return 0

    def l_scrollleft(self, ex, v):
        pass

    def g_scrollheight(self, ex):
        return self.height

    def l_scrollheight(self, ex, v):
        pass

    def g_scrollwidth(self, ex):
        return self.width

    def l_scrollwidth(self, ex, v):
        pass

    def g_zoom(self, ex):
        return 100

    def l_zoom(self, ex, v):
        pass

    def g_cycle(self, ex):
        return 0

    def l_cycle(self, ex, v):
        pass

    def g_show(self, ex, modal=MISSING):
        """Shows the form. A modal one (the default, unless the designer's ShowModal is False)
        keeps ``Show`` waiting while the person uses it, until code hides or unloads it."""
        if not self.loaded:
            raise VbaError(-2147418105, "Automation error: the form was unloaded (Unload) and can't be shown.")
        want_modal = self.design.show_modal if _m(modal) else _long(modal) != 0
        if self.shown:
            if want_modal and not self.modal:
                raise VbaError(400, "Form already displayed; can't show modally")
            return EMPTY
        reg = ex.host.form_registry(ex)
        reg.add(self)
        self.shown = True
        self.modal = want_modal
        reg.keep_shown(self)
        self.fire(ex, None, "Activate")
        if want_modal and self.shown and self.loaded:
            run_modal(ex, self)
        return EMPTY

    def g_hide(self, ex):
        if self.shown:
            self.shown = False
            self.modal = False
            if self.registry is not None:
                self.registry.drop_shown(self)
        return EMPTY

    def query_unload(self, ex, mode: int) -> bool:
        """QueryClose (``mode`` 0: the person closed the form, 1: ``Unload`` in code), then
        unloads unless the handler cancelled. Returns whether the form was unloaded."""
        if not self.loaded:
            return True
        cancel = Slot(V.Integer(0))
        self.fire(ex, None, "QueryClose", [cancel, Slot(V.Integer(mode))])
        try:
            cancelled = V.to_long(slot_value(cancel)) != 0
        except VbaError:
            cancelled = True
        if cancelled:
            return False
        self.unload(ex)
        return True

    def unload(self, ex) -> None:
        """Unloads the form: UserForm_Terminate, then it is gone (``Unload`` without QueryClose)."""
        if not self.loaded:
            return
        owner = self.owner
        if owner is not None and not owner.terminated:
            owner.terminated = True
            self.fire(ex, None, "Terminate")
        self.loaded = False
        self.shown = False
        self.modal = False
        if self.registry is not None:
            self.registry.forget(self)

    # -- for the browser

    def snapshot(self) -> Dict[str, Any]:
        return {
            "id": self.id,
            "name": self.name,
            "caption": self.caption,
            "width": self.width,
            "height": self.height,
            "insideWidth": max(0.0, self.width - FRAME_WIDTH),
            "insideHeight": max(0.0, self.height - FRAME_HEIGHT),
            "left": self.left,
            "top": self.top,
            "startUp": self.start_up,
            "modal": self.modal,
            "enabled": self.enabled,
            "backColor": css_color(self.back_color),
            "foreColor": css_color(self.fore_color),
            "font": self.font.snapshot(),
            "picture": picture_url(self.picture),
            "events": self.handlers(None),
            "active": self.active.name if self.active is not None else "",
            "controls": [c.snapshot() for c in self.controls],
        }


def _int(v: Any, default: int) -> int:
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def run_modal(ex, form: UserForm) -> None:
    """A modal form on screen: asks the person (``host.prompt`` of kind "form") for what they
    do, one event at a time, and runs it, until the form is hidden or unloaded. MsgBox and
    InputBox wait for the person the same way. With nobody to ask (a script, a test without
    answers) the form is closed as with its close button."""
    while form.shown and form.loaded and form.registry is not None and form.registry.alive(form):
        ans = ex.host.prompt(ex, "form", "", form.caption, 0, None, formId=form.id)
        if not isinstance(ans, dict):
            if not form.query_unload(ex, 0):
                form.unload(ex)
            return
        form.dispatch(ex, ans)


class FormRegistry:
    """The UserForms one person has loaded: the default instances (``UserForm1`` used by name)
    and the forms on screen. A project whose code changed (a new runtime) or ``End`` drops them
    all, as resetting the project does in Excel."""

    def __init__(self) -> None:
        self.rt: Any = None
        self.defaults: Dict[str, Any] = {}
        self.loaded: Dict[int, UserForm] = {}
        # Instances of the forms on screen, kept alive while they are.
        self.on_screen: Dict[int, Any] = {}

    def bind(self, rt: Any) -> "FormRegistry":
        if self.rt is not rt:
            self.rt = rt
            self.defaults.clear()
            self.loaded.clear()
            self.on_screen.clear()
        return self

    def alive(self, form: UserForm) -> bool:
        owner = form.owner
        return form.loaded and owner is not None and owner.gen == self.rt.generation

    def default(self, ex, cls) -> Any:
        key = cls.name.lower()
        inst = self.defaults.get(key)
        if inst is not None and self.alive(inst.form):
            return inst
        inst = ex.rt.make_form(ex, cls)
        self.defaults[key] = inst
        self.add(inst.form)
        inst.form.start(ex)
        return inst

    def add(self, form: UserForm) -> None:
        form.registry = self
        self.loaded[form.id] = form

    def keep_shown(self, form: UserForm) -> None:
        owner = form.owner
        if owner is not None:
            self.on_screen[form.id] = owner

    def drop_shown(self, form: UserForm) -> None:
        self.on_screen.pop(form.id, None)

    def forget(self, form: UserForm) -> None:
        self.loaded.pop(form.id, None)
        self.on_screen.pop(form.id, None)
        for k, inst in list(self.defaults.items()):
            if inst.form is form:
                del self.defaults[k]

    def get(self, form_id: int) -> Optional[UserForm]:
        f = self.loaded.get(form_id)
        return f if f is not None and self.alive(f) else None

    def shown(self) -> List[UserForm]:
        """The forms on screen, in the order they were loaded."""
        for fid, f in list(self.loaded.items()):
            if not self.alive(f):
                self.forget(f)
        return [f for f in self.loaded.values() if f.shown]


def design_snapshot(design: FormDesign) -> Dict[str, Any]:
    """A form as designed, for a preview (no code runs)."""

    class _Nobody:
        cls = None

    nobody = _Nobody()
    f = UserForm(None, nobody, design)
    f.loading = False
    first = f._tab_order()
    f.active = first[0] if first else None
    snap = f.snapshot()
    del nobody
    return snap


# ---------------------------------------------------------------- UserForms, DataObject


class LoadedForms(HostObject):
    """VBA's ``UserForms``: the forms loaded now (``UserForms(0)``, ``UserForms.Count``), and
    ``UserForms.Add("UserForm1")`` to load one by name."""

    TYPE_NAME = "Object"
    DEFAULT = "Item"
    PARAMS = {"add": ("bstrName",), "item": ("Index",)}

    def __init__(self, reg: FormRegistry):
        self.reg = reg

    def _items(self) -> List[Any]:
        return [f.owner for f in self.reg.loaded.values() if self.reg.alive(f)]

    def g_count(self, ex):
        return len(self._items())

    def g_item(self, ex, index=MISSING):
        items = self._items()
        i = _long(index)
        if not 0 <= i < len(items):
            raise VbaError(9, "Subscript out of range")
        return items[i]

    def vba_enum(self, ex):
        return iter(self._items())

    def g_add(self, ex, name=MISSING):
        cls = ex.rt.module(_str(name))
        if cls is None or cls.kind != "form":
            raise VbaError(424, f"Object required: there is no UserForm named {_str(name)}")
        return ex.rt.new_instance(ex, cls)


class DataObject(HostObject):
    """``MSForms.DataObject``: text for the clipboard. The clipboard is the person's own in
    ReflexGrid (a macro can't reach the computer's)."""

    TYPE_NAME = "DataObject"
    PARAMS = {"settext": ("Text", "Format"), "gettext": ("Format",), "getformat": ("Format",)}

    def __init__(self) -> None:
        self.text: Optional[str] = None

    def g_settext(self, ex, text=MISSING, fmt=MISSING):
        self.text = _str(text)
        return EMPTY

    def g_gettext(self, ex, fmt=MISSING):
        if self.text is None:
            raise VbaError(-2147221404, "DataObject:GetText Invalid FORMATETC structure")
        return self.text

    def g_getformat(self, ex, fmt=MISSING):
        return self.text is not None and (_m(fmt) or _long(fmt) in (1, 13))

    def g_putinclipboard(self, ex):
        ex.host.set_clipboard_text(ex, self.text or "")
        return EMPTY

    def g_getfromclipboard(self, ex):
        self.text = ex.host.clipboard_text(ex)
        return EMPTY

    def g_clear(self, ex):
        self.text = None
        return EMPTY

    def g_startdrag(self, ex, effect=MISSING):
        return V.Integer(0)
