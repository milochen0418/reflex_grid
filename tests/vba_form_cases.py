"""UserForm cases run by both ReflexGrid and Microsoft Excel (``scripts/excel_vba_oracle_mac.py
--forms`` and ``--form-events``): forms ReflexGrid writes, loaded and described control by control,
and the events a form's code fires, in order. Excel's answers are in ``vba_forms_expected.json``."""

import base64

from gridcore.engine import Session
from gridcore.model.forms import ControlDesign, FontDesign, FormDesign, new_control
from gridcore.xlsx.oforms import vbframe, write_design, write_frx

# A 1x1 PNG (made here, not taken from anywhere).
PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d4948445200000001000000010806000000"
    "1f15c4890000000d49444154789c6360f8cfc0f01f0005000201a5f4e9"
    "b80000000049454e44ae426082"
)


def every_control() -> FormDesign:
    return FormDesign(
        "Everything",
        caption="All controls",
        width=420,
        height=300,
        show_modal=False,
        start_up_position=0,
        left=30,
        top=40,
        controls=(
            ControlDesign("NameBox", "TextBox", 10, 10, 120, 18, tab_index=0, value="Ann", max_length=20, control_source="Sheet1!B1", tip="Your name"),
            ControlDesign("Secret", "TextBox", 10, 34, 120, 18, tab_index=1, password_char="*", locked=True),
            ControlDesign("Notes", "TextBox", 10, 58, 120, 40, tab_index=2, multi_line=True, word_wrap=True, enter_key_behavior=True, scroll_bars=2),
            ControlDesign("Fruit", "ComboBox", 140, 10, 100, 18, tab_index=3, style=2, row_source="Sheet1!D1:D3", match_entry=1, list_rows=5),
            ControlDesign("Picks", "ListBox", 140, 34, 100, 60, tab_index=4, multi_select=1, column_count=2, bound_column=2, match_entry=0, column_heads=True),
            ControlDesign("Rush", "CheckBox", 250, 10, 100, 18, tab_index=5, caption="Rush order", value="1", back_color=0x8000000F, fore_color=0x80000012, alignment=0),
            ControlDesign("Maybe", "CheckBox", 250, 30, 100, 18, tab_index=6, caption="Maybe", value="", triple_state=True, back_color=0x8000000F, fore_color=0x80000012),
            ControlDesign("Small", "OptionButton", 250, 50, 80, 18, tab_index=7, caption="Small", value="0", group_name="Size", back_color=0x8000000F, fore_color=0x80000012),
            ControlDesign("Large", "OptionButton", 250, 70, 80, 18, tab_index=8, caption="Large", value="1", group_name="Size", back_color=0x8000000F, fore_color=0x80000012),
            ControlDesign("Bold", "ToggleButton", 340, 50, 30, 30, tab_index=9, caption="B", value="0", text_align=2, back_color=0x8000000F, fore_color=0x80000012),
            ControlDesign("OK", "CommandButton", 250, 250, 72, 24, tab_index=10, caption="OK", default=True, text_align=2, font=FontDesign("Arial", 10, bold=True)),
            ControlDesign("Cancel", "CommandButton", 330, 250, 72, 24, tab_index=11, caption="Cancel", cancel=True, text_align=2, accelerator="C"),
            ControlDesign("Hint", "Label", 10, 110, 200, 30, tab_index=12, tab_stop=False, caption="Fill in the form", word_wrap=True, fore_color=0x000000FF, back_style_opaque=False, text_align=2),
            ControlDesign("Box", "Frame", 10, 150, 150, 80, tab_index=13, caption="Options", special_effect=3, children=(
                ControlDesign("Gift", "CheckBox", 6, 12, 100, 18, tab_index=0, caption="Gift wrap", value="0", back_color=0x8000000F, fore_color=0x80000012),
            )),
            # Pages are as large as Excel makes them in a 160 x 120 MultiPage.
            ControlDesign("Book", "MultiPage", 170, 110, 160, 120, tab_index=14, position=1, children=(
                ControlDesign("First", "Page", 1.5, 15.76, 156.98, 102.73, caption="First", children=(
                    ControlDesign("Qty", "TextBox", 6, 6, 50, 18, tab_index=0, value="1"),
                )),
                ControlDesign("Second", "Page", 1.5, 15.76, 156.98, 102.73, tab_index=1, caption="Second", children=(
                    ControlDesign("Go", "CommandButton", 6, 6, 60, 24, tab_index=0, caption="Go", text_align=2),
                )),
            )),
            ControlDesign("Tabs", "TabStrip", 10, 240, 150, 40, tab_index=15, tabs=("One", "Two", "Three"), position=2),
            ControlDesign("Magnify", "ScrollBar", 340, 110, 12.76, 63.78, tab_index=16, min=10, max=400, position=100, small_change=5, large_change=50, orientation=0),
            ControlDesign("Count", "SpinButton", 360, 110, 12.76, 25.51, tab_index=17, min=1, max=10, position=3),
            ControlDesign("Logo", "Image", 340, 180, 40, 40, tab_index=18, tab_stop=False, picture=PNG, picture_size_mode=3),
        ),
    )


ORDER = FormDesign("OrderForm", caption="Order", width=240, height=150, controls=(
    new_control("TextBox1", "TextBox", left=10, top=10, width=120, tab_index=0),
    new_control("ComboBox1", "ComboBox", left=10, top=34, width=120, tab_index=1),
    new_control("CheckBox1", "CheckBox", left=10, top=58, width=120, tab_index=2, caption="Rush"),
    new_control("CommandButton1", "CommandButton", left=10, top=110, tab_index=3, caption="OK", default=True),
    new_control("CommandButton2", "CommandButton", left=90, top=110, tab_index=4, caption="Cancel", cancel=True),
))


def form_files(design: FormDesign, code: str = ""):
    """The .frm text and .frx bytes the VBA editor would export for ``design``."""
    items = write_design(design)
    text = vbframe(design).replace("End\r\n", f'   OleObjectBlob   =   "{design.name}.frx":0000\r\nEnd\r\n', 1)
    text += f'Attribute VB_Name = "{design.name}"\r\n' + code.replace("\n", "\r\n")
    return text, write_frx(items, design)


def add_form(s, design: FormDesign, code: str = ""):
    text, frx = form_files(design, code)
    return s.run("macro.importModule", fileName=design.name + ".frm", text=text, frx=base64.b64encode(frx).decode())


# ---------------------------------------------------------------- forms, as Excel describes them

FORMS_MODULE = '''Dim R As Long

Sub Emit(ParamArray parts())
    Dim i As Long, s As String
    For i = 0 To UBound(parts)
        If i > 0 Then s = s & "|"
        s = s & parts(i)
    Next
    R = R + 1
    Worksheets(1).Cells(R, 1).Value = "'" & s
End Sub

Function Num(x) As String
    Num = Format(x, "0.00")
End Function

Function Nz(v) As String
    If IsNull(v) Then Nz = "Null" Else Nz = CStr(v)
End Function

Function Extra(c As Object) As String
    Dim t As String
    On Error Resume Next
    Select Case TypeName(c)
    Case "TextBox"
        t = c.Text & ";" & c.MaxLength & ";" & c.MultiLine & ";" & c.Locked & ";" & c.PasswordChar & ";" & c.ControlSource & ";" & c.TextAlign
    Case "ComboBox"
        t = c.Text & ";" & c.Style & ";" & c.ListCount & ";" & c.MatchEntry & ";" & c.ListRows & ";" & c.RowSource
    Case "ListBox"
        t = c.ListCount & ";" & c.MultiSelect & ";" & c.ColumnCount & ";" & c.BoundColumn & ";" & c.ColumnHeads & ";" & c.MatchEntry
    Case "CheckBox"
        t = c.Caption & ";" & Nz(c.Value) & ";" & c.TripleState & ";" & c.Alignment
    Case "OptionButton"
        t = c.Caption & ";" & Nz(c.Value) & ";" & c.GroupName
    Case "ToggleButton"
        t = c.Caption & ";" & Nz(c.Value)
    Case "CommandButton"
        t = c.Caption & ";" & c.Default & ";" & c.Cancel & ";" & c.Accelerator & ";" & c.Font.Name & ";" & c.Font.Size & ";" & c.Font.Bold
    Case "Label"
        t = c.Caption & ";" & c.WordWrap & ";" & c.ForeColor & ";" & c.BackStyle & ";" & c.TextAlign
    Case "Frame"
        t = c.Caption & ";" & c.Controls.Count & ";" & c.SpecialEffect
    Case "MultiPage"
        t = c.Value & ";" & c.Pages.Count & ";" & c.Pages(0).Caption & ";" & c.Pages(1).Caption & ";" & c.Pages(1).Name & ";" & c.Pages(0).Controls.Count
    Case "TabStrip"
        t = c.Value & ";" & c.Tabs.Count & ";" & c.Tabs(0).Caption & ";" & c.Tabs(c.Tabs.Count - 1).Caption
    Case "ScrollBar"
        t = c.Min & ";" & c.Max & ";" & c.Value & ";" & c.SmallChange & ";" & c.LargeChange & ";" & c.Orientation
    Case "SpinButton"
        t = c.Min & ";" & c.Max & ";" & c.Value & ";" & c.SmallChange
    Case "Image"
        t = c.PictureSizeMode & ";" & c.BorderStyle
    End Select
    If Err.Number <> 0 Then t = t & "#E" & Err.Number
    Extra = t
End Function

Sub Describe(f As Object)
    Dim c As Object
    Emit "form", f.Name, f.Caption, f.Controls.Count, f.StartUpPosition
    For Each c In f.Controls
        Emit c.Name, TypeName(c), Num(c.Left), Num(c.Top), Num(c.Width), Num(c.Height), c.Visible, c.Enabled, c.TabIndex, c.ControlTipText, Extra(c)
    Next
End Sub

Sub RunAll()
    R = 0
    Load Everything
    Describe Everything
    Unload Everything
    Load OrderForm
    Describe OrderForm
    Unload OrderForm
    Emit "done"
End Sub
'''



EVENTS_FORM_CODE = '''Private Sub UserForm_Initialize(): Note "Initialize": End Sub
Private Sub UserForm_Activate(): Note "Activate": End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer): Note "QueryClose" & CloseMode: End Sub
Private Sub UserForm_Terminate(): Note "Terminate": End Sub
Private Sub Text_Change(): Note "Text.Change": End Sub
Private Sub Check_Change(): Note "Check.Change": End Sub
Private Sub Check_Click(): Note "Check.Click": End Sub
Private Sub OptA_Change(): Note "OptA.Change": End Sub
Private Sub OptA_Click(): Note "OptA.Click": End Sub
Private Sub OptB_Change(): Note "OptB.Change": End Sub
Private Sub OptB_Click(): Note "OptB.Click": End Sub
Private Sub Toggle_Change(): Note "Toggle.Change": End Sub
Private Sub Toggle_Click(): Note "Toggle.Click": End Sub
Private Sub List_Change(): Note "List.Change": End Sub
Private Sub List_Click(): Note "List.Click": End Sub
Private Sub Combo_Change(): Note "Combo.Change": End Sub
Private Sub Combo_Click(): Note "Combo.Click": End Sub
Private Sub Scroll_Change(): Note "Scroll.Change": End Sub
Private Sub Spin_Change(): Note "Spin.Change": End Sub
Private Sub Pages_Change(): Note "Pages.Change": End Sub
Private Sub Strip_Change(): Note "Strip.Change": End Sub
Private Sub Strip_Click(ByVal Index As Long): Note "Strip.Click" & Index: End Sub
'''

EVENTS_MODULE = '''Public Log As String

Sub Note(t)
    Log = Log & t & ";"
End Sub

Sub Mark(t)
    Static r As Long
    r = r + 1
    Worksheets(1).Cells(r, 1).Value = "'" & t & "|" & Log
    Log = ""
End Sub

Sub RunAll()
    Load Events
    Mark "load"
    Events.Text.Text = "a"
    Mark "text"
    Events.Text.Value = "a"
    Mark "same text"
    Events.Check.Value = True
    Mark "check"
    Events.Check.Value = True
    Mark "same check"
    Events.OptB.Value = True
    Mark "option"
    Events.Toggle.Value = True
    Mark "toggle"
    Events.List.AddItem "x"
    Events.List.AddItem "y"
    Mark "additem"
    Events.List.ListIndex = 1
    Mark "list"
    Events.List.ListIndex = 1
    Mark "same list"
    Events.Combo.AddItem "p"
    Events.Combo.AddItem "q"
    Events.Combo.ListIndex = 1
    Mark "combo"
    Events.Combo.Text = "zz"
    Mark "combo text"
    Events.Combo.Clear
    Mark "combo clear"
    Events.Scroll.Value = 5
    Mark "scroll"
    Events.Spin.Value = 3
    Mark "spin"
    Events.Pages.Value = 1
    Mark "pages"
    Events.Strip.Value = 1
    Mark "strip"
    Events.Show vbModeless
    Mark "show"
    Events.Hide
    Mark "hide"
    Unload Events
    Mark "unload"
    Worksheets(1).Cells(99, 1).Value = "'done"
End Sub
'''


def events_design():
    return FormDesign("Events", caption="Events", width=300, height=200, show_modal=False, controls=(
        new_control("Text", "TextBox", left=6, top=6),
        new_control("Check", "CheckBox", left=6, top=30),
        new_control("OptA", "OptionButton", left=6, top=50, value="1"),
        new_control("OptB", "OptionButton", left=6, top=70),
        new_control("Toggle", "ToggleButton", left=120, top=6),
        new_control("List", "ListBox", left=120, top=36, height=40),
        new_control("Combo", "ComboBox", left=120, top=80),
        new_control("Scroll", "ScrollBar", left=210, top=6),
        new_control("Spin", "SpinButton", left=230, top=6),
        new_control("Pages", "MultiPage", left=6, top=110, width=120, height=80),
        new_control("Strip", "TabStrip", left=140, top=110, width=120, height=60),
    ))


def forms_workbook() -> Session:
    """Two forms ReflexGrid wrote (every kind of control; an order form imported as .frm/.frx)
    and the module describing them."""
    s = Session()
    s.new_workbook()
    add_form(s, every_control())
    add_form(s, ORDER)
    s.run("macro.setModule", name="Module1", code=FORMS_MODULE)
    return s


def events_workbook() -> Session:
    s = Session()
    s.new_workbook()
    add_form(s, events_design(), EVENTS_FORM_CODE)
    s.run("macro.setModule", name="Module1", code=EVENTS_MODULE)
    return s


def results(s: Session) -> list:
    """Column A of the first sheet, down to the first empty cell."""
    out = []
    for i in range(1, 200):
        v = s.run("cell.get", cell=f"A{i}")["value"]
        if v is None:
            break
        out.append(str(v))
    return out


# ---------------------------------------------------------------- a form made with the designer

DESIGNER_OPS = [
    [{"op": "add", "kind": k, "left": 6 + (i % 3) * 110, "top": 6 + (i // 3) * 40} for i, k in enumerate(
        ["Label", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton", "CommandButton", "TabStrip", "ScrollBar", "SpinButton", "Image"])]
    + [{"op": "set", "control": "", "key": "width", "value": 360}, {"op": "set", "control": "", "key": "height", "value": 330},
       {"op": "add", "kind": "Frame", "left": 6, "top": 170, "width": 160, "height": 70},
       {"op": "add", "kind": "MultiPage", "left": 180, "top": 170, "width": 160, "height": 120}],
    [
        {"op": "add", "kind": "CheckBox", "parent": "Frame1", "left": 6, "top": 12},
        {"op": "add", "kind": "TextBox", "parent": "MultiPage1", "left": 6, "top": 6},
        {"op": "set", "control": "CommandButton1", "key": "caption", "value": "Go"},
        {"op": "set", "control": "CommandButton1", "key": "default", "value": True},
        {"op": "set", "control": "TextBox1", "key": "value", "value": "hello"},
        {"op": "set", "control": "TextBox1", "key": "max_length", "value": "12"},
        {"op": "set", "control": "ComboBox1", "key": "style", "value": "2"},
        {"op": "set", "control": "ListBox1", "key": "multi_select", "value": "1"},
        {"op": "set", "control": "CheckBox1", "key": "triple_state", "value": True},
        {"op": "set", "control": "ScrollBar1", "key": "max", "value": "500"},
        {"op": "set", "control": "Label1", "key": "font.bold", "value": True},
        {"op": "set", "control": "Page2", "key": "caption", "value": "Second"},
        {"op": "set", "control": "MultiPage1", "key": "position", "value": 1},
        {"op": "set", "control": "", "key": "caption", "value": "Designed"},
    ],
    [
        {"op": "geometry", "control": "TextBox1", "left": 160, "top": 70},
        {"op": "add", "kind": "OptionButton", "parent": "Frame1", "left": 6, "top": 36},
        {"op": "delete", "control": "ToggleButton1"},
    ],
]


def designer_workbook() -> Session:
    """A form made in ReflexGrid's designer (Insert › UserForm, then the designer's changes) and
    the module describing it."""
    s = Session()
    s.new_workbook()
    s.run("macro.addModule", kind="form")
    for ops in DESIGNER_OPS:
        s.run("macro.editForm", name="UserForm1", ops=ops)
    s.run("macro.setModule", name="Module1", code=FORMS_MODULE.replace("Everything", "UserForm1").split("    Load OrderForm")[0] + '    Emit "done"\nEnd Sub\n')
    return s
