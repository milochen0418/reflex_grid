"""UserForms: designs read from and written to [MS-OFORMS] storages, .frm/.frx files and .xlsm
packages; the Forms object model VBA code uses (UserForm, its controls, Controls, Pages…); the
form's events and life cycle (Initialize, Show, Hide, Unload, QueryClose, Terminate; Change,
Click, Exit, KeyPress…); modal forms pausing a run like MsgBox, modeless ones staying open with
their events running as macros; and each person's forms on a shared workbook."""

import base64
import dataclasses
import io
import json
import zipfile
from pathlib import Path

import pytest

from gridcore.engine import CollabSession, DocState, EngineError, Session, SharedBook, open_bytes
from gridcore.model.forms import ControlDesign, FormDesign, new_control
from gridcore.model.workbook import Workbook
from gridcore.xlsx.oforms import MORPH_DATA, VBFRAME, control_record, design_streams, read_design, read_record, write_design, write_record
from tests.vba_form_cases import ORDER, add_form, designer_workbook, events_workbook, every_control, form_files, forms_workbook, results

EXPECTED_PATH = Path(__file__).parent / "vba_forms_expected.json"
EXPECTED = json.loads(EXPECTED_PATH.read_text("utf-8")) if EXPECTED_PATH.exists() else {}

def _same(a, b, path=""):
    """Field by field (the IDs Excel numbers are left out; sizes and positions are stored in
    HIMETRIC, a hundredth of a millimetre, so they come back within a hundredth of a point)."""
    assert type(a) is type(b)
    for f in dataclasses.fields(a):
        if f.name in ("id", "warnings"):
            continue
        x, y = getattr(a, f.name), getattr(b, f.name)
        if f.name in ("controls", "children"):
            assert [c.name for c in x] == [c.name for c in y], path
            for cx, cy in zip(x, y):
                _same(cx, cy, f"{path}/{cx.name}")
        elif f.name in ("left", "top", "width", "height"):
            assert abs(x - y) <= 0.02, f"{path or a.name}.{f.name}: {x!r} != {y!r}"
        else:
            assert x == y, f"{path or a.name}.{f.name}: {x!r} != {y!r}"


def new():
    s = Session()
    s.new_workbook()
    return s


def val(s, a):
    return s.run("cell.get", cell=a)["value"]


def macro(s, code, name="M", **p):
    s.run("macro.setModule", name="Module1", code=code)
    assert s.run("macro.list")["diagnostics"] == []
    return s.run("macro.run", name=name, **p)


# ---------------------------------------------------------------- the file format


def test_written_forms_read_back_as_designed():
    d = every_control()
    items = write_design(d)
    paths = [p for p, _ in items]
    assert {"/", "f", "o", "\x01CompObj", VBFRAME} <= set(paths)
    assert any(p.endswith("/x") for p in paths)  # the MultiPage's pages
    back = read_design("Everything", design_streams(items))
    assert back.warnings == ()
    _same(d, back)


@pytest.mark.parametrize("design,excel", [
    # Records Excel's designer wrote for new controls (its padding bytes zeroed).
    (new_control("TextBox1", "TextBox"),
     "0002140001010080000000001b48802cec0900007b020000000218003500000006000080a5000000000200005461686f6d610000"),
    (new_control("Label1", "Label"),
     "0002180028000000060000804c6162656c310000ec0900007b020000000218003500000006000080a5000000000200005461686f6d610000"),
    (new_control("CommandButton1", "CommandButton", height=24.01),
     "00022000280000000e000080436f6d6d616e64427574746f6e310000ec0900004f030000000218007500000006000080a5000000000203005461686f6d610000"),
    (new_control("CheckBox1", "CheckBox"),
     "000234004601c080000000000f00008012000080040000000100008009000080e20e00007b02000030000000436865636b426f7831000000000218003500000006000080a5000000000200005461686f6d610000"),
    (new_control("ComboBox1", "ComboBox"),
     "0002180041010580000000001b48802c03010200ec0900007b020000000218003500000006000080a5000000000200005461686f6d610000"),
    (new_control("SpinButton1", "SpinButton", width=12.76, height=25.51),
     "0002100008080000ffffffffc201000084030000"),
    (new_control("ListBox1", "ListBox"),
     "00021400600101800000000003020000ec090000ec090000000218003500000006000080a5000000000200005461686f6d610000"),
    (new_control("ScrollBar1", "ScrollBar", width=12.76, height=63.78),
     "0002100008200000ffffffffc2010000ca080000"),
    (new_control("OptionButton1", "OptionButton"),
     "000238004601c080000000000f0000801200008005000000010000800d000080e20e00007b020000300000004f7074696f6e427574746f6e31000000000218003500000006000080a5000000000200005461686f6d610000"),
    (new_control("ToggleButton1", "ToggleButton", width=36, height=40),
     "000238004601c080000000000f0000801200008006000000010000800d000080f60400008305000030000000546f67676c65427574746f6e31000000000218007500000006000080a5000000000203005461686f6d610000"),
])
def test_records_are_laid_out_as_excel_writes_them(design, excel):
    assert write_record(control_record(design)).hex() == excel
    spec = control_record(design).spec
    rec, _ = read_record(bytes.fromhex(excel), spec, "cp1252")
    assert write_record(rec).hex() == excel


def test_strings_are_compressed_when_they_fit_one_byte():
    for caption in ("Café", "地球"):
        d = ControlDesign("C", "CheckBox", caption=caption, value="0")
        rec, _ = read_record(write_record(control_record(d)), MORPH_DATA, "cp1252")
        assert rec.get("Caption") == caption
    assert b"Caf\xe9" in write_record(control_record(ControlDesign("C", "CheckBox", caption="Café")))


def test_frm_and_frx_files_import_and_export():
    s = new()
    text, frx = form_files(every_control(), "Private Sub OK_Click()\n    Me.Hide\nEnd Sub\n")
    assert frx[:4] == b"LB\x08\x00" and int.from_bytes(frx[4:8], "little") == len(frx) - 24
    assert s.run("macro.importModule", fileName="Everything.frm", text=text, frx=base64.b64encode(frx).decode()) == {"name": "Everything", "kind": "form"}
    with pytest.raises(EngineError):
        s.run("macro.importModule", fileName="Everything.frm", text=text)  # without its .frx
    out = s.run("macro.exportModule", name="Everything")
    assert out["fileName"] == "Everything.frm" and out["frxName"] == "Everything.frx"
    assert 'OleObjectBlob   =   "Everything.frx":0000' in out["text"] and "Private Sub OK_Click()" in out["text"]
    s2 = new()
    assert s2.run("macro.importModule", fileName=out["fileName"], text=out["text"], frx=out["frx"])["name"] == "Everything"
    p1, p2 = s.doc().wb.vba.module("Everything"), s2.doc().wb.vba.module("Everything")
    _same(read_design("Everything", design_streams(p1.designer)), read_design("Everything", design_streams(p2.designer)))
    assert p2.code == p1.code
    # Importing again picks a free name.
    assert s2.run("macro.importModule", fileName=out["fileName"], text=out["text"], frx=out["frx"])["name"] == "Everything1"
    assert s2.run("macro.formDesign", name="Everything1")["form"]["name"] == "Everything1"


def test_xlsm_files_keep_forms_and_their_code_can_be_edited():
    openpyxl = pytest.importorskip("openpyxl")
    s = new()
    add_form(s, ORDER, "Private Sub CommandButton1_Click()\n    Me.Hide\nEnd Sub\n")
    before = s.doc().wb.vba.module("OrderForm").designer
    s.run("macro.setModule", name="OrderForm", code="Private Sub CommandButton1_Click()\n    Unload Me\nEnd Sub\n")
    assert s.doc().wb.vba.module("OrderForm").designer == before
    data = base64.b64decode(s.run("file.saveBytes")["data"])
    s2 = Session()
    assert open_bytes(s2, data, "Book1.xlsm").warnings == []
    m = s2.doc().wb.vba.module("OrderForm")
    assert m.kind == "form" and "Unload Me" in m.code
    _same(ORDER, read_design("OrderForm", design_streams(m.designer)))
    wb = openpyxl.load_workbook(io.BytesIO(data), keep_vba=True)
    try:
        assert "xl/vbaProject.bin" in wb.vba_archive.namelist()
    finally:
        wb.vba_archive.close()
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        assert b"OrderForm" in z.read("xl/vbaProject.bin")


def test_a_form_whose_design_cant_be_read_still_has_its_code():
    s = new()
    add_form(s, ORDER, "Public Function Twice(x)\n    Twice = 2 * x\nEnd Function\n")

    def damage(ctx):
        p = ctx.wb.vba
        m = p.module("OrderForm")
        ctx.wb.vba = p.with_module(dataclasses.replace(m, designer=tuple((k, b"\x00\x04\xff\xff" if k == "f" else v) for k, v in m.designer)))

    from gridcore.engine.session import edit

    edit(s, damage)
    r = s.run("macro.formDesign", name="OrderForm")
    assert r["warnings"] and "can't be read" in r["warnings"][0]
    r = macro(s, "Sub M()\n    Range(\"A1\").Value = OrderForm.Twice(21)\nEnd Sub\n")
    assert r["status"] == "done" and val(s, "A1") == 42


# ---------------------------------------------------------------- running forms


def test_a_modal_form_runs_its_events_and_writes_the_workbook():
    s = new()
    add_form(s, ORDER, '''
Private Sub UserForm_Initialize()
    ComboBox1.AddItem "Apple"
    ComboBox1.AddItem "Pear"
    Log "Initialize"
End Sub
Private Sub UserForm_Activate()
    Log "Activate"
End Sub
Private Sub TextBox1_Change()
    Log "Change " & TextBox1.Text
End Sub
Private Sub CheckBox1_Click()
    Log "Click " & CheckBox1.Value
End Sub
Private Sub CommandButton1_Click()
    Range("A1").Value = TextBox1.Value
    Range("A2").Value = ComboBox1.Value
    Range("A3").Value = CheckBox1.Value
    Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Log "QueryClose " & CloseMode
End Sub
Private Sub UserForm_Terminate()
    Log "Terminate"
End Sub
Private Sub Log(t)
    Dim r As Long
    r = Range("C1").CurrentRegion.Rows.Count + 1
    If IsEmpty(Range("C1").Value) Then r = 1
    Cells(r, 3).Value = t
End Sub
''')
    r = macro(s, "Sub M()\n    OrderForm.Show\n    Range(\"A4\").Value = \"after\"\nEnd Sub\n", answers=[
        {"event": "change", "control": "TextBox1", "value": "hello"},
        {"event": "click", "control": "ComboBox1", "value": "Pear"},
        {"event": "click", "control": "CheckBox1"},
        {"event": "click", "control": "CommandButton1"},
    ])
    assert r["status"] == "done", r.get("error")
    assert [val(s, a) for a in ("A1", "A2", "A3", "A4")] == ["hello", "Pear", True, "after"]
    log = [val(s, f"C{i}") for i in range(1, 8)]
    assert log == ["Initialize", "Activate", "Change hello", "Click True", "QueryClose 1", "Terminate", None]
    assert r["forms"] == []  # unloaded
    s.run("edit.undo")
    assert val(s, "A1") is None and val(s, "C1") is None  # the whole run is one undo step


def test_hide_keeps_the_form_and_unload_makes_a_new_one():
    s = new()
    add_form(s, ORDER, "Public Count As Long\nPrivate Sub UserForm_Initialize()\n    Count = Count + 1\nEnd Sub\nPrivate Sub CommandButton1_Click()\n    Me.Hide\nEnd Sub\n")
    r = macro(s, '''
Sub M()
    OrderForm.TextBox1.Text = "kept"
    OrderForm.Show
    Range("A1").Value = OrderForm.TextBox1.Text
    Range("A2").Value = OrderForm.Count
    Unload OrderForm
    Range("A3").Value = OrderForm.TextBox1.Text
    Load OrderForm
    Range("A4").Value = OrderForm.Visible
End Sub''', answers=[{"event": "enter"}])
    assert r["status"] == "done", r.get("error")
    assert [val(s, a) for a in ("A1", "A2", "A3", "A4")] == ["kept", 1, None, False]


def test_closing_with_the_x_unloads_unless_query_close_cancels():
    s = new()
    add_form(s, ORDER, '''
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu And CheckBox1.Value = False Then Cancel = True
End Sub
Private Sub CommandButton2_Click()
    Unload Me
End Sub''')
    r = macro(s, '''
Sub M()
    OrderForm.Show
    Range("A1").Value = "closed"
End Sub''', answers=[{"event": "close"}, {"event": "click", "control": "CheckBox1"}, {"event": "close"}])
    assert r["status"] == "done" and val(s, "A1") == "closed"
    r = s.run("macro.run", name="M", answers=[{"event": "close"}, {"event": "escape"}])
    assert r["status"] == "done"


def test_new_instances_and_unloaded_forms():
    s = new()
    add_form(s, ORDER, "Public Tagged As String\nPrivate Sub CommandButton2_Click()\n    Unload Me\nEnd Sub\nPrivate Sub CommandButton1_Click()\n    Me.Hide\nEnd Sub\n")
    r = macro(s, '''
Sub M()
    Dim a As OrderForm, b As New OrderForm
    Set a = New OrderForm
    a.Caption = "A"
    b.Caption = "B"
    Range("A1").Value = a.Caption & b.Caption & OrderForm.Caption
    Range("A2").Value = TypeName(a) & " " & (TypeOf a Is OrderForm) & " " & (TypeOf a.TextBox1 Is MSForms.TextBox) & " " & (TypeOf a.TextBox1 Is Control)
    a.Show
    On Error Resume Next
    Range("A3").Value = a.TextBox1.Text
    Range("A4").Value = Err.Number
End Sub''', answers=[{"event": "escape"}])
    assert r["status"] == "done", r.get("error")
    assert val(s, "A1") == "ABOrder"
    assert val(s, "A2") == "OrderForm True True True"
    assert val(s, "A3") is None and val(s, "A4") == -2147418105  # the form was unloaded


def test_lists_and_their_cells():
    s = new()
    for i, v in enumerate(["Apple", "Pear", "Plum"], 1):
        s.run("cell.set", cell=f"D{i}", input=v)
        s.run("cell.set", cell=f"E{i}", input=str(i * 10))
    design = FormDesign("Lists", controls=(
        new_control("Bound", "ComboBox", row_source="Sheet1!D1:E3", column_count=2, bound_column=2, control_source="Sheet1!F1"),
        new_control("Many", "ListBox", multi_select=1),
        new_control("One", "ListBox"),
        new_control("Who", "TextBox", control_source="Sheet1!F2"),
        new_control("Done", "CommandButton", caption="Done"),
    ))
    add_form(s, design, '''
Private Sub UserForm_Initialize()
    Many.List = Array("a", "b", "c", "d")
    One.ColumnCount = 2
    One.AddItem "x"
    One.List(0, 1) = "first"
    One.AddItem "y"
    One.List(1, 1) = "second"
    One.AddItem "w", 0
End Sub
Private Sub Done_Click()
    Range("A1").Value = Bound.ListCount & " " & Bound.Value & " " & Bound.Text & " " & Bound.Column(0)
    Dim i As Long, t As String
    For i = 0 To Many.ListCount - 1
        If Many.Selected(i) Then t = t & Many.List(i)
    Next
    Range("A2").Value = t & " " & IsNull(Many.Value)
    Range("A3").Value = One.ListIndex & " " & One.Value & " " & One.Column(1) & " " & One.List(2, 0)
    One.RemoveItem 0
    Range("A4").Value = One.ListCount & " " & One.ListIndex
    Unload Me
End Sub''')
    r = macro(s, "Sub M()\n    Lists.Show\nEnd Sub\n", answers=[
        {"event": "click", "control": "Bound", "value": "Pear"},
        {"event": "click", "control": "Many", "value": [1, 3]},
        {"event": "click", "control": "One", "value": 1},
        {"event": "change", "control": "Who", "value": "Zoe"},
        {"event": "click", "control": "Done"},
    ])
    assert r["status"] == "done", r.get("error")
    assert val(s, "A1") == "3 20 Pear Pear"
    assert val(s, "A2") == "bd True"
    assert val(s, "A3") == "1 x first y"
    assert val(s, "A4") == "2 0"
    assert val(s, "F1") == 20 and val(s, "F2") == "Zoe"  # ControlSource cells


def test_toggles_numbers_and_pages():
    s = new()
    design = FormDesign("Bits", controls=(
        new_control("Frame1", "Frame", children=(
            new_control("A", "OptionButton", value="1"),
            new_control("B", "OptionButton", value="0"),
        )),
        new_control("C", "OptionButton", value="1"),
        new_control("D", "OptionButton", value="0", group_name="g"),
        new_control("E", "OptionButton", value="1", group_name="g"),
        new_control("Maybe", "CheckBox", value="", triple_state=True),
        new_control("Scroll", "ScrollBar", min=10, max=20, position=15),
        new_control("Spin", "SpinButton", min=0, max=3, position=3),
        new_control("Book", "MultiPage", children=(new_control("P1", "Page", caption="One"), new_control("P2", "Page", caption="Two"))),
    ))
    add_form(s, design, '''
Public Changes As String
Private Sub Book_Change()
    Changes = Changes & "page" & Book.Value & ";"
End Sub
Private Sub Spin_SpinUp()
    Changes = Changes & "up;"
End Sub
Private Sub Scroll_Change()
    Changes = Changes & "zoom" & Scroll.Value & ";"
End Sub
Private Sub UserForm_Click()
    Me.Hide
End Sub''')
    r = macro(s, '''
Sub M()
    With Bits
        Range("A1").Value = IsNull(.Maybe.Value) & " " & .A.Value & .B.Value & .C.Value & .D.Value & .E.Value
        .Show
        Range("A2").Value = .A.Value & " " & .B.Value & " " & .C.Value & " " & .D.Value & " " & .E.Value
        Range("A3").Value = .Spin.Value & " " & .Scroll.Value & " " & .Book.Value & " " & .Book.Pages.Count & " " & .Book.Pages(1).Caption & " " & .Book.SelectedItem.Name
        On Error Resume Next
        .Scroll.Value = 99
        Range("A4").Value = Err.Number
        Err.Clear
        Range("A5").Value = .Changes
    End With
End Sub''', answers=[
        {"event": "click", "control": "B"},
        {"event": "click", "control": "D"},
        {"event": "spin", "control": "Spin", "up": True},
        {"event": "scroll", "control": "Scroll", "value": 18},
        {"event": "click", "control": "Book", "value": 1},
        {"event": "click", "control": "Maybe", "value": True},
        {"event": "click"},  # on the form itself
    ])
    assert r["status"] == "done", r.get("error")
    assert val(s, "A1") == "True TrueFalseTrueFalseTrue"
    # B turned A off (same frame); C stays (another container); D turned E off (same group).
    assert val(s, "A2") == "False True True True False"
    assert val(s, "A3") == "3 18 1 2 Two P2"
    assert val(s, "A4") == 380
    assert val(s, "A5") == "up;zoom18;page1;"


def test_keys_focus_and_validation():
    s = new()
    design = FormDesign("Keys", controls=(
        new_control("Digits", "TextBox", tab_index=0),
        new_control("Age", "TextBox", tab_index=1),
        new_control("Save", "CommandButton", tab_index=2, caption="Save"),
    ))
    add_form(s, design, '''
Private Sub Digits_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Age_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNumeric(Age.Text) Then Cancel = True
End Sub
Private Sub Age_AfterUpdate()
    Range("B1").Value = "updated " & Age.Text
End Sub
Private Sub Save_Click()
    Range("A1").Value = Digits.Text & " " & Age.Text & " " & Me.ActiveControl.Name
    Unload Me
End Sub''')
    events = [{"event": "focus", "control": "Digits"}]
    events += [{"event": "keypress", "control": "Digits", "key": k, "selStart": 9} for k in "4x2"]
    events += [
        {"event": "focus", "control": "Age"},
        {"event": "change", "control": "Age", "value": "abc"},
        {"event": "click", "control": "Save"},  # Exit cancels: no click
        {"event": "change", "control": "Age", "value": "42"},
        {"event": "click", "control": "Save"},
    ]
    r = macro(s, "Sub M()\n    Keys.Show\nEnd Sub\n", answers=events)
    assert r["status"] == "done", r.get("error")
    assert val(s, "A1") == "42 42 Save"
    assert val(s, "B1") == "updated 42"


def test_controls_collection_and_dynamic_controls_with_events():
    s = new()
    s.run("macro.addModule", kind="class", name="ButtonHandler")
    s.run("macro.setModule", name="ButtonHandler", code="Public WithEvents Btn As MSForms.CommandButton\nPrivate Sub Btn_Click()\n    Range(\"A1\").Value = \"clicked \" & Btn.Caption\nEnd Sub\n")
    add_form(s, ORDER, '''
Private handlers As New Collection
Private Sub UserForm_Initialize()
    Dim b As MSForms.CommandButton, h As ButtonHandler
    Set b = Me.Controls.Add("Forms.CommandButton.1", "Extra")
    b.Caption = "More"
    b.Move 100, 60, 60, 24
    Set h = New ButtonHandler
    Set h.Btn = b
    handlers.Add h
    Dim c As Control, names As String
    For Each c In Me.Controls
        names = names & TypeName(c) & ":" & c.Name & " "
    Next
    Range("A2").Value = Trim(names)
    Range("A3").Value = Me.Controls.Count & " " & Me("TextBox1").Name & " " & Me!CheckBox1.Caption & " " & Controls(0).Name
End Sub
Private Sub CommandButton2_Click()
    Unload Me
End Sub''')
    r = macro(s, "Sub M()\n    OrderForm.Show\nEnd Sub\n", answers=[{"event": "click", "control": "Extra"}, {"event": "escape"}])
    assert r["status"] == "done", r.get("error")
    assert val(s, "A1") == "clicked More"
    assert val(s, "A2") == "TextBox:TextBox1 ComboBox:ComboBox1 CheckBox:CheckBox1 CommandButton:CommandButton1 CommandButton:CommandButton2 CommandButton:Extra"
    assert val(s, "A3") == "6 TextBox1 Rush TextBox1"


def test_forms_reach_their_snapshot_for_the_browser():
    s = new()
    s.run("cell.set", cell="B1", input="Ann")  # NameBox's ControlSource
    add_form(s, every_control(), "Private Sub OK_Click()\n    Me.Hide\nEnd Sub\nPrivate Sub NameBox_Change()\nEnd Sub\n")
    r = s.run("macro.showForm", name="Everything")
    assert r["status"] == "done"  # ShowModal is False in its design: the form stays open
    (f,) = r["forms"]
    assert f["caption"] == "All controls" and f["modal"] is False
    by = {c["name"]: c for c in f["controls"]}
    assert by["NameBox"]["text"] == "Ann" and by["NameBox"]["events"] == ["change"]
    assert by["OK"]["default"] is True and by["OK"]["events"] == ["click"]
    assert by["Fruit"]["style"] == 2 and by["Picks"]["multiSelect"] == 1
    assert by["Logo"]["picture"].startswith("data:image/png;base64,")
    assert [p["caption"] for p in by["Book"]["children"]] == ["First", "Second"] and by["Book"]["value"] == 1
    assert by["Hint"]["foreColor"] == "#FF0000" and by["Hint"]["opaque"] is False
    assert s.run("macro.forms")["forms"][0]["id"] == f["id"]
    d = s.run("macro.formDesign", name="Everything")
    assert d["warnings"] == [] and [c["name"] for c in d["form"]["controls"]] == [c["name"] for c in f["controls"]]


def test_a_modeless_form_runs_its_events_after_the_macro():
    s = new()
    add_form(s, ORDER, '''
Private Sub CommandButton1_Click()
    Range("A1").Value = TextBox1.Text
    If CheckBox1.Value Then MsgBox "Rush!"
End Sub
Private Sub CommandButton2_Click()
    Unload Me
End Sub''')
    r = macro(s, "Sub M()\n    OrderForm.Show vbModeless\n    Range(\"A2\").Value = \"macro ended\"\nEnd Sub\n")
    assert val(s, "A2") == "macro ended"
    (f,) = r["forms"]
    r = s.run("macro.formEvent", form=f["id"], event={"event": "click", "control": "CommandButton1", "values": {"TextBox1": "typed", "CheckBox1": True}})
    assert r["status"] == "done" and val(s, "A1") == "typed"
    assert [m["prompt"] for m in r["messages"]] == ["Rush!"]
    assert r["forms"][0]["controls"][0]["text"] == "typed"
    s.run("edit.undo")  # the event's run is its own undo step
    assert val(s, "A1") is None and val(s, "A2") == "macro ended"
    r = s.run("macro.formEvent", form=f["id"], event={"event": "escape"})
    assert r["forms"] == []
    # A form that is gone: nothing happens.
    assert s.run("macro.formEvent", form=f["id"], event={"event": "click", "control": "CommandButton1"})["forms"] == []


def test_a_modal_form_pauses_an_interactive_run():
    s = new()
    s.interactive_macros = True
    add_form(s, ORDER, '''
Private Sub CommandButton1_Click()
    If MsgBox("Save " & TextBox1.Text & "?", vbYesNo) = vbYes Then Range("A1").Value = TextBox1.Text
    Unload Me
End Sub''')
    r = macro(s, "Sub M()\n    Range(\"A2\").Value = \"before\"\n    OrderForm.Show\n    Range(\"A3\").Value = \"after\"\nEnd Sub\n")
    assert r["status"] == "prompt" and r["prompt"]["kind"] == "form"
    fid = r["prompt"]["formId"]
    assert [f["id"] for f in r["forms"]] == [fid] and r["forms"][0]["modal"] is True
    assert val(s, "A2") == "before"  # what ran before is in the workbook
    with pytest.raises(EngineError):
        s.run("macro.run", name="M")
    r = s.run("macro.formEvent", form=fid, event={"event": "change", "control": "TextBox1", "value": "x"})
    assert r["status"] == "prompt" and r["forms"][0]["controls"][0]["text"] == "x"
    r = s.run("macro.formEvent", form=fid, event={"event": "click", "control": "CommandButton1"})
    assert r["prompt"]["kind"] == "msgbox" and r["forms"][0]["id"] == fid  # the form stays under the MsgBox
    r = s.run("macro.answer", value=6)
    assert r["status"] == "done" and r["forms"] == []
    assert [val(s, a) for a in ("A1", "A2", "A3")] == ["x", "before", "after"]
    # Stopping the macro takes its modal form away.
    r = s.run("macro.run", name="M")
    assert r["status"] == "prompt"
    assert s.run("macro.stop")["forms"] == []
    assert s.run("macro.forms") == {"forms": [], "waiting": None}


def test_each_person_has_their_own_forms_on_a_shared_workbook():
    b = SharedBook(DocState(Workbook(), "Shared"), "b1")
    amy = b.join("amy", CollabSession(), "Amy")
    ben = b.join("ben", CollabSession(), "Ben")
    for v in (amy, ben):
        v.session.interactive_macros = True

    def act(v, cid, **p):
        with b.act(v) as s:
            return s.execute(cid, p)

    with b.act(amy) as s:
        add_form(s, ORDER, "Private Sub CommandButton1_Click()\n    Range(\"A1\").Value = Application.UserName & \": \" & TextBox1.Text\nEnd Sub\n")
    r = act(amy, "macro.showForm", name="OrderForm", modal=False)
    (f,) = r["forms"]
    assert act(ben, "macro.forms")["forms"] == []
    act(ben, "macro.showForm", name="OrderForm", modal=False)
    (g,) = act(ben, "macro.forms")["forms"]
    assert g["id"] != f["id"]
    act(ben, "macro.formEvent", form=g["id"], event={"event": "click", "control": "CommandButton1", "values": {"TextBox1": "b"}})
    act(amy, "macro.formEvent", form=f["id"], event={"event": "click", "control": "CommandButton1", "values": {"TextBox1": "a"}})
    assert act(ben, "cell.get", cell="A1")["value"] == "Amy: a"
    # Ben can't drive Amy's form: his session doesn't have it.
    r = act(ben, "macro.formEvent", form=f["id"], event={"event": "click", "control": "CommandButton1"})
    assert [x["id"] for x in r["forms"]] == [g["id"]]
    assert act(amy, "cell.get", cell="A1")["value"] == "Amy: a"
    # Changing the code resets the project: everybody's forms go away, as in Excel.
    act(amy, "macro.setModule", name="Module1", code="Sub X()\nEnd Sub\n")
    assert act(amy, "macro.forms")["forms"] == [] and act(ben, "macro.forms")["forms"] == []


def test_end_unloads_every_form_and_functions_cant_show_forms():
    s = new()
    add_form(s, ORDER, "")
    r = macro(s, '''
Sub M()
    OrderForm.Show vbModeless
End Sub
Sub Stop1()
    End
End Sub
Function F()
    OrderForm.Show
    F = 1
End Function''')
    assert len(r["forms"]) == 1
    assert s.run("macro.run", name="Stop1")["forms"] == []
    s.run("cell.set", cell="A1", input="=F()")
    assert s.run("cell.get", cell="A1")["text"] == "#VALUE!"


def test_userforms_collection_and_the_data_object():
    s = new()
    add_form(s, ORDER, "")
    r = macro(s, '''
Sub M()
    Dim f As Object, d As New MSForms.DataObject, e As Object
    Load OrderForm
    Set f = UserForms.Add("OrderForm")
    f.Caption = "Second"
    Range("A1").Value = UserForms.Count & " " & UserForms(1).Caption & " " & TypeName(UserForms(0))
    d.SetText "copied"
    d.PutInClipboard
    Set e = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    e.GetFromClipboard
    Range("A2").Value = e.GetText & " " & e.GetFormat(1)
    e.Clear
    On Error Resume Next
    Range("A3").Value = e.GetText
    Range("A4").Value = Err.Number
End Sub''')
    assert r["status"] == "done", r.get("error")
    assert val(s, "A1") == "2 Second OrderForm"
    assert val(s, "A2") == "copied True"
    assert val(s, "A3") is None and val(s, "A4") == -2147221404


# ---------------------------------------------------------------- Excel's answers


@pytest.mark.skipif(not EXPECTED, reason="Excel's answers not captured yet (scripts/excel_vba_oracle_mac.py --forms --save)")
def test_forms_as_excel_describes_them():
    """Excel loaded these forms (written by ReflexGrid) and described every control; so does
    ReflexGrid, from the same module."""
    s = forms_workbook()
    assert s.run("macro.run", name="RunAll")["status"] == "done"
    assert results(s) == EXPECTED["forms"]


@pytest.mark.skipif(not EXPECTED, reason="Excel's answers not captured yet (scripts/excel_vba_oracle_mac.py --form-events --save)")
def test_form_events_as_excel_fires_them():
    s = events_workbook()
    assert s.run("macro.run", name="RunAll")["status"] == "done"
    assert results(s) == EXPECTED["events"]


# ---------------------------------------------------------------- the form designer


def _designer(s, name):
    return dict(s.doc().wb.vba.module(name).designer)


def test_a_new_form_is_designed_control_by_control():
    s = new()
    assert s.run("macro.addModule", kind="form") == {"name": "UserForm1", "kind": "form"}
    view = s.run("macro.formDesign", name="UserForm1")
    assert view["editable"] and view["form"]["caption"] == "UserForm1" and view["form"]["controls"] == []
    assert [r["label"] for r in view["props"]][:4] == ["(Name)", "Caption", "Width", "Height"]
    r = s.run("macro.editForm", name="UserForm1", ops=[
        {"op": "add", "kind": "Frame", "left": 6, "top": 6, "width": 150, "height": 60},
        {"op": "add", "kind": "MultiPage", "left": 6, "top": 70},
        {"op": "add", "kind": "TextBox", "left": 170, "top": 6},
        {"op": "add", "kind": "CommandButton", "left": 170, "top": 30},
    ])
    assert r["added"] == ["Frame1", "MultiPage1", "TextBox1", "CommandButton1"] and r["control"] == "CommandButton1"
    r = s.run("macro.editForm", name="UserForm1", ops=[
        {"op": "add", "kind": "CheckBox", "parent": "Frame1", "left": 6, "top": 12},
        {"op": "add", "kind": "OptionButton", "parent": "MultiPage1", "left": 6, "top": 6},
        {"op": "set", "control": "CommandButton1", "key": "caption", "value": "OK"},
        {"op": "set", "control": "CommandButton1", "key": "default", "value": True},
        {"op": "set", "control": "TextBox1", "key": "font.bold", "value": "True"},
        {"op": "set", "control": "TextBox1", "key": "back_color", "value": "#FFFF00"},
        {"op": "set", "control": "", "key": "caption", "value": "Order"},
        {"op": "geometry", "control": "TextBox1", "left": 172, "top": 8, "width": 90, "height": 20},
        {"op": "set", "control": "Page2", "key": "caption", "value": "Extras"},
        {"op": "set", "control": "MultiPage1", "key": "position", "value": 1},
    ], control="TextBox1")
    d = read_design("UserForm1", design_streams(tuple(_designer(s, "UserForm1").items())))
    assert d.warnings == () and d.caption == "Order"
    assert [c.name for c in d.controls] == ["Frame1", "MultiPage1", "TextBox1", "CommandButton1"]
    frame, mp, box, ok = d.controls
    assert [c.name for c in frame.children] == ["CheckBox1"]
    assert [p.caption for p in mp.children] == ["Page1", "Extras"] and mp.position == 1
    assert [c.name for c in mp.children[0].children] == ["OptionButton1"]  # it went on the page on top then
    assert ok.caption == "OK" and ok.default and box.font.bold and box.back_color == 0x0000FFFF
    assert (box.left, box.top, box.width, box.height) == pytest.approx((172, 8, 90, 20), abs=0.02)
    ids = [c.id for c in d.walk()]
    assert len(ids) == len(set(ids))
    # The form works: code for the button, run, events.
    s.run("macro.setModule", name="UserForm1", code='Private Sub CommandButton1_Click()\n    Range("A1").Value = TextBox1.Text & "/" & CheckBox1.Value\n    Unload Me\nEnd Sub\n')
    r = macro(s, "Sub M()\n    UserForm1.Show\nEnd Sub\n", answers=[
        {"event": "change", "control": "TextBox1", "value": "hi"},
        {"event": "click", "control": "CheckBox1"},
        {"event": "enter"},
    ])
    assert r["status"] == "done" and val(s, "A1") == "hi/True"
    # Removing controls (a container with what is in it), and undo.
    r = s.run("macro.editForm", name="UserForm1", ops=[{"op": "delete", "control": "Frame1"}])
    assert [c["name"] for c in r["form"]["controls"]] == ["MultiPage1", "TextBox1", "CommandButton1"]
    assert not any(p.startswith(f"i{frame.id:02d}/") for p in _designer(s, "UserForm1"))
    s.run("edit.undo")
    assert [c["name"] for c in s.run("macro.formDesign", name="UserForm1")["form"]["controls"]][0] == "Frame1"


def test_designer_changes_leave_the_rest_of_a_form_as_it_was():
    """Editing one control rewrites its record and site only, keeping fields ReflexGrid doesn't
    model (here a mouse pointer and an IME mode, as Excel stores them)."""
    from gridcore.xlsx.oforms import CLASSES, COMMAND_BUTTON, read_form_stream
    from gridcore.xlsx.oforms_edit import FormEditor

    items = dict(write_design(ORDER))
    # What Excel may store that the design doesn't model: CommandButton1 with MousePointer 99 and
    # an IME mode in its VariousPropertyBits.
    fs = read_form_stream(items["f"], "cp1252")
    o = items["o"]
    off = 0
    spans = {}
    for site in fs.sites:
        spans[site.name] = (off, off + site.object_stream_size)
        off += site.object_stream_size
    a, b = spans["CommandButton1"]
    rec, _ = read_record(o[a:b], COMMAND_BUTTON, "cp1252")
    rec.mask |= (1 << 6) | (1 << 2)
    rec.values["MousePointer"] = 99
    rec.values["VariousPropertyBits"] = 0x1B | (3 << 15)
    data = write_record(rec)
    items["o"] = o[:a] + data + o[b:]
    # (the site's size must follow)
    from gridcore.xlsx.oforms_edit import Level

    lv = Level(items["f"], "cp1252")
    i = [se.site.name for se in lv.sites].index("CommandButton1")
    lv.sites[i].site.object_stream_size = len(data)
    lv.sites[i].dirty = True
    items["f"] = lv.to_bytes()
    before = tuple(sorted(items.items()))

    ed = FormEditor("OrderForm", before)
    assert ed.items() == before  # nothing changed: byte for byte
    ed.set("CommandButton1", {"caption": "Go", "left": 20.0})
    ed.set("TextBox1", {"width": 150.0})
    after = dict(ed.items())
    rec2, _ = read_record(_record(after, "CommandButton1"), COMMAND_BUTTON, "cp1252")
    assert rec2.values["MousePointer"] == 99 and rec2.values["VariousPropertyBits"] & (3 << 15) == 3 << 15
    assert rec2.values["Caption"] == "Go"
    # The other controls' records are untouched.
    for name in ("ComboBox1", "CheckBox1", "CommandButton2"):
        assert _record(after, name) == _record(dict(before), name)
    d = read_design("OrderForm", design_streams(tuple(after.items())))
    assert d.find("CommandButton1").left == pytest.approx(20, abs=0.02) and d.find("TextBox1").width == pytest.approx(150, abs=0.02)


def _record(items, name):
    from gridcore.xlsx.oforms import CONTAINERS, read_form_stream

    fs = read_form_stream(items["f"], "cp1252")
    off = 0
    for site in fs.sites:
        if site.cache_index in CONTAINERS:
            continue
        if site.name == name:
            return items["o"][off:off + site.object_stream_size]
        off += site.object_stream_size
    raise KeyError(name)


def test_designer_refuses_what_excel_would():
    s = new()
    s.run("macro.addModule", kind="form")
    s.run("macro.editForm", name="UserForm1", ops=[{"op": "add", "kind": "TextBox"}, {"op": "add", "kind": "Label"}])
    for op in ({"op": "set", "control": "Label1", "key": "name", "value": "TextBox1"},
               {"op": "set", "control": "Label1", "key": "name", "value": "1st"},
               {"op": "set", "control": "Label1", "key": "width", "value": "0"},
               {"op": "set", "control": "Label1", "key": "multi_line", "value": True},
               {"op": "add", "kind": "Calendar"},
               {"op": "add", "kind": "TextBox", "parent": "TextBox1"},
               {"op": "delete", "control": "Nothing"}):
        with pytest.raises(EngineError):
            s.run("macro.editForm", name="UserForm1", ops=[op])
    # A property set to what it is changes nothing (and makes no undo step).
    n = len(s.doc().undo)
    s.run("macro.editForm", name="UserForm1", ops=[{"op": "set", "control": "Label1", "key": "caption", "value": "Label1"}])
    assert len(s.doc().undo) == n


@pytest.mark.skipif("designer" not in EXPECTED, reason="Excel's answers not captured yet (scripts/excel_vba_oracle_mac.py --form-designer --save)")
def test_a_designed_form_as_excel_describes_it():
    """A form made with the designer (every kind of control, in a Frame and on a page too,
    properties changed, one moved, one removed): Excel loads it and describes it as ReflexGrid."""
    s = designer_workbook()
    assert s.run("macro.run", name="RunAll")["status"] == "done"
    assert results(s) == EXPECTED["designer"]


# ---------------------------------------------------------------- renaming a form


def test_a_renamed_form_keeps_its_design_runs_and_saves_under_the_new_name():
    """Rename in the toolbar, or (Name) in the Properties window: the module, its VB_Name, its
    design's header and its storage in the .xlsm take the new name; code naming the form isn't
    rewritten, as in Excel."""
    begin = "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrder \r\n"

    def has(session, name):
        return session.doc().wb.vba.module(name) is not None

    s = new()
    add_form(s, ORDER, 'Private Sub CommandButton1_Click()\n    Range("A1").Value = "clicked"\n    Unload Me\nEnd Sub\n')
    assert s.run("macro.renameModule", name="OrderForm", newName="frmOrder") == {"name": "frmOrder"}
    m = s.doc().wb.vba.module("frmOrder")
    assert m is not None and m.kind == "form" and not has(s, "OrderForm")
    assert 'Attribute VB_Name = "frmOrder"' in m.attributes
    header = dict(m.designer)[VBFRAME].decode("cp1252")
    assert begin in header and "OrderForm" not in header
    _same(dataclasses.replace(ORDER, name="frmOrder"), read_design("frmOrder", design_streams(m.designer)))
    exported = s.run("macro.exportModule", name="frmOrder")
    assert (exported["fileName"], exported["frxName"]) == ("frmOrder.frm", "frmOrder.frx")
    assert begin in exported["text"] and 'Attribute VB_Name = "frmOrder"' in exported["text"]
    # It runs under its new name.
    r = macro(s, "Sub M()\n    frmOrder.Show\nEnd Sub\n", answers=[{"event": "click", "control": "CommandButton1"}])
    assert r["status"] == "done", r.get("error")
    assert val(s, "A1") == "clicked"
    # Saved and opened again: the form's storage is under the new name, the old one is gone.
    data = base64.b64decode(s.run("file.saveBytes")["data"])
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        raw = z.read("xl/vbaProject.bin")
    assert b"frmOrder" in raw and b"OrderForm" not in raw
    s2 = Session()
    assert open_bytes(s2, data, "Book1.xlsm").warnings == []
    again = s2.doc().wb.vba.module("frmOrder")
    assert again.kind == "form" and not has(s2, "OrderForm")
    _same(dataclasses.replace(ORDER, name="frmOrder"), read_design("frmOrder", design_streams(again.designer)))
    # The Properties window's (Name) does the same, through macro.editForm.
    r = s.run("macro.editForm", name="frmOrder", ops=[{"op": "set", "key": "name", "value": "frmLast"}])
    assert r["name"] == "frmLast" and has(s, "frmLast") and not has(s, "frmOrder")
    name_row = next(row for row in r["props"] if row["key"] == "name")
    assert (name_row["type"], name_row["value"]) == ("text", "frmLast")
    # Refused: a taken name, a bad name, a sheet module; undo takes a rename back.
    with pytest.raises(EngineError, match="already exists"):
        s.run("macro.renameModule", name="frmLast", newName="Module1")
    with pytest.raises(EngineError, match="start with a letter"):
        s.run("macro.editForm", name="frmLast", ops=[{"op": "set", "key": "name", "value": "1st"}])
    with pytest.raises(EngineError, match="can't be renamed"):
        s.run("macro.renameModule", name="Sheet1", newName="Data")
    s.run("edit.undo")
    assert has(s, "frmOrder") and not has(s, "frmLast")
