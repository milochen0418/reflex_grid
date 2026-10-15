"""Excel macros on the engine: the macro.* commands, the Excel object model VBA code uses, worksheet
functions written in VBA, event macros, dialogs, limits, undo (alone and on shared workbooks),
.xlsm files, and the sandbox."""

import base64
import io

import pytest

from gridcore.engine import CollabSession, DocState, EngineError, Session, SharedBook, open_bytes
from gridcore.model.workbook import Workbook


def new():
    s = Session()
    s.new_workbook()
    return s


def val(s, a):
    return s.run("cell.get", cell=a)["value"]


def text(s, a):
    return s.run("cell.get", cell=a)["text"]


def run(s, code, macro="M", **p):
    s.run("macro.setModule", name="Module1", code=code)
    assert s.run("macro.list")["diagnostics"] == []
    r = s.run("macro.run", name=macro, **p)
    assert r["status"] == "done", r.get("error")
    return r["output"]


# ---------------------------------------------------------------- the project


def test_a_new_project_has_excels_modules():
    s = new()
    r = s.run("macro.list")
    assert not r["hasProject"]
    s.run("macro.setModule", name="Module1", code="Sub Hello()\n    Range(\"A1\").Value = 42\nEnd Sub\n")
    r = s.run("macro.list")
    assert [(m["name"], m["kind"], m["sheet"]) for m in r["modules"]] == [("ThisWorkbook", "document", "(workbook)"), ("Sheet1", "document", "Sheet1"), ("Module1", "standard", "")]
    assert [m["name"] for m in r["macros"]] == ["Hello"]
    assert s.doc().wb.code_name == "ThisWorkbook" and s.doc().wb.sheets[0].code_name == "Sheet1"
    s.run("macro.run", name="Hello")
    assert val(s, "A1") == 42


def test_editing_code_is_an_undoable_edit():
    s = new()
    s.run("macro.setModule", name="Module1", code="Sub A()\nEnd Sub\n")
    v1 = s.run("macro.getModule", name="Module1")["version"]
    s.run("macro.setModule", name="Module1", code="Sub B()\nEnd Sub\n", version=v1)
    with pytest.raises(EngineError, match="changed"):
        s.run("macro.setModule", name="Module1", code="Sub C()\nEnd Sub\n", version=v1)
    s.run("edit.undo")
    assert "Sub A()" in s.run("macro.getModule", name="Module1")["code"]
    s.run("edit.undo")
    assert not s.run("macro.list")["hasProject"]


def test_modules_are_added_renamed_exported_and_imported():
    s = new()
    s.run("macro.addModule", kind="class", name="Counter", code="Private n As Long\nPublic Sub Add()\n    n = n + 1\nEnd Sub\nPublic Property Get Value() As Long\n    Value = n\nEnd Property\n")
    out = run(s, "Sub M()\n    Dim c As New Counter\n    c.Add: c.Add\n    Debug.Print c.Value\nEnd Sub\n")
    assert out == [" 2 "]
    exported = s.run("macro.exportModule", name="Counter")
    assert exported["fileName"] == "Counter.cls"
    assert exported["text"].startswith("VERSION 1.0 CLASS\r\nBEGIN\r\n") and "\r\nAttribute VB_Name = \"Counter\"\r\n" in exported["text"]
    assert "\n" not in exported["text"].replace("\r\n", "")
    s.run("macro.renameModule", name="Counter", newName="Tally")
    with pytest.raises(EngineError):
        s.run("macro.renameModule", name="Sheet1", newName="Other")
    with pytest.raises(EngineError):
        s.run("macro.deleteModule", name="ThisWorkbook")
    s.run("macro.deleteModule", name="Tally")
    r = s.run("macro.importModule", fileName="Counter.cls", text=exported["text"])
    assert r == {"name": "Counter", "kind": "class"}
    bas = 'Attribute VB_Name = "Helpers"\r\nPublic Function Twice(x)\r\n    Twice = x * 2\r\nEnd Function\r\n'
    assert s.run("macro.importModule", fileName="whatever.bas", text=bas)["name"] == "Helpers"
    assert s.run("macro.importModule", fileName="Helpers.bas", text=bas)["name"] == "Helpers1"
    s.run("cell.set", cell="A1", input="=Twice(4)")
    assert val(s, "A1") == 8
    with pytest.raises(EngineError):
        s.run("macro.importModule", fileName="UserForm1.frm", text="VERSION 5.00\r\n")


def test_compile_errors_are_reported_per_module():
    s = new()
    s.run("macro.setModule", name="Module1", code="Sub Good()\nEnd Sub\n")
    s.run("macro.addModule", name="Broken", code="Sub Bad()\n    If x Then\nEnd Sub\n")
    d = s.run("macro.list")["diagnostics"]
    assert len(d) == 1 and d[0]["module"] == "Broken" and d[0]["line"] >= 2
    r = s.run("macro.run", name="Good")
    assert r["status"] == "done"
    s.run("macro.setModule", name="Module1", code="Sub Good()\n    Bad\nEnd Sub\n")
    assert s.run("macro.run", name="Good")["status"] == "error"


def test_removing_the_project():
    s = new()
    s.run("macro.setModule", name="Module1", code="Sub A()\nEnd Sub\n")
    s.run("macro.removeAll")
    assert not s.run("macro.list")["hasProject"]
    assert s.run("file.saveBytes")["name"].endswith(".xlsx")


# ---------------------------------------------------------------- running


def test_cells_values_and_formulas():
    s = new()
    out = run(s, '''
Sub M()
    Dim i As Long
    For i = 1 To 5
        Cells(i, 1).Value = i * i
    Next
    Range("B1").Formula = "=SUM(A1:A5)"
    Range("C1:C2").Value = Range("A1:A2").Value
    Range("D1").Value = DateSerial(2024, 1, 15)
    Range("E1").Value = "'007"
    Debug.Print Range("B1").Value, Cells(Rows.Count, 1).End(xlUp).Row, Range("D1").Text, TypeName(Range("D1").Value)
    Debug.Print WorksheetFunction.Sum(Range("A1:A5")), Application.CountA(Range("A:A")), Evaluate("SUM(A1:A2)")
End Sub''')
    assert out[0].split() == ["55", "5", "1/15/2024", "Date"]
    assert out[1].split() == ["55", "5", "5"]
    assert val(s, "C2") == 4 and text(s, "D1") == "1/15/2024" and val(s, "E1") == "007"


def test_a_whole_run_is_one_undo_step():
    s = new()
    run(s, "Sub M()\n    Range(\"A1\").Value = 1\n    Range(\"A1\").Font.Bold = True\n    Range(\"A2\").Formula = \"=A1+1\"\nEnd Sub\n")
    assert val(s, "A2") == 2
    s.run("edit.undo")
    assert val(s, "A1") is None and val(s, "A2") is None
    assert "Sub M()" in s.run("macro.getModule", name="Module1")["code"]


def test_with_blocks_and_named_arguments():
    s = new()
    out = run(s, '''
Sub Out(ParamArray xs())
    Dim i, t
    For i = 0 To UBound(xs)
        t = t & xs(i) & "|"
    Next
    Debug.Print t
End Sub
Sub M()
    Dim src As Worksheet, ws As Worksheet
    Set src = ActiveSheet
    With Range("A1").Font
        .Bold = True
        Out .Bold, .Size
    End With
    Set ws = Worksheets.Add(After:=src)
    ws.Name = "Data"
    Out Worksheets.Count, ws.Index, ActiveSheet.Name
End Sub''')
    assert out == ["True|11|", "2|2|Data|"]


def test_enum_types_and_qualified_constants():
    s = new()
    out = run(s, '''
Function Dir1(Optional d As XlDirection = xlUp) As Long
    Dir1 = d
End Function
Sub M()
    Dim d As XlDirection, r As VbMsgBoxResult
    d = xlDown
    Debug.Print TypeName(d), d, XlDirection.xlToLeft, Excel.xlUp, Excel.XlDirection.xlToRight, Dir1(), TypeName(vbYes), vbYes * 10000, xlAdd, xlMinimum
End Sub''')
    assert out[0].split() == ["Long", "-4121", "-4159", "-4162", "-4161", "-4162", "Long", "60000", "2", "4"]



def test_conditional_compilation_takes_the_windows_64_bit_branches():
    # ReflexGrid compiles as 64-bit VBA 7 for Windows; a branch not taken is never parsed (here, text that
    # isn't VBA); #Const names exist only for the directives; VBA.vbTab is the library's constant.
    s = new()
    out = run(s, '''#Const LEVEL = 2
#If VBA7 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If
#If Mac Then
    this line is not VBA: only Excel for Mac would read it
#End If

Function Platform() As String
#If Mac Then
    Platform = "mac"
#ElseIf Win64 And VBA7 Then
    Platform = "win64"
#Else
    Platform = "win32"
#End If
End Function

Function Pick() As String
#If LEVEL > 1 And Not Mac Then ' a comment
  #If LEVEL = 3 Then
    Pick = "three"
  #Else
    Pick = "more than one"
  #End If
#Else
    Pick = "one"
#End If
End Function

Sub M()
    Debug.Print Platform()
    Debug.Print Pick()
    Debug.Print IsEmpty(LEVEL)
    Debug.Print Len("a" & VBA.vbTab & "b")
End Sub''')
    assert [o.strip() for o in out] == ["win64", "more than one", "True", "3"]


def test_conditional_compilation_errors_keep_their_lines():
    s = new()
    s.run("macro.setModule", name="Module1", code="Sub M()\n#If VBA7 Then\n    x = 1\n#End If\n    If y Then\nEnd Sub\n")
    d = s.run("macro.list")["diagnostics"]
    assert len(d) == 1 and d[0]["line"] in (5, 6)
    s.run("macro.setModule", name="Module1", code="Sub M()\n#If Win64 Then\n    x = 1\nEnd Sub\n")
    d = s.run("macro.list")["diagnostics"]
    assert d[0]["line"] == 2 and "#End If" in d[0]["message"]
    s.run("macro.setModule", name="Module1", code="Sub M()\n#Else\nEnd Sub\n")
    assert s.run("macro.list")["diagnostics"][0]["line"] == 2
    s.run("macro.setModule", name="Module1", code="Sub M()\n#If VBA7\n#End If\nEnd Sub\n")
    assert "Then" in s.run("macro.list")["diagnostics"][0]["message"]


def test_the_projects_conditional_compilation_arguments():
    from dataclasses import replace

    from gridcore.vba.conditional import project_constants

    assert project_constants("DEBUG_MODE = 1 : Level=-2: bad") == {"debug_mode": 1, "level": -2}
    s = new()
    s.run("macro.setModule", name="Module1", code='Sub M()\n#If DEBUG_MODE Then\n    Debug.Print "debug"\n#Else\n    Debug.Print "release"\n#End If\nEnd Sub\n')
    assert s.run("macro.run", name="M")["output"] == ["release"]
    wb = s.doc().wb
    wb.vba = replace(wb.vba, constants="DEBUG_MODE = 1")
    assert s.run("macro.run", name="M")["output"] == ["debug"]


def test_find_replace_and_wildcards():
    s = new()
    out = run(s, '''
Sub M()
    Range("A1:A4").Value = Application.Transpose(Array("Apple", "apple", "pineapple", "APPLE PIE"))
    Debug.Print Range("A1:A4").Find("apple", LookAt:=xlWhole, MatchCase:=True).Address, Range("A1:A4").Find("a*e", LookAt:=xlWhole).Address
    Debug.Print Range("A1:A4").Find("apple", SearchDirection:=xlPrevious, LookAt:=xlPart).Address, Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Debug.Print Range("A1:A4").Find("zzz") Is Nothing, Range("A1:A4").Replace("apple", "pear")
End Sub''')
    assert out[0].split() == ["$A$2", "$A$2"]
    assert out[1].split() == ["$A$4", "4"]
    assert out[2].split() == ["True", "True"]
    assert val(s, "A3") == "pinepear"


def test_errors_and_handlers():
    s = new()
    s.run("macro.setModule", name="Module1", code='''
Sub M()
    On Error Resume Next
    Worksheets("Nope").Activate
    Debug.Print Err.Number
    Err.Clear
    Range("A1").Formula = "=SUM(1,"
    Debug.Print Err.Number, IsEmpty(Range("A1").Value)
    On Error GoTo 0
    Range("A2").Offset(-5, 0).Value = 1
End Sub''')
    r = s.run("macro.run", name="M")
    assert r["status"] == "error"
    assert [x.split() for x in r["output"]] == [["9"], ["1004", "True"]]
    assert r["error"]["number"] == 1004 and r["error"]["line"] == 10 and r["error"]["proc"] == "M"


def test_dialogs_get_scripted_answers_or_their_defaults():
    s = new()
    code = '''
Sub M()
    If MsgBox("Go on?", vbYesNo + vbQuestion) = vbYes Then Range("A1").Value = "yes" Else Range("A1").Value = "no"
    Range("A2").Value = InputBox("Name?", , "Bob")
    MsgBox "Done"
End Sub'''
    s.run("macro.setModule", name="Module1", code=code)
    r = s.run("macro.run", name="M", answers=[6, "Ann"])
    assert (val(s, "A1"), val(s, "A2")) == ("yes", "Ann")
    assert [m["prompt"] for m in r["messages"]] == ["Done"]
    s.run("macro.run", name="M")
    assert (val(s, "A1"), val(s, "A2")) == ("yes", "Bob")


def test_interactive_dialogs_pause_the_run():
    s = new()
    s.interactive_macros = True
    s.run("macro.setModule", name="Module1", code='''
Sub M()
    Range("A1").Value = 1
    If MsgBox("Go on?", vbYesNo) = vbYes Then Range("A2").Value = InputBox("Value?")
    MsgBox "Bye"
End Sub''')
    r = s.run("macro.run", name="M")
    assert r["status"] == "prompt" and r["prompt"]["kind"] == "msgbox" and r["prompt"]["buttons"] == 4
    assert val(s, "A1") == 1  # what ran before the dialog is in the workbook
    with pytest.raises(EngineError):
        s.run("macro.run", name="M")
    assert s.run("macro.answer", value=6)["prompt"]["kind"] == "inputbox"
    assert s.run("macro.answer", value="hello")["prompt"]["prompt"] == "Bye"
    assert s.run("macro.answer", value=1)["status"] == "done"
    assert val(s, "A2") == "hello"
    s.run("edit.undo")
    assert val(s, "A1") is None and val(s, "A2") is None


def test_stopping_and_limits():
    s = new()
    s.run("macro.setModule", name="Module1", code="Sub Forever()\n    Do\n    Loop\nEnd Sub\n")
    r = s.run("macro.run", name="Forever", timeLimit=0.5)
    assert r["status"] == "error" and r["error"]["kind"] == "timeout"
    s.interactive_macros = True
    r = s.run("macro.run", name="Forever")
    assert r["status"] == "busy"
    r = s.run("macro.stop")
    assert r["status"] == "cancelled"
    assert s.macro_run is None
    s.run("macro.setModule", name="Module1", code="Sub Deep(n)\n    Deep n + 1\nEnd Sub\nSub M()\n    Deep 1\nEnd Sub\n")
    r = s.run("macro.run", name="M")
    assert r["status"] == "error" and r["error"]["number"] == 28


def test_module_state_lasts_until_end_or_an_edit():
    s = new()
    s.run("macro.setModule", name="Module1", code="Dim n As Long\nSub M()\n    n = n + 1\n    Range(\"A1\").Value = n\nEnd Sub\nSub Reset()\n    End\nEnd Sub\n")
    s.run("macro.run", name="M")
    s.run("macro.run", name="M")
    assert val(s, "A1") == 2
    s.run("macro.run", name="Reset")
    s.run("macro.run", name="M")
    assert val(s, "A1") == 1


def test_immediate_window():
    s = new()
    s.run("macro.setModule", name="Module1", code="Function Twice(x)\n    Twice = x * 2\nEnd Function\n")
    assert s.run("macro.immediate", code="? Twice(21)")["output"] == [" 42 "]
    s.run("macro.immediate", code='Range("A1").Value = 5')
    assert val(s, "A1") == 5
    r = s.run("macro.immediate", code="? 1 +")
    assert r["status"] == "error" and r["error"]["line"] == 1


def test_manual_calculation_and_calculate():
    s = new()
    out = run(s, '''
Sub M()
    Application.Calculation = xlCalculationManual
    Range("A1").Value = 1
    Range("A2").Formula = "=A1*2"
    Range("A1").Value = 5
    Debug.Print Range("A2").Value
    Application.Calculate
    Debug.Print Range("A2").Value
End Sub''')
    assert [x.strip() for x in out] == ["2", "10"]
    assert s.doc().wb.calc.mode != "manual"  # a run puts automatic calculation back


def test_calculate_now_in_manual_mode():
    s = new()
    s.run("cell.set", cell="A1", input="1")
    s.run("cell.set", cell="A2", input="=A1*2")
    s.run("formulas.calculationOptions", mode="manual")
    s.run("cell.set", cell="A1", input="5")
    assert val(s, "A2") == 2
    s.run("formulas.calculateNow")
    assert val(s, "A2") == 10


def test_protection_and_user_interface_only():
    s = new()
    s.run("macro.setModule", name="Module1", code='''
Sub Locked()
    ActiveSheet.Protect
    Range("A1").Value = 1
End Sub
Sub UiOnly()
    ActiveSheet.Unprotect
    ActiveSheet.Protect UserInterfaceOnly:=True
    Range("A1").Value = 2
    Range("A1:A2").Sort Key1:=Range("A1")
    Debug.Print ActiveSheet.ProtectionMode
End Sub''')
    r = s.run("macro.run", name="Locked")
    assert r["status"] == "error" and r["error"]["number"] == 1004
    r = s.run("macro.run", name="UiOnly")
    assert r["status"] == "done" and r["output"] == ["True"]
    assert val(s, "A1") == 2
    with pytest.raises(EngineError):  # people still can't
        s.run("cell.set", cell="A1", input="3")


def test_conditional_formats():
    s = new()
    out = run(s, '''
Sub M()
    Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))
    With Range("A1:A3").FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=1")
        .Interior.Color = RGB(255, 199, 206)
        .Font.Bold = True
    End With
    Debug.Print Range("A1:A3").FormatConditions.Count, Range("A1").FormatConditions(1).Formula1, Range("A1").FormatConditions(1).Interior.Color
End Sub''')
    assert out[0].split() == ["1", "=1", "13551615"]
    cf = s.doc().wb.sheets[0].cond_formats[0]
    assert cf.rule.type == "cellIs" and cf.rule.style.font.bold


def test_class_events_and_terminate():
    s = new()
    s.run("macro.addModule", kind="class", name="Counter", code='''
Public Event Changed(ByVal v As Long)
Private n As Long
Public Sub Bump()
    n = n + 1
    RaiseEvent Changed(n)
End Sub
Private Sub Class_Initialize()
    n = 100
End Sub
Private Sub Class_Terminate()
    Module1.log = Module1.log & "T" & n & ","
End Sub
''')
    s.run("macro.addModule", kind="class", name="Listener", code='''
Private WithEvents src As Counter
Public total As Long
Public Sub Attach(x As Counter)
    Set src = x
End Sub
Public Sub Detach()
    Set src = Nothing
End Sub
Private Sub src_Changed(ByVal v As Long)
    total = total + v
End Sub
''')
    out = run(s, '''
Public log As String
Sub Scoped()
    Dim y As New Counter
    y.Bump
End Sub
Sub M()
    Dim c As New Counter, l As New Listener
    l.Attach c
    c.Bump
    c.Bump
    l.Detach
    c.Bump
    Debug.Print l.total
    Dim x As Counter
    Set x = New Counter
    Set x = Nothing
    Debug.Print log
    Scoped
    Debug.Print log
End Sub''')
    assert [x.strip() for x in out] == ["203", "T100,", "T100,T101,"]


# ---------------------------------------------------------------- worksheet functions in VBA


def test_vba_functions_in_formulas():
    s = new()
    s.run("macro.setModule", name="Module1", code='''
Function MySum(r As Range) As Double
    Dim c As Range
    For Each c In r
        If IsNumeric(c.Value) Then MySum = MySum + c.Value
    Next
End Function
Function Greet(Optional who As String = "World") As String
    Greet = "Hello, " & who
End Function
Function Seq3()
    Seq3 = Array(1, 2, 3)
End Function
Function Bad()
    Bad = 1 / 0
End Function
Function Where()
    Where = Application.Caller.Address
End Function
Function Sneaky()
    Range("Z1").Value = 1
    Sneaky = 1
End Function''')
    for a, t in [("A1", "1"), ("A2", "2"), ("A3", "x"), ("B1", "=MySum(A1:A3)"), ("B2", "=Greet()"), ("B3", '=Greet("Ann")'), ("B4", "=Bad()"), ("B5", "=Where()"), ("B6", "=NoSuch()"), ("B7", "=Sneaky()"), ("C1", "=Seq3()")]:
        s.run("cell.set", cell=a, input=t)
    assert val(s, "B1") == 3
    s.run("cell.set", cell="A1", input="10")
    assert val(s, "B1") == 12  # recalculated with its arguments
    assert (val(s, "B2"), val(s, "B3")) == ("Hello, World", "Hello, Ann")
    assert (text(s, "B4"), val(s, "B5"), text(s, "B6"), text(s, "B7")) == ("#VALUE!", "$B$5", "#NAME?", "#VALUE!")
    assert val(s, "Z1") is None
    assert (val(s, "C1"), val(s, "D1"), val(s, "E1")) == (1, 2, 3)


def test_functions_follow_code_changes():
    s = new()
    s.run("macro.setModule", name="Module1", code="Function F(x)\n    F = x + 1\nEnd Function\n")
    s.run("cell.set", cell="A1", input="=F(1)")
    assert val(s, "A1") == 2
    s.run("macro.setModule", name="Module1", code="Function F(x)\n    F = x + 100\nEnd Function\n")
    assert val(s, "A1") == 101
    s.run("macro.setModule", name="Module1", code="Sub Other()\nEnd Sub\n")
    assert text(s, "A1") == "#NAME?"


# ---------------------------------------------------------------- events


EVENTS = '''
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("A1:A10")) Is Nothing Then
        Application.EnableEvents = False
        Target.Offset(0, 1).Value = "changed " & Target.Address(0, 0)
        Application.EnableEvents = True
    End If
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Range("D1").Value = Target.Address(0, 0)
End Sub
'''


def test_events_run_only_when_enabled():
    s = new()
    s.run("macro.setModule", name="Sheet1", code=EVENTS)
    s.run("cell.set", cell="A2", input="5")
    assert val(s, "B2") is None
    s.run("macro.enableEvents", on=True)
    s.run("cell.set", cell="A3", input="5")
    assert val(s, "B3") == "changed A3"
    s.run("edit.undo")  # the edit and what its event did go together
    assert val(s, "A3") is None and val(s, "B3") is None
    s.run("selection.set", range="C3:C4")
    assert val(s, "D1") == "C3:C4"
    s.run("edit.undo")
    assert val(s, "D1") is None


def test_macro_edits_fire_events_too():
    s = new()
    s.run("macro.setModule", name="ThisWorkbook", code='''
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Module1.log = Module1.log & Target.Address(0, 0) & ","
End Sub''')
    out = run(s, '''
Public log As String
Sub M()
    Range("A1").Value = 1
    Range("B1:C2").Value = 2
    Range("D1").ClearContents
    Range("A1").Font.Bold = True
    Application.EnableEvents = False
    Range("E1").Value = 1
    Application.EnableEvents = True
    Debug.Print log
End Sub''')
    assert out == ["A1,B1:C2,D1,"]


def test_workbook_open_runs_when_macros_are_enabled():
    s = new()
    s.run("macro.setModule", name="ThisWorkbook", code="Private Sub Workbook_Open()\n    Sheet1.Range(\"A1\").Value = \"opened\"\nEnd Sub\n")
    assert s.run("macro.open")["status"] == "skipped"
    s.run("macro.enableEvents", on=True)
    s.run("macro.open")
    assert val(s, "A1") == "opened"


# ---------------------------------------------------------------- shared workbooks


def _book():
    return SharedBook(DocState(Workbook(), "Shared"), "b1")


def _join(b, key):
    v = b.join(key, CollabSession(), key.title())
    v.session.interactive_macros = True
    return v


def _act(b, v, cid, **p):
    with b.act(v) as s:
        return s.execute(cid, p)


def test_a_paused_run_on_a_shared_workbook():
    b = _book()
    amy, ben = _join(b, "amy"), _join(b, "ben")
    _act(b, amy, "macro.setModule", name="Module1", code='''
Sub Fill()
    Range("A1").Value = 1
    If MsgBox("More?", vbYesNo) = vbYes Then Range("A2").Value = 2
    Range("A3").Value = Application.UserName
End Sub''')
    assert _act(b, amy, "macro.run", name="Fill")["status"] == "prompt"
    assert _act(b, ben, "cell.get", cell="A1")["value"] == 1  # others see it while Amy decides
    _act(b, ben, "cell.set", cell="C1", input="ben")
    assert _act(b, amy, "macro.answer", value=6)["status"] == "done"
    assert _act(b, amy, "cell.get", cell="A3")["value"] == "Amy"
    # Ben's edit came between: Amy's run is two steps of her own history, Ben's stays his.
    assert [t.label for t in amy.undo][-2:] == ["Run Macro", "Answer Macro"]
    _act(b, amy, "edit.undo")
    _act(b, amy, "edit.undo")
    assert _act(b, amy, "cell.get", cell="A1")["value"] is None
    assert _act(b, amy, "cell.get", cell="C1")["value"] == "ben"


def test_an_uninterrupted_run_is_one_step_on_a_shared_workbook():
    b = _book()
    amy = _join(b, "amy")
    _act(b, amy, "macro.setModule", name="Module1", code="Sub Fill()\n    Range(\"A1\").Value = 1\n    If MsgBox(\"More?\", vbYesNo) = vbYes Then Range(\"A2\").Value = 2\nEnd Sub\n")
    _act(b, amy, "macro.run", name="Fill")
    _act(b, amy, "macro.answer", value=6)
    assert [t.label for t in amy.undo] == ["Edit VBA Code", "Run Macro"]
    _act(b, amy, "edit.undo")
    assert _act(b, amy, "cell.get", cell="A2")["value"] is None


def test_events_on_a_shared_workbook():
    b = _book()
    amy, ben = _join(b, "amy"), _join(b, "ben")
    _act(b, amy, "macro.setModule", name="Sheet1", code="Private Sub Worksheet_Change(ByVal Target As Range)\n    Application.EnableEvents = False\n    Target.Offset(0, 1).Value = UCase(Target.Value)\n    Application.EnableEvents = True\nEnd Sub\n")
    _act(b, amy, "macro.enableEvents", on=True)
    _act(b, ben, "cell.set", cell="A1", input="hi")
    assert _act(b, amy, "cell.get", cell="B1")["value"] == "HI"
    _act(b, ben, "edit.undo")
    assert _act(b, amy, "cell.get", cell="B1")["value"] is None


# ---------------------------------------------------------------- files


def test_xlsm_round_trip_keeps_and_runs_the_macros():
    s = new()
    s.run("macro.setModule", name="Module1", code="Sub Hello()\n    Range(\"A1\").Value = \"hi\"\nEnd Sub\nFunction Twice(x)\n    Twice = x * 2\nEnd Function\n")
    s.run("cell.set", cell="B1", input="=Twice(21)")
    saved = s.run("file.saveBytes")
    assert saved["name"].endswith(".xlsm") and saved["mime"] == "application/vnd.ms-excel.sheet.macroEnabled.12"
    data = base64.b64decode(saved["data"])
    s2 = Session()
    d = open_bytes(s2, data, "Book1.xlsm")
    assert d.warnings == []
    assert val(s2, "B1") == 42
    assert [m["name"] for m in s2.run("macro.list")["macros"]] == ["Hello"]
    s2.run("macro.run", name="Hello")
    assert val(s2, "A1") == "hi"
    plain = base64.b64decode(s.run("file.saveBytes", format="xlsx")["data"])
    import zipfile

    with zipfile.ZipFile(io.BytesIO(data)) as z:
        assert "xl/vbaProject.bin" in z.namelist()
    with zipfile.ZipFile(io.BytesIO(plain)) as z:
        assert "xl/vbaProject.bin" not in z.namelist()


def test_openpyxl_reads_the_xlsm():
    openpyxl = pytest.importorskip("openpyxl")
    s = new()
    s.run("macro.setModule", name="Module1", code="Sub A()\nEnd Sub\n")
    data = base64.b64decode(s.run("file.saveBytes")["data"])
    wb = openpyxl.load_workbook(io.BytesIO(data), keep_vba=True)
    try:
        assert wb.vba_archive is not None and "xl/vbaProject.bin" in wb.vba_archive.namelist()
    finally:
        wb.vba_archive.close()


# ---------------------------------------------------------------- the sandbox


@pytest.mark.parametrize("stmt,number", [
    ('Shell "rm -rf /tmp/x"', 445),
    ('Open "/etc/passwd" For Input As #1', 445),
    ('Kill "/tmp/x"', 445),
    ('x = Dir("/etc/*")', 445),
    ('Set o = CreateObject("WScript.Shell")', 429),
    ('Set o = CreateObject("Scripting.FileSystemObject")', 429),
    ('Set o = CreateObject("MSXML2.XMLHTTP")', 429),
    ('x = MessageBoxA(0, "x", "y", 0)', 445),
    ('SendKeys "abc"', 445),
    ('Workbooks.Open "/tmp/a.xlsx"', 445),
    ('ThisWorkbook.SaveAs "/tmp/a.xlsm"', 445),
    ('x = ThisWorkbook.VBProject.Name', 445),
])
def test_macros_cant_reach_outside_the_workbook(stmt, number):
    s = new()
    s.run("macro.setModule", name="Module1", code=f'''
Private Declare PtrSafe Function MessageBoxA Lib "user32" (ByVal h As LongPtr, ByVal t As String, ByVal c As String, ByVal f As Long) As Long
Sub M()
    Dim x, o
    On Error Resume Next
    {stmt}
    Debug.Print Err.Number
End Sub''')
    r = s.run("macro.run", name="M")
    assert r["output"] == [f" {number} "]


def test_environment_is_not_visible():
    s = new()
    assert run(s, 'Sub M()\n    Debug.Print "[" & Environ("HOME") & "]"\nEnd Sub\n') == ["[]"]
