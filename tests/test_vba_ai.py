"""AI in the VBA editor: what it is asked about, how the questions are worded and the answers read
(reflex_grid/vba_ai.py). No model is asked here; the browser suite vba_ai drives the editor against
a stand-in hub."""

from gridcore.engine import Session
from reflex_grid import vba_ai


def workbook():
    s = Session()
    s.new_workbook()
    s.run("range.setValues", range="A1", values=[["Customer", "Qty"], ["Ann", 3]])
    s.run("insert.table", range="A1:B2", header=True, name="Orders")
    s.run("macro.setModule", name="Module1", code='Sub Hello()\n    Range("A1").Value = 42\nEnd Sub\n\nPrivate Sub Hidden()\nEnd Sub\n\nFunction Twice(x)\n    Twice = x * 2\nEnd Function\n')
    return s


# ---------------------------------------------------------------- an answer in a module

MODULE = """Option Explicit

Dim total As Long

' Says hello.
Sub Hello()
    MsgBox "Hi"
End Sub

Function Twice(x)
    Twice = x * 2
End Function

Property Get Size()
    Size = 1
End Property

Property Let Size(v)
End Property
"""


def test_what_an_answer_leaves_out_stays():
    new = vba_ai.merged(MODULE, "Sub Added()\n    x = 1\nEnd Sub", line=len(MODULE.split("\n")))
    assert new == MODULE + "\nSub Added()\n    x = 1\nEnd Sub\n"
    # A model that answers with nothing new changes nothing.
    assert vba_ai.merged(MODULE, "' nothing to do") == MODULE


def test_a_procedure_takes_the_place_of_its_name():
    new = vba_ai.merged(MODULE, 'sub hello()\n    MsgBox "Yo"\nEnd Sub')
    assert new == MODULE.replace('Sub Hello()\n    MsgBox "Hi"\nEnd Sub', 'sub hello()\n    MsgBox "Yo"\nEnd Sub')
    assert "' Says hello." in new, "the comment above it stays when the answer has none"
    new = vba_ai.merged(MODULE, "' Says it louder.\nSub Hello()\nEnd Sub")
    assert "' Says hello." not in new and "' Says it louder.\nSub Hello()\nEnd Sub\n\nFunction Twice" in new
    # A Sub may become a Function; a Property's Get and Let are two.
    assert "Sub Twice" in vba_ai.merged(MODULE, "Sub Twice()\nEnd Sub") and "Function Twice" not in vba_ai.merged(MODULE, "Sub Twice()\nEnd Sub")
    new = vba_ai.merged(MODULE, "Property Let Size(v)\n    total = v\nEnd Property")
    assert new.count("Property Get Size") == 1 and "    Size = 1" in new and "    total = v" in new


def test_a_new_procedure_goes_where_the_cursor_is():
    block = "Sub New1()\nEnd Sub"
    at = MODULE.split("\n").index('    MsgBox "Hi"') + 1
    assert 'End Sub\n\nSub New1()\nEnd Sub\n\nFunction Twice' in vba_ai.merged(MODULE, block, line=at)
    # Among the declarations, where a module opens: at its end.
    assert vba_ai.merged(MODULE, block, line=1).endswith("Property Let Size(v)\nEnd Property\n\nSub New1()\nEnd Sub\n")
    assert vba_ai.merged(MODULE, block).endswith("End Property\n\nSub New1()\nEnd Sub\n")
    # On the blank line between two procedures.
    at = MODULE.split("\n").index("Function Twice(x)")
    assert 'End Sub\n\nSub New1()\nEnd Sub\n\nFunction Twice' in vba_ai.merged(MODULE, block, line=at)


def test_declarations_go_up_with_the_others():
    new = vba_ai.merged(MODULE, "Dim total As Long\nConst RATE = 5\n\nSub New1()\nEnd Sub")
    assert new.count("Dim total As Long") == 1
    assert "Dim total As Long\nConst RATE = 5\n\n' Says hello." in new
    new = vba_ai.merged("Dim a\n\nSub A()\nEnd Sub\n", "Option Explicit\nPrivate Type Pair\n    x As Long\nEnd Type")
    assert new == "Option Explicit\nDim a\nPrivate Type Pair\n    x As Long\nEnd Type\n\nSub A()\nEnd Sub\n"
    # In a module with no procedures yet.
    assert vba_ai.merged("Dim a\n", "Dim b") == "Dim a\nDim b\n"


def test_remove_takes_a_procedure_out():
    new = vba_ai.merged(MODULE, "' REMOVE Function Twice\n' remove Size\nSub Twice2()\nEnd Sub")
    assert "Twice(x)" not in new and "Property" not in new and "Sub Twice2()" in new
    assert "End Sub\n\nSub Twice2()" in new
    # Renaming: the new name, and the old one removed.
    new = vba_ai.merged("Sub A()\nEnd Sub\n\nSub B()\nEnd Sub\n", "' REMOVE A\nSub C()\nEnd Sub", line=1)
    assert new == "Sub C()\nEnd Sub\n\nSub B()\nEnd Sub\n" or new == "Sub B()\nEnd Sub\n\nSub C()\nEnd Sub\n"


def test_statements_with_a_dim_are_not_declarations():
    code = "Dim a\n\nSub A()\n    x = 1\nEnd Sub\n"
    assert vba_ai.merged(code, "    Dim y\n    y = 2", line=4) == "Dim a\n\nSub A()\n    x = 1\n    Dim y\n    y = 2\nEnd Sub\n"


def test_an_open_procedure_is_replaced_up_to_the_next():
    code = "Sub Broken()\n    x = 1\n\nSub Next1()\nEnd Sub\n"
    assert vba_ai.merged(code, "Sub Broken()\n    x = 1\nEnd Sub") == "Sub Broken()\n    x = 1\nEnd Sub\n\nSub Next1()\nEnd Sub\n"


def test_statements_alone_go_after_the_cursor():
    code = "Sub A()\n    x = 1\nEnd Sub\n"
    assert vba_ai.merged(code, "    y = 2", line=2) == "Sub A()\n    x = 1\n    y = 2\nEnd Sub\n"
    # On an empty line: in its place.
    assert vba_ai.merged("Sub A()\n\nEnd Sub\n", "    y = 2", line=2) == "Sub A()\n    y = 2\nEnd Sub\n"


def test_an_empty_module_takes_the_answer_as_it_is():
    assert vba_ai.merged("", "Option Explicit\n\nSub A()\nEnd Sub") == "Option Explicit\n\nSub A()\nEnd Sub"
    assert vba_ai.merged("\n", "' A note\nSub A()\nEnd Sub") == "' A note\nSub A()\nEnd Sub"


# ---------------------------------------------------------------- what to write


def test_selected_lines_are_replaced_as_they_are():
    code = "Sub A()\nEnd Sub\n\nSub B()\nEnd Sub\n"
    t = vba_ai.target_of(code, line=4, first=4, last=5)
    assert (t.first, t.last) == (4, 5)
    assert t.result("Sub C()\nEnd Sub") == "Sub A()\nEnd Sub\n\nSub C()\nEnd Sub\n"
    # Nothing written back: the lines go.
    assert t.result("") == "Sub A()\nEnd Sub\n\n"
    # Past the end of the module: its last line.
    assert vba_ai.target_of(code, first=99, last=120).first == 5
    # Going on from an answer: the lines it became.
    longer = t.after("Sub C()\n    x = 1\nEnd Sub")
    assert (longer.first, longer.last) == (4, 6) and longer.result("Sub D()\nEnd Sub") == "Sub A()\nEnd Sub\n\nSub D()\nEnd Sub\n"


def test_selected_lines_stay_as_far_in_as_they_were():
    code = "Sub A()\n    x = 1\n    y = 2\nEnd Sub\n"
    t = vba_ai.target_of(code, 2, 2, 3)
    assert t.result("For i = 1 To 2\n\n    x = i\nNext") == "Sub A()\n    For i = 1 To 2\n\n        x = i\n    Next\nEnd Sub\n"
    assert t.result("    z = 3") == "Sub A()\n    z = 3\nEnd Sub\n"


def test_without_a_selection_the_answer_goes_in_by_name():
    t = vba_ai.target_of(MODULE, line=1)
    assert t.first == 0 and t.result("Sub Hello()\nEnd Sub").count("Sub Hello") == 1
    # Going on from an answer: it stays in.
    more = t.after("Sub New1()\nEnd Sub")
    assert "Sub New1" in more.code and "Sub New1" in more.result("Sub New2()\nEnd Sub")


def test_a_compile_error_is_fixed_in_its_procedure():
    code = "' Note\nSub A()\n    x = 1\nEnd Sub\n\nSub B()\n    If x Then\nEnd Sub\n"
    t = vba_ai.fix_target(code, 8)
    assert (t.first, t.last) == (6, 8)
    assert t.result("Sub B()\nEnd Sub") == "' Note\nSub A()\n    x = 1\nEnd Sub\n\nSub B()\nEnd Sub\n"
    # With the comment above it.
    assert vba_ai.fix_target(code, 3).first == 1
    # A procedure left open shows on the next one's first line.
    t = vba_ai.fix_target("Sub A()\n    x = 1\n\nSub B()\nEnd Sub\n", 4)
    assert (t.first, t.last) == (1, 5)


def test_a_run_time_error_may_come_from_anywhere():
    code = "Sub A()\n    B\nEnd Sub\n\nSub B()\n    x = 1 / 0\nEnd Sub\n"
    assert vba_ai.fix_target(code, 6, compile_error=False).first == 0
    assert vba_ai.fix_target(code, 0).first == 0


def test_an_error_outside_the_procedures_is_fixed_where_it_is():
    code = "Option Explicit\nbad line\n\nSub A()\n    x = 1\nEnd Sub\n"
    t = vba_ai.fix_target(code, 2)
    assert (t.first, t.last) == (1, 3)
    assert vba_ai.fix_target("Sub A()\nEnd Sub\n\nstray\n", 4).first == 3


def test_the_procedure_at_a_line():
    lines = ["Option Explicit", "", "' Note", "Public Static Function F(x)", "  F = 1", "End Function", "", "Property Get P()", "End Property"]
    assert vba_ai.procedure_at(lines, 5) == (4, 6)
    assert vba_ai.procedure_at(lines, 3) == (4, 6)
    assert vba_ai.procedure_at(lines, 9) == (8, 9)
    assert vba_ai.procedure_at(lines, 7) is None
    assert vba_ai.procedure_at(lines, 1) is None


# ---------------------------------------------------------------- what the AI is told


def test_the_brief_names_what_exists():
    wb = workbook().doc().wb
    brief = vba_ai.workbook_brief(wb, "Other")
    assert 'Sheet "Sheet1", code name Sheet1, cells A1:B2' in brief
    assert "table Orders at A1:B2, columns: Customer, Qty" in brief
    assert "A1: Customer | B1: Qty" in brief and "A2: Ann | B2: 3" in brief
    assert "Module Module1 (standard): Sub Hello(); Function Twice(x)" in brief
    assert "Hidden" not in brief
    # The module being written isn't repeated: it is in the question.
    assert "Module Module1" not in vba_ai.workbook_brief(wb, "module1")


def test_a_forms_controls_are_named():
    s = workbook()
    s.run("macro.addModule", kind="form", name="OrderForm")
    s.run("macro.editForm", name="OrderForm", ops=[{"op": "add", "kind": "TextBox", "left": 10, "top": 10}, {"op": "add", "kind": "CommandButton", "left": 10, "top": 40}])
    assert "UserForm OrderForm, controls: TextBox1 (TextBox), CommandButton1 (CommandButton" in vba_ai.workbook_brief(s.doc().wb, "OrderForm")


def test_the_question_for_a_module_and_for_selected_lines():
    code = "Sub A()\n    x = 1\nEnd Sub\n\nSub B()\nEnd Sub\n"
    text = vba_ai.write_prompt("Sheet", "Module1", "standard", vba_ai.target_of(code, 2), "加上註解")
    assert '<module name="Module1" kind="module">\nSub A()' in text
    assert 'The person\'s cursor is on line 2, in A: "this" and "here" mean A.' in text
    assert "<request>\n加上註解\n</request>" in text and "<target" not in text
    assert "The person's cursor is on line 4.\n" in vba_ai.write_prompt("Sheet", "Module1", "standard", vba_ai.target_of(code, 4), "x")
    # A new module has no cursor to speak of.
    assert "cursor" not in vba_ai.write_prompt("Sheet", "OrderForm", "form", vba_ai.target_of("", 1), "x")
    text = vba_ai.write_prompt("Sheet", "Sheet1", "document", vba_ai.target_of(code, 5, 5, 6), "x")
    assert '<target lines="5-6">\nSub B()\nEnd Sub\n</target>' in text
    assert '<module_before name="Sheet1" kind="code of a sheet or the workbook">' in text and "Sub A()" in text


def test_a_big_module_is_shown_around_the_cursor():
    filler = "\n".join(f"' line {i}" for i in range(vba_ai.WHOLE_MODULE_LINES))
    code = f"Sub First()\nEnd Sub\n\n{filler}\n\nPrivate Sub Far()\n    x = 1\nEnd Sub\n"
    at = code.split("\n").index("    x = 1") + 1
    text = vba_ai.write_prompt("Sheet", "Module1", "standard", vba_ai.target_of(code, at), "x")
    assert ", in Far:" in text and "shown=\"lines" in text
    assert "    x = 1" in text and "' line 5\n" not in text
    assert "<its_procedures>\nSub First()\nPrivate Sub Far()\n</its_procedures>" in text


def test_errors_are_shown_with_their_line():
    code = "Sub A()\n    x = \nEnd Sub\n"
    assert vba_ai.fix_request("Compile error: Expected: expression", 2, code) == "Fix this error: Compile error: Expected: expression\nLine 2 of the module: x =\nKeep what the code is meant to do; change only what is needed."
    assert "Line 9" not in vba_ai.fix_request("Run-time error", 9, code)
    wrong = [{"line": 2, "message": "Expected: expression"}]
    text = vba_ai.retry_request(wrong, code, vba_ai.target_of(code))
    assert text.startswith("That doesn't compile. Line 2 of the module: Expected: expression (x =)")
    assert "each procedure you add or change, whole" in text
    assert "the target's lines" in vba_ai.retry_request(wrong, code, vba_ai.target_of(code, 2, 1, 3))


# ---------------------------------------------------------------- reading the answers


def test_the_code_of_an_answer():
    assert vba_ai.code_of("```vba\nSub A()\nEnd Sub\n```\nDone.") == "Sub A()\nEnd Sub"
    assert vba_ai.code_of("Here it is:\n\n```vb\nSub A()\nEnd Sub\n```") == "Sub A()\nEnd Sub"
    assert vba_ai.code_of("```\nSub A()\n```") == "Sub A()"
    # Still being written.
    assert vba_ai.code_of("```vba\nSub A()\n    x") == "Sub A()\n    x"
    assert vba_ai.code_of("```vba\nSub A()\n``") == "Sub A()"
    assert vba_ai.code_of("``") == ""
    # A model that left the fence out; Attribute lines don't belong in a module's code.
    assert vba_ai.code_of("Option Explicit\n\nSub A()\nEnd Sub\n") == "Option Explicit\n\nSub A()\nEnd Sub"
    assert vba_ai.code_of('```vba\nAttribute VB_Name = "Module1"\nSub A()\nEnd Sub\n```') == "Sub A()\nEnd Sub"
    # Words are an answer, not code.
    assert vba_ai.code_of("這段程式把 42 寫到 A1。") == ""
    assert vba_ai.code_of("It writes 42 into A1.") == ""


def test_a_suggestion_is_trimmed_to_what_helps():
    prefix = 'Sub A()\n    Range("A1").'
    assert vba_ai.clean_inline("Value = 1", prefix, "\nEnd Sub") == "Value = 1"
    assert vba_ai.clean_inline("```vba\nValue = 1\n```", prefix, "") == "Value = 1"
    assert vba_ai.clean_inline('    Range("A1").Value = 1', prefix, "") == "Value = 1"
    assert vba_ai.clean_inline("Value = 1)", prefix, ")\nEnd Sub") == "Value = 1"
    assert vba_ai.clean_inline("Here is the code:\n```\nx\n```", prefix, "") == ""
    assert vba_ai.clean_inline("Sure, you can set it", prefix, "") == ""
    assert vba_ai.clean_inline("", prefix, "") == ""
    many = "\n".join(f"x{i} = 1" for i in range(20))
    assert len(vba_ai.clean_inline(many, prefix, "").split("\n")) == vba_ai.INLINE_LINES


def test_a_model_that_writes_the_code_out_again():
    prefix = "Sub Test()\n    Dim ws As Worksheet\n    Set ws = Worksheets("
    suffix = "\nEnd Sub\n"
    echo = 'Sub Test()\n    Dim ws As Worksheet\n    Set ws = Worksheets("訂單")\nEnd Sub'
    assert vba_ai.clean_inline(echo, prefix, suffix) == '"訂單")'
    # Nothing added to it: nothing to suggest.
    assert vba_ai.clean_inline(prefix + suffix, prefix, suffix) == ""
    # It may put the spaces right.
    messy = "Sub Test()\n    Dim ws As Worksheet\n        Set ws = Worksheets("
    assert vba_ai.clean_inline(echo, messy, suffix) == '"訂單")'
    # It may start further down, or write out part of what follows.
    assert vba_ai.clean_inline('Set ws = Worksheets("訂單")\nEnd Sub', prefix, suffix) == '"訂單")'
    assert vba_ai.clean_inline('Worksheets("訂單")', prefix, suffix) == '"訂單")'
    long_tail = "\nEnd Sub\n\nSub Other()\n    x = 1\nEnd Sub\n"
    assert vba_ai.clean_inline(echo, prefix, long_tail) == '"訂單")'
    # What follows the cursor isn't a suggestion.
    assert vba_ai.clean_inline("End Sub", "Sub A()\n    x.", "\nEnd Sub\n") == ""
    # Lines that belong to it stay.
    body = 'Sheet1.Range("A" & i).Value = i\n    Next i'
    assert vba_ai.clean_inline(f"Sub T()\n    For i = 1 To 3\n        {body}\nEnd Sub", "Sub T()\n    For i = 1 To 3\n        ", "\nEnd Sub\n") == body
    # Two characters in common aren't an echo.
    assert vba_ai.clean_inline("x = x + 1", "Sub A()\n    Dim x\n    x", "\nEnd Sub") == "x = x + 1"


# ---------------------------------------------------------------- trying an answer


def test_an_answer_is_compiled_without_touching_the_workbook():
    s = workbook()
    project = s.doc().wb.vba
    before = project.module("Module1").code
    assert vba_ai.compile_problems(project, "Module1", "Sub Good()\nEnd Sub\n") == []
    assert vba_ai.compile_problems(project, "module1", "Sub Bad()\n") == [{"line": 1, "message": "Expected: End Sub"}]
    assert vba_ai.compile_problems(project, "Module1", "Sub Bad()\n    Dim x As\nEnd Sub\n") == [{"line": 2, "message": "Expected: type name"}]
    assert s.doc().wb.vba is project and project.module("Module1").code == before
    # A module to come, and a workbook with no macros yet.
    assert vba_ai.compile_problems(project, "Fresh", "Sub Good()\nEnd Sub\n") == []
    assert vba_ai.compile_problems(None, "", "Sub Bad(\n")


def test_only_the_errors_a_change_brought_are_held_against_it():
    s = workbook()
    project = s.doc().wb.vba
    before = "Sub Open1()\n"
    # It had this error already: not the answer's, unless the answer is to fix it.
    after = "Sub Fine()\nEnd Sub\n\nSub Open1()\n"
    problems, ours = vba_ai.new_problems(project, "Module1", before, after)
    assert problems and not ours
    assert vba_ai.new_problems(project, "Module1", before, after, fixing=True)[1] == problems
    # A new error is the answer's.
    problems, ours = vba_ai.new_problems(project, "Module1", "Sub A()\nEnd Sub\n", "Sub A()\n    Dim x As\nEnd Sub\n")
    assert ours == problems == [{"line": 2, "message": "Expected: type name"}]
    assert vba_ai.new_problems(project, "Module1", "Sub A()\nEnd Sub\n", "Sub B()\nEnd Sub\n") == ([], [])


def test_a_suggestion_the_compiler_stumbles_on_is_dropped():
    project = workbook().doc().wb.vba
    # Fine, and fine while the code around it is unfinished.
    assert not vba_ai.breaks(project, "Module1", 'Sub A()\n    Range("A1").', "Value = 1", "\nEnd Sub\n")
    assert not vba_ai.breaks(project, "Module1", "Sub A()\n    For i = 1 To ", "3", "\nEnd Sub\n")
    assert not vba_ai.breaks(project, "Module1", 'Sub A()\n    Set ws = Worksheets(', '"訂單"', "\nEnd Sub\n")
    # An error on a line it brings.
    assert vba_ai.breaks(project, "Module1", "Sub A()\n    x = 1", "\n    Sub Inner()", "\nEnd Sub\n")
    assert vba_ai.breaks(project, "Module1", "Sub A()\n", "    x = = 1", "\nEnd Sub\n")


def test_an_error_in_another_module_is_not_this_answers():
    s = workbook()
    s.run("macro.setModule", name="Other", code="Sub Broken(\n")
    assert vba_ai.compile_problems(s.doc().wb.vba, "Module1", "Sub Good()\nEnd Sub\n") == []
