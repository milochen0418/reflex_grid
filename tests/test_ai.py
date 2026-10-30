"""AI through the AI Hub: how AI Fill and formula suggestions word their questions and read the
answers (reflex_grid/ai.py), and the tools the hub's assistant reads workbooks with
(reflex_grid/ai_tools.py). No model is asked here; the browser suite ai_features drives both
features against a stand-in hub."""

import asyncio
from types import SimpleNamespace

import pytest

from gridcore.core.addr import RangeRef
from gridcore.engine import CollabSession
from reflex_grid import ai, ai_tools
from reflex_grid.books import BookStore


# ---------------------------------------------------------------- AI Fill


def test_fill_prompt_names_each_cell_and_numbers_the_rows():
    text = ai.fill_prompt("Positive or negative?", ["Comment", "Stars"], [["Great pen", "5"], ["", ""], ["Leaks\nbadly", "1"]], title=True)
    assert "Instruction: Positive or negative?" in text
    assert "Title:" in text
    assert "1. Comment: Great pen | Stars: 5" in text
    assert "2. (empty)" in text
    assert "3. Comment: Leaks badly | Stars: 1" in text


def test_parse_fill_reads_numbered_lines_and_the_title():
    reply = "Title: **Sentiment**\n1. Positive\n2)  \"Negative\"\n3. -\n\n7. out of range\n2. again"
    title, answers = ai.parse_fill(reply, 3)
    assert title == "Sentiment"
    assert answers == {1: "Positive", 2: "Negative", 3: ""}
    # Rows left out are missing; a reply without numbers but one line per row is taken in order.
    assert ai.parse_fill("1. Yes\n3. No", 3)[1] == {1: "Yes", 3: "No"}
    assert ai.parse_fill("Yes\nNo", 2)[1] == {1: "Yes", 2: "No"}
    assert ai.parse_fill("```\n1：好\n2、壞\n```", 2)[1] == {1: "好", 2: "壞"}


def test_answers_go_in_as_values_never_as_formulas():
    assert ai.as_input("42") == "42"
    assert ai.as_input("Positive") == "Positive"
    assert ai.as_input("=1+1") == "'=1+1"
    assert ai.as_input("-abc") == "'-abc"
    assert ai.as_input("a\nb") == "a b"


def sheet_with(rows):
    s = CollabSession()
    s.new_workbook()
    s.run("range.setValues", range="A1", values=rows)
    wb = s.doc().wb
    return wb, wb.sheet(0)


def test_headers_names_and_a_free_column_for_the_results():
    wb, sh = sheet_with([["Comment", "Stars", "Note"], ["Great", 5, "x"], ["Bad", 1, None]])
    rng = RangeRef.parse("A1:B3")
    assert ai.looks_like_header(sh, rng)
    assert ai.column_names(wb, sh, rng, True) == ["Comment", "Stars"]
    # Selected under the headers: the row above names the columns.
    below = RangeRef.parse("A2:B3")
    assert not ai.looks_like_header(sh, below)
    assert ai.column_names(wb, sh, below, False) == ["Comment", "Stars"]
    # Column C holds a value in row 2, so the results go to D.
    assert ai.free_column(sh, rng, (1, 2)) == 3
    wb2, sh2 = sheet_with([["This pen leaks all over my shirt and I want my money back"], ["Fine"]])
    assert not ai.looks_like_header(sh2, RangeRef.parse("A1:A2"))
    assert ai.column_names(wb2, sh2, RangeRef.parse("A1:A2"), False) == ["A"]


# ---------------------------------------------------------------- formula suggestions


def test_neighbourhood_shows_the_cells_around_with_formulas_and_far_headers():
    rows = [["Item", "Qty", "Price", "Total"]] + [[f"I{i}", i, 2, f"=B{i + 2}*C{i + 2}"] for i in range(20)]
    wb, sh = sheet_with(rows)
    text = ai.neighbourhood(wb, 0, (21, 3))
    assert 'Sheet "Sheet1", cell D22' in text
    assert "\nA\tB\tC\tD" not in text and "\tA\tB\tC\tD" in text
    assert "1\tItem\tQty\tPrice\tTotal" in text  # the headers, though row 1 is out of the window
    assert "[=B21*C21]" in text
    assert "<<here>>" in text


def test_parse_suggestion():
    assert ai.parse_suggestion("FORMULA: =SUM(B2:B9)\nWHY: Adds the sales.") == ("=SUM(B2:B9)", "Adds the sales.")
    assert ai.parse_suggestion("**Formula:** `=AVERAGE(C2:C5)`\n**Why:** mean price") == ("=AVERAGE(C2:C5)", "mean price")
    assert ai.parse_suggestion("```excel\n=B2*C2\n```") == ("=B2*C2", "")
    assert ai.parse_suggestion("I can't tell.") == ("", "")
    assert "corrected" in ai.suggest_prompt("ctx", "=FOO(1) gives #NAME?")


# ---------------------------------------------------------------- the assistant's tools


@pytest.fixture
def store(tmp_path, monkeypatch):
    s = BookStore(tmp_path)
    monkeypatch.setattr(ai_tools, "STORE", s)
    return s


def run(coro):
    return asyncio.run(coro)


def test_tools_read_a_workbook(store):
    book_id = store.create_sample("Sales").id
    overview = run(ai_tools.workbook(None, book_id))
    assert overview["title"] == "Sales"
    sales = overview["sheets"][0]
    assert sales["name"] == "Sales" and sales["tables"][0]["columns"][:2] == ["Region", "Product"]
    assert overview["vba"] == []
    assert "runs Excel VBA itself" in run(ai_tools.guide(None))
    text = run(ai_tools.read(None, book_id, "Sales!H1:I3"))
    assert text.splitlines()[0].startswith("Sales!H1:I3 (3 rows × 2 columns")
    assert "2\tNorth\t$5,390.25 [=SUMIFS(Table1[Revenue],Table1[Region],H2)]" in text
    assert "[=" not in run(ai_tools.read(None, book_id, "H1:I3", formulas=False))
    found = run(ai_tools.find(None, book_id, "sumifs"))
    assert found.splitlines()[0].startswith("Sales!I2\t")


def test_a_big_read_comes_in_parts(store, monkeypatch):
    monkeypatch.setattr(ai_tools, "READ_MAX_CELLS", 12)
    book_id = store.create_sample("Sales").id
    text = run(ai_tools.read(None, book_id, "A1:F25"))
    assert text.splitlines()[0] == "Sales!A1:F2 (2 rows × 6 columns, formulas in brackets); 23 more rows: read Sales!A3:F25 next; chart 1 covers K1:R15"


def test_tools_say_what_went_wrong(store):
    book_id = store.create_sample("Sales").id
    with pytest.raises(ValueError, match="no workbook"):
        run(ai_tools.workbook(None, "nope"))
    with pytest.raises(ValueError, match="no sheet named 'Nope'"):
        run(ai_tools.read(None, book_id, "Nope!A1"))
    with pytest.raises(ValueError, match="isn't a reference"):
        run(ai_tools.read(None, book_id, "not a ref"))
    assert run(ai_tools.read(None, book_id, "Z100:Z200")).endswith("empty (the sheet's cells end at A1:J25).")
    assert run(ai_tools.find(None, book_id, "nowhere to be found")) == "No cell contains 'nowhere to be found'."


# ---------------------------------------------------------------- the tools that change a workbook


@pytest.fixture
def tab(store, monkeypatch):
    """A workbook with a person's tab on it (the room names the tab), and what was published."""

    async def fake_push(book_id):
        pushed.append(book_id)

    pushed = []
    monkeypatch.setattr(ai_tools, "push", fake_push)
    meta = store.create_blank()
    book = store.open(meta.id)
    view = book.join("tab1", CollabSession(), "Milo")
    with book.act(view) as s:
        s.run("range.setValues", range="A1", values=[["Item", "Qty"], ["Pen", 3], ["Ink", 5]])
    return SimpleNamespace(book=book, view=view, room=f"{meta.id}~tab1", pushed=pushed)


def cell(t, a1):
    with t.book.act(t.view) as s:
        return s.run("cell.get", cell=a1)


def undo(t):
    with t.book.act(t.view) as s:
        s.run("edit.undo")


def step(command, **params):
    return {"command": command, "params": params}


def test_edit_runs_the_commands_as_the_tab_in_one_undo_step(tab):
    out = run(ai_tools.edit(None, tab.room, "Add a total", [step("cell.set", cell="A4", input="Total"), step("cell.set", cell="B4", input="=SUM(B2:B3)"), step("home.bold", range="A4:B4")]))
    assert out["done"] == "3 of 3 steps" and out["changed"] == ["Sheet1!A4:B4"] and "failed" not in out
    assert cell(tab, "B4")["value"] == 8
    assert tab.pushed == [tab.book.key]
    undo(tab)  # the person's Ctrl+Z takes all of it back
    assert cell(tab, "A4")["value"] is None and cell(tab, "B4")["value"] is None
    assert cell(tab, "B3")["value"] == 5


def test_edit_refuses_what_isnt_the_ais_before_doing_anything(tab):
    for command, word in (("file.saveAs", "isn't open"), ("edit.undo", "isn't open"), ("macro.run", "grid_vba_set"), ("cell.sett", "did you mean cell.set")):
        with pytest.raises(ValueError, match=word):
            run(ai_tools.edit(None, tab.room, "x", [step("cell.set", cell="C1", input="no"), step(command)]))
    assert cell(tab, "C1")["value"] is None
    with pytest.raises(ValueError, match="no steps"):
        run(ai_tools.edit(None, tab.room, "x", []))


def test_a_failed_step_leaves_everything_as_it_was_and_says_why(tab):
    rev, wb, undos = tab.book.rev, tab.book.doc.wb, len(tab.view.undo)
    out = run(ai_tools.edit(None, tab.room, "x", [step("cell.set", cell="C1", input="kept?"), step("home.insertSheet", name="Extra"), step("cell.set", cell="C2", input="=1+"), step("cell.set", cell="C3", input="late")]))
    assert out["done"] == "0 of 4 steps" and "step 3 (cell.set)" in out["failed"] and "nothing was changed" in out["failed"]
    assert "changed" not in out and "results" not in out
    assert tab.book.rev == rev and tab.book.doc.wb is wb and len(tab.view.undo) == undos
    assert cell(tab, "C1")["value"] is None and tab.view.sheet == "Sheet1" and tab.book.ai_changes == []


def test_error_values_are_reported(tab):
    out = run(ai_tools.edit(None, tab.room, "x", [step("cell.set", cell="C1", input="=1/0")]))
    assert out["error_values"] == ["Sheet1!C1: #DIV/0!"]


def test_what_a_formula_spills_is_read_and_counts_as_changed(tab):
    out = run(ai_tools.edit(None, tab.room, "Summary", [step("cell.set", cell="D1", input="=GROUPBY(A2:A3,B2:B3,SUM)")]))
    assert out["changed"] == ["Sheet1!D1:E3"]
    text = run(ai_tools.read(None, tab.room, "Sheet1!D1:E3"))
    assert "1\tInk [=GROUPBY(A2:A3,B2:B3,SUM)]\t5" in text and "2\tPen\t3" in text and "3\tTotal\t8" in text


def test_a_change_from_a_tab_that_is_gone_is_made_as_ai(tab):
    out = run(ai_tools.edit(None, tab.room.split("~")[0] + "~gone", "x", [step("cell.set", cell="C1", input="hi")]))
    assert out["done"] == "1 of 1 steps" and cell(tab, "C1")["value"] == "hi"
    assert tab.book.ai_view.name == "AI" and tab.book.ai_view.undo and "ai" not in tab.book.views
    assert not any(t.label == "Enter Cell" for t in tab.view.undo[1:])


def test_an_accepted_change_is_the_persons_and_shows_as_the_ais(tab):
    run(ai_tools.edit(None, tab.room, "Add a total", [step("cell.set", cell="A4", input="Total"), step("cell.set", cell="B4", input="=SUM(B2:B3)")]))
    change = tab.book.ai_changes[-1]
    assert change.peer == "tab1" and change.name == "AI for Milo" and change.summary == "Add a total" and change.areas == [("Sheet1", [3, 0, 3, 1])]
    assert tab.view.undo[-1] is change.step  # on the person's undo, not the AI's
    # Everyone's grid lights it up in the AI's colour, the person's too.
    rev = change.revs[-1]
    row = ai.seen_change(tab.book, {"rev": rev, "author": "tab1", "name": "Milo", "color": "#123456", "areas": []})
    assert (row["author"], row["name"], row["color"], row["areas"]) == ("ai", "AI for Milo", ai.AI_COLOR, [["Sheet1", [3, 0, 3, 1]]])
    other = {"rev": rev - 1, "author": "tab1", "name": "Milo", "color": "#123456", "areas": []}
    assert ai.seen_change(tab.book, other) == other
    # Its box is on the grid for a moment; it isn't one of the people here.
    box = ai.presence(tab.book, now=change.at + 1)
    assert len(box) == 1 and box[0]["id"] == "ai" and box[0]["sheet"] == "Sheet1" and box[0]["ranges"] == [[3, 0, 3, 1]] and box[0]["color"] == ai.AI_COLOR
    assert ai.presence(tab.book, now=change.at + ai.AI_MARK_SECONDS + 1) == []
    assert list(tab.book.views) == ["tab1"]


def test_a_note_says_what_its_person_can_take_back_while_that_stands(tab):
    run(ai_tools.edit(None, tab.room, "First", [step("cell.set", cell="C1", input="a")]))
    run(ai_tools.edit(None, tab.room, "Second", [step("cell.set", cell="C2", input="b")]))
    first, second = tab.book.ai_changes
    assert ai.notes(tab.book) == [{"rev": second.revs[-1], "peer": "tab1", "summary": "Second"}]
    undo(tab)  # Ctrl+Z: the note goes back to the change before, and the AI's box leaves
    assert [n["summary"] for n in ai.notes(tab.book)] == ["First"]
    assert ai.presence(tab.book, now=first.at) == []
    first.closed = True
    assert ai.notes(tab.book) == []
    with tab.book.act(tab.view) as s:
        s.run("edit.redo")
    assert ai.notes(tab.book) == []  # a redo is the person's own step


def test_nothing_is_noted_when_nothing_changed(tab):
    run(ai_tools.edit(None, tab.room, "Look around", [step("selection.set", range="B2")]))
    assert tab.book.ai_changes == []
    out = run(ai_tools.edit(None, tab.room, "x", [step("cell.set", cell="not a cell!", input="x")]))
    assert out["done"] == "0 of 1 steps" and "failed" in out
    assert tab.book.ai_changes == []


def test_a_change_with_moved_cells_says_where_it_landed(tab):
    out = run(ai_tools.edit(None, tab.room, "Sort", [step("data.sort", range="A1:B3", header=True, keys=[{"col": 1, "order": "desc"}])]))
    assert cell(tab, "A2")["value"] == "Ink"
    assert out["changed"] == ["Sheet1!A2:B3"]
    assert tab.book.ai_changes[-1].areas == [("Sheet1", [1, 0, 2, 1])]
    assert tab.book.changes[-1].areas == []  # the log itself has no areas once cells moved


def test_vba_and_forms_are_noted_too(tab):
    run(ai_tools.vba_set(None, tab.room, "Add a macro", "Module1", "Sub Hello()\nEnd Sub\n"))
    assert [c.summary for c in tab.book.ai_changes] == ["Add a macro"]
    assert ai.notes(tab.book)[0]["summary"] == "Add a macro"


def test_the_ais_own_change_is_anyones_to_take_back(tab):
    run(ai_tools.edit(None, tab.room.split("~")[0] + "~gone", "From nowhere", [step("cell.set", cell="C1", input="hi")]))
    assert [(n["peer"], n["summary"]) for n in ai.notes(tab.book)] == [("ai", "From nowhere")]
    assert ai.seen_change(tab.book, {"rev": tab.book.rev, "author": "ai", "name": "", "color": "", "areas": []})["name"] == "AI"


def test_the_workbook_says_which_cells_a_chart_covers(tab):
    run(ai_tools.edit(None, tab.room, "Chart", [step("insert.chart", kind="line", range="A1:B3", title="Trend")]))
    charts = run(ai_tools.workbook(None, tab.room))["sheets"][0]["charts"]
    assert len(charts) == 1 and charts[0]["title"] == "Trend" and charts[0]["covers"].startswith("D1:")
    run(ai_tools.edit(None, tab.room, "Move it", [step("object.move", id=charts[0]["id"], row=4, col=7)]))
    assert run(ai_tools.workbook(None, tab.room))["sheets"][0]["charts"][0]["covers"].startswith("H5:")


def test_cells_a_chart_hides_are_reported(tab):
    out = run(ai_tools.edit(None, tab.room, "Chart", [step("insert.chart", kind="line", range="A1:B3", title="Trend")]))
    assert "hidden" not in out
    out = run(ai_tools.edit(None, tab.room, "Summary under the chart", [step("cell.set", cell="E2", input="x")]))
    assert "Trend" in out["hidden"][0] and "Sheet1!D1:" in out["hidden"][0]
    assert "hidden" not in run(ai_tools.edit(None, tab.room, "Beside it", [step("cell.set", cell="A20", input="y")]))
    # Moved out from under it: nothing is hidden, though the change lies on both sides of the chart.
    out = run(ai_tools.edit(None, tab.room, "Move it", [step("edit.clearContents", range="E2"), step("cell.set", cell="A21", input="x")]))
    assert "hidden" not in out and sorted(out["changed"]) == ["Sheet1!A21", "Sheet1!E2"]


def test_reading_says_what_charts_cover(tab):
    run(ai_tools.edit(None, tab.room, "Chart", [step("insert.chart", kind="line", range="A1:B3", title="Trend")]))
    assert "chart 1 covers D1:" in run(ai_tools.read(None, tab.room, "Sheet1!A1:B2")).splitlines()[0]


def test_changes_that_lie_apart_stay_apart(tab):
    out = run(ai_tools.edit(None, tab.room, "Two places", [step("cell.set", cell="D1", input="a"), step("cell.set", cell="E1", input="b"), step("cell.set", cell="A9", input="c")]))
    assert sorted(out["changed"]) == ["Sheet1!A9", "Sheet1!D1:E1"]
    assert sorted(b for _, b in tab.book.ai_changes[-1].areas) == [[0, 3, 0, 4], [8, 0, 8, 0]]
    assert sorted(b for b in ai.presence(tab.book)[0]["ranges"]) == [[0, 3, 0, 4], [8, 0, 8, 0]]
    # Scattered, one box around them.
    far = [step("cell.set", cell=f"H{i * 3 + 1}", input="x") for i in range(10)]
    assert run(ai_tools.edit(None, tab.room, "Scattered", far))["changed"] == ["Sheet1!H1:H28"]


def test_a_macro_comes_back_out_after_the_person_went_on(tab):
    run(ai_tools.vba_set(None, tab.room, "Add a macro", "Module1", "Sub Hello()\nEnd Sub\n"))
    change = tab.book.ai_changes[-1]
    with tab.book.act(tab.view) as s:
        s.run("cell.set", cell="F1", input="mine")
    with tab.book.act(tab.view):
        tab.book.take_back(tab.view, change.step)
    assert cell(tab, "F1")["value"] == "mine"
    assert not run(ai_tools.workbook(None, tab.room))["vba"]
    assert ai.notes(tab.book) == []


def test_a_new_chart_that_lands_on_cells_says_so(tab):
    run(ai_tools.edit(None, tab.room, "Notes", [step("cell.set", cell="E2", input="a note")]))
    out = run(ai_tools.edit(None, tab.room, "Chart", [step("insert.chart", kind="line", range="A1:B3", title="Trend")]))
    assert "Trend" in out["hidden"][0]


def test_commands_lists_what_grid_edit_can_run(store):
    all_of_them = run(ai_tools.commands(None))
    assert "cell.set — " in all_of_them and "home.insertRows — " in all_of_them
    assert "macro.run" not in all_of_them and "file.save" not in all_of_them and "edit.undo" not in all_of_them
    sort = run(ai_tools.commands(None, "sort"))
    assert "data.sort — " in sort and "cell.set" not in sort
    assert "No command" in run(ai_tools.commands(None, "zzz"))


def test_vba_set_makes_and_replaces_a_module_and_reports_compile_errors(tab):
    with pytest.raises(ValueError, match="no VBA yet"):
        run(ai_tools.vba_get(None, tab.room, "Module1"))
    out = run(ai_tools.vba_set(None, tab.room, "Add a macro", "Helpers", "Sub Hi()\n    Range(\"D1\").Value = \"hi\"\nEnd Sub\n"))
    assert out == {"module": "Helpers", "made": True, "compile_errors": []}
    got = run(ai_tools.vba_get(None, tab.room, "Helpers"))
    assert got["kind"] == "standard" and 'Range("D1")' in got["code"]
    out = run(ai_tools.vba_set(None, tab.room, "Break it", "Helpers", "Sub Bad()\n    If x Then\nEnd Sub\n"))
    assert out["made"] is False and out["compile_errors"] and out["compile_errors"][0].startswith("Helpers line ")
    undo(tab)
    assert "Sub Hi()" in run(ai_tools.vba_get(None, tab.room, "Helpers"))["code"]
    assert any(m["name"] == "Helpers" for m in run(ai_tools.workbook(None, tab.room))["vba"])
    with pytest.raises(ValueError, match="no module named 'Nope'"):
        run(ai_tools.vba_get(None, tab.room, "Nope"))


FORM_CODE = '''Private Sub btnAdd_Click()
    Dim ws As Worksheet, r As Long
    Set ws = Worksheets(1)
    r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(r, 1).Value = txtName.Text
    ws.Cells(r, 2).Value = CDbl(txtQty.Text)
    ws.Cells(r, 3).Formula = "=B" & r & "*2"
    Unload Me
End Sub
'''


def test_the_ai_builds_a_userform_that_adds_a_row(tab):
    """The request that started this: a form to type a new row into, made by the AI's three
    calls, then run by the person."""
    out = run(ai_tools.form_design(None, tab.room, "A form to add an item", "frmAddItem", caption="Add item", width=240, height=120, controls=[
        {"name": "lblName", "kind": "Label", "caption": "Item", "left": 10, "top": 12, "width": 60},
        {"name": "txtName", "kind": "TextBox", "left": 80, "top": 10, "width": 140},
        {"name": "txtQty", "kind": "TextBox", "left": 80, "top": 36, "width": 140, "text": "1"},
        {"name": "btnAdd", "kind": "CommandButton", "caption": "Add", "left": 80, "top": 70, "default": True},
    ]))
    assert out["controls"] == ["lblName", "txtName", "txtQty", "btnAdd"] and out["caption"] == "Add item" and out["compile_errors"] == []
    assert run(ai_tools.vba_set(None, tab.room, "The form's button", "frmAddItem", FORM_CODE, kind="form"))["made"] is False
    assert run(ai_tools.vba_set(None, tab.room, "A macro to open it", "Module1", "Sub AddItem()\n    frmAddItem.Show\nEnd Sub\n"))["compile_errors"] == []
    got = run(ai_tools.vba_get(None, tab.room, "frmAddItem"))
    assert got["kind"] == "form" and got["form"]["caption"] == "Add item" and got["form"]["controls"][3]["caption"] == "Add"
    with tab.book.act(tab.view) as s:
        r = s.run("macro.run", name="AddItem", answers=[{"event": "change", "control": "txtName", "value": "Cap"}, {"event": "change", "control": "txtQty", "value": "7"}, {"event": "click", "control": "btnAdd"}])
    assert r["status"] == "done", r.get("error")
    assert cell(tab, "A4")["value"] == "Cap" and cell(tab, "B4")["value"] == 7 and cell(tab, "C4")["value"] == 14
    # Changed later, by name: what isn't given stays; a removed control goes.
    out = run(ai_tools.form_design(None, tab.room, "Rename the button's caption", "frmAddItem", controls=[{"name": "btnAdd", "caption": "Save"}], remove=["lblName"]))
    assert out["controls"] == ["txtName", "txtQty", "btnAdd"]
    design = run(ai_tools.vba_get(None, tab.room, "frmAddItem"))["form"]
    assert design["controls"][2]["caption"] == "Save" and abs(design["controls"][2]["left"] - 80) <= 0.02
    with pytest.raises(ValueError, match="say its kind"):
        run(ai_tools.form_design(None, tab.room, "x", "frmAddItem", controls=[{"name": "newOne"}]))
