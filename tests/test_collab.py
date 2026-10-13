"""Several people on one workbook: shared cells, separate views, and undo that takes back only
your own changes."""

import pytest

from gridcore.engine import CollabSession, DocState, EngineError, SharedBook, geometry, viewport
from gridcore.core.addr import RangeRef
from gridcore.model.workbook import Workbook


def new_book():
    return SharedBook(DocState(Workbook(), "Shared"), "b1")


def join(b, key):
    return b.join(key, CollabSession(), key.title())


def run(b, v, cid, **p):
    with b.act(v) as s:
        return s.execute(cid, p)


def val(b, v, a):
    return run(b, v, "cell.get", cell=a)["value"]


def bold(b, v, a):
    with b.act(v) as s:
        d = s.doc()
        return d.wb.styles.get(d.sheet.style_id(_rc(a))).font.bold


def _rc(a):
    r = RangeRef.parse(a)
    return (r.r0, r.c0)


def test_everyone_edits_the_same_cells():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    run(b, amy, "cell.set", cell="A1", input="1")
    run(b, ben, "cell.set", cell="B1", input="=A1+1")
    assert val(b, amy, "B1") == 2
    run(b, amy, "cell.set", cell="A1", input="10")
    assert val(b, ben, "B1") == 11
    assert b.rev == 3
    assert [c.author for c in b.changes] == ["amy", "ben", "amy"]
    assert b.changes[-1].areas == [("Sheet1", [0, 0, 0, 0])]


def test_reading_changes_nothing():
    b = new_book()
    amy = join(b, "amy")
    run(b, amy, "cell.get", cell="A1")
    run(b, amy, "selection.set", range="C3")
    assert b.rev == 0 and not amy.undo


def test_each_person_has_their_own_sheet_and_selection():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    run(b, ben, "selection.set", range="C3")
    run(b, amy, "home.insertSheet", index=1)
    assert amy.sheet == "Sheet2" and ben.sheet == "Sheet1"
    assert ben.selection.active == (2, 2)
    # Without a cell, a command works on the person's own selection and sheet.
    run(b, ben, "cell.set", input="ben")
    run(b, amy, "cell.set", input="amy")
    with b.act(ben) as s:
        wb = s.doc().wb
        assert wb.sheets[0].value((2, 2)) == "ben"
        assert wb.sheets[1].value((0, 0)) == "amy"
    run(b, ben, "sheet.activate", sheet="Sheet2")
    assert ben.sheet == "Sheet2" and amy.sheet == "Sheet2"
    run(b, ben, "sheet.activate", sheet="Sheet1")
    assert ben.selection.active == (2, 2), "the selection on a sheet is remembered"


def test_undo_takes_back_only_your_own_changes():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    run(b, amy, "cell.set", cell="A1", input="1")
    run(b, ben, "cell.set", cell="B1", input="2")
    run(b, amy, "cell.set", cell="A2", input="3")
    run(b, amy, "edit.undo")
    assert val(b, amy, "A2") is None and val(b, amy, "B1") == 2 and val(b, amy, "A1") == 1
    run(b, amy, "edit.undo")
    assert val(b, amy, "A1") is None and val(b, amy, "B1") == 2
    run(b, ben, "edit.undo")
    assert val(b, ben, "B1") is None
    with pytest.raises(EngineError):
        run(b, ben, "edit.undo")
    run(b, amy, "edit.redo")
    assert val(b, amy, "A1") == 1
    run(b, amy, "edit.redo")
    assert val(b, amy, "A2") == 3 and not amy.redo


def test_undo_keeps_what_others_changed_since():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    run(b, amy, "range.setValues", range="A1", values=[[1, 2]])
    run(b, ben, "cell.set", cell="A1", input="9")
    r = run(b, amy, "edit.undo")
    assert r["kept"] == 1
    assert val(b, amy, "A1") == 9, "Ben's newer value stays"
    assert val(b, amy, "B1") is None, "the rest of Amy's change is undone"
    assert amy.session.messages and "kept" in amy.session.messages[-1]


def test_formulas_recalculate_after_a_selective_undo():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    run(b, ben, "cell.set", cell="C1", input="=A1*2")
    run(b, amy, "cell.set", cell="A1", input="5")
    run(b, ben, "cell.set", cell="D1", input="x")
    assert val(b, ben, "C1") == 10
    run(b, amy, "edit.undo")
    assert val(b, ben, "C1") == 0 and val(b, ben, "D1") == "x"


def test_formats_rows_and_merges_undo_selectively():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    run(b, amy, "home.bold", range="A1")
    run(b, amy, "home.mergeCells", range="B2:C2")
    run(b, amy, "home.rowHeight", rows="5", height=40)
    run(b, ben, "cell.set", cell="E9", input="ben")
    run(b, amy, "edit.undo")
    run(b, amy, "edit.undo")
    run(b, amy, "edit.undo")
    with b.act(amy) as s:
        sh = s.doc().sheet
        assert sh.merges == [] and sh.row_height(4) == sh.default_row_height
    assert not bold(b, amy, "A1")
    assert val(b, ben, "E9") == "ben"


def test_exact_undo_of_your_own_last_change_whatever_it_was():
    b = new_book()
    amy = join(b, "amy")
    run(b, amy, "cell.set", cell="A2", input="x")
    run(b, amy, "home.insertRows", rows="1:1")
    assert val(b, amy, "A3") == "x"
    run(b, amy, "edit.undo")
    assert val(b, amy, "A2") == "x" and val(b, amy, "A3") is None
    run(b, amy, "edit.redo")
    assert val(b, amy, "A3") == "x"


def test_moved_cells_stop_older_undo_until_the_move_is_undone():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    run(b, amy, "cell.set", cell="A1", input="amy")
    run(b, ben, "home.insertRows", rows="1:1")
    assert val(b, amy, "A2") == "amy"
    with pytest.raises(EngineError, match="moved cells"):
        run(b, amy, "edit.undo")
    assert val(b, amy, "A2") == "amy" and amy.undo, "nothing happened and the step is still there"
    run(b, ben, "edit.undo")  # Ben's insert was the last change: exact undo
    run(b, amy, "edit.undo")
    assert val(b, amy, "A1") is None


def test_a_sheet_added_at_the_end_moves_nothing():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    run(b, amy, "cell.set", cell="A1", input="amy")
    run(b, ben, "home.insertSheet", index=1)
    run(b, amy, "edit.undo")
    assert val(b, amy, "A1") is None
    run(b, amy, "edit.redo")
    run(b, ben, "home.insertSheet", index=0)
    with pytest.raises(EngineError, match="moved cells"):
        run(b, amy, "edit.undo")


def test_a_structural_change_by_you_cant_be_undone_after_others_edit():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    run(b, amy, "home.insertRows", rows="2:2")
    run(b, ben, "cell.set", cell="D4", input="1")
    with pytest.raises(EngineError):
        run(b, amy, "edit.undo")
    assert val(b, ben, "D4") == 1


def test_a_new_change_clears_your_redo_but_not_others():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    run(b, amy, "cell.set", cell="A1", input="1")
    run(b, ben, "cell.set", cell="B1", input="2")
    run(b, amy, "edit.undo")
    run(b, ben, "edit.undo")
    run(b, ben, "cell.set", cell="C1", input="3")
    assert amy.redo and not ben.redo
    run(b, amy, "edit.redo")
    assert val(b, amy, "A1") == 1


def test_undo_returns_to_where_you_were():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    run(b, amy, "selection.set", range="D4")
    run(b, amy, "cell.set", input="x")
    run(b, amy, "selection.set", range="A10")
    run(b, ben, "cell.set", cell="Z1", input="y")
    run(b, amy, "edit.undo")
    assert amy.selection.active == (3, 3)


def test_renamed_and_deleted_sheets():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    run(b, amy, "selection.set", range="E5")
    run(b, ben, "sheet.rename", sheet=0, name="Data")
    assert amy.sheet == "Data"
    run(b, amy, "cell.get", cell="A1")
    assert amy.selection.active == (4, 4), "a rename keeps people where they were"
    run(b, ben, "home.insertSheet", index=1, name="Notes")
    run(b, amy, "sheet.activate", sheet="Notes")
    run(b, ben, "home.deleteSheet", sheet="Notes")
    run(b, amy, "cell.set", input="here")
    assert amy.sheet == "Data", "a deleted sheet sends people to a neighbour"
    with b.act(ben) as s:
        assert s.doc().wb.sheets[0].value((4, 4)) == "here", "where they last were on it"


def test_zoom_and_formulas_view_belong_to_each_person():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    run(b, amy, "cell.set", cell="A1", input="=1+1")
    run(b, amy, "view.zoom", percent=150)
    run(b, amy, "formulas.showFormulas")
    assert b.rev == 1, "view changes don't change the workbook"
    with b.act(amy) as s:
        sh = s.doc().sheet
        assert sh.zoom == 100 and amy.zoom_of(sh) == 150 and amy.formulas_of(sh)
        vp = viewport(s.doc().wb, 0, RangeRef(0, 0, 5, 5), show_formulas=amy.formulas_of(sh))
        assert vp["cells"][0]["t"] == "=1+1"
    with b.act(ben) as s:
        sh = s.doc().sheet
        assert ben.zoom_of(sh) == 100 and not ben.formulas_of(sh)
        assert geometry(sh)["zoom"] == 100


def test_files_are_handled_by_the_app():
    b = new_book()
    amy = join(b, "amy")
    for cid, p in (("file.open", {"path": "/etc/hosts"}), ("file.saveAs", {"path": "/tmp/x.xlsx"}), ("file.new", {}), ("window.activate", {"index": 0})):
        with pytest.raises(EngineError):
            run(b, amy, cid, **p)


def test_commands_run_inside_act_only():
    b = new_book()
    amy = join(b, "amy")
    with pytest.raises(EngineError):
        amy.session.execute("cell.set", {"cell": "A1", "input": "1"})
    with b.act(amy):
        with pytest.raises(RuntimeError):
            with b.act(amy):
                pass


def test_listeners_hear_every_change():
    b = new_book()
    amy = join(b, "amy")
    heard = []
    b.listeners.append(lambda book: heard.append(book.rev))
    run(b, amy, "cell.set", cell="A1", input="1")
    run(b, amy, "selection.set", range="B2")
    run(b, amy, "edit.undo")
    assert heard == [1, 2]


def test_idle_people_are_dropped_and_can_come_back():
    b = new_book()
    amy, ben = join(b, "amy"), join(b, "ben")
    ben.last_seen -= 100
    assert b.prune(30) == ["ben"]
    assert list(b.views) == ["amy"]
    run(b, ben, "cell.set", cell="A1", input="back")
    assert "ben" in b.views and val(b, amy, "A1") == "back"


def test_a_tab_that_reloads_takes_its_view_back():
    b = new_book()
    amy = join(b, "amy")
    run(b, amy, "cell.set", cell="A1", input="1")
    b.leave("amy")
    again = b.join("amy", amy.session, "Amy", view=amy)
    assert again is amy and again.undo
    run(b, again, "edit.undo")
    assert val(b, again, "A1") is None


def test_failed_command_keeps_the_workbook():
    b = new_book()
    amy = join(b, "amy")
    run(b, amy, "cell.set", cell="A1", input="1")
    with pytest.raises(EngineError):
        run(b, amy, "cell.set", cell="A2", input="=IF(1)")
    assert b.rev == 1 and len(amy.undo) == 1
