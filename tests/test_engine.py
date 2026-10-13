"""Engine commands end to end: entering data, recalculation, undo/redo, clipboard, structure
edits, formatting, sort/filter, tables, names, conditional formats and the viewport."""

import pytest

from gridcore.engine import EngineError, Session, geometry, objects, ui_state, viewport
from gridcore.core.addr import RangeRef


def new():
    s = Session()
    s.new_workbook()
    return s


def val(s, a):
    return s.run("cell.get", cell=a)["value"]


def text(s, a):
    return s.run("cell.get", cell=a)["text"]


def inp(s, a):
    return s.run("cell.get", cell=a)["input"]


def test_enter_and_recalc():
    s = new()
    s.run("cell.set", cell="A1", input="10")
    s.run("cell.set", cell="A2", input="32")
    s.run("cell.set", cell="A3", input="=SUM(A1:A2")  # missing paren closed
    assert inp(s, "A3") == "=SUM(A1:A2)"
    assert val(s, "A3") == 42
    s.run("cell.set", cell="A1", input="20")
    assert val(s, "A3") == 52
    s.run("cell.set", cell="B1", input="50%")
    assert val(s, "B1") == 0.5 and text(s, "B1") == "50%"
    s.run("cell.set", cell="B2", input="3/14/2026")
    assert text(s, "B2") == "3/14/2026"
    s.run("cell.set", cell="B3", input="'123")
    assert val(s, "B3") == "123" and inp(s, "B3") == "'123"
    with pytest.raises(EngineError):
        s.run("cell.set", cell="C1", input="=IF(1)")


def test_undo_redo():
    s = new()
    s.run("cell.set", cell="A1", input="1")
    s.run("cell.set", cell="A1", input="2")
    s.run("edit.undo")
    assert val(s, "A1") == 1
    s.run("edit.redo")
    assert val(s, "A1") == 2
    s.run("edit.undo", steps=2)
    assert val(s, "A1") is None


def test_copy_paste_relative():
    s = new()
    s.run("range.setValues", range="A1", values=[[1, 2], [3, 4]])
    s.run("cell.set", cell="C1", input="=A1+B1")
    s.run("edit.copy", range="C1")
    s.run("selection.set", range="C2")
    s.run("edit.paste")
    assert inp(s, "C2") == "=A2+B2" and val(s, "C2") == 7
    # External text pastes as values.
    s.run("selection.set", range="E1")
    s.run("edit.paste", text="x\t5\n6\t=E2*2\n")
    assert val(s, "E1") == "x" and val(s, "F1") == 5 and val(s, "F2") == 12


def test_cut_paste_moves_references():
    s = new()
    s.run("cell.set", cell="A1", input="5")
    s.run("cell.set", cell="B1", input="=A1*2")
    s.run("edit.cut", range="A1")
    s.run("selection.set", range="D4")
    s.run("edit.paste")
    assert inp(s, "B1") == "=D4*2" and val(s, "B1") == 10
    assert val(s, "A1") is None


def test_insert_delete_rows_adjust_formulas():
    s = new()
    s.run("range.setValues", range="A1", values=[[1], [2], [3]])
    s.run("cell.set", cell="B1", input="=SUM(A1:A3)")
    s.run("home.insertRows", rows="2")
    assert inp(s, "B1") == "=SUM(A1:A4)"
    assert val(s, "A3") == 2
    s.run("cell.set", cell="C6", input="=A1")
    s.run("home.deleteRows", rows="1")
    # B1 itself was deleted; C6 moved up and lost its target.
    assert inp(s, "B1") == ""
    assert inp(s, "C5") == "=#REF!" and val(s, "C5") == "#REF!"
    s2 = new()
    s2.run("range.setValues", range="A1", values=[[1, 2, 3]])
    s2.run("cell.set", cell="A2", input="=C1")
    s2.run("home.deleteColumns", cols="B")
    assert inp(s2, "A2") == "=B1" and val(s2, "A2") == 3
    s2.run("home.deleteColumns", cols="B")
    assert val(s2, "A2") == "#REF!"


def test_fill_series_and_autofill():
    s = new()
    s.run("range.setValues", range="A1", values=[[1], [3]])
    s.run("edit.autoFill", source="A1:A2", target="A1:A5")
    assert [val(s, f"A{i}") for i in range(1, 6)] == [1, 3, 5, 7, 9]
    s.run("cell.set", cell="B1", input="Mon")
    s.run("edit.autoFill", source="B1", target="B1:B3")
    assert [val(s, f"B{i}") for i in range(1, 4)] == ["Mon", "Tue", "Wed"]
    s.run("cell.set", cell="C1", input="Item 1")
    s.run("edit.autoFill", source="C1", target="C1:C3")
    assert val(s, "C3") == "Item 3"
    s.run("cell.set", cell="D1", input="=A1*2")
    s.run("selection.set", range="D1:D3")
    s.run("edit.fillDown")
    assert inp(s, "D3") == "=A3*2" and val(s, "D3") == 10


def test_formatting_and_styles():
    s = new()
    s.run("cell.set", cell="A1", input="1234.5")
    s.run("selection.set", range="A1")
    s.run("home.bold")
    s.run("home.comma")
    assert text(s, "A1").strip() == "1,234.50"
    s.run("home.increaseDecimal")
    assert text(s, "A1").strip() == "1,234.500"
    s.run("home.numberFormat", format="percent")
    assert text(s, "A1") == "123450.00%"
    st = ui_state(s)["ribbon"]
    assert st["bold"] is True
    s.run("home.borders", preset="outside", range="A1:B2")
    s.run("home.fillColor", color="#FFFF00")
    s.run("home.mergeCenter", range="C1:D1")
    assert s.doc().sheet.merges == [RangeRef(0, 2, 0, 3)]
    s.run("home.cellStyle", name="Good", range="E1")
    s.run("home.columnWidth", width=120, cols="B")
    assert s.doc().sheet.col_width(1) == 120


def test_sort_and_filter():
    s = new()
    s.run("range.setValues", range="A1", values=[["Name", "Score"], ["b", 3], ["a", 1], ["c", 2]])
    s.run("selection.set", range="B2")
    s.run("data.sortDescending")
    assert [val(s, f"A{i}") for i in range(2, 5)] == ["b", "c", "a"]
    s.run("data.sortAscending", col="A")
    assert [val(s, f"A{i}") for i in range(2, 5)] == ["a", "b", "c"]
    s.run("data.filter")
    s.run("data.filterColumn", col="B", criterion={"kind": "custom", "conditions": [{"op": "greaterThan", "value": "1"}]})
    sh = s.doc().sheet
    assert sh.is_row_hidden(1) and not sh.is_row_hidden(2)
    vals = s.run("data.filterValues", col="A")
    assert [v["text"] for v in vals["values"]] == ["b", "c"]
    s.run("data.clearFilter")
    assert not s.doc().sheet.is_row_hidden(1)


def test_tables_and_structured_refs():
    s = new()
    s.run("range.setValues", range="A1", values=[["Item", "Qty"], ["x", 2], ["y", 3]])
    s.run("selection.set", range="A1")
    r = s.run("insert.table")
    assert r["name"] == "Table1"
    s.run("cell.set", cell="D1", input="=SUM(Table1[Qty])")
    assert val(s, "D1") == 5
    # Typing below the table extends it.
    s.run("cell.set", cell="A4", input="z")
    s.run("cell.set", cell="B4", input="4")
    assert val(s, "D1") == 9
    s.run("selection.set", range="A2")
    s.run("table.totalRow")
    assert val(s, "B5") == 9
    # Renaming the header renames the column; formulas still work.
    s.run("cell.set", cell="B1", input="Count")
    assert s.doc().sheet.tables[0].columns[1].name == "Count"


def test_names_and_autosum():
    s = new()
    s.run("range.setValues", range="A1", values=[[1], [2], [3]])
    s.run("selection.set", range="A1:A3")
    s.run("formulas.defineName", name="Data")
    s.run("cell.set", cell="B1", input="=SUM(Data)")
    assert val(s, "B1") == 6
    s.run("selection.set", range="A4")
    s.run("formulas.autoSum")
    assert inp(s, "A4") == "=SUM(A1:A3)"
    with pytest.raises(EngineError):
        s.run("formulas.defineName", name="A1")
    names = s.run("formulas.nameManager")["names"]
    assert names[0]["name"] == "Data" and names[0]["value"].startswith("{")


def test_sheets_rename_and_delete():
    s = new()
    s.run("home.insertSheet", name="Data")
    s.run("cell.set", cell="A1", input="7")
    s.run("sheet.activate", sheet="Sheet1")
    s.run("cell.set", cell="A1", input="=Data!A1*2")
    assert val(s, "A1") == 14
    s.run("sheet.rename", sheet="Data", name="Input Data")
    assert inp(s, "A1") == "='Input Data'!A1*2"
    s.run("home.deleteSheet", sheet="Input Data")
    assert val(s, "A1") == "#REF!"


def test_conditional_formats_and_viewport():
    s = new()
    s.run("range.setValues", range="A1", values=[[200], [400], [510]])
    s.run("home.conditionalFormat", range="A1:A3", rule={"type": "cellIs", "operator": "greaterThan", "value": "300", "preset": "lightRedFill"})
    s.run("home.conditionalFormat", range="A1:A3", rule={"type": "dataBar"})
    vp = viewport(s.doc().wb, 0, RangeRef(0, 0, 30, 10))
    cells = {(c["r"], c["c"]): c for c in vp["cells"]}
    assert "bg" not in cells[(0, 0)].get("x", {})
    assert cells[(1, 0)]["x"]["bg"] == "#FFC7CE"
    assert "bar" in cells[(2, 0)]["x"]
    assert cells[(0, 0)]["a"] == "r" and cells[(0, 0)]["t"] == "200"
    g = geometry(s.doc().sheet)
    assert g["defaultColWidth"] == 64
    ui = ui_state(s)
    assert ui["nameBox"] == "A1"


def test_validation_and_find_replace():
    s = new()
    s.run("data.validation", range="A1:A5", kind="list", f1="Yes,No")
    with pytest.raises(EngineError):
        s.run("cell.set", cell="A1", input="Maybe")
    s.run("cell.set", cell="A1", input="yes")
    assert val(s, "A1") == "yes"
    s.run("selection.set", range="A2")
    assert ui_state(s)["dropdown"] == ["Yes", "No"]
    s.run("range.setValues", range="B1", values=[["apple"], ["banana"], ["apple pie"]])
    r = s.run("edit.find", what="apple", all=True)
    assert [m["cell"] for m in r["matches"]] == ["B1", "B3"]
    s.run("edit.replace", what="apple", **{"with": "pear"}, all=True)
    assert val(s, "B3") == "pear pie"


def test_charts_and_objects():
    s = new()
    s.run("range.setValues", range="A1", values=[["", "Sales"], ["Q1", 10], ["Q2", 20]])
    s.run("selection.set", range="A1:B3")
    s.run("insert.chart", kind="columnClustered")
    ob = objects(s.doc().wb, 0)
    ch = ob["charts"][0]
    assert ch["categories"] == ["Q1", "Q2"] and ch["series"][0]["values"] == [10.0, 20.0]
    assert ch["title"] == "Sales"


def test_xlsx_round_trip_through_commands():
    s = new()
    s.run("range.setValues", range="A1", values=[["n"], [1], [2]])
    s.run("cell.set", cell="B1", input="=SEQUENCE(3)")
    s.run("home.bold", range="A1")
    data = s.run("file.saveBytes")
    import base64

    s2 = Session()
    s2.run("file.open", data=data["data"], name="x.xlsx")
    assert val(s2, "A3") == 2
    assert val(s2, "B3") == 3
    assert ui_state(s2)["ribbon"]["bold"] is True


def test_protection():
    s = new()
    s.run("cell.set", cell="A1", input="1")
    s.run("home.lockCell", range="B1", on=False)
    s.run("review.protectSheet", password="pw")
    with pytest.raises(EngineError):
        s.run("cell.set", cell="A1", input="2")
    s.run("cell.set", cell="B1", input="ok")
    with pytest.raises(EngineError):
        s.run("review.unprotectSheet", password="nope")
    s.run("review.unprotectSheet", password="pw")
    s.run("cell.set", cell="A1", input="2")
    assert val(s, "A1") == 2
