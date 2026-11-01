"""A workbook on other apps' pages: references that follow the data, how a selection gets named,
the read-only table page and its live socket, and a form's view, which is nobody on the book."""

import base64

import pytest
from starlette.applications import Starlette
from starlette.routing import Route, WebSocketRoute
from starlette.testclient import TestClient

from gridcore.core.addr import RangeRef
from gridcore.engine import CollabSession, EngineError, Session, open_bytes
from gridcore.engine.embed import TABLE_ROWS, forms, last_row, main_block, names, preview, resolve, suggest, table_view, tables
from gridcore.model.forms import FormDesign, new_control
from gridcore.xlsx.oforms import vbframe, write_design, write_frx
from reflex_grid import books, embed
from reflex_grid.books import BookStore

FORM = FormDesign("EntryForm", caption="New entry", width=200, height=60, controls=(
    new_control("TextBox1", "TextBox", left=8, top=8, width=120, tab_index=0),
    new_control("CommandButton1", "CommandButton", left=136, top=6, width=56, tab_index=1, caption="Add", default=True),
))
FORM_CODE = '''Private Sub CommandButton1_Click()
    Dim r As Long
    r = Cells(Rows.Count, 1).End(xlUp).Row + 1
    Cells(r, 1).Value = TextBox1.Value
    Unload Me
End Sub
'''


def sheet(*rows, name="Data"):
    s = Session()
    s.new_workbook()
    s.run("sheet.rename", name=name)
    for r, row in enumerate(rows, 1):
        for c, v in enumerate(row):
            if v != "":
                s.run("cell.set", cell=f"{'ABCDEF'[c]}{r}", input=str(v))
    return s


def with_form(s):
    text = vbframe(FORM).replace("End\r\n", f'   OleObjectBlob   =   "{FORM.name}.frx":0000\r\nEnd\r\n', 1)
    text += f'Attribute VB_Name = "{FORM.name}"\r\n' + FORM_CODE.replace("\n", "\r\n")
    s.run("macro.importModule", fileName=f"{FORM.name}.frm", text=text, frx=base64.b64encode(write_frx(write_design(FORM), FORM)).decode())
    return s


def wb_of(s):
    return s.doc().wb


def test_a_table_follows_itself():
    s = sheet(["Name", "Qty"], ["a", 1], ["b", 2])
    s.run("insert.table", range="A1:B3", header=True, name="Orders")
    assert resolve(wb_of(s), "Orders") == (0, RangeRef(0, 0, 2, 1))
    s.run("cell.set", cell="A4", input="c")
    assert resolve(wb_of(s), "Orders") == (0, RangeRef(0, 0, 3, 1)), "a row typed below joins the table"


def test_open_ended_cells_end_where_the_data_ends():
    s = sheet(["Name", "Qty"], ["a", 1], ["b", 2])
    assert resolve(wb_of(s), "Data!A1:B")[1] == RangeRef(0, 0, 2, 1)
    s.run("cell.set", cell="B9", input="7")
    assert resolve(wb_of(s), "Data!A1:B")[1] == RangeRef(0, 0, 8, 1)
    assert resolve(wb_of(s), "Data!A1:A")[1] == RangeRef(0, 0, 2, 0), "only the columns named count"
    assert resolve(wb_of(s), "Data!A1:B3")[1] == RangeRef(0, 0, 2, 1), "with a last row: these cells"
    assert resolve(wb_of(s), "Data!A:B")[1] == RangeRef(0, 0, 8, 1), "whole columns: where there is data"


def test_an_empty_open_range_is_its_first_row():
    assert resolve(wb_of(sheet(["x"])), "Data!C5:D")[1] == RangeRef(4, 2, 4, 3)


def test_formulas_waiting_for_data_are_not_data():
    s = sheet(["Name", "Twice"], ["a", "=IF(A2=\"\",\"\",2)"], ["", "=IF(A3=\"\",\"\",2)"], ["", "=IF(A4=\"\",\"\",2)"])
    assert resolve(wb_of(s), "Data!A1:B")[1].r1 == 1
    s.run("cell.set", cell="A3", input="b")
    assert resolve(wb_of(s), "Data!A1:B")[1].r1 == 2


def test_a_spill_counts_as_data():
    s = sheet(["N"], ["=SEQUENCE(4)"])
    assert resolve(wb_of(s), "Data!A1:A")[1].r1 == 4


def test_sheet_names_are_spelled_as_in_excel():
    s = sheet(["Name"], ["a"], name="Q1 銷售")
    assert resolve(wb_of(s), "'Q1 銷售'!A1:A")[1] == RangeRef(0, 0, 1, 0)
    assert suggest(wb_of(s), 0, RangeRef(0, 0, 1, 0))["ref"] == "'Q1 銷售'!A1:A"


def test_references_that_stand_for_nothing():
    wb = wb_of(sheet(["x"]))
    for bad in ("", "   ", "Nope!A1:B", "NoSuchName", "1+1"):
        with pytest.raises(EngineError):
            resolve(wb, bad)


def test_hidden_sheets_are_not_shown():
    s = sheet(["x"])
    s.run("home.insertSheet", name="Secret")
    s.run("cell.set", cell="A1", input="1234")
    s.run("sheet.hide")
    with pytest.raises(EngineError):
        resolve(wb_of(s), "Secret!A1:A")
    with pytest.raises(EngineError):
        resolve(wb_of(s), "Secret!A1:A1")


def test_a_defined_name():
    s = sheet(["x", 1], ["y", 2])
    s.run("selection.set", range="A1:B2")
    s.run("formulas.defineName", name="Block")
    assert resolve(wb_of(s), "Block") == (0, RangeRef(0, 0, 1, 1))
    assert [n["name"] for n in names(wb_of(s))] == ["Block"]


def test_a_selection_of_a_table_means_the_table():
    s = sheet(["Name", "Qty"], ["a", 1], ["b", 2])
    s.run("insert.table", range="A1:B3", header=True, name="Orders")
    wb = wb_of(s)
    for sel in (RangeRef(1, 1, 1, 1), RangeRef(0, 0, 2, 1), RangeRef(1, 0, 2, 1)):
        got = suggest(wb, 0, sel)
        assert got["kind"] == "table" and got["ref"] == "Orders" and got["grows"]
    assert suggest(wb, 0, RangeRef(0, 0, 2, 0))["kind"] == "cells", "one of its columns is cells"


def test_a_selection_down_to_the_end_of_the_data_grows():
    wb = wb_of(sheet(["Name", "Qty"], ["a", 1], ["b", 2]))
    got = suggest(wb, 0, RangeRef(0, 0, 2, 1))
    assert (got["ref"], got["grows"], got["fixed"], got["growing"]) == ("Data!A1:B", True, "Data!A1:B3", "Data!A1:B")


def test_a_selection_above_the_end_stays_those_cells():
    wb = wb_of(sheet(["Name", "Qty"], ["a", 1], ["b", 2]))
    got = suggest(wb, 0, RangeRef(0, 0, 1, 1))
    assert (got["ref"], got["grows"]) == ("Data!A1:B2", False)


def test_one_cell_means_the_block_around_it():
    wb = wb_of(sheet(["Name", "Qty"], ["a", 1], ["b", 2], ["", ""], ["other", ""]))
    got = suggest(wb, 0, RangeRef(1, 1, 1, 1))
    assert got["area"] == "A1:B3"
    assert not got["grows"], "data further down: the block isn't the last thing in its columns"


def test_a_single_row_is_not_a_list():
    wb = wb_of(sheet(["Total", 12]))
    got = suggest(wb, 0, RangeRef(0, 0, 0, 1))
    assert (got["ref"], got["grows"], got["growing"]) == ("Data!A1:B1", False, "")


def test_where_a_sheet_starts():
    wb = wb_of(sheet(["Report 2026"], [""], ["Name", "Qty"], ["a", 1], ["b", 2]))
    assert main_block(wb.sheets[0]) == RangeRef(2, 0, 4, 1), "the list, not the title above it"


def test_what_a_workbook_offers():
    s = with_form(sheet(["Name", "Qty"], ["a", 1]))
    s.run("insert.table", range="A1:B2", header=True, name="Orders")
    wb = wb_of(s)
    assert forms(wb) == [{"name": "EntryForm", "caption": "New entry"}]
    assert tables(wb) == [{"name": "Orders", "sheet": "Data", "area": "A1:B2", "rows": 2}]
    info = s.run("embed.info")
    assert info["selection"]["ref"] == "Orders" and info["forms"] == forms(wb)
    assert forms(wb_of(sheet(["x"]))) == []


def test_the_preview_is_text():
    wb = wb_of(sheet(["", "Name"], ["", "=1+1"]))
    p = preview(wb, 0)
    assert (p["r0"], p["c0"], p["rows"], p["more"]) == (0, 1, [["Name"], ["2"]], False)


def test_a_table_view_shows_the_cells_as_they_look():
    s = sheet(["Name", "Pay"], ["<a>", 1234.5])
    v = table_view(wb_of(s), "Data!A1:B")
    assert "&lt;a&gt;" in v["html"] and "<a>" not in v["html"] and (v["rows"], v["shown"]) == (2, 2)


def test_a_long_table_shows_its_first_row_and_the_latest():
    s = sheet(["Head"])
    wb = wb_of(s)
    sm = s.doc().wb.sheets[0]
    from gridcore.model.cell import Cell

    for r in range(1, TABLE_ROWS + 50):
        sm.cells.set((r, 0), Cell(f"row{r}", None, 0))
    last = last_row(sm, 0, 0, 0)
    v = table_view(wb, "Data!A1:A")
    assert (v["rows"], v["shown"]) == (last + 1, TABLE_ROWS)
    assert "Head" in v["html"] and f"row{last}<" in v["html"] and "row1<" not in v["html"]


# ---------------------------------------------------------------- the read-only page


@pytest.fixture
def store(tmp_path, monkeypatch):
    s = BookStore(tmp_path)
    monkeypatch.setattr(embed, "STORE", s)
    embed._DRAWN.clear()
    return s


@pytest.fixture
def client():
    app = Starlette(routes=[Route(embed.TABLE_PATH, embed.table_page, methods=["GET"]), WebSocketRoute(embed.LIVE_PATH, embed.live)])
    with TestClient(app) as c:
        yield c


def book_with(store, book_id="b1"):
    data = base64.b64decode(sheet(["Name", "Qty"], ["a", 1]).run("file.saveBytes")["data"])
    return store.open(store.create(open_bytes(Session(), data, "t.xlsx"), "T", book_id=book_id).id)


def edit(book, cell, text):
    view = book.join("tab", CollabSession(), "Amy")
    with book.act(view) as s:
        s.run("cell.set", cell=cell, input=text)
    book.leave("tab")


def test_the_page_shows_the_cells_and_nothing_to_edit_with(store, client):
    book_with(store)
    r = client.get(embed.TABLE_PATH, params={"book": "b1", "range": "Data!A1:B"})
    assert r.status_code == 200 and "no-store" in r.headers["cache-control"]
    body = r.text
    assert ">Name<" in body and "Read only" in body and "2 rows" in body
    assert not any(tag in body for tag in ("<input", "<textarea", "<select", "<form", "contenteditable", "<button"))


def test_the_page_can_only_be_read(store, client):
    book_with(store)
    assert client.post(embed.TABLE_PATH, params={"book": "b1", "range": "A1"}).status_code == 405


def test_an_address_that_leads_nowhere_creates_nothing(store, client):
    r = client.get(embed.TABLE_PATH, params={"book": "nothing", "range": "A1"})
    assert r.status_code == 404 and embed.GONE.replace("'", "&#x27;") in r.text and 'data-gone="1"' in r.text
    assert not store.exists("nothing") and not store.get("nothing")
    assert client.get(embed.TABLE_PATH, params={"book": "../etc", "range": "A1"}).status_code == 404


def test_a_bad_reference_says_so(store, client):
    book_with(store)
    assert "Reference isn" in client.get(embed.TABLE_PATH, params={"book": "b1", "range": "Nope"}).text


def test_text_in_a_reference_stays_text(store, client):
    book_with(store)
    body = client.get(embed.TABLE_PATH, params={"book": "b1", "range": "<script>x</script>"}).text
    assert "<script>x" not in body


def test_the_socket_carries_each_change(store, client):
    book = book_with(store)
    with client.websocket_connect(f"{embed.LIVE_PATH}?book=b1&range=Data!A1:B") as ws:
        first = ws.receive_json()
        assert ">a<" in first["html"] and first["note"] == "2 rows"
        assert book.watchers, "watching keeps the workbook open"
        edit(book, "A3", "new")
        second = ws.receive_json()
        assert ">new<" in second["html"] and second["note"] == "3 rows"
    assert not book.watchers


def test_a_change_elsewhere_sends_nothing(store, client):
    book = book_with(store)
    with client.websocket_connect(f"{embed.LIVE_PATH}?book=b1&range=Data!A1:A2") as ws:
        ws.receive_json()
        edit(book, "D9", "elsewhere")
        edit(book, "A2", "seen")
        assert ">seen<" in ws.receive_json()["html"], "the first message after is the change that shows"


def test_the_socket_says_when_the_workbook_goes(store, client):
    book_with(store)
    with client.websocket_connect(f"{embed.LIVE_PATH}?book=b1&range=A1") as ws:
        ws.receive_json()
        store.delete("b1")
        got = ws.receive_json()
        assert got["gone"] and got["error"] == embed.GONE


def test_the_socket_to_nowhere_creates_nothing(store, client):
    with client.websocket_connect(f"{embed.LIVE_PATH}?book=nothing&range=A1") as ws:
        assert ws.receive_json()["gone"]
    assert not store.exists("nothing")


def test_everyone_on_the_same_cells_shares_a_drawing(store, monkeypatch):
    book = book_with(store)
    calls = []
    real = embed.table_view
    monkeypatch.setattr(embed, "table_view", lambda wb, ref: calls.append(ref) or real(wb, ref))
    assert embed._drawn(book, "Data!A1:B") == embed._drawn(book, "Data!A1:B")
    assert calls == ["Data!A1:B"]
    edit(book, "A3", "x")
    embed._drawn(book, "Data!A1:B")
    assert len(calls) == 2


def test_a_watched_book_stays_open(store):
    book = book_with(store)
    book.watchers.add(lambda: None)
    book.last_active -= books.IDLE_CLOSE + 1
    assert store.close_idle() == [] and store.get("b1") is book
    book.watchers.clear()
    assert store.close_idle() == ["b1"]


def test_a_form_on_a_page_is_nobody_on_the_book(store):
    book = book_with(store)
    amy = book.join("amy", CollabSession(), "Amy")
    amy.info["user"] = "amy"
    form = book.join("form", CollabSession(), "Ben")
    form.info.update(user="ben", embed=True)
    assert [v.key for v in book.people()] == ["amy"]
    assert [p["user"] for p in store.viewers()["b1"]] == ["amy"]


def test_directive():
    assert embed.directive("b1", "form", "OrderForm") == "{%grid b1 form=OrderForm %}"
    assert embed.directive("b1", "range", "'Q1 Sales'!A1:D") == "{%grid b1 range='Q1 Sales'!A1:D %}"


def test_choices_from_a_workbook():
    s = with_form(sheet(["Name", "Qty"], ["a", 1]))
    s.run("insert.table", range="A1:B2", header=True, name="Orders")
    rows = embed.embed_choices("b1", s.run("embed.info"))
    assert [r["code"] for r in rows] == ["{%grid b1 range=Orders %}", "{%grid b1 form=EntryForm %}"]


# ---------------------------------------------------------------- a form that adds a row


def test_a_form_shown_modeless_adds_to_the_cells_a_page_shows():
    s = with_form(sheet(["Name"], ["a"]))
    s.interactive_macros = True
    r = s.run("macro.showForm", name="EntryForm", modal=False)
    assert r["status"] == "done" and s.macro_run is None, "no run waits: nothing holds a macro thread"
    form = r["forms"][0]
    r = s.run("macro.formEvent", form=form["id"], event={"event": "click", "control": "CommandButton1", "values": {"TextBox1": "b"}})
    assert r["forms"] == [], "OK closed it"
    assert ">b<" in table_view(wb_of(s), "Data!A1:A")["html"]
    r = s.run("macro.showForm", name="EntryForm", modal=False)
    assert r["forms"][0]["id"] != form["id"], "the next entry starts on a new form"
