"""The shared-workbook store: workbooks as .xlsx files plus metadata, opened once and shared,
saved as they change, renamed, deleted, closed when idle."""

import io

import openpyxl
import pytest

from gridcore.engine import CollabSession, EngineError
from reflex_grid.books import COLORS, DEFAULT_TITLE, BookStore, clean_title, valid_id


@pytest.fixture
def store(tmp_path):
    return BookStore(tmp_path)


def edit(book, key, cell, text):
    view = book.join(key, CollabSession(), key.title())
    with book.act(view) as s:
        s.run("cell.set", cell=cell, input=text)
    return view


def test_ids_and_titles():
    assert valid_id("a1b2-C3") and not valid_id("") and not valid_id("../x") and not valid_id("a_b") and not valid_id("x" * 65)
    assert clean_title("  Q1   plan ") == "Q1 plan"
    assert clean_title("   ") == DEFAULT_TITLE
    assert len(clean_title("x" * 500)) == 120


def test_a_new_address_creates_a_blank_book(store):
    book = store.open("fresh1", creator={"email": "amy@example.com", "name": "Amy"})
    assert store.open("fresh1") is book, "one shared book per id"
    meta = store.meta("fresh1")
    assert meta.title == DEFAULT_TITLE and meta.created_by_email == "amy@example.com" and meta.sheets == ["Sheet1"]
    assert (store.dir / "fresh1.xlsx").exists() and (store.dir / "fresh1.json").exists()
    with pytest.raises(EngineError):
        store.open("no/such")


def test_changes_are_saved_as_xlsx_and_read_back(store):
    meta = store.create_blank({"email": "", "name": ""}, "Budget")
    book = store.open(meta.id)
    edit(book, "amy", "A1", "12")
    edit(book, "ben", "A2", "=A1*2")
    assert book.dirty and book._timer is not None, "a save is scheduled"
    assert store.save(book)
    assert not book.dirty and book.meta.last_editor == "Ben"
    wb = openpyxl.load_workbook(io.BytesIO((store.dir / f"{meta.id}.xlsx").read_bytes()))
    assert wb["Sheet1"]["A2"].value == "=A1*2"
    # Another process (a restart) reads the same workbook.
    again = BookStore(store.home).open(meta.id)
    with again.act(again.join("cat", CollabSession())) as s:
        assert s.run("cell.get", cell="A2")["value"] == 24
    assert again.meta.title == "Budget"


def test_list_rename_and_delete(store):
    a = store.create_blank(title="First")
    b = store.create_sample("Sales")
    assert b.title == "Sales" and b.sheets and b.size > 0
    assert [m.id for m in store.list()][0] == b.id, "most recent first"
    book = store.open(a.id)
    store.rename(a.id, "  Renamed  ")
    assert book.meta.title == "Renamed" and book.doc.name == "Renamed"
    assert BookStore(store.home).meta(a.id).title == "Renamed"
    store.delete(a.id)
    assert book.deleted and store.get(a.id) is None and not store.exists(a.id)
    assert [m.id for m in store.list()] == [b.id]


def test_upload_titles_the_book_after_the_file(store):
    meta, warnings = store.create_from_file(b"Item,Qty\nPen,3\n", "stock count.csv")
    assert meta.title == "stock count" and warnings == []
    book = store.open(meta.id)
    with book.act(book.join("amy", CollabSession())) as s:
        assert s.run("cell.get", cell="B2")["value"] == 3
    with pytest.raises(EngineError):
        store.create_from_file(b"\x00\x01 not a workbook", "junk.xlsx")


def test_people_and_colours(store):
    book = store.open("people1")
    for key, user in (("t1", "amy@x"), ("t2", "ben@x"), ("t3", "amy@x")):
        v = book.join(key, CollabSession(), user)
        v.info.update(user=user, color=book.color_for(user))
    colors = {v.key: v.info["color"] for v in book.views.values()}
    assert colors["t1"] == colors["t3"] != colors["t2"], "one colour per person"
    assert colors["t1"] == COLORS[0]
    viewers = store.viewers()["people1"]
    assert sorted(p["user"] for p in viewers) == ["amy@x", "ben@x"]


def test_idle_books_are_saved_and_closed(store):
    book = store.open("idle1")
    v = edit(book, "amy", "A1", "keep")
    book.leave(v.key)
    if book._timer is not None:
        book._timer.cancel()
        book._timer = None
    assert store.close_idle(now=book.last_active + 10) == [], "not idle long enough"
    assert store.close_idle(now=book.last_active + 16 * 60) == ["idle1"]
    assert store.get("idle1") is None and book.closed
    reopened = store.open("idle1")
    with reopened.act(reopened.join("ben", CollabSession())) as s:
        assert s.run("cell.get", cell="A1")["value"] == "keep"
