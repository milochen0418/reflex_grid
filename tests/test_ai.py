"""AI through the AI Hub: how AI Fill and formula suggestions word their questions and read the
answers (reflex_grid/ai.py), and the tools the hub's assistant reads workbooks with
(reflex_grid/ai_tools.py). No model is asked here; the browser suite ai_features drives both
features against a stand-in hub."""

import asyncio

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
    assert text.splitlines()[0] == "Sales!A1:F2 (2 rows × 6 columns, formulas in brackets); 23 more rows: read Sales!A3:F25 next"


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
