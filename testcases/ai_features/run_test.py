"""AI through the AI Hub, against a stand-in hub (testcases/_lib/ai_stub.py): AI Fill writes a
value for each selected row into the next free column, with a title, in one undo step; typing
"=" suggests a formula that Tab takes and Esc dismisses (not while typing on, nor when
suggestions are off); the page tells the hub's assistant what it shows; the assistant's tools
read the workbook over the platform's tool routes; a change the assistant makes (the hub has
asked for the person's Accept by then) reaches everyone, lights up as the AI's, its box stays out
of the people here, and the person's Undo takes it back whatever they did after it.

Needs the server started with ``DDNS_AI_HUB_URL`` (where this suite runs the stand-in, e.g.
http://127.0.0.1:8399) and ``DDNS_AUTH_SECRET``, as the CI workflow does for this suite; without
them it is skipped. ``BACKEND_URL`` is the server's backend (default: BASE_URL's port + 5000)."""

import os
import sys
import time
from pathlib import Path
from urllib.parse import urlparse

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _lib import ai_stub  # noqa: E402
from _lib.grid import Grid, new_book_id, run  # noqa: E402

COMMENTS = ["Great pen, love it", "It leaks everywhere", "Arrived on Tuesday", "Bad ink, broke in a day", "Good value"]


def backend_url() -> str:
    if os.getenv("BACKEND_URL"):
        return os.getenv("BACKEND_URL").rstrip("/")
    u = urlparse(os.getenv("BASE_URL", "http://127.0.0.1:3000"))
    return f"{u.scheme}://{u.hostname}:{(u.port or 3000) + 5000}"


def wait_for(fn, what: str, timeout: float = 10.0):
    end = time.time() + timeout
    while time.time() < end:
        got = fn()
        if got:
            return got
        time.sleep(0.1)
    raise AssertionError(f"timed out waiting for {what}")


def ai_fill(g: Grid) -> None:
    page = g.page
    g.type_in("A1", "Comment")
    for i, text in enumerate(COMMENTS, start=2):
        g.type_in(f"A{i}", text)
    g.select("A1:A6")
    page.click("#tab-data")
    page.click("#rb-ai-fill")
    page.locator("#dlg-aiFill").wait_for()
    assert "Sheet1!A1:A6" in page.locator("#ai-fill-range").inner_text()
    assert page.locator("#dlg-header").get_attribute("aria-checked") == "true", "the first row should read as headers"
    assert page.locator("#dlg-target").input_value() == "B"
    # Nothing asked: an error in the dialog.
    page.click("#dlg-ok")
    page.locator("#dlg-error").wait_for()
    page.fill("#dlg-instruction", "Positive, Negative or Neutral?")
    page.click("#dlg-ok")
    page.locator("#dlg-aiFill").wait_for(state="detached")
    page.locator("#ai-fill-progress").wait_for()
    g.shot("ai_fill_running")
    page.locator("#ai-fill-progress").wait_for(state="detached", timeout=20000)
    for a1, want in (("B1", "Sentiment"), ("B2", "Positive"), ("B3", "Negative"), ("B4", "Neutral"), ("B5", "Negative"), ("B6", "Positive")):
        g.wait_text(a1, want)
    fills = [a for a in ai_stub.Stub.asks if a["system"].startswith("You fill one column")]
    assert len(fills) == 1 and fills[0]["role"] == "fast" and fills[0]["auth"], fills
    assert "1. Comment: Great pen, love it" in fills[0]["prompt"], fills[0]["prompt"]
    g.shot("ai_fill_done")
    # One undo takes the whole fill back, the title too.
    g.click("E10")
    g.key("Mod+z")
    wait_for(lambda: not g.text("B2") and not g.text("B1"), "undo to empty B1:B6")
    g.key("Mod+y")
    g.wait_text("B6", "Positive")
    # Column B is taken now: the next AI Fill proposes C.
    g.select("A1:A6")
    page.click("#rb-ai-fill")
    page.locator("#dlg-aiFill").wait_for()
    assert page.locator("#dlg-target").input_value() == "C"
    page.fill("#dlg-target", "B")
    page.fill("#dlg-instruction", "Again")
    page.click("#dlg-ok")
    assert "already holds values" in page.locator("#dlg-error").inner_text()
    page.click("#dlg-cancel")
    page.locator("#dlg-aiFill").wait_for(state="detached")
    page.wait_for_function("() => !document.querySelector('[role=dialog]')")
    g.wait_idle()


def suggestions(g: Grid) -> None:
    page = g.page
    g.type_in("D1", "Qty")
    for i, n in enumerate((1, 2, 3, 4), start=2):
        g.type_in(f"D{i}", str(n))
    # "=" and a moment: the suggestion, checked by the engine (its result shown), Tab takes it.
    g.click("D6")
    page.keyboard.type("=")
    page.locator(".rg-ai-hint").wait_for()
    page.locator(".rg-ai-formula").wait_for()
    assert page.locator(".rg-ai-formula").inner_text() == "=SUM(D2:D5)"
    assert page.locator(".rg-ai-shown").inner_text() == "→ 10"
    g.shot("suggestion")
    page.keyboard.press("Tab")
    assert page.locator(".rg-ai-hint").count() == 0
    assert g.formula_bar() == "=SUM(D2:D5)", g.formula_bar()
    page.keyboard.press("Enter")
    g.wait_text("D6", "10")
    # Esc dismisses the suggestion; a second Esc cancels the entry.
    g.click("D8")
    page.keyboard.type("=")
    page.locator(".rg-ai-formula").wait_for()
    page.keyboard.press("Escape")
    assert page.locator(".rg-ai-hint").count() == 0
    assert "rg-editing" in (page.locator("#rg-editor").get_attribute("class") or "")
    page.keyboard.press("Escape")
    g.wait_idle()
    assert "rg-editing" not in (page.locator("#rg-editor").get_attribute("class") or "")
    # Typing on past "=" asks nothing.
    asked = len([a for a in ai_stub.Stub.asks if a["system"].startswith("You suggest")])
    g.click("D9")
    page.keyboard.type("=D2*3")
    time.sleep(1.0)
    assert len([a for a in ai_stub.Stub.asks if a["system"].startswith("You suggest")]) == asked
    page.keyboard.press("Enter")
    g.wait_text("D9", "3")
    # Esc while it is still being thought of cancels the entry at once.
    ai_stub.Stub.suggest_delay = 3.0
    g.click("D11")
    page.keyboard.type("=")
    page.locator(".rg-ai-waiting").wait_for()
    page.keyboard.press("Escape")
    g.wait_idle()
    assert "rg-editing" not in (page.locator("#rg-editor").get_attribute("class") or "")
    assert page.locator(".rg-ai-hint").count() == 0
    ai_stub.Stub.suggest_delay = 0.8
    # Off: "=" asks nothing.
    asked = len([a for a in ai_stub.Stub.asks if a["system"].startswith("You suggest")])
    page.click("#tab-formulas")
    page.click("#rb-ai-suggest")
    g.click("D10")
    page.keyboard.type("=")
    time.sleep(1.2)
    assert page.locator(".rg-ai-hint").count() == 0
    assert len([a for a in ai_stub.Stub.asks if a["system"].startswith("You suggest")]) == asked
    page.keyboard.press("Escape")
    page.click("#rb-ai-suggest")


def context_and_tools(g: Grid) -> None:
    page = g.page
    g.click("D6")
    g.wait_idle(0.5)
    ctx = page.evaluate("() => window.__ddnsContext().text")
    for part in ('sheet "Sheet1"', "active cell D6", "holds the formula =SUM(D2:D5)", "shows 10"):
        assert part in ctx, ctx
    assert page.locator("#ai-assist").is_visible(), "the title bar's AI button is missing"

    import httpx
    from reflex_ddns_auth.tools import call_token

    base = backend_url()
    manifest = httpx.get(f"{base}/_ddns_tool/manifest", timeout=10).json()
    names = {t["name"] for t in manifest["tools"]}
    assert {"grid.guide", "grid.workbook", "grid.read", "grid.find", "grid.commands", "grid.vba_get"} <= names, names
    # What changes a workbook waits for the person's Accept on the hub.
    assert {t["name"] for t in manifest["tools"] if t["kind"] != "read"} == {"grid.edit", "grid.vba_set", "grid.form_design"}
    assert all(t["kind"] == "danger" for t in manifest["tools"] if t["kind"] != "read")
    assert "fast" in manifest["ai_uses"], manifest["ai_uses"]
    headers = {"Authorization": f"Bearer {call_token(manifest.get('app') or 'grid', 'ai_features suite')}"}
    r = httpx.post(f"{base}/_ddns_tool/call/grid.read", json={"args": {"room": g.book, "ref": "Sheet1!A1:D6"}}, headers=headers, timeout=10).json()
    assert r["ok"], r
    assert "1\tComment\tSentiment\t\tQty" in r["result"], r["result"]
    assert "2\tGreat pen, love it\tPositive\t\t1" in r["result"], r["result"]
    assert "10 [=SUM(D2:D5)]" in r["result"], r["result"]
    refused = httpx.post(f"{base}/_ddns_tool/call/grid.read", json={"args": {"room": g.book, "ref": "A1"}}, timeout=10)
    assert refused.status_code == 401


def tool(name: str, **args):
    """Calls a tool the way the hub does, once the person has accepted the change."""
    import httpx
    from reflex_ddns_auth.tools import call_token

    base = backend_url()
    app = httpx.get(f"{base}/_ddns_tool/manifest", timeout=10).json().get("app") or "grid"
    headers = {"Authorization": f"Bearer {call_token(app, 'ai_features suite')}"}
    r = httpx.post(f"{base}/_ddns_tool/call/{name}", json={"args": args}, headers=headers, timeout=20).json()
    assert r["ok"], r
    return r["result"]


def step(command: str, **params) -> dict:
    return {"command": command, "params": params}


SORT = step("data.sort", range="A1:B5", keys=[{"col": "B", "order": "desc"}])
VIOLET = "rgb(124, 58, 237)"


def lit_in_ai_colour(g: Grid) -> bool:
    return g.page.evaluate("(c) => [...document.querySelectorAll('.rg-flash')].some((e) => getComputedStyle(e).backgroundColor === c)", VIOLET)


def ai_box_cell(g: Grid) -> str:
    box = g.page.locator(".rg-peer[data-peer^='AI']")
    return box.first.get_attribute("data-peer-cell") if box.count() else ""


def undo_button(g: Grid) -> int:
    return g.page.locator("#ai-change-undo").count()


def changes_by_the_ai(g: Grid) -> None:
    base_url = os.getenv("BASE_URL", "http://127.0.0.1:3000").rstrip("/")
    book = new_book_id()
    g.open(base_url, book)
    context = g.page.context.browser.new_context(viewport={"width": 1400, "height": 860})
    ben = Grid(context.new_page(), g.out)
    ben.open(base_url, book)
    people = lambda who: who.page.locator("#people-count").inner_text()  # noqa: E731
    wait_for(lambda: people(g) == "2 here", "2 here")
    room = g.page.locator("#ai-assist").get_attribute("data-room")
    assert room.startswith(f"{book}~"), room

    # It reaches everyone, in the AI's colour and with its box, also for the one who accepted it.
    rows = [["Item", "Amount"], ["Pen", 82000], ["Ink", 125000], ["Cap", 230000], ["Box", 64000]]
    tool("grid.edit", room=room, summary="Enter the sales", steps=[step("range.setValues", range="A1", values=rows)])
    ben.wait_text("B4", "230000")
    wait_for(lambda: lit_in_ai_colour(g), "the AI's colour on the accepter's grid")
    wait_for(lambda: lit_in_ai_colour(ben), "the AI's colour on the other person's grid")
    wait_for(lambda: ai_box_cell(g) == "A1" and ai_box_cell(ben) == "A1", "the AI's box on both grids")
    assert people(g) == "2 here" and people(ben) == "2 here", "the AI isn't one of the people here"
    assert g.page.locator(".rg-person").count() == 2
    # The note is the accepter's alone.
    wait_for(lambda: undo_button(g) == 1, "the Undo note")
    assert "Enter the sales" in g.page.locator("#ai-change-text").inner_text()
    assert undo_button(ben) == 0
    g.shot("ai_change_note")
    ben.shot("ai_change_seen_by_the_other")

    # Dismissed, the change stays and Ctrl+Z still has it.
    g.page.locator("#ai-change-close").click()
    wait_for(lambda: undo_button(g) == 0, "the note closing")
    assert ben.text("B4") == "230000"

    # Undo takes the AI's change back and leaves what came after it.
    tool("grid.edit", room=room, summary="Add a total", steps=[step("cell.set", cell="A6", input="Total"), step("cell.set", cell="B6", input="=SUM(B2:B5)")])
    ben.wait_text("B6", "501000")
    g.type_in("D1", "mine")
    wait_for(lambda: undo_button(g) == 1, "the Undo note")
    g.page.locator("#ai-change-undo").click()
    wait_for(lambda: not ben.text("B6") and not g.text("A6"), "the total leaving")
    assert g.wait_text("D1", "mine") and ben.wait_text("D1", "mine")
    wait_for(lambda: undo_button(g) == 0 and not ai_box_cell(g) and not ai_box_cell(ben), "the note and the box leaving")

    # The whole request of the plan: one call, one Accept, one Undo.
    tool("grid.edit", room=room, summary="Table, total row, biggest first, over 100k in red", steps=[
        step("insert.table", range="A1:B5", header=True, name="Sales"),
        step("table.totalRow", on=True, table="Sales"),
        step("table.totalFunction", col="Amount", function="sum", table="Sales"),
        SORT,
        step("home.conditionalFormat", range="B2:B5", rule={"type": "cellIs", "operator": "greaterThan", "value": 100000, "preset": "lightRed"}),
    ])
    ben.wait_text("B2", "230000")
    assert ben.wait_text("B6", "501000") and tool("grid.workbook", room=room)["sheets"][0]["tables"]
    ben.shot("ai_change_table_seen_by_the_other")
    wait_for(lambda: undo_button(g) == 1, "the Undo note")
    g.page.locator("#ai-change-undo").click()
    ben.wait_text("B2", "82000")
    wait_for(lambda: not ben.text("B6"), "the total row leaving")
    assert not tool("grid.workbook", room=room)["sheets"][0]["tables"]

    # After moved cells a change can't be matched any more: it says so, and stays.
    tool("grid.edit", room=room, summary="Sort", steps=[SORT])
    ben.wait_text("B2", "230000")
    wait_for(lambda: undo_button(g) == 1, "the Undo note")
    ben.type_in("D2", "theirs")
    g.wait_text("D2", "theirs")
    g.page.locator("#ai-change-undo").click()
    g.page.get_by_text("cells have moved").first.wait_for(state="visible")
    wait_for(lambda: undo_button(g) == 0, "the note closing")
    assert ben.text("B2") == "230000"

    assert not ben.console_errors, ben.console_errors[:3]
    context.close()


def body(g: Grid) -> None:
    ai_fill(g)
    suggestions(g)
    context_and_tools(g)
    changes_by_the_ai(g)


if __name__ == "__main__":
    hub = os.getenv("DDNS_AI_HUB_URL", "")
    if not hub or not os.getenv("DDNS_AUTH_SECRET"):
        print("[ai_features] SKIP: start the server with DDNS_AI_HUB_URL (where this suite runs its stand-in hub) and DDNS_AUTH_SECRET")
        sys.exit(0)
    ai_stub.start(hub)
    sys.exit(run("ai_features", body))
