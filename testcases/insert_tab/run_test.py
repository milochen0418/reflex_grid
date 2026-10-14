"""Insert tab, button by button: every chart kind of the chart menus and All Charts, sparklines of
each type, links (their ScreenTip on hover, Ctrl/Cmd+click), comment threads (post, reply,
resolve) and notes shown on hover, Insert Function, tables and pictures."""

import struct
import sys
import time
import zlib
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _lib.grid import MOD, Grid, run  # noqa: E402

CHARTS = {
    "Column": ["Clustered Column", "Stacked Column", "100% Stacked Column", "Clustered Bar", "Stacked Bar", "100% Stacked Bar"],
    "Line": ["Line", "Stacked Line", "Line with Markers", "Area", "Stacked Area"],
    "Pie": ["Pie", "Doughnut"],
    "Scatter": ["Scatter", "Scatter with Lines", "Radar"],
}


def png(w: int, h: int) -> bytes:
    """A plain green ``w`` × ``h`` PNG."""

    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)

    rows = b"".join(b"\x00" + b"\x4e\xa7\x2e" * w for _ in range(h))
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b"")


def insert(g: Grid, button: str, at: str = "") -> None:
    """Selects ``at`` (a cell or range), then clicks Insert › ``button``."""
    if at:
        g.select(at)
    g.page.click("#tab-insert")
    g.page.get_by_role("button", name=button, exact=True).click()


def open_menu(g: Grid, title: str, at: str = "") -> list:
    """Selects ``at``, opens the Insert tab's ``title`` dropdown and returns its items."""
    if at:
        g.select(at)
    g.page.click("#tab-insert")
    g.page.click(f'button[title="{title}"]')
    items = g.page.locator("[role=menuitem]")
    items.first.wait_for(state="visible")
    return [t.strip() for t in items.all_inner_texts()]


def hover(g: Grid, a1: str, what: str) -> str:
    """Rests the pointer on a cell, moving it a little as a hand does until the ``what`` popup
    shows, and returns the popup's text."""
    x, y, w, h = g.cell_rect(a1)
    tip = g.page.locator(what).first
    for i in range(20):
        g.page.mouse.move(x + w / 3 + i % 3, y + h / 2)
        if tip.is_visible():
            break
        time.sleep(0.2)
    tip.wait_for(state="visible", timeout=1000)
    text = " ".join(tip.inner_text().split())
    g.page.mouse.move(x + w / 3, y + h * 6)
    return text


def toast(g: Grid, text: str) -> None:
    g.page.locator("[data-sonner-toast]", has_text=text).first.wait_for(state="visible", timeout=6000)


def body(g: Grid) -> None:
    page = g.page
    rows = [["Region", "Q1", "Q2", "Q3"], ["North", "10", "-12", "5"], ["South", "20", "18", "-3"], ["East", "15", "22", "9"], ["West", "8", "-11", "14"]]
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            g.type_in(f"{'ABCD'[c]}{r + 1}", v)

    # Every kind of the four chart menus draws a chart (each is deleted again).
    for title, labels in CHARTS.items():
        for i, label in enumerate(labels):
            got = open_menu(g, title, "A1:D5")
            assert got == labels, f"{title} menu: {got}"
            page.locator("[role=menuitem]").nth(i).click()
            # The closed menu gives the keyboard back to its button: let it, before clicking the chart.
            page.locator("[role=menu]").wait_for(state="detached")
            chart = page.locator(".rg-chart")
            chart.wait_for(state="visible")
            assert chart.locator("svg path, svg rect, svg circle, svg polyline, svg polygon").count() > 4, f"{label}: nothing drawn"
            chart.click(position={"x": 8, "y": 8})
            page.locator("#tab-chart-design").wait_for(state="visible")
            page.keyboard.press("Delete")
            chart.wait_for(state="detached")
    # All Charts lists them too.
    insert(g, "All Charts", "A1:D5")
    dialog = page.locator("#dlg-insertChart")
    dialog.wait_for(state="visible")
    for label in ("100% Stacked Bar", "Stacked Line"):
        assert dialog.get_by_text(label, exact=True).count() == 1, f"All Charts misses {label}"
    dialog.get_by_text("Stacked Line", exact=True).click()
    dialog.locator("input").last.fill("Trend")
    g.dialog_ok("insertChart")
    page.locator(".rg-chart", has_text="Trend").wait_for(state="visible")
    page.locator(".rg-chart").click(position={"x": 8, "y": 8})
    page.locator("#tab-chart-design").wait_for(state="visible")
    page.keyboard.press("Delete")
    page.locator(".rg-chart").wait_for(state="detached")

    # Sparklines: Line, Column and Win/Loss, each from the menu.
    for kind, col, mark in (("Line", "E", "path"), ("Column", "F", "rect"), ("Win/Loss", "G", "rect")):
        assert open_menu(g, "Sparklines", "B2:D5") == ["Line", "Column", "Win/Loss"]
        page.get_by_role("menuitem", name=kind, exact=True).click()
        dialog = page.locator("#dlg-sparkline")
        dialog.wait_for(state="visible")
        assert dialog.locator("button[role=combobox]").inner_text().strip() == kind
        dialog.locator("input").nth(1).fill(f"{col}2:{col}5")
        g.dialog_ok("sparkline")
        for r in range(2, 6):
            page.locator(f'[data-spark="{col}{r}"] svg {mark}').first.wait_for(state="visible")
    g.shot("sparklines")

    # Links: the ScreenTip (or the address) shows on hover; Ctrl/Cmd+click follows.
    insert(g, "Link", "A8")
    page.locator("#dlg-hyperlink").wait_for(state="visible")
    page.locator("#dlg-hyperlink input").first.fill("Example")
    page.fill("#dlg-target", "https://example.com/")
    page.locator("#dlg-hyperlink input").last.fill("Our example site")
    g.dialog_ok("hyperlink")
    g.wait_text("A8", "Example")
    tip = hover(g, "A8", ".rg-link-tip")
    assert "Our example site" in tip and "click to follow the link" in tip, tip
    insert(g, "Link", "A9")
    page.locator("#dlg-hyperlink").wait_for(state="visible")
    page.locator("#dlg-hyperlink input").first.fill("Go to K40")
    page.fill("#dlg-target", "Sheet1!K40")
    g.dialog_ok("hyperlink")
    g.wait_text("A9", "Go to K40")
    assert "Sheet1!K40" in hover(g, "A9", ".rg-link-tip")
    page.evaluate("() => { window.open = (u) => { window.__opened = u; return null; }; }")
    x, y, w, h = g.cell_rect("A8")
    page.keyboard.down(MOD)
    page.mouse.click(x + 12, y + h / 2)
    page.keyboard.up(MOD)
    page.wait_for_function("() => window.__opened === 'https://example.com/'")
    x, y, w, h = g.cell_rect("A9")
    page.keyboard.down(MOD)
    page.mouse.click(x + 12, y + h / 2)
    page.keyboard.up(MOD)
    page.wait_for_function("() => document.querySelector('#rg-namebox').value === 'K40'")
    page.keyboard.press(f"{MOD}+Home")
    g.wait_idle(0.5)

    # A comment thread: post, see it on hover, reply, resolve.
    insert(g, "Comment", "B8")
    dialog = page.locator("#dlg-note")
    dialog.wait_for(state="visible")
    assert dialog.locator(".rg-thread-msg").count() == 0 and page.locator("#dlg-ok").inner_text() == "Post"
    page.fill("#dlg-text", "Is Q2 final?")
    g.dialog_ok("note")
    assert "Is Q2 final?" in hover(g, "B8", ".rg-note")
    insert(g, "Comment", "B8")
    dialog.wait_for(state="visible")
    assert dialog.locator(".rg-thread-msg").count() == 1 and page.locator("#dlg-text").input_value() == ""
    assert page.locator("#dlg-ok").inner_text() == "Reply"
    page.fill("#dlg-text", "Yes, final.")
    g.dialog_ok("note")
    thread = hover(g, "B8", ".rg-note")
    assert "Is Q2 final?" in thread and "Yes, final." in thread, thread
    insert(g, "Comment", "B8")
    dialog.wait_for(state="visible")
    assert dialog.locator(".rg-thread-msg").count() == 2
    g.shot("comment_thread")
    page.click("#dlg-resolve")
    dialog.wait_for(state="detached")
    page.wait_for_function("() => !document.querySelector('[role=dialog]')")
    assert "(resolved)" in hover(g, "B8", ".rg-note")
    # A cell holds a comment or a note, not both.
    insert(g, "Note", "B8")
    toast(g, "already has a comment")
    assert page.locator("#dlg-note").count() == 0
    x, y, w, h = g.cell_rect("B8")
    page.mouse.click(x + w / 2, y + h / 2, button="right")
    menu = page.locator(".rg-menu")
    menu.wait_for(state="visible")
    assert menu.get_by_text("Reply to Comment", exact=True).count() == 1 and menu.get_by_text("Delete Comment", exact=True).count() == 1
    page.keyboard.press("Escape")
    menu.wait_for(state="detached")

    # A note shows when the pointer is anywhere on its cell.
    insert(g, "Note", "C8")
    page.locator("#dlg-note").wait_for(state="visible")
    page.fill("#dlg-text", "Checked by Amy")
    g.dialog_ok("note")
    assert "Checked by Amy" in hover(g, "C8", ".rg-note")
    insert(g, "Comment", "C8")
    toast(g, "already has a note")

    # Insert Function starts the formula in the cell.
    insert(g, "Function", "E8")
    page.locator("#dlg-insertFunction").wait_for(state="visible")
    page.fill("#dlg-fn-search", "AVERAGE")
    page.locator("#dlg-insertFunction").get_by_text("AVERAGE", exact=True).first.click()
    g.dialog_ok("insertFunction")
    g.wait_formula_bar("=AVERAGE(")
    page.keyboard.type("B2:B5)")
    page.keyboard.press("Enter")
    g.wait_text("E8", "13.25")

    # A table, and a picture.
    insert(g, "Table", "A1:D5")
    page.locator("#dlg-insertTable").wait_for(state="visible")
    assert page.locator("#dlg-table-range").input_value() == "A1:D5"
    g.dialog_ok("insertTable")
    page.locator(".rg-filter-btn").nth(3).wait_for(state="visible")
    insert(g, "Pictures", "I10")
    page.locator('#dlg-insertPicture input[type="file"]').set_input_files({"name": "green.png", "mimeType": "image/png", "buffer": png(90, 60)})
    page.locator(".rg-image").wait_for(state="visible")
    time.sleep(0.3)
    g.shot("insert_tab")


if __name__ == "__main__":
    sys.exit(run("insert_tab", body))
