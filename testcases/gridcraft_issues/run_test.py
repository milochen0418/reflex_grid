"""The browser side of bugs GridCraft (the Rust spreadsheet gridcore follows) fixed: a header click in
formula point mode, Enter pasting a copy, typing after the name box, Ctrl+Shift+Arrow revealing the
moving end, the last row reached and drawn, a new chart selected so Delete removes it,
double-clicking the fill handle, a validation's input message under its open list, Center Across
Selection, Shrink to Fit and a merged area's borders, links that may not run script, panning the
grid with a finger, a multi-line value in an unwrapped cell on one line, Page Up into frozen rows
scrolling the rows below back to the top, and Shift+Arrow on a whole column keeping the view. Then
GridCraft's issues closed 2026-10-27 to 2026-11-06: AutoSum's caret, Ctrl+., Ctrl+Shift+Down while
pointing, a spilled cell's formula, dropping onto data (a move asks, a copy doesn't), a linked cell's menu and dialog, Find in a
selected range, copying quotes, hidden sheet tabs and scroll bars, pasting a picture, doughnut rings,
Set Precision as Displayed and a web page saved as .xls."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl  # noqa: E402
from openpyxl.styles import Alignment, Border, Side  # noqa: E402
from openpyxl.worksheet.datavalidation import DataValidation  # noqa: E402

from _lib.grid import MOD, Grid, run  # noqa: E402


def header(g: Grid, kind: str, label: str):
    """The center of a column (``cols``) or row (``rows``) header."""
    heads = g.page.locator(".rg-colhead" if kind == "cols" else ".rg-rowhead")
    for i in range(heads.count()):
        el = heads.nth(i)
        if el.inner_text().strip() == label:
            b = el.bounding_box()
            return b["x"] + b["width"] / 2, b["y"] + b["height"] / 2
    raise AssertionError(f"no {kind} header {label}")


def editor(g: Grid) -> str:
    return g.page.locator("#rg-editor").input_value()


def toast(g: Grid, text: str) -> None:
    g.page.locator("[data-sonner-toast]", has_text=text).first.wait_for(state="visible", timeout=6000)


def body(g: Grid) -> None:
    page = g.page
    for r in range(1, 5):
        g.type_in(f"C{r}", str(r))
        g.type_in(f"D{r}", str(r * 10))

    # GridCraft #691: in point mode a column header inserts C:C, a row header 4:4; a drag across
    # headers extends the reference.
    g.click("A1")
    page.keyboard.type("=SUM(")
    page.mouse.click(*header(g, "cols", "C"))
    g.wait_idle(0.2)
    assert editor(g) == "=SUM(C:C", editor(g)
    page.mouse.move(*header(g, "cols", "C"))
    page.mouse.down()
    page.mouse.move(*header(g, "cols", "D"), steps=5)
    page.mouse.up()
    g.wait_idle(0.2)
    assert editor(g) == "=SUM(C:D", editor(g)
    page.keyboard.type(")")
    page.keyboard.press("Enter")
    g.wait_text("A1", "110")
    g.click("A6")
    page.keyboard.type("=SUM(")
    page.mouse.click(*header(g, "rows", "2"))
    g.wait_idle(0.2)
    assert editor(g) == "=SUM(2:2", editor(g)
    page.keyboard.type(")")
    page.keyboard.press("Enter")
    g.wait_text("A6", "22")

    # GridCraft #845: Enter pastes the copied cells and ends copy mode.
    g.drag("C1", "C2")
    g.key("Mod+c")
    g.click("F1")
    page.keyboard.press("Enter")
    g.wait_text("F2", "2")
    assert g.text("F1") == "1"
    g.click("H1")
    page.keyboard.press("Enter")
    g.wait_idle(0.4)
    assert g.text("H1") in (None, "") and g.name_box() == "H2", (g.text("H1"), g.name_box())

    # Typing right after Enter in the name box goes into the cell it went to (the name box's own
    # text selection used to take the keyboard back).
    g.select("M300")
    page.keyboard.type("end")
    page.keyboard.press("Enter")
    g.select("M1")
    page.keyboard.type("top")
    page.keyboard.press("Enter")
    g.wait_text("M1", "top")
    assert g.name_box() == "M2", g.name_box()

    # GridCraft #762: Ctrl+Shift+Down scrolls to the end it extends to, not the anchor.
    g.select("M1")
    page.focus("#rg-editor")
    g.key(f"{MOD}+Shift+ArrowDown")
    page.wait_for_function("() => [...document.querySelectorAll('.rg-rowhead')].some((e) => e.innerText.trim() === '300')", timeout=8000)
    g.cell_rect("M300")

    # GridCraft #687: a selection down to the last row shows there, outline and all. Browsers stop
    # scrolling near 15 million px (row 750,000 at 100%); the rows past that are squeezed in.
    g.select("K5")
    page.focus("#rg-editor")
    g.key(f"{MOD}+Shift+ArrowDown")
    page.wait_for_function("() => [...document.querySelectorAll('.rg-rowhead')].some((e) => e.innerText.trim() === '1048576')", timeout=8000)
    g.wait_idle(0.5)
    sel = page.locator(".rg-selection").bounding_box()
    view = page.locator(".rg-scroller").bounding_box()
    assert view["y"] < sel["y"] + sel["height"] <= view["y"] + view["height"] + 2, (sel, view)
    g.select("D900000")
    page.keyboard.type("far")
    page.keyboard.press("Enter")
    g.wait_text("D900000", "far")
    g.key(f"{MOD}+Home")
    page.wait_for_function("() => [...document.querySelectorAll('.rg-rowhead')].some((e) => e.innerText.trim() === '1')", timeout=8000)
    g.select("A1")
    g.wait_idle(0.5)

    # GridCraft #668: a chart just inserted is selected, so Delete removes it, not its data.
    g.select("C1:D4")
    page.click("#tab-insert")
    page.click('button[title="Column"]')
    page.locator("[role=menuitem]").first.click()
    page.locator("[role=menu]").wait_for(state="detached")
    chart = page.locator(".rg-chart")
    chart.wait_for(state="visible")
    page.locator("#tab-chart-design").wait_for(state="visible")
    page.focus("#rg-editor")
    page.keyboard.press("Delete")
    chart.wait_for(state="detached")
    g.wait_text("C1", "1")
    g.wait_text("D4", "40")

    # GridCraft #507: double-clicking the fill handle fills down as far as the data beside it.
    g.type_in("E1", "=D1+1")
    g.click("E1")
    x, y, w, h = g.cell_rect("E1")
    page.mouse.dblclick(x + w - 1, y + h - 1)
    g.wait_text("E4", "41")
    assert g.text("E5") in (None, ""), g.text("E5")

    # GridCraft #777: a link can't run script; Ctrl/Cmd+click refuses it.
    page.evaluate("() => { window.__opened = null; window.open = (u) => { window.__opened = u; return null; }; }")
    g.select("A10")
    page.click("#tab-insert")
    page.get_by_role("button", name="Link", exact=True).click()
    page.locator("#dlg-hyperlink").wait_for(state="visible")
    page.locator("#dlg-hyperlink input").first.fill("Click me")
    page.fill("#dlg-target", "javascript:alert(document.cookie)//x.y")
    g.dialog_ok("hyperlink")
    g.wait_text("A10", "Click me")
    x, y, w, h = g.cell_rect("A10")
    page.keyboard.down(MOD)
    page.mouse.click(x + 12, y + h / 2)
    page.keyboard.up(MOD)
    toast(g, "opens only web")
    assert page.evaluate("() => window.__opened") is None
    g.shot("grid")

    # GridCraft #505: a list's input message hides while its list is open (and #479, #375 below).
    src = g.out / "validation.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    dv = DataValidation(type="list", formula1='"North,South,East"', showInputMessage=True, promptTitle="Region", prompt="Pick a region")
    ws.add_data_validation(dv)
    dv.add("B2")
    ws["D5"] = "Quarterly report"
    for c in "DEFG":
        ws[f"{c}5"].alignment = Alignment(horizontal="centerContinuous")
    ws["D7"] = 1234567.891
    ws["D7"].number_format = "#,##0.00"
    ws["D7"].alignment = Alignment(shrink_to_fit=True)
    ws["D8"] = 1234567.891
    ws["D8"].number_format = "#,##0.00"
    ws.column_dimensions["D"].width = 8
    ws["F10"] = "merged"
    ws.merge_cells("F10:H11")
    ws["J2"] = "one\ntwo"
    ws["J3"] = "one\ntwo"
    ws["J3"].alignment = Alignment(wrap_text=True)
    ws.freeze_panes = "A2"
    red = Side(style="medium", color="FF0000")
    for r in (10, 11):
        for c in "FGH":
            ws[f"{c}{r}"].border = Border(top=red if r == 10 else None, bottom=red if r == 11 else None, left=red if c == "F" else None, right=red if c == "H" else None)
    wb.save(src)
    page.click("#tab-file")
    page.click("#bs-open")
    page.locator("#dlg-open").wait_for(state="visible")
    page.locator('#dlg-open input[type="file"]').set_input_files(str(src))
    g.wait_title("validation")
    g.wait_idle(0.8)
    g.click("B2")
    msg = page.locator(".rg-input-msg")
    msg.wait_for(state="visible")
    page.locator(".rg-list-btn").click()
    page.locator(".rg-list").wait_for(state="visible")
    msg.wait_for(state="detached")
    page.keyboard.press("Escape")
    page.locator(".rg-list").wait_for(state="detached")
    msg.wait_for(state="visible")

    # GridCraft #479: Center Across Selection centres over the empty cells to its right that have it.
    d5 = page.locator('.rg-pane [data-cell="D5"]').bounding_box()
    x, y, w, h = g.cell_rect("D5")
    gx, gy, gw, gh = g.cell_rect("G5")
    assert abs(d5["x"] - x) < 2 and abs(d5["x"] + d5["width"] - (gx + gw)) < 3, (d5, x, gx + gw)
    # GridCraft #375: Shrink to Fit draws a smaller font where the number would show ###.
    assert g.text("D7") == "1,234,567.89" and g.text("D8").startswith("###"), (g.text("D7"), g.text("D8"))
    size = page.evaluate("() => parseFloat(getComputedStyle(document.querySelector('.rg-pane [data-cell=\"D7\"]')).fontSize)")
    assert size < 14, size
    # GridCraft #376: a merged area's right and bottom borders, kept on its right and bottom cells.
    g.click("B14")
    canvas = page.locator(".rg-view canvas").bounding_box()

    def pixel(x, y):
        return page.evaluate(
            "([x, y]) => { const cv = document.querySelector('.rg-view canvas'); const r = cv.getBoundingClientRect(); const k = cv.width / r.width;"
            " return [...cv.getContext('2d').getImageData(Math.round(x * k), Math.round(y * k), 1, 1).data].slice(0, 3); }",
            [x - canvas["x"], y - canvas["y"]],
        )

    fx, fy, fw, fh = g.cell_rect("F10")
    hx, hy, hw, hh = g.cell_rect("H11")
    mid_x, mid_y = (fx + hx + hw) / 2, (fy + hy + hh) / 2
    for side, (px_, py_) in {"left": (fx - 1, mid_y), "top": (mid_x, fy - 1), "right": (hx + hw - 1, mid_y), "bottom": (mid_x, hy + hh - 1)}.items():
        assert pixel(px_, py_) == [255, 0, 0], (side, pixel(px_, py_))

    # GridCraft 4faa3ff0: with Wrap Text off a multi-line value shows on one line, as in Excel; the
    # breaks stay in the value.
    assert page.locator('.rg-pane [data-cell="J2"]').inner_text() == "onetwo"
    g.click("J2")
    g.wait_formula_bar("one\ntwo")
    # GridCraft 8eda5775: Page Up into the frozen row scrolls the rows below back to the top (row 10,
    # Page Down, Page Up, Page Up used to end on row 1 with rows 2-9 out of sight).
    g.select("B10")
    page.focus("#rg-editor")
    page.keyboard.press("PageDown")
    g.wait_idle(0.3)
    page.keyboard.press("PageUp")
    g.wait_idle(0.3)
    page.keyboard.press("PageUp")
    g.wait_idle(0.5)
    assert g.name_box() == "B1", g.name_box()
    page.wait_for_function("() => document.querySelector('.rg-scroller').scrollTop === 0 && [...document.querySelectorAll('.rg-rowhead')].some((e) => e.innerText.trim() === '2')", timeout=8000)
    # GridCraft #2451: Shift+Arrow on a whole column keeps the view on the rows it showed; it used
    # to scroll to row 1,048,576, the extension's far end.
    page.mouse.click(*header(g, "cols", "C"))
    g.wait_idle(0.3)
    page.focus("#rg-editor")
    page.keyboard.press("Shift+ArrowRight")
    g.wait_idle(0.6)
    assert page.evaluate("() => document.querySelector('.rg-scroller').scrollTop") < 50
    assert g.name_box().startswith("C"), g.name_box()
    page.locator(".rg-selection").wait_for(state="visible")
    g.click("B2")
    g.wait_idle(0.3)

    # GridCraft #1907: a finger dragging on the grid pans it; it doesn't select cells.
    ctx = page.context.browser.new_context(viewport={"width": 900, "height": 700}, has_touch=True, is_mobile=False)
    try:
        touch = ctx.new_page()
        t = Grid(touch, g.out)
        t.open(g.page.url.split("/book/")[0])
        assert t.name_box() == "A1"
        box = touch.locator(".rg-scroller").bounding_box()
        # The finger's own events (Input.dispatchTouchEvent), not Input.synthesizeScrollGesture: in headless
        # Chromium on Linux (CI) the synthesized gesture's moves never reach the page as touchmove, so nothing
        # scrolls; the same gesture scrolls on macOS.
        cdp = ctx.new_cdp_session(touch)
        x, y = box["x"] + 300, box["y"] + 400
        cdp.send("Input.dispatchTouchEvent", {"type": "touchStart", "touchPoints": [{"x": x, "y": y}]})
        for i in range(1, 16):
            cdp.send("Input.dispatchTouchEvent", {"type": "touchMove", "touchPoints": [{"x": x, "y": y - 20 * i}]})
        cdp.send("Input.dispatchTouchEvent", {"type": "touchEnd", "touchPoints": []})
        touch.wait_for_function("() => document.querySelector('.rg-scroller').scrollTop > 100", timeout=8000)
        t.wait_idle(0.5)
        assert t.name_box() == "A1", t.name_box()
    finally:
        ctx.close()

    november(g)


def menu_items(g: Grid, a1: str) -> list:
    """The labels of a cell's context menu (it stays open)."""
    x, y = g._exact_center(a1)
    g.page.mouse.click(x, y, button="right")
    g.page.locator(".rg-menu-item").first.wait_for(state="visible")
    return [g.page.locator(".rg-menu-item").nth(i).inner_text().split("\n")[0].strip() for i in range(g.page.locator(".rg-menu-item").count())]


def menu_pick(g: Grid, label: str) -> None:
    g.page.locator(".rg-menu-item", has_text=label).first.dispatch_event("mousedown")
    g.wait_idle(0.3)


def november(g: Grid) -> None:
    """GridCraft's issues closed from 2026-10-27 to 2026-11-06, on a workbook of their own."""
    page = g.page
    base = page.url.split("/book/")[0]
    g.open(base)

    # GridCraft #2778: AutoSum with nothing to add up opens =SUM() with the caret between the
    # parentheses, ready to point (nothing happened before).
    g.click("B2")
    page.keyboard.press("Alt+Equal")
    page.wait_for_function("() => document.querySelector('#rg-editor').value === '=SUM()'", timeout=8000)
    g.wait_idle(0.2)
    assert page.evaluate("() => document.querySelector('#rg-editor').selectionStart") == 5
    page.keyboard.type("3,4")
    page.keyboard.press("Enter")
    g.wait_text("B2", "7")

    # GridCraft #1442: Ctrl+. walks the corners of the selection clockwise; the range stays selected.
    for a, t in (("D4", "tl"), ("F4", "tr"), ("F6", "br"), ("D6", "bl")):
        g.type_in(a, t)
    g.select("D4:F6")
    page.focus("#rg-editor")
    seen = []
    for _ in range(4):
        g.key(f"{MOD}+Period")
        g.wait_idle(0.3)
        seen.append(g.formula_bar())
    assert seen == ["tr", "br", "bl", "tl"], seen
    page.locator(".rg-selection").wait_for(state="visible")

    # GridCraft #2918: Ctrl+Shift+Down while pointing extends the pointed range to the edge of the data.
    for r in range(1, 6):
        g.type_in(f"H{r}", str(r))
    g.click("J1")
    page.keyboard.type("=SUM(")
    g.click("H1")
    page.wait_for_function("() => document.querySelector('#rg-editor').value === '=SUM(H1'", timeout=8000)
    page.keyboard.press(f"{MOD}+Shift+ArrowDown")
    page.wait_for_function("() => document.querySelector('#rg-editor').value === '=SUM(H1:H5'", timeout=8000)
    page.keyboard.type(")")
    page.keyboard.press("Enter")
    g.wait_text("J1", "15")

    # GridCraft #1297: a cell showing a spilled value shows its anchor's formula, greyed.
    g.type_in("L1", "=SEQUENCE(3)")
    g.click("L2")
    page.wait_for_function("() => document.querySelector('#rg-formula-bar').placeholder === '=SEQUENCE(3)'", timeout=8000)
    assert g.formula_bar() == ""
    g.click("L1")
    g.wait_formula_bar("=SEQUENCE(3)")
    assert not page.locator("#rg-formula-bar").get_attribute("placeholder")

    # GridCraft #3136: dropping cells onto data asks first: Cancel keeps both, OK replaces.
    g.type_in("A10", "keep")
    g.type_in("C10", "move")

    def drag_border(src: str, dst: str) -> None:
        g.click(src)
        x, y, w, h = g.cell_rect(src)
        page.mouse.move(x + 1, y + h / 2)
        page.mouse.down()
        tx, ty = g._exact_center(dst)
        page.mouse.move(tx + 1, ty, steps=8)
        page.mouse.up()
        g.wait_idle(0.3)

    drag_border("C10", "A10")
    page.locator("#dlg-confirm").wait_for(state="visible")
    assert "There's already data here" in page.locator("#dlg-confirm-text").inner_text()
    page.click("#dlg-cancel")
    page.locator("#dlg-confirm").wait_for(state="detached")
    g.wait_text("A10", "keep")
    assert g.text("C10") == "move"
    drag_border("C10", "A10")
    page.locator("#dlg-confirm").wait_for(state="visible")
    g.dialog_ok("confirm")
    g.wait_text("A10", "move")
    assert g.text("C10") in (None, ""), g.text("C10")
    # A copy (Ctrl or ⌘ held; Option in Excel for Mac) replaces the data without asking, as Excel does.
    g.type_in("A11", "keep2")
    g.type_in("C11", "copy2")
    g.click("C11")
    x, y, w, h = g.cell_rect("C11")
    page.mouse.move(x + 1, y + h / 2)
    page.keyboard.down(MOD)
    page.mouse.down()
    tx, ty = g._exact_center("A11")
    page.mouse.move(tx + 1, ty, steps=8)
    page.mouse.up()
    page.keyboard.up(MOD)
    g.wait_text("A11", "copy2")
    assert g.text("C11") == "copy2" and page.locator("#dlg-confirm").count() == 0

    # GridCraft #2876 and #2120: a linked cell's menu edits, opens or removes its link; editing
    # starts from its text, address and ScreenTip, and keeps a formula there.
    page.evaluate("() => { window.__opened = null; window.open = (u) => { window.__opened = u; return null; }; }")
    g.type_in("A12", "=1+2")
    g.click("A12")
    page.click("#tab-insert")
    page.get_by_role("button", name="Link", exact=True).click()
    page.locator("#dlg-hyperlink").wait_for(state="visible")
    assert page.input_value("#dlg-text") == "3"
    page.fill("#dlg-target", "https://example.com/a")
    page.fill("#dlg-tooltip", "tip")
    g.dialog_ok("hyperlink")
    labels = menu_items(g, "A12")
    assert "Edit Hyperlink…" in labels and "Open Hyperlink" in labels and "Remove Hyperlink" in labels and "Link…" not in labels, labels
    menu_pick(g, "Edit Hyperlink…")
    page.locator("#dlg-hyperlink").wait_for(state="visible")
    assert page.locator("#dlg-hyperlink").get_by_text("Edit Hyperlink").count() >= 1
    assert (page.input_value("#dlg-text"), page.input_value("#dlg-target"), page.input_value("#dlg-tooltip")) == ("3", "https://example.com/a", "tip")
    page.fill("#dlg-target", "https://example.com/b")
    g.dialog_ok("hyperlink")
    g.click("A12")
    g.wait_formula_bar("=1+2")
    menu_items(g, "A12")
    menu_pick(g, "Open Hyperlink")
    page.wait_for_function("() => window.__opened === 'https://example.com/b'", timeout=8000)
    menu_items(g, "A12")
    menu_pick(g, "Remove Hyperlink")
    assert "Link…" in menu_items(g, "A12")
    page.keyboard.press("Escape")
    g.click("B12")

    # GridCraft #3224: with a range selected, Find All lists only the matches in it.
    for a in ("N1", "N2", "O1"):
        g.type_in(a, "pear")
    for sel, count in (("N1:N2", "2"), ("P5", "3")):
        g.select(sel)
        page.focus("#rg-editor")
        g.key("Mod+f")
        page.locator("#dlg-find").wait_for(state="visible")
        page.fill("#dlg-what", "pear")
        page.click("#dlg-find-all")
        page.wait_for_function("c => (document.querySelector('#dlg-status') || {}).innerText === c + ' cell(s) found'", arg=count, timeout=8000)
        page.keyboard.press("Escape")
        page.locator("#dlg-find").wait_for(state="detached")

    # GridCraft #2225: a cell with quotes is copied as it is; only tabs and line breaks quote.
    page.evaluate("() => navigator.clipboard.writeText('')")
    g.type_in("Q1", 'say "hi"')
    g.click("Q1")
    g.key("Mod+c")
    page.wait_for_function("async () => (await navigator.clipboard.readText()).startsWith('say')", timeout=8000)
    assert page.evaluate("() => navigator.clipboard.readText()").rstrip("\n") == 'say "hi"'
    page.keyboard.press("Escape")

    # GridCraft #3101: View › Show hides and shows the sheet tabs and the scroll bars.
    page.click("#tab-view")
    page.click("#rb-sheet-tabs")
    page.locator("#rg-sheet-tab-bar").wait_for(state="detached")
    page.click("#rb-sheet-tabs")
    page.locator("#rg-sheet-tab-bar").wait_for(state="visible")
    page.click("#rb-vscroll")
    page.wait_for_function("() => document.querySelector('.rg-scroller').classList.contains('rg-no-vbar')", timeout=8000)
    page.click("#rb-vscroll")
    page.wait_for_function("() => !document.querySelector('.rg-scroller').classList.contains('rg-no-vbar')", timeout=8000)

    # GridCraft #699: a picture on the clipboard (no text) is pasted at the active cell.
    g.click("S3")
    page.evaluate(
        """async () => {
            const c = document.createElement('canvas'); c.width = 30; c.height = 20;
            const x = c.getContext('2d'); x.fillStyle = '#c00'; x.fillRect(0, 0, 30, 20);
            const blob = await new Promise((r) => c.toBlob(r, 'image/png'));
            const dt = new DataTransfer(); dt.items.add(new File([blob], 'shot.png', { type: 'image/png' }));
            document.querySelector('#rg-editor').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
        }"""
    )
    page.locator(".rg-image").wait_for(state="visible")
    page.keyboard.press("Escape")
    g.click("B25")

    # GridCraft #1823: every series of a doughnut is a ring.
    g.type_in("D20", "k")
    g.type_in("E20", "a")
    g.type_in("F20", "b")
    for r, (k, a, b) in enumerate((("p", 1, 2), ("q", 3, 4)), start=21):
        g.type_in(f"D{r}", k)
        g.type_in(f"E{r}", str(a))
        g.type_in(f"F{r}", str(b))
    g.select("D20:F22")
    page.click("#tab-insert")
    page.click('button[title="Pie"]')
    page.locator("[role=menuitem]", has_text="Doughnut").click()
    page.locator("[role=menu]").wait_for(state="detached")
    chart = page.locator(".rg-chart").last
    chart.wait_for(state="visible")
    page.wait_for_function("() => [...document.querySelectorAll('.rg-chart')].some((c) => c.querySelectorAll('path[fill-rule=evenodd]').length === 4)", timeout=8000)
    page.focus("#rg-editor")
    page.keyboard.press("Delete")
    page.wait_for_function("() => !document.querySelectorAll('.rg-chart path[fill-rule=evenodd]').length", timeout=8000)

    # GridCraft #1343: Set Precision as Displayed asks first, then numbers are stored as they show.
    g.type_in("H20", "$1.234")
    g.type_in("I20", "=H20*1000")
    g.wait_text("I20", "1234")
    page.click("#tab-formulas")
    page.click('button[title="Calculation Options"]')
    page.locator("[role=menuitem]", has_text="Set Precision as Displayed").click()
    page.locator("#dlg-confirm").wait_for(state="visible")
    assert "permanently lose accuracy" in page.locator("#dlg-confirm-text").inner_text()
    g.dialog_ok("confirm")
    g.wait_text("I20", "1230")
    page.click('button[title="Calculation Options"]')
    page.locator("[role=menuitem]", has_text="Use Full Precision").click()
    g.wait_idle(0.5)
    g.type_in("H20", "$1.234")
    g.wait_text("I20", "1234")
    g.shot("november")

    # GridCraft #3166: a web page saved as .xls opens its table (it opened as lines of markup).
    src = g.out / "orders.xls"
    src.write_text(
        '<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="UTF-8"></head><body><table>'
        "<tr><th>Item</th><th>Qty</th></tr><tr><td>Kiwi</td><td>12</td></tr><tr><td>Plum</td><td x:num=\"4.5\">4,50</td></tr></table></body></html>",
        encoding="utf-8",
    )
    page.click("#tab-file")
    page.click("#bs-open")
    page.locator("#dlg-open").wait_for(state="visible")
    page.locator('#dlg-open input[type="file"]').set_input_files(str(src))
    g.wait_title("orders")
    g.wait_idle(0.8)
    g.wait_text("A1", "Item")
    g.wait_text("B3", "4.5")
    assert g.text("A2") == "Kiwi"


if __name__ == "__main__":
    sys.exit(run("gridcraft_issues", body))
