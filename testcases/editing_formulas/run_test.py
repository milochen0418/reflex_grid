"""Typing, formulas and recalculation, the formula bar, autocomplete, AutoSum, fill handle,
copy/paste with relative references, undo/redo and formula errors."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _lib.grid import MOD, Grid, run  # noqa: E402


def body(g: Grid) -> None:
    page = g.page
    # Values and a formula; recalculation after an edit.
    g.click("A1")
    for v in ("10", "20", "30"):
        page.keyboard.type(v)
        page.keyboard.press("Enter")
    page.keyboard.type("=SUM(A1:A3)")
    page.keyboard.press("Enter")
    g.wait_text("A4", "60")
    g.type_in("A2", "25")
    g.wait_text("A4", "65")
    g.click("A4")
    g.wait_formula_bar("=SUM(A1:A3)")
    assert g.name_box() == "A4"

    # Typed input is recognised like Excel: percent, date, currency, text with apostrophe.
    g.type_in("B1", "50%")
    g.wait_text("B1", "50%")
    g.type_in("B2", "3/14/2026")
    g.wait_text("B2", "3/14/2026")
    g.type_in("B3", "$1,234.5")
    g.wait_text("B3", "$1,234.50")
    g.type_in("B4", "'007")
    g.wait_text("B4", "007")

    # Missing parentheses are closed; F2 edits the formula, not the value.
    g.type_in("C1", "=ROUND(PI(),2")
    g.wait_text("C1", "3.14")
    g.click("C1")
    page.keyboard.press("F2")
    g.wait_idle(0.2)
    assert page.locator("#rg-editor").input_value() == "=ROUND(PI(),2)", page.locator("#rg-editor").input_value()
    page.keyboard.press("Escape")

    # Function autocomplete.
    g.click("C2")
    page.keyboard.type("=AVER")
    page.locator(".rg-ac").wait_for(state="visible")
    page.keyboard.press("Tab")
    page.keyboard.type("A1:A3)")
    page.keyboard.press("Enter")
    g.wait_text("C2", "21.66667")  # General fits the standard column width, as in Excel

    # AutoSum below a column of numbers.
    g.click("A5")
    page.keyboard.press("Alt+Equal")
    g.wait_text("A5", "130")
    g.click("A5")
    g.wait_formula_bar("=SUM(A1:A4)")

    # Fill handle: 1, 2 -> series to 6.
    g.type_in("D1", "1")
    g.type_in("D2", "2")
    g.drag("D1", "D2")
    hx, hy = g._exact_center("D2")
    box = page.locator(".rg-colhead", has_text="D").first.bounding_box()
    handle_x = box["x"] + box["width"] - 1
    handle_y = hy + 9
    page.mouse.move(handle_x, handle_y)
    page.mouse.down()
    _, y6 = g._exact_center("D6")
    page.mouse.move(handle_x, y6, steps=8)
    page.mouse.up()
    g.wait_text("D6", "6")

    # Copy a formula and paste it: references move.
    g.type_in("E1", "=D1*10")
    g.click("E1")
    g.key("Mod+c")
    g.drag("E2", "E4")
    g.key("Mod+v")
    g.wait_text("E4", "40")
    g.click("E4")
    g.wait_formula_bar("=D4*10")

    # Undo and redo.
    g.key("Mod+z")
    g.wait_idle(0.5)
    assert g.text("E4") in (None, ""), f"E4 after undo: {g.text('E4')!r}"
    g.key("Mod+y")
    g.wait_text("E4", "40")

    # A formula with the wrong number of arguments is refused and the editor reopens.
    g.click("F1")
    page.keyboard.type("=IF(1)")
    page.keyboard.press("Enter")
    page.get_by_text("IF has too few arguments").first.wait_for(state="visible")
    g.wait_idle(0.5)
    assert "rg-editing" in (page.locator("#rg-editor").get_attribute("class") or ""), "editor did not reopen"
    page.keyboard.press("Escape")

    # Dynamic arrays spill.
    g.type_in("G1", "=SEQUENCE(3)")
    g.wait_text("G3", "3")

    # IME input (e.g. Zhuyin/Pinyin): composition edits in place, the committed text lands.
    cdp = page.context.new_cdp_session(page)
    g.click("H1")
    for partial in ("n", "ni", "ni hao"):
        cdp.send("Input.imeSetComposition", {"text": partial, "selectionStart": len(partial), "selectionEnd": len(partial)})
    cdp.send("Input.insertText", {"text": "你好"})
    page.keyboard.press("Enter")
    g.wait_text("H1", "你好")
    assert g.name_box() == "H2", g.name_box()

    # A big block. Pasted from another app: more than Reflex's default 1 MB websocket limit, past
    # which the paste was dropped without a word. Then copied in ReflexGrid (off-screen, so the
    # server makes the clipboard text) and pasted elsewhere, by token: the cells don't go back.
    rows, cols = 15000, 10
    big = "\n".join("\t".join(f"b{r * cols + c}" for c in range(cols)) for r in range(rows)) + "\n"
    assert len(big) > 1_000_000, len(big)
    g.select("J1")
    page.focus("#rg-editor")
    page.evaluate(
        """(text) => { const el = document.getElementById('rg-editor'); el.focus(); const dt = new DataTransfer();
        dt.setData('text/plain', text); el.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); }""",
        big,
    )
    g.wait_text("J20", "b190", timeout=60)
    g.wait_text("S20", "b199")
    # The pasted block is selected: copy it.
    page.focus("#rg-editor")
    g.key("Mod+c")
    page.wait_for_function("(n) => navigator.clipboard.readText().then((t) => t.length === n)", arg=len(big), timeout=60000, polling=250)
    g.select("U1")
    page.focus("#rg-editor")
    g.key("Mod+v")
    g.wait_text("U20", "b190", timeout=60)
    g.select(f"AD{rows}")
    g.wait_formula_bar(f"b{rows * cols - 1}")
    # After Esc (the copy is cancelled), Ctrl+V still pastes what the system clipboard holds, as text.
    g.select("J2")
    page.focus("#rg-editor")
    g.key("Mod+c")
    page.wait_for_function("() => navigator.clipboard.readText().then((t) => t === 'b10\\n')", timeout=30000, polling=250)
    page.focus("#rg-editor")
    g.key("Escape")
    g.select("AF1")
    page.focus("#rg-editor")
    g.key("Mod+v")
    g.wait_formula_bar("b10")
    g.shot("editing")


if __name__ == "__main__":
    sys.exit(run("editing_formulas", body))
