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
    g.shot("editing")


if __name__ == "__main__":
    sys.exit(run("editing_formulas", body))
