"""Formatting from the ribbon and Format Cells: bold, number formats, decimals, merge & center,
column width, wrap text, and the Format Cells dialog."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _lib.grid import Grid, run  # noqa: E402


def weight(g: Grid, a1: str) -> str:
    return g.page.locator(f'.rg-pane [data-cell="{a1}"]').first.evaluate("e => getComputedStyle(e).fontWeight")


def body(g: Grid) -> None:
    page = g.page
    g.type_in("A1", "1234.5")
    g.click("A1")
    page.click("#rb-bold")
    g.wait_idle(0.6)
    assert weight(g, "A1") == "700", weight(g, "A1")
    # Bold again toggles off.
    page.click("#rb-bold")
    g.wait_idle(0.6)
    assert weight(g, "A1") == "400", weight(g, "A1")

    # Number formats from the ribbon.
    page.locator('[aria-label="Comma Style"], button:has(svg.lucide-hash)').first.click()
    g.wait_text("A1", "1,234.50")
    page.locator("button:has(svg.lucide-decimals-arrow-right)").first.click()
    g.wait_text("A1", "1,234.500")
    page.locator("button:has(svg.lucide-percent)").first.click()
    g.wait_text("A1", "123450%")

    # Format Cells dialog: a custom code.
    g.click("A1")
    g.key("Mod+1")
    page.locator("#dlg-formatCells").wait_for(state="visible")
    page.fill("#dlg-numFmt", "0.000")
    g.dialog_ok("formatCells")
    g.wait_text("A1", "1234.500")

    # Merge & Center: B2:D2 becomes one cell centred over three columns.
    g.type_in("B2", "Title")
    g.select("B2:D2")
    page.click("#rb-merge")
    page.get_by_role("menuitem", name="Merge & Center").click()
    g.wait_idle(0.8)
    box = page.locator('.rg-pane [data-cell="B2"]').first.bounding_box()
    assert box["width"] > 150, f"merged cell width {box['width']}"
    g.click("C2")
    assert g.name_box() == "B2", g.name_box()

    # Wrap text grows the row.
    g.type_in("F3", "A long sentence that will wrap inside the cell")
    g.click("F3")
    page.click("#rb-wrap")
    g.wait_idle(0.8)
    rh = [page.locator(".rg-rowhead").nth(i) for i in range(6)]
    h3 = next(el.bounding_box()["height"] for el in rh if el.inner_text().strip() == "3")
    assert h3 > 30, f"row 3 height {h3}"

    # Column width by dragging the header border.
    head = page.locator(".rg-colhead").nth(7)  # H
    hb = head.bounding_box()
    x = hb["x"] + hb["width"] + 0.5
    page.mouse.move(x, hb["y"] + hb["height"] / 2)
    page.mouse.down()
    page.mouse.move(x + 60, hb["y"] + hb["height"] / 2, steps=6)
    page.mouse.up()
    g.wait_idle(0.8)
    w = page.locator(".rg-colhead").nth(7).bounding_box()["width"]
    assert w > hb["width"] + 40, f"column H width {w} (was {hb['width']})"

    # Borders and fill run without errors (drawn on canvas).
    g.select("B5:D7")
    page.click("#rb-borders")
    page.get_by_role("menuitem", name="All Borders").click()
    g.wait_idle(0.5)
    page.click("#rb-fill")
    page.locator('[title="#FFFF00"]').first.click()
    g.wait_idle(0.5)
    g.shot("formatting")


if __name__ == "__main__":
    sys.exit(run("formatting_ribbon", body))
