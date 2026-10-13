"""Sheets and data: sample workbooks, tables, charts, AutoFilter, sort, data validation lists,
sheet tabs (add, rename, switch)."""

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _lib.grid import Grid, run  # noqa: E402


def open_sample(g: Grid, name: str) -> None:
    g.page.click("#tab-file")
    g.page.click(f"#sample-{name.lower()}")
    g.page.locator("#doc-title", has_text=f"{name}.xlsx").wait_for()
    g.wait_idle(0.8)


def body(g: Grid) -> None:
    page = g.page
    # Sheet tabs.
    page.click("#add-sheet")
    page.locator(".rg-sheet-tab", has_text="Sheet2").wait_for()
    page.locator(".rg-sheet-tab", has_text="Sheet2").dblclick()
    page.locator("#dlg-renameSheet").wait_for(state="visible")
    page.fill("#dlg-sheet-name", "Data")
    g.dialog_ok("renameSheet")
    page.locator(".rg-sheet-tab", has_text="Data").wait_for()
    page.locator(".rg-sheet-tab", has_text="Sheet1").click()
    g.wait_idle(0.5)

    # Sales sample: a table with a calculated column, a chart and filter buttons.
    open_sample(g, "Sales")
    g.wait_text("A1", "Region")
    g.wait_text("F2", "$1,368.00")
    assert page.locator(".rg-chart").count() == 1, "chart missing"
    assert page.locator(".rg-chart svg rect").count() >= 4
    assert page.locator(".rg-filter-btn").count() >= 6
    # Filter Region: keep only North.
    page.locator('.rg-filter-btn[data-filter="A1"]').click()
    page.locator("#dlg-filter").wait_for(state="visible")
    dialog = page.locator("#dlg-filter")
    for region in ("South", "East", "West"):
        dialog.get_by_text(region, exact=True).click()
    g.dialog_ok("filter")
    g.wait_idle(1.0)
    visible = [page.locator(".rg-rowhead").nth(i).inner_text() for i in range(4)]
    assert visible[:3] == ["1", "2", "6"], f"rows after filter: {visible}"
    assert g.text("A6") == "North"
    # Clear it again.
    page.locator('.rg-filter-btn[data-filter="A1"]').click()
    page.locator("#dlg-filter").wait_for(state="visible")
    page.get_by_role("button", name="Clear Filter").click()
    g.wait_idle(1.0)
    g.wait_text("A3", "South")

    # Sort the summary table descending by Share (column J).
    g.click("J2")
    page.click("#tab-data")
    page.click("#rb-sort-desc")
    g.wait_idle(0.8)
    g.wait_text("H2", "North")
    page.click("#tab-home")

    # Grades sample: list validation with an in-cell dropdown drives XLOOKUP.
    open_sample(g, "Grades")
    g.wait_text("J2", "76.3")
    g.click("I2")
    page.locator(".rg-list-btn").wait_for(state="visible")
    page.locator(".rg-list-btn").click()
    page.locator(".rg-list-item", has_text="Harper").click()
    g.wait_text("I2", "Harper")
    g.wait_text("J2", "82.5")
    # Typing a value outside the list is refused.
    g.click("I2")
    page.keyboard.type("Nobody")
    page.keyboard.press("Enter")
    page.get_by_text("doesn't match the data validation").first.wait_for(state="visible")
    page.keyboard.press("Escape")
    g.wait_text("I2", "Harper")
    g.shot("grades")


if __name__ == "__main__":
    sys.exit(run("sheets_data", body))
