"""The page loads with the ribbon, formula bar, grid, sheet tabs and status bar."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _lib.grid import Grid, run  # noqa: E402


def body(g: Grid) -> None:
    page = g.page
    for tab in ("file", "home", "insert", "page-layout", "formulas", "data", "review", "view"):
        assert page.locator(f"#tab-{tab}").is_visible(), f"ribbon tab {tab} missing"
    assert page.locator("#rb-bold").is_visible(), "Bold button missing"
    assert g.name_box() == "A1", f"name box shows {g.name_box()!r}"
    assert page.locator("#rg-formula-bar").is_visible()
    heads = [page.locator(".rg-colhead").nth(i).inner_text() for i in range(5)]
    assert heads == ["A", "B", "C", "D", "E"], heads
    assert page.locator(".rg-rowhead").first.inner_text() == "1"
    assert page.locator(".rg-sheet-tab").first.inner_text().strip() == "Sheet1"
    assert page.locator("#status-mode").inner_text() == "Ready"
    assert page.locator("#doc-title").inner_text().startswith("Book")
    g.shot("home")


if __name__ == "__main__":
    sys.exit(run("smoke_home", body))
