"""Sheets and data: sample workbooks, tables, charts, moving a chart and a picture, AutoFilter, sort,
data validation lists, sheet tabs (add, rename, switch)."""

import struct
import sys
import time
import zlib
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _lib.grid import Grid, run  # noqa: E402


def open_sample(g: Grid, name: str) -> None:
    """File › New › <sample>: a new shared workbook, opened in this tab."""
    g.page.click("#tab-file")
    g.page.click(f"#sample-{name.lower()}")
    g.wait_title(name)
    g.wait_idle(0.8)


def png(w: int, h: int) -> bytes:
    """A plain green ``w`` × ``h`` PNG."""

    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)

    rows = b"".join(b"\x00" + b"\x4e\xa7\x2e" * w for _ in range(h))
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b"")


def wait_box(obj, want: dict, what: str, timeout: float = 8.0) -> None:
    """Waits until a chart or picture is drawn at ``want`` (x, y, width, height, to the pixel)."""
    end = time.time() + timeout
    while True:
        b = obj.bounding_box()
        if all(abs(b[k] - want[k]) <= 1 for k in ("x", "y", "width", "height")):
            return
        if time.time() > end:
            raise AssertionError(f"{what}: at {b}, expected {want}")
        time.sleep(0.1)


def drag_object(g: Grid, obj, dx: float, dy: float) -> dict:
    """Drags a chart or picture by its body: it moves by (dx, dy) and keeps its size, both while it
    is dragged and once the server has it (only the handles resize). Returns where it is now."""
    page = g.page
    b0 = obj.bounding_box()
    want = {**b0, "x": b0["x"] + dx, "y": b0["y"] + dy}
    x, y = b0["x"] + b0["width"] / 2, b0["y"] + b0["height"] / 2
    page.mouse.move(x, y)
    page.mouse.down()
    page.mouse.move(x + dx, y + dy, steps=8)
    wait_box(obj, want, "while dragging", 2.0)
    page.mouse.up()
    wait_box(obj, want, "after the drop")
    return want


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

    # Dragging a chart or a picture by its body moves it without resizing it, either way; each drag
    # is one undo step.
    chart = page.locator(".rg-chart").first
    home = chart.bounding_box()
    away = drag_object(g, chart, 120, 60)
    drag_object(g, chart, -120, -60)
    g.key("Mod+z")
    wait_box(chart, away, "after one undo")
    g.key("Mod+z")
    wait_box(chart, home, "after two undos")
    g.click("H8")
    page.click("#tab-insert")
    page.get_by_role("button", name="Pictures", exact=True).click()
    page.locator('#dlg-insertPicture input[type="file"]').set_input_files({"name": "green.png", "mimeType": "image/png", "buffer": png(120, 80)})
    picture = page.locator(".rg-image").first
    picture.wait_for(state="visible")
    g.wait_idle(0.5)
    drag_object(g, picture, 70, 50)
    drag_object(g, picture, -70, -50)
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
