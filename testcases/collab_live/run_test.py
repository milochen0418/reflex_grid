"""Three people on one workbook, live: everyone sees the others (avatars, "3 here") and where
they are (their selection boxes, with names), edits reach everyone at once, a cell someone else
changed lights up, each person has their own sheet, undo takes back only your own changes, the
title is shared, and leaving (or closing the tab) shows at once."""

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _lib.grid import Grid, run_people  # noqa: E402


def wait_until(fn, what: str, timeout: float = 8.0):
    end = time.time() + timeout
    last = None
    while time.time() < end:
        last = fn()
        if last:
            return last
        time.sleep(0.1)
    raise AssertionError(f"timed out waiting for {what} (last: {last!r})")


def people_here(g: Grid) -> str:
    return g.page.locator("#people-count").inner_text()


def peer_boxes(g: Grid):
    """{name: cell} of the other people's active cells drawn on this person's grid."""
    loc = g.page.locator(".rg-peer")
    return {loc.nth(i).get_attribute("data-peer"): loc.nth(i).get_attribute("data-peer-cell") for i in range(loc.count())}


def active_tab(g: Grid) -> str:
    return g.page.evaluate("() => { for (const t of document.querySelectorAll('.rg-sheet-tab')) if (getComputedStyle(t).fontWeight === '600') return t.innerText.trim(); return ''; }")


def body(grids) -> None:
    amy, ben, cat = grids
    names = {}
    for g in grids:
        wait_until(lambda: people_here(g) == "3 here", "3 here", 15)
        names[id(g)] = g.page.locator("#me-name").inner_text()
    assert amy.page.locator(".rg-person").count() == 3

    # Edits reach everyone; formulas recalculate for everyone.
    amy.type_in("B2", "40")
    ben.wait_text("B2", "40")
    cat.wait_text("B2", "40")
    ben.type_in("B3", "2")
    cat.type_in("B4", "=B2+B3")
    amy.wait_text("B4", "42")
    ben.wait_text("B4", "42")
    # ...and light up for the others for a moment.
    amy.type_in("C5", "fresh")
    ben.page.locator(".rg-flash").first.wait_for(state="attached", timeout=3000)

    # Where everyone is: each sees the others' selection, with their name.
    ben.click("D7")
    cat.drag("F2", "G4")
    wait_until(lambda: peer_boxes(amy).get(names[id(ben)]) == "D7", "Ben's box at D7")
    wait_until(lambda: peer_boxes(amy).get(names[id(cat)]) == "F2", "Cat's box at F2")
    assert amy.page.locator(".rg-peer-range").count() >= 1, "Cat's range F2:G4 isn't drawn"
    assert amy.page.locator(".rg-peer-tag", has_text=names[id(ben)]).count() == 1
    amy.shot("amy_sees_others")

    # Typing shows: Cat's box goes dashed with a pencil while she types.
    cat.click("H9")
    cat.page.keyboard.type("draft")
    wait_until(lambda: amy.page.locator(f'.rg-peer[data-peer="{names[id(cat)]}"][data-editing="1"]').count() == 1, "Cat typing")
    cat.page.keyboard.press("Escape")
    wait_until(lambda: amy.page.locator(f'.rg-peer[data-peer="{names[id(cat)]}"][data-editing="0"]').count() == 1, "Cat done")

    # Each person has their own sheet.
    ben.page.click("#add-sheet")
    ben.page.locator(".rg-sheet-tab", has_text="Sheet2").wait_for()
    amy.page.locator(".rg-sheet-tab", has_text="Sheet2").wait_for(timeout=8000)
    assert active_tab(ben) == "Sheet2" and active_tab(amy) == "Sheet1", (active_tab(ben), active_tab(amy))
    ben.type_in("A1", "on two")
    amy.type_in("A1", "on one")
    assert amy.text("A1") == "on one"
    ben.wait_text("A1", "on two")
    wait_until(lambda: names[id(ben)] not in peer_boxes(amy), "Ben off Sheet1 on Amy's grid")

    # Undo takes back only your own changes: Amy's B2 goes, Ben's B3 and Cat's B4 stay.
    amy.key("Mod+z")  # her A1 on Sheet1
    amy.key("Mod+z")  # her C5
    amy.key("Mod+z")  # her B2
    amy.wait_text("B4", "2")
    assert amy.text("B2") is None and amy.text("B3") == "2", (amy.text("B2"), amy.text("B3"))
    cat.wait_text("B4", "2")
    wait_until(lambda: cat.text("B2") is None, "B2 cleared on Cat's grid")
    ben.page.locator(".rg-sheet-tab", has_text="Sheet2").click()
    ben.wait_text("A1", "on two")

    # The title is everyone's.
    t = cat.page.locator("#doc-title")
    t.click()
    t.fill("Team plan")
    t.press("Enter")
    wait_until(lambda: amy.title() == "Team plan", "Amy sees the new title")
    wait_until(lambda: ben.title() == "Team plan", "Ben sees the new title")

    # Clicking an avatar jumps to where that person is.
    cat.click("K12")
    amy.page.locator(f'.rg-person[data-name="{names[id(cat)]}"]').click()
    wait_until(lambda: amy.name_box() == "K12", "Amy jumped to Cat's cell")

    # Leaving shows at once: going back to the workbook list, or closing the tab.
    ben.page.click("#home-link")
    ben.page.locator("#manager").wait_for()
    wait_until(lambda: people_here(amy) == "2 here", "2 here after Ben left", 10)
    cat.page.close(run_before_unload=True)
    wait_until(lambda: people_here(amy) == "1 here", "1 here after Cat closed her tab", 8)
    amy.shot("after_leave")


if __name__ == "__main__":
    sys.exit(run_people("collab_live", body, count=3))
