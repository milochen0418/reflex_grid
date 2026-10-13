"""Helpers shared by the E2E suites: open the app, address cells, read what the grid shows."""

from __future__ import annotations

import os
import sys
import time
import traceback
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Callable, List, Optional

from playwright.sync_api import Page, sync_playwright


def env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() not in {"0", "false", "no", "off"}


def col_index(letters: str) -> int:
    c = 0
    for ch in letters.upper():
        c = c * 26 + (ord(ch) - 64)
    return c - 1


def split_a1(a1: str):
    i = 0
    while i < len(a1) and a1[i].isalpha():
        i += 1
    return int(a1[i:]) - 1, col_index(a1[:i])


MOD = "Meta" if sys.platform == "darwin" else "Control"


def new_book_id() -> str:
    return "t" + uuid.uuid4().hex[:10]


class Grid:
    """A thin driver for the sheet view."""

    def __init__(self, page: Page, out: Path):
        self.page = page
        self.out = out
        self.console_errors: List[str] = []
        page.on("console", lambda m: self.console_errors.append(m.text) if m.type == "error" else None)
        page.on("pageerror", lambda e: self.console_errors.append(f"pageerror: {e}"))

    # -------------------------------------------------------------- app

    def open(self, base_url: str, book: Optional[str] = None) -> None:
        """Opens the workbook ``book`` (an id), else a new blank one of its own (a fresh address
        creates the workbook)."""
        self.book = book or new_book_id()
        self.page.goto(f"{base_url}/book/{self.book}", wait_until="networkidle")
        self.page.locator(".rg-view").wait_for(state="visible")
        self.page.locator(".rg-colhead").first.wait_for(state="visible")
        self.wait_idle()

    def title(self) -> str:
        return self.page.locator("#doc-title").input_value()

    def wait_title(self, expected: str, timeout: float = 20.0) -> None:
        """Waits for the workbook title (e.g. after opening a sample or a file, which opens a
        new workbook)."""
        self.page.wait_for_function(
            "t => { const el = document.querySelector('#doc-title'); return el && el.value === t; }",
            arg=expected,
            timeout=timeout * 1000,
        )
        self.page.locator(".rg-colhead").first.wait_for(state="visible")

    def wait_idle(self, t: float = 0.35) -> None:
        time.sleep(t)

    def shot(self, name: str) -> None:
        self.page.screenshot(path=str(self.out / f"{name}.png"))

    # -------------------------------------------------------------- geometry

    @staticmethod
    def _col_label(c: int) -> str:
        s = ""
        c += 1
        while c > 0:
            m = (c - 1) % 26
            s = chr(65 + m) + s
            c = (c - 1) // 26
        return s

    def click(self, a1: str, **kw) -> None:
        # Exact header match: "A" would also match "AA".
        x, y = self._exact_center(a1)
        self.page.mouse.click(x, y, **kw)
        self.wait_idle(0.15)

    def _exact_center(self, a1: str):
        r, c = split_a1(a1)
        label = self._col_label(c)
        heads = self.page.locator(".rg-colhead")
        cb = None
        for i in range(heads.count()):
            el = heads.nth(i)
            if el.inner_text().strip() == label:
                cb = el.bounding_box()
                break
        rows = self.page.locator(".rg-rowhead")
        rb = None
        for i in range(rows.count()):
            el = rows.nth(i)
            if el.inner_text().strip() == str(r + 1):
                rb = el.bounding_box()
                break
        if cb is None or rb is None:
            raise AssertionError(f"cell {a1} is not visible")
        return cb["x"] + cb["width"] / 2, rb["y"] + rb["height"] / 2

    def drag(self, a: str, b: str) -> None:
        x0, y0 = self._exact_center(a)
        x1, y1 = self._exact_center(b)
        self.page.mouse.move(x0, y0)
        self.page.mouse.down()
        self.page.mouse.move(x1, y1, steps=6)
        self.page.mouse.up()
        self.wait_idle(0.2)

    def dialog_ok(self, name: str) -> None:
        """Clicks OK in a dialog and waits until it has closed."""
        self.page.click("#dlg-ok")
        self.page.locator(f"#dlg-{name}").wait_for(state="detached")
        self.page.wait_for_function("() => !document.querySelector('[role=dialog]')")
        self.wait_idle(0.2)

    def select(self, ref: str) -> None:
        """Select a cell or range through the name box."""
        box = self.page.locator("#rg-namebox")
        box.click()
        box.fill(ref)
        box.press("Enter")
        self.wait_idle()

    # -------------------------------------------------------------- typing

    def type_in(self, a1: str, text: str, key: str = "Enter") -> None:
        self.click(a1)
        self.page.keyboard.type(text)
        self.page.keyboard.press(key)
        self.wait_idle()

    def key(self, combo: str) -> None:
        self.page.keyboard.press(combo.replace("Mod", MOD))
        self.wait_idle()

    # -------------------------------------------------------------- reading

    def text(self, a1: str) -> Optional[str]:
        loc = self.page.locator(f'.rg-pane [data-cell="{a1}"]')
        if loc.count() == 0:
            return None
        # Accounting-style formats render as left and right parts; compare the visible text.
        return " ".join(loc.first.inner_text().split())

    def wait_text(self, a1: str, expected: str, timeout: float = 8.0) -> str:
        end = time.time() + timeout
        got = None
        while time.time() < end:
            got = self.text(a1)
            if got == expected:
                return got
            time.sleep(0.1)
        raise AssertionError(f"{a1}: expected {expected!r}, got {got!r}")

    def formula_bar(self) -> str:
        return self.page.locator("#rg-formula-bar").input_value()

    def name_box(self) -> str:
        return self.page.locator("#rg-namebox").input_value()

    def wait_formula_bar(self, expected: str, timeout: float = 8.0) -> None:
        end = time.time() + timeout
        got = None
        while time.time() < end:
            got = self.formula_bar()
            if got == expected:
                return
            time.sleep(0.1)
        raise AssertionError(f"formula bar: expected {expected!r}, got {got!r}")


@contextmanager
def session(default_out: str):
    base_url = os.getenv("BASE_URL", "http://127.0.0.1:3000").rstrip("/")
    out = Path(os.getenv("OUTPUT_DIR", default_out))
    out.mkdir(parents=True, exist_ok=True)
    headless = env_bool("HEADLESS", True)
    slow = int(os.getenv("PW_SLOWMO_MS", "0"))
    timeout_ms = int(os.getenv("PW_TIMEOUT_MS", "30000"))
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, slow_mo=slow)
        context = browser.new_context(viewport={"width": 1400, "height": 860}, permissions=["clipboard-read", "clipboard-write"], accept_downloads=True)
        page = context.new_page()
        page.set_default_timeout(timeout_ms)
        grid = Grid(page, out)
        grid.open(base_url)
        try:
            yield grid
        except Exception:
            grid.shot("failure")
            raise
        finally:
            browser.close()


def run(name: str, body: Callable[[Grid], None]) -> int:
    """Runs a suite body; prints PASS/FAIL and returns the exit code."""
    try:
        with session(f"testcases/{name}/output") as g:
            body(g)
            errors = [e for e in g.console_errors if "Download the React DevTools" not in e]
            if errors:
                raise AssertionError("console errors: " + " | ".join(errors[:5]))
    except Exception:
        traceback.print_exc()
        print(f"[{name}] FAIL")
        return 1
    print(f"[{name}] PASS")
    return 0


# A fake microphone, for suites that make calls.
MEDIA_ARGS = ["--use-fake-ui-for-media-stream", "--use-fake-device-for-media-stream", "--autoplay-policy=no-user-gesture-required"]


@contextmanager
def people(default_out: str, count: int = 2, media: bool = False):
    """``count`` people, each in a browser context of their own (separate sessions, as on
    separate computers), on the same new workbook. ``media`` gives them a fake microphone."""
    base_url = os.getenv("BASE_URL", "http://127.0.0.1:3000").rstrip("/")
    out = Path(os.getenv("OUTPUT_DIR", default_out))
    out.mkdir(parents=True, exist_ok=True)
    headless = env_bool("HEADLESS", True)
    slow = int(os.getenv("PW_SLOWMO_MS", "0"))
    timeout_ms = int(os.getenv("PW_TIMEOUT_MS", "30000"))
    book = new_book_id()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, slow_mo=slow, args=MEDIA_ARGS if media else [])
        perms = ["clipboard-read", "clipboard-write"] + (["microphone"] if media else [])
        grids: List[Grid] = []
        try:
            for _ in range(count):
                context = browser.new_context(viewport={"width": 1280, "height": 800}, permissions=perms, accept_downloads=True)
                page = context.new_page()
                page.set_default_timeout(timeout_ms)
                g = Grid(page, out)
                g.base_url = base_url
                g.open(base_url, book)
                grids.append(g)
            yield grids
        except Exception:
            for i, g in enumerate(grids):
                try:
                    g.shot(f"failure_{i}")
                except Exception:
                    pass
            raise
        finally:
            browser.close()


def run_people(name: str, body: Callable[[List[Grid]], None], count: int = 2, ignore: Optional[Callable[[str], bool]] = None, media: bool = False) -> int:
    """Runs a suite body with ``count`` people on one workbook; prints PASS/FAIL. ``ignore``
    tells console errors the suite expects."""
    try:
        with people(f"testcases/{name}/output", count, media) as grids:
            body(grids)
            for i, g in enumerate(grids):
                errors = [e for e in g.console_errors if "Download the React DevTools" not in e and not (ignore and ignore(e))]
                if errors:
                    raise AssertionError(f"console errors (person {i}): " + " | ".join(errors[:5]))
    except Exception:
        traceback.print_exc()
        print(f"[{name}] FAIL")
        return 1
    print(f"[{name}] PASS")
    return 0
