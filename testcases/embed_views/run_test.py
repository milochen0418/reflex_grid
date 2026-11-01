"""A workbook on other apps' pages, the pages as another app's frames load them: a table to read and
a UserForm to use. The form's new row reaches the table and the workbook's own tab at once, a form that
unloaded itself opens again for the next entry, a MsgBox shows over the form, the person on the form
isn't one of the people on the workbook, the table has nothing to edit with, the cells of a growing list
follow it, Insert › Embed gives the lines to paste, grid.pick chooses a form, a table or dragged cells, and
an address that leads nowhere says so and creates nothing."""

import base64
import os
import sys
import time
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from _lib.grid import Grid, run  # noqa: E402

from gridcore.engine import Session  # noqa: E402
from gridcore.model.forms import FormDesign, new_control  # noqa: E402
from gridcore.xlsx.oforms import vbframe, write_design, write_frx  # noqa: E402

BASE = os.getenv("BASE_URL", "http://127.0.0.1:3000").rstrip("/")
# The backend's own routes: nginx serves them at the same address, the dev server doesn't.
BACKEND = os.getenv("BACKEND_URL", BASE.replace(":3000", ":8000")).rstrip("/")

ENTRY = FormDesign("EntryForm", caption="New order", width=264, height=108, controls=(
    new_control("Label1", "Label", left=12, top=12, width=54, caption="Name:"),
    new_control("TextBox1", "TextBox", left=72, top=10, width=170, tab_index=0),
    new_control("Label2", "Label", left=12, top=38, width=54, caption="Qty:"),
    new_control("TextBox2", "TextBox", left=72, top=36, width=60, tab_index=1),
    new_control("CommandButton1", "CommandButton", left=96, top=72, width=72, tab_index=2, caption="Add", default=True),
    new_control("CommandButton2", "CommandButton", left=180, top=72, width=72, tab_index=3, caption="Check"),
))

ENTRY_CODE = '''Private Sub CommandButton1_Click()
    Dim row As ListRow
    Set row = Worksheets("Orders").ListObjects("Orders").ListRows.Add
    row.Range(1, 1).Value = TextBox1.Value
    row.Range(1, 2).Value = CLng(TextBox2.Value)
    Dim r As Long
    r = Worksheets("Log").Cells(Rows.Count, 1).End(xlUp).Row + 1
    Worksheets("Log").Cells(r, 1).Value = TextBox1.Value
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    MsgBox "Checked " & TextBox1.Text
End Sub
'''


def build_xlsm(path: Path) -> None:
    s = Session()
    s.new_workbook()
    s.run("sheet.rename", name="Orders")
    for cell, text in (("A1", "Name"), ("B1", "Qty"), ("A2", "Amy"), ("B2", "3")):
        s.run("cell.set", cell=cell, input=text)
    s.run("insert.table", range="A1:B2", header=True, name="Orders")
    s.run("home.insertSheet", name="Log")
    s.run("cell.set", cell="A1", input="Who")
    s.run("cell.set", cell="A2", input="Amy")
    text = vbframe(ENTRY).replace("End\r\n", f'   OleObjectBlob   =   "{ENTRY.name}.frx":0000\r\nEnd\r\n', 1)
    text += f'Attribute VB_Name = "{ENTRY.name}"\r\n' + ENTRY_CODE.replace("\n", "\r\n")
    s.run("macro.importModule", fileName=f"{ENTRY.name}.frm", text=text, frx=base64.b64encode(write_frx(write_design(ENTRY), ENTRY)).decode())
    assert s.run("macro.list")["diagnostics"] == []
    s.run("sheet.activate", sheet="Orders")
    path.write_bytes(base64.b64decode(s.run("file.saveBytes")["data"]))


def table_url(book: str, ref: str) -> str:
    return f"{BACKEND}/_reflex_grid/embed/table?book={book}&range={quote(ref)}"


def form_url(book: str, form: str) -> str:
    return f"{BASE}/intent/grid-form?book={book}&form={form}"


def rows_of(page) -> list:
    return [r.split("\t") for r in page.locator("tr").all_inner_texts()]


def wait_rows(page, n: int, timeout: float = 8.0) -> None:
    page.wait_for_function("(n) => document.querySelectorAll('tr').length === n", arg=n, timeout=timeout * 1000)


def body(g: Grid) -> None:
    page = g.page
    xlsm = g.out / "orders.xlsm"
    build_xlsm(xlsm)
    g.open(BASE)
    page.click("#tab-file")
    page.click("#bs-open")
    page.locator("#dlg-open").wait_for(state="visible")
    page.locator('#dlg-open input[type="file"]').set_input_files(str(xlsm))
    g.wait_title("orders")
    g.wait_idle(1.0)
    book = page.url.rstrip("/").rsplit("/", 1)[-1]

    # The lines to paste, from the workbook.
    page.click("#tab-insert")
    page.click("#rb-embed")
    page.locator("#dlg-embed").wait_for(state="visible")
    codes = [i.input_value() for i in page.locator("#dlg-embed input").all()]
    assert codes == [f"{{%grid {book} range=Orders %}}", f"{{%grid {book} form=EntryForm %}}"], codes
    g.shot("embed_dialog")
    page.keyboard.press("Escape")
    page.click("#tab-home")

    # A table to read: nothing on it to edit with.
    table = page.context.new_page()
    table.goto(table_url(book, "Orders"))
    assert rows_of(table) == [["Name", "Qty"], ["Amy", "3"]], rows_of(table)
    assert table.locator("#foot").inner_text().replace("\n", " ").endswith("2 rows")
    assert table.locator("input, textarea, select, button, [contenteditable]").count() == 0
    log = page.context.new_page()
    log.goto(table_url(book, "Log!A1:A"))
    assert len(rows_of(log)) == 2

    # A form to use: the row reaches the table, and the workbook's tab, at once.
    form = page.context.new_page()
    form.goto(form_url(book, "EntryForm"))
    shown = form.locator('.rg-userform[data-form="EntryForm"]')
    shown.wait_for(state="visible")
    assert form.locator(".rg-sheet, .rg-view, .rg-ribbon").count() == 0, "a form brings no sheet with it"
    assert form.locator("[data-form-close]").count() == 0
    form.locator('[data-control="TextBox1"]').fill("Ben")
    form.locator('[data-control="TextBox2"]').fill("5")
    start = time.time()
    form.locator('[data-control="CommandButton1"]').click()
    wait_rows(table, 3)
    took = time.time() - start
    assert took < 3, f"{took:.2f} s to reach the table"
    assert rows_of(table)[-1] == ["Ben", "5"], rows_of(table)
    wait_rows(log, 3)
    g.wait_text("A3", "Ben")
    table.locator("tr.fresh").wait_for(state="attached")
    table.screenshot(path=str(g.out / "table.png"))

    # The form unloaded itself: the next entry starts on a new one.
    form.wait_for_function("""() => {
        const t = document.querySelector('[data-control="TextBox1"]');
        return t && t.value === "";
    }""", timeout=8000)

    # A MsgBox over the form, then back to it.
    form.locator('[data-control="TextBox1"]').fill("Cleo")
    form.locator('[data-control="CommandButton2"]').click()
    form.locator("#mb-prompt").wait_for(state="visible")
    assert form.locator("#mb-prompt").inner_text() == "Checked Cleo"
    form.screenshot(path=str(g.out / "form_msgbox.png"))
    form.locator("#mb-ok").click()
    form.locator("#mb-prompt").wait_for(state="detached")
    assert form.locator('[data-control="TextBox1"]').input_value() == "Cleo"

    # Whoever is on the form isn't on the workbook.
    page.bring_to_front()
    assert page.locator("#people-count").inner_text() == "1 here", "only this tab is here"

    # Edits made on the workbook show too, and the page goes with the workbook's data.
    g.type_in("A2", "Amy Lee")
    table.wait_for_function("() => document.body.innerText.includes('Amy Lee')", timeout=8000)

    # Choosing: a form, a table, cells dragged.
    pick = page.context.new_page()
    pick.goto(f"{BASE}/intent/grid-pick")
    pick.locator(f'[data-book="{book}"]').wait_for(state="visible")
    assert pick.locator("[data-book]").count() >= 2, "the blank workbook this suite started on, and this one"
    pick.locator("#pick-search").fill("ORDERS")
    pick.wait_for_function("() => document.querySelectorAll('[data-book]').length === 1", timeout=8000)
    pick.locator(f'[data-book="{book}"]').click()
    pick.locator('[data-pick="EntryForm"]').wait_for(state="visible")
    pick.locator('[data-pick="Orders"]').wait_for(state="visible")
    assert pick.locator("#pick-ref").inner_text() == "Orders", "it starts on the sheet's table"
    pick.locator('[data-sheet="Log"]').click()
    pick.wait_for_function("() => document.getElementById('pick-ref').innerText === 'Log!A1:A'", timeout=8000)
    cell = lambda r, c: pick.locator(f'[data-cell="{r},{c}"]').bounding_box()
    a, b = cell(0, 0), cell(1, 0)
    pick.mouse.move(a["x"] + 4, a["y"] + 4)
    pick.mouse.down()
    pick.mouse.move(b["x"] + 4, b["y"] + 4, steps=3)
    pick.mouse.up()
    pick.wait_for_function("() => document.getElementById('pick-ref').innerText === 'Log!A1:A2'", timeout=8000)
    pick.locator("#pick-grow").wait_for(state="visible")
    assert pick.locator("[data-cell] input, [data-cell] textarea").count() == 0
    pick.close()

    # Addresses that lead nowhere.
    gone = page.context.new_page()
    gone.goto(table_url("t0000000000", "A1"))
    assert "isn't there" in gone.locator("#error").inner_text()
    gone.goto(form_url("t0000000000", "EntryForm"))
    gone.locator("#embed-error").wait_for(state="visible")
    gone.goto(form_url(book, "NoForm"))
    gone.locator("#embed-error").wait_for(state="visible")
    assert "NoForm" in gone.locator("#embed-error").inner_text()
    gone.goto(f"{BASE}/")
    gone.locator("#book-search").wait_for(state="visible")
    time.sleep(1.0)
    assert "t0000000000" not in gone.content(), "no workbook was made for it"
    gone.close()


if __name__ == "__main__":
    sys.exit(run("embed_views", body))
