"""UserForms in the browser: open an .xlsm whose VBA project has UserForms (built here with
gridcore), run the macro that shows one and use it — type, pick from a drop-down list, tick a box,
a KeyPress handler that drops letters, a MsgBox over the form, the OK button whose Click writes
the cells and unloads the form — then a modeless form beside the grid, the VBA editor's form view
(double-click a control for its event procedure, Run Form), designing a new form (controls from
the toolbox, the Properties window, moving one, its code, running it), and the forms kept in the
downloaded .xlsm. The editor stays light when the browser remembers a dark colour mode."""

import base64
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from _lib.grid import MOD, Grid, run  # noqa: E402

from gridcore.engine import Session, open_bytes  # noqa: E402
from gridcore.model.forms import FormDesign, new_control  # noqa: E402
from gridcore.xlsx.oforms import design_streams, read_design, vbframe, write_design, write_frx  # noqa: E402

ORDER = FormDesign("OrderForm", caption="Order", width=264, height=156, controls=(
    new_control("Label1", "Label", left=12, top=12, width=54, caption="Name:"),
    new_control("TextBox1", "TextBox", left=72, top=10, width=150, tab_index=0),
    new_control("Label2", "Label", left=12, top=38, width=54, caption="Fruit:"),
    new_control("ComboBox1", "ComboBox", left=72, top=36, width=150, tab_index=1, style=2),
    new_control("CheckBox1", "CheckBox", left=72, top=62, width=150, tab_index=2, caption="Rush order"),
    new_control("Label3", "Label", left=12, top=90, width=54, caption="Qty:"),
    new_control("TextBox2", "TextBox", left=72, top=88, width=60, tab_index=3),
    new_control("CommandButton3", "CommandButton", left=12, top=122, width=72, tab_index=4, caption="Greet"),
    new_control("CommandButton1", "CommandButton", left=96, top=122, width=72, tab_index=5, caption="OK", default=True),
    new_control("CommandButton2", "CommandButton", left=180, top=122, width=72, tab_index=6, caption="Cancel", cancel=True),
))

ORDER_CODE = '''Private Sub UserForm_Initialize()
    ComboBox1.AddItem "Apple"
    ComboBox1.AddItem "Pear"
    ComboBox1.AddItem "Plum"
End Sub

Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub CommandButton1_Click()
    Range("A1").Value = TextBox1.Value
    Range("A2").Value = ComboBox1.Value
    Range("A3").Value = CheckBox1.Value
    Range("A4").Value = TextBox2.Value
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub CommandButton3_Click()
    MsgBox "Hello " & TextBox1.Text
End Sub
'''

NOTES = FormDesign("NoteForm", caption="Notes", width=200, height=60, show_modal=False, start_up_position=0, left=900, top=120, controls=(
    new_control("TextBox1", "TextBox", left=8, top=8, width=120, tab_index=0),
    new_control("CommandButton1", "CommandButton", left=136, top=6, width=56, tab_index=1, caption="Add"),
    new_control("Label1", "Label", left=8, top=36, width=180, caption="0 notes"),
))

NOTES_CODE = '''Private Count As Long

Private Sub CommandButton1_Click()
    Count = Count + 1
    Cells(Count, 3).Value = TextBox1.Text
    Label1.Caption = Count & " notes"
    TextBox1.Text = ""
End Sub
'''

MODULE = '''Sub ShowOrder()
    OrderForm.Show
    Range("B1").Value = "after"
End Sub

Sub ShowNotes()
    NoteForm.Show
End Sub
'''


def build_xlsm(path: Path) -> None:
    """A workbook with the two forms, as the VBA editor would have them, saved as .xlsm."""
    s = Session()
    s.new_workbook()
    for design, code in ((ORDER, ORDER_CODE), (NOTES, NOTES_CODE)):
        text = vbframe(design).replace("End\r\n", f'   OleObjectBlob   =   "{design.name}.frx":0000\r\nEnd\r\n', 1)
        text += f'Attribute VB_Name = "{design.name}"\r\n' + code.replace("\n", "\r\n")
        frx = write_frx(write_design(design), design)
        s.run("macro.importModule", fileName=f"{design.name}.frm", text=text, frx=base64.b64encode(frx).decode())
    s.run("macro.setModule", name="Module1", code=MODULE)
    assert s.run("macro.list")["diagnostics"] == []
    path.write_bytes(base64.b64decode(s.run("file.saveBytes")["data"]))


def form(g: Grid, name: str):
    return g.page.locator(f'.rg-userform[data-form="{name}"]')


def ctl(g: Grid, form_name: str, control: str):
    return g.page.locator(f'.rg-userform[data-form="{form_name}"] [data-control="{control}"]')


def wait_value(loc, expected: str, timeout: float = 8.0) -> None:
    end = time.time() + timeout
    got = None
    while time.time() < end:
        got = loc.input_value()
        if got == expected:
            return
        time.sleep(0.1)
    raise AssertionError(f"expected {expected!r}, got {got!r}")


def run_macro(g: Grid, name: str) -> None:
    page = g.page
    page.keyboard.press("Alt+F8")
    page.locator("#dlg-macros").wait_for(state="visible")
    page.locator(".dlg-macro", has_text=name).first.click()
    page.click("#dlg-macro-run")
    page.locator("#dlg-macros").wait_for(state="detached")
    # The dialog covers the page until its closing animation has ended.
    page.wait_for_function("() => !document.querySelector('.rt-BaseDialogScroll') && getComputedStyle(document.body).pointerEvents !== 'none'")


def light_in_dark_mode(g: Grid) -> None:
    """A browser in dark mode that remembers a "system" colour mode (left by an older version)
    still shows the VBA editor light, with its dark text readable."""
    page = g.page
    page.emulate_media(color_scheme="dark")
    page.evaluate("() => localStorage.setItem('theme', 'system')")
    g.open(os.getenv("BASE_URL", "http://127.0.0.1:3000").rstrip("/"), g.book)
    page.click("#tab-developer")
    page.click("#rb-vba")
    page.locator("#dlg-vba").wait_for(state="visible")
    time.sleep(0.4)
    bg = page.evaluate("() => getComputedStyle(document.getElementById('dlg-vba')).backgroundColor")
    assert bg == "rgb(255, 255, 255)", f"the editor is {bg} in dark mode"
    g.shot("editor_dark_mode")
    page.click("#vba-close")
    page.locator("#dlg-vba").wait_for(state="detached")
    page.emulate_media(color_scheme="light")


def design_a_form(g: Grid) -> None:
    """Insert › UserForm, a TextBox and a CommandButton from the toolbox, the button's caption in
    the Properties window, the box moved, the form renamed through its (Name), the button's Click
    by double-clicking it, Run Form."""
    page = g.page
    page.click("#tab-developer")
    page.click("#rb-vba")
    page.locator("#dlg-vba").wait_for(state="visible")
    page.click("#vba-add-form")
    page.locator('.rg-form-designer[data-form="UserForm1"]').wait_for(state="visible")
    client = page.locator(".rg-form-designer-client")

    def place(kind: str, x: float, y: float) -> None:
        page.click(f"#vba-tool-{kind}")
        page.locator(f'.rg-form-designer-client[data-tool="{kind}"]').wait_for()
        b = client.bounding_box()
        page.mouse.click(b["x"] + x, b["y"] + y)
        page.locator(f'.rg-form-designer [data-control="{kind}1"]').wait_for()

    place("TextBox", 40, 30)
    place("CommandButton", 40, 90)
    caption = page.locator("#vba-prop-caption input[name=value]")
    caption.fill("Save")
    caption.press("Enter")
    page.wait_for_function("() => document.querySelector('.rg-form-designer [data-control=\"CommandButton1\"]').innerText.includes('Save')")
    box = page.locator('.rg-form-designer [data-control="TextBox1"]').bounding_box()
    page.mouse.move(box["x"] + 10, box["y"] + 8)
    page.mouse.down()
    page.mouse.move(box["x"] + 70, box["y"] + 8, steps=6)
    page.mouse.up()
    page.wait_for_function("() => { const r = document.getElementById('vba-prop-left'); return r && parseFloat(r.querySelector('input').value) > 60; }")
    # The form itself (an empty spot), then its (Name) in the Properties window: the form, the
    # project tree and the editor take the new name.
    b = client.bounding_box()
    page.mouse.click(b["x"] + b["width"] - 12, b["y"] + 12)
    page.wait_for_function("() => (document.getElementById('vba-properties') || {}).innerText.includes('ShowModal')")
    name = page.locator("#vba-prop-name input[name=value]")
    name.fill("frmDesigned")
    name.press("Enter")
    page.locator('.rg-form-designer[data-form="frmDesigned"]').wait_for(state="visible")
    page.locator("#vba-mod-frmDesigned").wait_for(state="visible")
    assert page.locator("#vba-mod-UserForm1").count() == 0
    g.shot("designer")
    button = page.locator('.rg-form-designer [data-control="CommandButton1"]').bounding_box()
    page.mouse.dblclick(button["x"] + 10, button["y"] + 8)
    page.wait_for_function("() => { const a = document.getElementById('vba-code'); return a && a.value.includes('Private Sub CommandButton1_Click()'); }")
    area = page.locator("#vba-code")
    area.fill('Private Sub CommandButton1_Click()\n    Range("A6").Value = TextBox1.Value\n    Unload Me\nEnd Sub\n')
    area.press(f"{MOD}+s")
    page.wait_for_function("() => !document.querySelector('#vba-dirty')")
    page.click("#vba-run-form")
    form(g, "frmDesigned").wait_for(state="visible")
    ctl(g, "frmDesigned", "TextBox1").fill("designed here")
    ctl(g, "frmDesigned", "CommandButton1").click()
    form(g, "frmDesigned").wait_for(state="detached")
    page.locator("#dlg-vba").wait_for(state="visible")
    page.click("#vba-close")
    page.locator("#dlg-vba").wait_for(state="detached")
    g.wait_text("A6", "designed here")


def body(g: Grid) -> None:
    page = g.page
    light_in_dark_mode(g)
    xlsm = g.out / "forms.xlsm"
    build_xlsm(xlsm)
    page.click("#tab-file")
    page.click("#bs-open")
    page.locator("#dlg-open").wait_for(state="visible")
    page.locator('#dlg-open input[type="file"]').set_input_files(str(xlsm))
    g.wait_title("forms")
    g.wait_idle(1.0)

    # A modal form: the macro waits while the person fills it in.
    g.click("F10")
    run_macro(g, "ShowOrder")
    form(g, "OrderForm").wait_for(state="visible")
    assert page.locator(".rg-userform-backdrop").count() == 1
    g.shot("form")
    ctl(g, "OrderForm", "TextBox1").fill("Ann")
    ctl(g, "OrderForm", "ComboBox1").click()
    ctl(g, "OrderForm", "ComboBox1").locator('[data-row="1"]').click()
    page.wait_for_function("() => document.querySelector('.rg-userform [data-control=\"ComboBox1\"]').innerText.includes('Pear')")
    ctl(g, "OrderForm", "CheckBox1").click()
    page.wait_for_function("() => document.querySelector('.rg-userform [data-control=\"CheckBox1\"]').getAttribute('aria-checked') === 'true'")
    # KeyPress: the code drops what isn't a digit.
    box = ctl(g, "OrderForm", "TextBox2")
    box.click()
    page.keyboard.type("a1b2", delay=60)
    wait_value(box, "12")
    # A MsgBox over the form; the form stays.
    ctl(g, "OrderForm", "CommandButton3").click()
    page.locator("#dlg-macro-prompt").wait_for(state="visible")
    assert page.locator("#mb-prompt").inner_text().strip() == "Hello Ann"
    g.shot("msgbox_over_form")
    page.click("#mb-ok")
    page.locator("#dlg-macro-prompt").wait_for(state="detached")
    assert form(g, "OrderForm").is_visible()
    assert ctl(g, "OrderForm", "TextBox1").input_value() == "Ann"
    # OK: CommandButton1_Click writes the cells and unloads the form; the macro goes on.
    ctl(g, "OrderForm", "CommandButton1").click()
    form(g, "OrderForm").wait_for(state="detached")
    assert page.locator(".rg-userform-backdrop").count() == 0
    g.wait_text("A1", "Ann")
    g.wait_text("A2", "Pear")
    g.wait_text("A3", "TRUE")
    g.wait_text("A4", "12")
    g.wait_text("B1", "after")

    # Escape is the Cancel button.
    run_macro(g, "ShowOrder")
    form(g, "OrderForm").wait_for(state="visible")
    ctl(g, "OrderForm", "TextBox1").click()
    page.keyboard.press("Escape")
    form(g, "OrderForm").wait_for(state="detached")

    # A modeless form stays open beside the grid after its macro.
    run_macro(g, "ShowNotes")
    form(g, "NoteForm").wait_for(state="visible")
    assert page.locator(".rg-userform-backdrop").count() == 0
    g.type_in("E1", "grid still works")
    g.wait_text("E1", "grid still works")
    for note in ("first", "second"):
        ctl(g, "NoteForm", "TextBox1").fill(note)
        ctl(g, "NoteForm", "CommandButton1").click()
        page.wait_for_function("n => document.querySelector('.rg-userform [data-control=\"Label1\"]').innerText.trim() === n", arg=f"{'1' if note == 'first' else '2'} notes")
    g.wait_text("C1", "first")
    g.wait_text("C2", "second")
    assert ctl(g, "NoteForm", "TextBox1").input_value() == ""
    g.shot("modeless")
    page.locator('[data-form-close="NoteForm"]').click()
    form(g, "NoteForm").wait_for(state="detached")

    # The VBA editor: a form opens on its design; double-clicking a control opens its code.
    page.click("#tab-developer")
    page.click("#rb-vba")
    page.locator("#dlg-vba").wait_for(state="visible")
    page.click("#vba-mod-OrderForm")
    page.locator('#vba-form-pane .rg-form-designer[data-form="OrderForm"]').wait_for(state="visible")
    g.shot("form_view")
    page.locator('#vba-form-pane [data-control="Label1"]').dblclick()
    page.locator("#vba-code").wait_for(state="visible")
    page.wait_for_function("() => document.getElementById('vba-code').value.includes('Private Sub Label1_Click()')")
    # Run Form: the editor makes way for the form and comes back when it closes.
    page.click("#vba-run-form")
    form(g, "OrderForm").wait_for(state="visible")
    page.locator("#dlg-vba").wait_for(state="detached")
    page.locator('[data-form-close="OrderForm"]').click()
    form(g, "OrderForm").wait_for(state="detached")
    page.locator("#dlg-vba").wait_for(state="visible")
    page.click("#vba-close")
    page.locator("#dlg-vba").wait_for(state="detached")

    design_a_form(g)

    # Download: the forms are in the .xlsm, with the code as edited.
    with page.expect_download(timeout=20000) as dl:
        page.click("#qa-save")
    out = g.out / "forms_saved.xlsm"
    dl.value.save_as(out)
    s = Session()
    d = open_bytes(s, out.read_bytes(), "forms_saved.xlsm")
    assert d.warnings == [], d.warnings
    m = s.doc().wb.vba.module("OrderForm")
    assert m.kind == "form" and "Private Sub Label1_Click()" in m.code
    back = read_design("OrderForm", design_streams(m.designer))
    assert [c.name for c in back.controls] == [c.name for c in ORDER.controls] and back.warnings == ()
    assert s.doc().wb.vba.module("UserForm1") is None
    made = s.doc().wb.vba.module("frmDesigned")
    designed = read_design("frmDesigned", design_streams(made.designer))
    assert [(c.name, c.caption) for c in designed.controls] == [("TextBox1", ""), ("CommandButton1", "Save")] and designed.controls[0].left > 60


if __name__ == "__main__":
    sys.exit(run("vba_userforms", body))
