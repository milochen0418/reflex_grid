"""Excel macros in the browser: write a module in the VBA editor, run a macro with F5 and answer
its MsgBox and InputBox, use a VBA function in a formula, undo the whole run, run a macro from
the Macros dialog (Alt+F8), turn on event macros from the security bar, stop a long macro, and
download the workbook as .xlsm and open it again (openpyxl checks the file)."""

import io
import sys
import time
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl  # noqa: E402

from _lib.grid import MOD, Grid, run  # noqa: E402

CODE = '''Option Explicit

Sub Hello()
    Range("A1").Value = 21 * 2
    If MsgBox("Write a greeting?", vbYesNo + vbQuestion, "Greeting") = vbYes Then
        Range("A2").Value = "Hi " & InputBox("Your name?", "Greeting", "Bob")
    End If
    Range("A3").Formula = "=A1*2"
    MsgBox "Done"
End Sub

Sub Fill()
    Dim i As Long
    For i = 1 To 5
        Cells(i, 5).Value = i * i
    Next
End Sub

Sub Slow()
    Dim t As Double
    t = Timer
    Do While Timer - t < 20
    Loop
End Sub

Function Twice(x)
    Twice = x * 2
End Function
'''

EVENTS = '''Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 7 Then
        Application.EnableEvents = False
        Target.Offset(0, 1).Value = UCase(Target.Value)
        Application.EnableEvents = True
    End If
End Sub
'''


def wait_gone(page, selector: str, timeout: float = 15.0) -> None:
    page.locator(selector).wait_for(state="detached", timeout=timeout * 1000)
    # A dialog's overlay fades out after its content: a click meanwhile would land on it.
    page.wait_for_function(
        "() => [...document.querySelectorAll('.rt-BaseDialogOverlay')].every((o) => o.querySelector('.rt-BaseDialogContent'))",
        timeout=timeout * 1000,
    )


def wait_empty(g: Grid, a1: str, timeout: float = 8.0) -> None:
    end = time.time() + timeout
    while time.time() < end:
        if (g.text(a1) or "") == "":
            return
        time.sleep(0.1)
    raise AssertionError(f"{a1}: expected empty, got {g.text(a1)!r}")


def set_code(g: Grid, code: str) -> None:
    """Replaces the editor's text and applies it (Ctrl+S)."""
    area = g.page.locator("#vba-code")
    area.click()
    area.fill(code)
    area.press(f"{MOD}+s")
    g.page.wait_for_function("() => !document.querySelector('#vba-dirty')")
    g.wait_idle(0.4)


def caret_in(g: Grid, text: str) -> None:
    g.page.evaluate(
        "t => { const a = document.getElementById('vba-code'); const i = a.value.indexOf(t); a.focus(); a.setSelectionRange(i + 4, i + 4); }",
        text,
    )


def body(g: Grid) -> None:
    page = g.page
    # A new module from the Developer tab opens the editor on it.
    page.click("#tab-developer")
    page.click("#rb-add-module")
    page.locator("#dlg-vba").wait_for(state="visible")
    page.locator("#vba-mod-Module1").wait_for(state="visible")
    for name in ("ThisWorkbook", "Sheet1"):
        page.locator(f"#vba-mod-{name}").wait_for(state="visible")
    set_code(g, CODE)
    assert "error" not in (page.locator("#vba-status").inner_text() if page.locator("#vba-status").count() else ""), page.locator("#vba-status").inner_text()
    g.shot("editor")

    # F5 runs the Sub the cursor is in; its dialogs show and wait for an answer.
    caret_in(g, "Sub Hello")
    page.keyboard.press("F5")
    page.locator("#dlg-macro-prompt").wait_for(state="visible")
    assert "Write a greeting?" in page.locator("#mb-prompt").inner_text()
    g.shot("msgbox")
    page.click("#mb-yes")
    page.locator("#ib-input").wait_for(state="visible")
    assert page.locator("#ib-input").input_value() == "Bob"
    page.locator("#ib-input").fill("Ann")
    page.click("#ib-ok")
    page.locator("#mb-ok").wait_for(state="visible")
    assert page.locator("#mb-prompt").inner_text().strip() == "Done"
    page.click("#mb-ok")
    wait_gone(page, "#dlg-macro-prompt")
    page.click("#vba-close")
    wait_gone(page, "#dlg-vba")
    g.wait_text("A1", "42")
    g.wait_text("A2", "Hi Ann")
    g.wait_text("A3", "84")

    # A VBA Function in a formula.
    g.type_in("C1", "=Twice(A1)")
    g.wait_text("C1", "84")

    # One undo takes back the formula, the next the whole macro run.
    g.key("Mod+z")
    wait_empty(g, "C1")
    g.key("Mod+z")
    wait_empty(g, "A1")
    wait_empty(g, "A2")

    # Alt+F8: the Macros dialog.
    g.click("D10")
    page.keyboard.press("Alt+F8")
    page.locator("#dlg-macros").wait_for(state="visible")
    page.locator(".dlg-macro", has_text="Fill").click()
    page.click("#dlg-macro-run")
    wait_gone(page, "#dlg-macros")
    g.wait_text("E5", "25")

    # Event macros: off until someone turns them on (the security bar).
    page.click("#tab-developer")
    page.click("#rb-vba")
    page.locator("#dlg-vba").wait_for(state="visible")
    page.click("#vba-mod-Sheet1")
    g.wait_idle(0.5)
    set_code(g, EVENTS)
    page.click("#vba-close")
    wait_gone(page, "#dlg-vba")
    page.locator("#macro-bar").wait_for(state="visible")
    g.type_in("G1", "quiet")
    g.wait_idle(0.5)
    assert (g.text("H1") or "") == "", g.text("H1")
    page.click("#macro-enable")
    wait_gone(page, "#macro-bar")
    g.type_in("G2", "loud")
    g.wait_text("H2", "LOUD")
    g.shot("events")

    # A long macro runs in steps; Stop ends it.
    page.keyboard.press("Alt+F8")
    page.locator("#dlg-macros").wait_for(state="visible")
    page.locator(".dlg-macro", has_text="Slow").click()
    page.click("#dlg-macro-run")
    page.locator("#dlg-macro-running").wait_for(state="visible")
    g.shot("running")
    t0 = time.time()
    page.click("#macro-stop")
    wait_gone(page, "#dlg-macro-running")
    assert time.time() - t0 < 10, "Stop took too long"
    # The workbook still answers.
    g.type_in("A10", "after")
    g.wait_text("A10", "after")

    # Download: an .xlsm with the project, readable by openpyxl.
    with page.expect_download(timeout=20000) as dl:
        page.click("#qa-save")
    assert dl.value.suggested_filename.endswith(".xlsm"), dl.value.suggested_filename
    out = g.out / "macros.xlsm"
    dl.value.save_as(out)
    with zipfile.ZipFile(out) as z:
        assert "xl/vbaProject.bin" in z.namelist()
        ct = z.read("[Content_Types].xml").decode("utf-8")
        assert "application/vnd.ms-excel.sheet.macroEnabled.main+xml" in ct
    wb = openpyxl.load_workbook(io.BytesIO(out.read_bytes()), keep_vba=True)
    try:
        assert wb["Sheet1"]["E5"].value == 25
    finally:
        wb.vba_archive.close()

    # Opening the .xlsm again: the macros are there, and off until enabled.
    page.click("#tab-file")
    page.click("#bs-open")
    page.locator("#dlg-open").wait_for(state="visible")
    page.locator('#dlg-open input[type="file"]').set_input_files(str(out))
    g.wait_title("macros")
    g.wait_idle(1.0)
    page.locator("#macro-bar").wait_for(state="visible")
    page.keyboard.press("Alt+F8")
    page.locator("#dlg-macros").wait_for(state="visible")
    names = [x.strip() for x in page.locator(".dlg-macro").all_inner_texts()]
    assert names == ["Hello", "Fill", "Slow"], names
    page.click("#dlg-cancel")
    wait_gone(page, "#dlg-macros")
    g.click("B12")
    g.type_in("B12", "=Twice(4)")
    g.wait_text("B12", "8")
    g.shot("reopened")


if __name__ == "__main__":
    sys.exit(run("vba_macros", body))
