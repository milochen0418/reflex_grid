"""AI in the VBA editor, against a stand-in AI Hub (testcases/_lib/ai_stub.py): what the person
asks shows as it is written and then as a diff, Accept puts it in the module in one edit that the
code's Ctrl+Z takes back and Reject leaves the code alone; the request goes on from a change on
screen; selected lines are all the AI writes; a question gets an answer in words; an answer that
doesn't compile is asked again with the compiler's words; Fix with AI answers compile and run-time
errors; grey suggestions take Tab and Esc and can be turned off; and what is said into the
microphone (the browser's speech, faked) becomes the request, a pause sends it, and "accept" and
"reject" answer a change.

Needs the server started with ``DDNS_AI_HUB_URL`` (where this suite runs the stand-in, e.g.
http://127.0.0.1:8399) and ``DDNS_AUTH_SECRET``, as the CI workflow does for this suite; without
them it is skipped."""

import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _lib import ai_stub  # noqa: E402
from _lib.grid import MOD, Grid, run  # noqa: E402

CODE = 'Sub Hello()\n    Range("A1").Value = 42\nEnd Sub\n'

# The microphone, the browser's speech and the AI Hub's ears, faked: window.__said(words, final) is
# what the browser hears, window.__hub(message) what the hub's socket says, window.__pcm what it got.
FAKE_EARS = """
class FakeRec {
  start() { window.__rec = this; }
  stop() { window.__rec = null; }
}
window.SpeechRecognition = FakeRec;
window.webkitSpeechRecognition = FakeRec;
window.__said = (words, final) => {
  const result = [{ transcript: words }];
  result.isFinal = final;
  window.__rec.onresult({ resultIndex: 0, results: [result] });
};
navigator.mediaDevices.getUserMedia = async () => new AudioContext().createMediaStreamDestination().stream;
class FakeHub {
  constructor(url) {
    window.__hubUrl = url;
    window.__hubSocket = this;
    window.__pcm = 0;
    this.readyState = 1;
    setTimeout(() => { this.onopen && this.onopen({}); window.__hub({ type: "status", phase: "loading" }); }, 0);
  }
  send(data) { window.__pcm += data.byteLength; }
  close() { this.readyState = 3; }
}
window.__hub = (m) => window.__hubSocket.onmessage({ data: JSON.stringify(m) });
window.WebSocket = new Proxy(window.WebSocket, {
  construct: (Real, args) => (String(args[0]).includes("/api/talk/stt") ? new FakeHub(args[0]) : new Real(...args)),
});
"""


def vba_asks() -> list:
    return [a for a in ai_stub.Stub.asks if a["system"].startswith("You change one module of Excel VBA")]


def code_of(g: Grid) -> str:
    return g.page.locator("#vba-code").input_value()


def set_code(g: Grid, code: str) -> None:
    area = g.page.locator("#vba-code")
    area.click()
    area.fill(code)
    area.press(f"{MOD}+s")
    g.page.wait_for_function("() => !document.querySelector('#vba-dirty')")
    g.wait_idle(0.4)


def ask(g: Grid, words: str) -> None:
    box = g.page.locator("#vba-ai-prompt")
    box.fill(words)
    box.press("Enter")


def no_macros_yet(g: Grid) -> None:
    """A workbook without macros: what the AI writes becomes its first module."""
    page = g.page
    page.click("#rb-vba")
    page.locator("#dlg-vba").wait_for(state="visible")
    page.locator("#vba-ai-bar").wait_for()
    ask(g, "add a sub")
    page.locator("#vba-ai-diff").wait_for()
    page.click("#vba-ai-accept")
    page.locator("#vba-mod-Module1").wait_for()
    page.wait_for_function("() => document.getElementById('vba-code').value.includes('Sub Added()')")


def writes_and_accepts(g: Grid) -> None:
    page = g.page
    page.locator("#vba-code").click()
    page.keyboard.press(f"{MOD}+i")
    assert page.evaluate("document.activeElement.id") == "vba-ai-prompt"
    page.keyboard.type("add a sub")
    ai_stub.Stub.vba_line_delay = 0.3
    page.keyboard.press("Enter")
    # It shows as it is written, over the code, with Stop.
    page.locator("#vba-ai-writing").wait_for()
    page.locator("#vba-ai-stop").wait_for()
    page.wait_for_function("() => document.getElementById('vba-ai-writing') && document.getElementById('vba-ai-writing').textContent.includes('Sub Added')")
    page.locator("#vba-ai-diff").wait_for()
    ai_stub.Stub.vba_line_delay = 0.08
    g.shot("change_to_accept")
    assert page.locator("#vba-ai-diff .vba-diff-add").count() == 4, page.locator("#vba-ai-diff").inner_text()
    assert page.locator("#vba-ai-diff .vba-diff-del").count() == 0
    assert "compiles" in page.locator("#vba-ai-status").inner_text()
    assert code_of(g) == CODE.rstrip("\n") or code_of(g) == CODE, "nothing changes before Accept"
    asked = vba_asks()[-1]
    assert asked["role"] == "coder" and asked["auth"], asked
    assert '<module name="Module1" kind="module">' in asked["prompt"] and "Sub Hello()" in asked["prompt"]
    assert "The person's cursor is on line 4." in asked["prompt"]
    assert 'Sheet "Sheet1"' in asked["prompt"] and "<request>\nadd a sub\n</request>" in asked["prompt"]
    # Accept: in the module at once, as an edit of the code.
    page.click("#vba-ai-accept")
    page.locator("#vba-ai-pane").wait_for(state="detached")
    page.wait_for_function("() => document.getElementById('vba-code').value.includes('Sub Added()')")
    page.wait_for_function("() => !document.querySelector('#vba-dirty')")
    assert "Saved" in page.locator("#vba-status").inner_text()
    area = page.locator("#vba-code")
    area.press(f"{MOD}+z")
    page.wait_for_function("() => !document.getElementById('vba-code').value.includes('Sub Added()')")
    page.locator("#vba-dirty").wait_for()
    area.press(f"{MOD}+s")
    page.wait_for_function("() => !document.querySelector('#vba-dirty')")


def rejects_stops_and_goes_on(g: Grid) -> None:
    page = g.page
    before = code_of(g)
    ask(g, "add a sub")
    page.locator("#vba-ai-diff").wait_for()
    page.keyboard.press("Escape")
    page.locator("#vba-ai-pane").wait_for(state="detached")
    assert code_of(g) == before
    # A second request goes on from the change on screen; the diff stays against the code.
    ask(g, "add a sub")
    page.locator("#vba-ai-diff").wait_for()
    ask(g, "rename it")
    page.wait_for_function("() => document.getElementById('vba-ai-diff') && document.getElementById('vba-ai-diff').textContent.includes('Sub Renamed')")
    assert "That" not in vba_asks()[-1]["prompt"] and vba_asks()[-1]["prompt"].strip() == "rename it"
    page.click("#vba-ai-reject")
    page.locator("#vba-ai-pane").wait_for(state="detached")
    # Stop while it writes: nothing is left of it.
    ai_stub.Stub.vba_line_delay = 0.5
    ask(g, "add a sub")
    page.locator("#vba-ai-stop").click()
    page.locator("#vba-ai-pane").wait_for(state="detached")
    time.sleep(0.6)
    assert page.locator("#vba-ai-pane").count() == 0 and code_of(g) == before
    ai_stub.Stub.vba_line_delay = 0.08
    # Another module closes a change that was for this one.
    ask(g, "add a sub")
    page.locator("#vba-ai-diff").wait_for()
    page.click("#vba-mod-Sheet1")
    page.locator("#vba-ai-pane").wait_for(state="detached")
    page.click("#vba-mod-Module1")
    page.wait_for_function("() => document.getElementById('vba-code').value.includes('Sub Hello')")
    assert "Sub Added" not in code_of(g)


def selected_lines_and_questions(g: Grid) -> None:
    page = g.page
    set_code(g, 'Sub A()\n    x = 1\nEnd Sub\n\nSub B()\n    y = 2\nEnd Sub\n')
    # The second Sub is selected: only it is written.
    page.evaluate("() => { const a = document.getElementById('vba-code'); a.focus(); a.setSelectionRange(a.value.indexOf('Sub B'), a.value.length); }")
    page.keyboard.press(f"{MOD}+i")
    assert page.locator("#vba-ai-lines").inner_text() == "Lines 5–7"
    page.keyboard.type("change these")
    page.keyboard.press("Enter")
    page.locator("#vba-ai-diff").wait_for()
    assert page.locator("#vba-ai-lines").inner_text() == "Lines 5–7"
    asked = vba_asks()[-1]["prompt"]
    assert '<target lines="5-7">' in asked and "<module_before" in asked and "Sub A()" in asked.split("<target")[0]
    page.click("#vba-ai-accept")
    page.wait_for_function("() => document.getElementById('vba-code').value.includes('SUB B()')")
    assert code_of(g).startswith("Sub A()"), code_of(g)
    page.wait_for_function("() => !document.querySelector('#vba-dirty')")
    # A question gets words, and the code stays.
    set_code(g, CODE)
    ask(g, "what does this do?")
    page.locator("#vba-ai-answer").wait_for()
    assert "42 into cell A1" in page.locator("#vba-ai-answer").inner_text()
    assert page.locator("#vba-ai-pane").count() == 0
    g.shot("answer_in_words")
    page.click("#vba-ai-close")
    page.locator("#vba-ai-answer").wait_for(state="detached")
    # An answer that changes nothing says so.
    ask(g, "leave it as it is")
    page.locator("#vba-ai-answer").wait_for()
    assert "changes nothing" in page.locator("#vba-ai-answer").inner_text()
    page.keyboard.press("Escape")
    page.locator("#vba-ai-answer").wait_for(state="detached")


def compiles_before_it_shows(g: Grid) -> None:
    page = g.page
    n = len(vba_asks())
    ask(g, "broken first")
    page.locator("#vba-ai-diff").wait_for()
    assert len(vba_asks()) == n + 2, "an answer that doesn't compile is asked again"
    assert vba_asks()[-1]["prompt"].startswith("That doesn't compile. Line 6 of the module: Expected: type name (Dim x As)"), vba_asks()[-1]["prompt"]
    assert "compiles" in page.locator("#vba-ai-status").inner_text()
    page.click("#vba-ai-reject")
    page.locator("#vba-ai-pane").wait_for(state="detached")


def fixes_errors(g: Grid) -> None:
    page = g.page
    # A compile error: the button is beside it.
    area = page.locator("#vba-code")
    area.click()
    area.fill('Sub Broken()\n    Range("A1").Value = 1\n')
    page.click("#vba-compile")
    page.locator("#vba-ai-fix").wait_for()
    assert "Compile error" in page.locator("#vba-status").inner_text()
    g.shot("compile_error")
    page.click("#vba-ai-fix")
    page.locator("#vba-ai-diff").wait_for()
    asked = vba_asks()[-1]["prompt"]
    assert "Fix this error: Compile error: Expected: End Sub" in asked, asked
    g.shot("fix_diff")
    page.click("#vba-ai-accept")
    page.wait_for_function("() => document.getElementById('vba-code').value.includes('End Sub')")
    page.locator("#vba-ai-fix").wait_for(state="detached")
    page.click("#vba-compile")
    page.wait_for_function("() => document.getElementById('vba-status') && document.getElementById('vba-status').textContent.includes('No compile errors')")
    # The Problems tab's list has it too.
    area.fill('Sub Broken()\n    Range("A1").Value = 1\n')
    page.click("#vba-compile")
    page.locator("#vba-ai-fix").wait_for()
    page.get_by_role("tab", name="Problems").click()
    page.locator("#vba-ai-fix-problem").click()
    page.locator("#vba-ai-diff").wait_for()
    page.click("#vba-ai-accept")
    page.locator("#vba-ai-fix").wait_for(state="detached")
    page.get_by_role("tab", name="Immediate").click()
    # A run-time error: the macro stopped at a line, the same button.
    set_code(g, 'Sub Boom()\n    Worksheets("Nope").Range("A1").Value = 1\nEnd Sub\n')
    page.evaluate("() => { const a = document.getElementById('vba-code'); a.focus(); a.setSelectionRange(5, 5); }")
    page.keyboard.press("F5")
    page.locator("#vba-ai-fix").wait_for()
    assert "Subscript out of range" in page.locator("#vba-status").inner_text() or "error" in page.locator("#vba-status").inner_text().lower()
    page.click("#vba-ai-fix")
    page.locator("#vba-ai-diff").wait_for()
    assert "Line 2 of the module: Worksheets(\"Nope\")" in vba_asks()[-1]["prompt"], vba_asks()[-1]["prompt"]
    page.click("#vba-ai-accept")
    page.wait_for_function("() => document.getElementById('vba-code').value.includes('Worksheets(\"Sheet1\")')")
    page.wait_for_function("() => !document.querySelector('#vba-dirty')")
    page.locator("#vba-ai-fix").wait_for(state="detached")
    page.evaluate("() => { const a = document.getElementById('vba-code'); a.focus(); a.setSelectionRange(5, 5); }")
    page.keyboard.press("F5")
    page.wait_for_function("() => document.getElementById('vba-status') && document.getElementById('vba-status').textContent.includes('done')")


def grey_suggestions(g: Grid) -> None:
    page = g.page
    set_code(g, 'Sub Hello()\n    Range("A1").\nEnd Sub\n')
    at = code_of(g).index('.\n') + 1
    park = "at => { const a = document.getElementById('vba-code'); a.focus(); a.setSelectionRange(at, at); }"
    page.evaluate(park, at)
    # Tab takes it.
    page.keyboard.type(" ")
    page.keyboard.press("Backspace")
    page.locator(".vba-ghost").wait_for()
    assert page.locator(".vba-ghost").inner_text() == "Value = 1"
    g.shot("grey_suggestion")
    page.keyboard.press("Tab")
    page.wait_for_function("() => document.getElementById('vba-code').value.includes('.Value = 1')")
    assert page.locator(".vba-ghost").count() == 0
    # Esc dismisses it, the code stays.
    set_code(g, 'Sub Hello()\n    Range("A1").\nEnd Sub\n')
    page.evaluate(park, at)
    page.keyboard.type(" ")
    page.keyboard.press("Backspace")
    page.locator(".vba-ghost").wait_for()
    page.keyboard.press("Escape")
    page.locator(".vba-ghost").wait_for(state="detached")
    assert "Value = 1" not in code_of(g)
    # Off: nothing is asked.
    asked = len(ai_stub.Stub.asks)
    page.click("#vba-ai-suggest")
    page.wait_for_function("() => document.getElementById('vba-ai-suggest').dataset.on === '0'")
    page.evaluate(park, at)
    page.keyboard.type(" ")
    page.keyboard.press("Backspace")
    time.sleep(1.0)
    assert page.locator(".vba-ghost").count() == 0 and len(ai_stub.Stub.asks) == asked
    page.click("#vba-ai-suggest")
    page.wait_for_function("() => document.getElementById('vba-ai-suggest').dataset.on === '1'")
    area = page.locator("#vba-code")
    area.fill(CODE)
    area.press(f"{MOD}+s")


def speaking(g: Grid) -> None:
    page = g.page
    page.locator("#vba-code").click()
    page.keyboard.press("Alt+v")
    page.wait_for_function("() => document.getElementById('vba-ai-mic').dataset.on === '1'")
    page.wait_for_function("() => window.__rec")
    # Words appear as they are heard; the pause after them sends them.
    n = len(vba_asks())
    page.evaluate("() => window.__said('add', false)")
    page.wait_for_function("() => document.getElementById('vba-ai-prompt').value === 'add'")
    page.evaluate("() => window.__said('add a', true)")
    page.evaluate("() => window.__said('sub', true)")
    page.wait_for_function("() => document.getElementById('vba-ai-prompt').value === 'add a sub'")
    page.locator("#vba-ai-countdown").wait_for()
    g.shot("heard")
    page.locator("#vba-ai-diff").wait_for(timeout=8000)
    assert len(vba_asks()) == n + 1 and "<request>\nadd a sub\n</request>" in vba_asks()[-1]["prompt"]
    assert page.locator("#vba-ai-prompt").input_value() == ""
    # Words alone answer it.
    page.evaluate("() => window.__said('接受。', true)")
    page.wait_for_function("() => document.getElementById('vba-code').value.includes('Sub Added()')")
    page.locator("#vba-ai-pane").wait_for(state="detached")
    page.wait_for_function("() => !document.querySelector('#vba-dirty')")
    page.evaluate("() => window.__said('broken first', true)")
    page.locator("#vba-ai-diff").wait_for(timeout=8000)
    page.evaluate("() => window.__said('不要', true)")
    page.locator("#vba-ai-pane").wait_for(state="detached")
    # A word of those in a longer sentence is a request, and on its own with nothing to answer, nothing.
    page.evaluate("() => window.__said('好', true)")
    time.sleep(0.4)
    assert page.locator("#vba-ai-prompt").input_value() == ""
    page.evaluate("() => window.__said('好的 add a sub', true)")
    page.locator("#vba-ai-countdown").wait_for()
    # Typing takes over from the countdown.
    page.locator("#vba-ai-prompt").press("End")
    page.keyboard.type("!")
    page.locator("#vba-ai-countdown").wait_for(state="detached")
    page.locator("#vba-ai-prompt").fill("")
    page.keyboard.press("Alt+v")
    page.wait_for_function("() => document.getElementById('vba-ai-mic').dataset.on === '0'")


def the_hubs_ears(g: Grid) -> None:
    """Once the hub's ears are ready they take over: the microphone goes to them as sound, their
    words come back; closing the editor stops all that and leaves no change behind."""
    page = g.page
    page.locator("#vba-code").click()
    page.keyboard.press("Alt+v")
    page.wait_for_function("() => window.__rec")
    url = page.evaluate("() => decodeURIComponent(window.__hubUrl)")
    assert "/api/talk/stt" in url and "hint=" in url and "Sheet1" in url and "Module1" in url, url
    page.locator("#vba-ai-note").wait_for()
    page.evaluate("() => window.__hub({ type: 'status', phase: 'ready' })")
    page.wait_for_function("() => !window.__rec")
    page.locator("#vba-ai-note").wait_for(state="detached")
    page.wait_for_function("() => window.__pcm > 3200")
    page.evaluate("() => window.__hub({ type: 'subtitle', seg: 0, text: 'broken', final: false })")
    page.wait_for_function("() => document.getElementById('vba-ai-prompt').value === 'broken'")
    page.evaluate("() => window.__hub({ type: 'subtitle', seg: 0, text: 'broken first', final: true })")
    page.locator("#vba-ai-diff").wait_for(timeout=8000)
    page.click("#vba-close")
    page.locator("#dlg-vba").wait_for(state="detached")
    assert page.evaluate("() => window.__hubSocket.readyState") == 3, "closing the editor stops listening"
    page.click("#rb-vba")
    page.locator("#vba-code").wait_for()
    page.locator("#vba-ai-bar").wait_for()
    assert page.locator("#vba-ai-pane").count() == 0


def body(g: Grid) -> None:
    page = g.page
    page.add_init_script(FAKE_EARS)
    page.reload()
    g.wait_idle(1.0)
    page.click("#tab-developer")
    no_macros_yet(g)
    set_code(g, CODE)
    writes_and_accepts(g)
    rejects_stops_and_goes_on(g)
    selected_lines_and_questions(g)
    compiles_before_it_shows(g)
    fixes_errors(g)
    grey_suggestions(g)
    speaking(g)
    the_hubs_ears(g)


if __name__ == "__main__":
    hub = os.getenv("DDNS_AI_HUB_URL", "")
    if not hub or not os.getenv("DDNS_AUTH_SECRET"):
        print("[vba_ai] SKIP: start the server with DDNS_AI_HUB_URL (where this suite runs its stand-in hub) and DDNS_AUTH_SECRET")
        sys.exit(0)
    ai_stub.start(hub)
    sys.exit(run("vba_ai", body))
