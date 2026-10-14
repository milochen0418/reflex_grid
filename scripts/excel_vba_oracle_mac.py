"""Compares ReflexGrid's VBA interpreter with the VBA of the Microsoft Excel installed on this Mac
(opt-in: it opens Excel windows, like ``excel_check_mac.py``).

Each case is a VBA expression (or a small Function) evaluated by both: Excel runs the code from a
macro-enabled workbook that ReflexGrid wrote (so this also checks that Excel accepts ReflexGrid's
vbaProject.bin), and ReflexGrid runs the same module. Results are compared as
``TypeName|text`` (or ``E<number>`` for a run-time error).

    poetry run python scripts/excel_vba_oracle_mac.py [--keep DIR] [--only SUBSTRING]

Excel asks whether to enable the macros of the generated workbook; the script answers through
System Events (enable the terminal under Privacy & Security › Accessibility if it can't).
"""

from __future__ import annotations

import argparse
import base64
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import List, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from gridcore.engine import Session  # noqa: E402
from gridcore.model.vba import STANDARD, VbaModule, default_attributes, new_project  # noqa: E402

from tests.vba_excel_cases import CLASS1, EXPRESSIONS, FUNCTIONS, HELPERS, WINDOWS_ONLY, build_module, mac_cases, repr_vba  # noqa: E402,F401


ENABLE = """tell application "System Events"
tell process "Microsoft Excel"
repeat with w in windows
if exists button "Enable Macros" of w then
click button "Enable Macros" of w
return "clicked"
end if
end repeat
end tell
end tell
return "none"
"""


def osa(script: str, timeout: float = 60) -> str:
    try:
        p = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=timeout)
        return (p.stdout + p.stderr).strip()
    except subprocess.TimeoutExpired:
        return "TIMEOUT"


DIALOGS = """tell application "System Events"
tell process "Microsoft Excel"
repeat with w in windows
try
set t to (value of every static text of w) as text
if t contains "error" or t contains "Error" then
set b to name of every button of w
return t & " ## " & (b as text)
end if
end try
end repeat
end tell
end tell
return \"\""""

DISMISS = """tell application "System Events"
tell process "Microsoft Excel"
repeat with w in windows
repeat with nm in {"End", "OK"}
if exists button nm of w then
click button nm of w
return "ok"
end if
end repeat
end repeat
end tell
end tell
return \"\""""

RESET = """tell application "System Events"
tell process "Microsoft Excel"
repeat with w in windows
if name of w starts with "Microsoft Visual Basic" then
try
click button "Reset" of toolbar 4 of w
end try
try
click (first button of w whose subrole is "AXCloseButton")
end try
end if
end repeat
end tell
end tell"""


def run_watching(book: str, macro: str, limit: float = 600) -> str:
    """Runs a macro without blocking on VBA's dialogs: returns "" when it finished, else the text
    of the compile or run-time error dialog it raised (dismissed)."""
    proc = subprocess.Popen(["osascript", "-e", f'with timeout of {int(limit)} seconds\ntell application "Microsoft Excel" to run VB macro "{book}!Module1.{macro}"\nend timeout'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    start = time.time()
    while time.time() - start < limit:
        if proc.poll() is not None:
            return ""
        d = osa(DIALOGS, 10)
        if d and d != "TIMEOUT":
            osa(DISMISS, 10)
            time.sleep(0.5)
            osa(RESET, 10)
            proc.kill()
            return d.replace("\n", " ")
        time.sleep(0.5)
    proc.kill()
    return "TIMEOUT"


def excel_results(path: Path, n: int, macro: str = "WriteAll", columns: int = 1) -> List[str]:
    """Opens the workbook in Excel, runs ``macro`` and reads its results (the first ``columns``
    columns of rows 1..n of the first sheet, row by row)."""
    name = path.name
    osa(RESET, 10)
    subprocess.run(["open", "-a", "Microsoft Excel", str(path)])
    for _ in range(60):
        if osa(ENABLE, 10) == "clicked":
            break
        time.sleep(1)
    time.sleep(2)
    problem = run_watching(name, macro)
    if problem:
        where = osa(f'tell application "Microsoft Excel" to return string value of range "B1" of worksheet 1 of workbook "{name}"', 10)
        print(f"Excel stopped in {where}: {problem}")
        osa(f'tell application "Microsoft Excel" to close workbook "{name}" saving no', 30)
        return []
    last = "AB"[columns - 1]
    script = f'''tell application "Microsoft Excel"
set vals to value of range "A1:{last}{n}" of worksheet 1 of workbook "{name}"
end tell
set AppleScript's text item delimiters to (ASCII character 30)
set out to {{}}
repeat with r in vals
repeat with v in r
set end of out to (v as text)
end repeat
end repeat
return out as text'''
    res = osa(script, 120)
    osa(f'tell application "Microsoft Excel" to close workbook "{name}" saving no', 30)
    return res.split(chr(30))


# Excel for Mac answers these differently from Windows (ReflexGrid follows Windows): TypeName of
# an array is "V()" there instead of "Variant()".
MAC_QUIRKS = ("TypeName(Array(", "TypeName(Split(")


def _close(a: str, b: str) -> bool:
    """Doubles that differ only in the 15th significant digit (operation order)."""
    if not (a.startswith("Double|") and b.startswith("Double|")):
        return False
    try:
        x, y = float(a[7:]), float(b[7:])
    except ValueError:
        return False
    return abs(x - y) <= 2e-14 * max(1.0, abs(x))


def _us_dates(text: str) -> str:
    """This Mac may show dates as yyyy/m/d (its region); ReflexGrid uses the US m/d/yyyy."""
    import re

    return re.sub(r"\b(\d{4})/(\d{1,2})/(\d{1,2})\b", lambda m: f"{int(m.group(2))}/{int(m.group(3))}/{m.group(1)}", text)


def _mac_objects(text: str) -> str:
    """Excel for Mac's answers as Excel for Windows (US) gives them: dates in US order, arrays
    named Variant() (not V()), the built-in short date formats with four-digit years."""
    import re

    t = _us_dates(text).rstrip()
    t = re.sub(r"(?<![A-Za-z])V\(\)", "Variant()", t)
    t = re.sub(r"(?<![A-Za-z/])m/d/yy(?![y])", "m/d/yyyy", t)
    return t


def ours(code: str) -> List[str]:
    from gridcore.vba.compiler import compile_project
    from gridcore.vba.runtime import Execution

    # WriteAll needs Excel's object model; RunAll is the part both run.
    code = code.split("Sub WriteAll()")[0]
    rt = compile_project([("Module1", "standard", code, ()), ("Class1", "class", CLASS1, ())])
    if rt.errors:
        for e in rt.errors.values():
            print("compile error:", e)
        return []
    ex = Execution(rt, time_limit=120)
    r = rt.run(ex, "RunAll")
    return [str(x) for x in r.data]


# ---------------------------------------------------------------- object model cases


def object_workbook():
    """A session whose workbook holds the object-model cases (tests/vba_object_cases.py)."""
    from tests.vba_object_cases import build_module as build_objects

    code, names, this_workbook = build_objects()
    s = Session()
    s.new_workbook()
    s.run("macro.setModule", name="Module1", code=code)
    s.run("macro.setModule", name="ThisWorkbook", code=this_workbook)
    return s, names


def ours_objects() -> List[Tuple[str, str]]:
    s, names = object_workbook()
    r = s.run("macro.run", name="RunAll", timeLimit=600)
    if r["status"] != "done":
        print("ReflexGrid:", r.get("error"))
        return []
    out = []
    for i in range(1, len(names) + 1):
        a = s.run("cell.get", cell=f"'Results'!A{i}")["value"] or ""
        b = s.run("cell.get", cell=f"'Results'!B{i}")["value"] or ""
        out.append((a, b))
    return out


def main_objects(args) -> int:
    s, names = object_workbook()
    mine = ours_objects()
    if args.ours_only:
        for n, (a, b) in zip(names, mine):
            if args.only in n:
                print(f"{n}\n  {a}\n  {b}")
        return 0
    out = Path(args.keep or tempfile.mkdtemp(prefix="reflexgrid-vba-objects-"))
    out.mkdir(parents=True, exist_ok=True)
    path = out / "vba_objects.xlsm"
    path.write_bytes(base64.b64decode(s.run("file.saveBytes")["data"]))
    print(f"workbook: {path}")
    flat = excel_results(path, len(names), "RunAll", 2)
    if len(flat) != 2 * len(names):
        print(f"Excel returned {len(flat)} values for {len(names)} cases: {flat[:6]}")
        return 1
    excel = [(_mac_objects(flat[2 * i]), _mac_objects(flat[2 * i + 1])) for i in range(len(names))]
    mine = [(a.rstrip(), b.rstrip()) for a, b in mine]
    if args.save:
        import json

        Path(args.save).write_text(json.dumps(dict(zip(names, [list(x) for x in excel])), indent=1, ensure_ascii=False) + "\n", "utf-8")
        print(f"saved Excel's results to {args.save}")
    bad = 0
    for n, e, m in zip(names, excel, mine):
        if args.only and args.only not in n:
            continue
        if e != tuple(m):
            bad += 1
            print(f"DIFF {n}")
            if e[0] != m[0]:
                print(f"   excel: {e[0]}\n   ours:  {m[0]}")
            if e[1] != m[1]:
                print(f"   excel cells: {e[1]}\n   ours cells:  {m[1]}")
    print(f"{len(names) - bad} of {len(names)} cases match")
    return 0 if bad == 0 else 1


# ---------------------------------------------------------------- UserForms
#
# Forms ReflexGrid wrote (gridcore.xlsx.oforms) in a workbook Excel opens: Excel loads each one
# (Load, not Show, which would wait for a person) and describes it and every control, and so
# does ReflexGrid, from the same module.

# ---------------------------------------------------------------- UserForms
#
# Forms ReflexGrid wrote (gridcore.xlsx.oforms) in a workbook Excel opens: Excel loads each one (Load,
# not a modal Show, which would wait for a person) and describes it and every control; and the events
# a form's code fires, in order. ReflexGrid runs the same modules (tests/vba_form_cases.py).


def _save(args, key: str, values: List[str]) -> None:
    """Puts Excel's answers into the JSON file the unit tests read."""
    import json

    path = Path(args.save)
    data = json.loads(path.read_text("utf-8")) if path.exists() else {}
    data[key] = values
    path.write_text(json.dumps(data, indent=1, ensure_ascii=False) + "\n", "utf-8")
    print(f"saved Excel's answers to {path} ({key})")


def _forms_run(s) -> List[str]:
    from tests.vba_form_cases import results

    r = s.run("macro.run", name="RunAll", timeLimit=120)
    if r["status"] != "done":
        print("ReflexGrid:", r.get("error"))
        return []
    return results(s)


def _compare(excel: List[str], mine: List[str], what: str) -> int:
    bad = 0
    for i in range(max(len(excel), len(mine))):
        e = excel[i] if i < len(excel) else "(nothing)"
        m = mine[i] if i < len(mine) else "(nothing)"
        if e != m:
            bad += 1
            print(f"DIFF {what} {i + 1}\n   excel: {e}\n   ours:  {m}")
    print(f"{len(mine) - bad} of {len(mine)} {what} match")
    return 0 if bad == 0 else 1


def main_forms(args) -> int:
    from tests.vba_form_cases import forms_workbook

    return _forms_oracle(args, forms_workbook(), "forms", "vba_forms.xlsm")


def main_form_events(args) -> int:
    from tests.vba_form_cases import events_workbook

    return _forms_oracle(args, events_workbook(), "events", "vba_form_events.xlsm")


def main_form_designer(args) -> int:
    from tests.vba_form_cases import designer_workbook

    return _forms_oracle(args, designer_workbook(), "designer", "vba_form_designer.xlsm")


def _forms_oracle(args, s, key: str, file_name: str) -> int:
    assert s.run("macro.list")["diagnostics"] == []
    mine = _forms_run(s)
    if args.ours_only:
        print("\n".join(mine))
        return 0
    if not mine:
        return 1
    out = Path(args.keep or tempfile.mkdtemp(prefix="reflexgrid-vba-forms-"))
    out.mkdir(parents=True, exist_ok=True)
    path = out / file_name
    path.write_bytes(base64.b64decode(s.run("file.saveBytes")["data"]))
    print(f"workbook: {path}")
    excel = [x for x in excel_results(path, len(mine)) if x != ""]
    if args.save and excel:
        _save(args, key, excel)
    return _compare(excel, mine, key)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--keep")
    ap.add_argument("--only", default="")
    ap.add_argument("--ours-only", action="store_true", help="print ReflexGrid's results without Excel")
    ap.add_argument("--objects", action="store_true", help="the Excel object model cases (tests/vba_object_cases.py)")
    ap.add_argument("--save", help="with --objects, --forms or --form-events: write Excel's results to this JSON file")
    ap.add_argument("--forms", action="store_true", help="UserForms ReflexGrid wrote, loaded and described by Excel")
    ap.add_argument("--form-events", action="store_true", help="the events a form's code fires, and their order")
    ap.add_argument("--form-designer", action="store_true", help="a form made with ReflexGrid's form designer, loaded and described by Excel")
    args = ap.parse_args()
    if args.objects:
        return main_objects(args)
    if args.forms:
        return main_forms(args)
    if args.form_events:
        return main_form_events(args)
    if args.form_designer:
        return main_form_designer(args)
    code, names = build_module()
    mine = ours(code)
    if args.ours_only:
        for n, m in zip(names, mine):
            if args.only in n:
                print(f"{n:60} {m}")
        return 0
    if sys.platform != "darwin" or not Path("/Applications/Microsoft Excel.app").exists():
        print("This check needs Microsoft Excel on macOS.")
        return 2
    out = Path(args.keep or tempfile.mkdtemp(prefix="reflexgrid-vba-"))
    out.mkdir(parents=True, exist_ok=True)
    s = Session()
    s.new_workbook()
    p = new_project([]).with_module(VbaModule("Module1", STANDARD, code, default_attributes("Module1", STANDARD)))
    from gridcore.model.vba import CLASS

    p = p.with_module(VbaModule("Class1", CLASS, CLASS1.replace("Attribute Name.VB_UserMemId = 0\n", ""), default_attributes("Class1", CLASS), (("Name", "Attribute Name.VB_UserMemId = 0"),)))
    s.doc().wb.vba = p
    path = out / "vba_oracle.xlsm"
    path.write_bytes(base64.b64decode(s.run("file.saveBytes")["data"]))
    print(f"workbook: {path}")
    excel = excel_results(path, len(names))
    if len(excel) != len(names):
        print(f"Excel returned {len(excel)} results for {len(names)} cases: {excel[:5]}")
        return 1
    bad = 0
    excel = [_us_dates(x) for x in excel]
    for n, e, m in zip(names, excel, mine):
        if args.only and args.only not in n:
            continue
        if any(q in n for q in MAC_QUIRKS):
            continue
        if e != m and not _close(e, m):
            bad += 1
            print(f"DIFF {n:55} excel={e!r:40} ours={m!r}")
    print(f"{len(names) - bad} of {len(names)} cases match" if not args.only else f"{bad} differences")
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
