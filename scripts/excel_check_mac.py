"""Checks ReflexGrid's .xlsx files with the Microsoft Excel installed on this Mac (opt-in: it opens
Excel windows).

For each workbook it writes, the script: opens it in Excel (a repair prompt would block and be
reported), reads values, formulas and displayed text through AppleScript, lets Excel save it, and
reads Excel's file back with gridcore, comparing a full recalculation with Excel's cached results.

    poetry run python scripts/excel_check_mac.py [--keep DIR] [--record]

``--record`` keeps the results, with the Excel version and the commit, in ``docs/excel-oracle.json``
(``oracle_record.py``), which the Excel compatibility report publishes.

Functions this Excel installation doesn't have (older licences lack IFS, XLOOKUP, SEQUENCE…) show
as #NAME? in Excel; the script detects them first and reports those cells as skipped.
"""

from __future__ import annotations

import argparse
import base64
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Dict, List, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from gridcore.calc import Calc  # noqa: E402
from gridcore.core.value import CellError  # noqa: E402
from gridcore.engine import Session  # noqa: E402
from gridcore.engine.samples import new_sample  # noqa: E402
from gridcore.core.addr import a1  # noqa: E402
from gridcore.xlsx import read_xlsx  # noqa: E402
from scripts.oracle_record import RESULTS, ROOT, record  # noqa: E402

CELL_SCRIPT = """on run argv
	tell application "Microsoft Excel"
		set r to range (item 3 of argv) of worksheet (item 2 of argv) of workbook (item 1 of argv)
		set v to value of r
		set f to formula of r
		try
			set t to string value of r
		on error
			set t to "?"
		end try
		return (v as text) & "|" & f & "|" & t
	end tell
end run"""

PROBE = ["IFS(TRUE,1)", "XLOOKUP(1,{1},{1})", "SEQUENCE(1)", "TEXTJOIN(\"\",TRUE,\"a\")", "LET(x,1,x)", "FILTER({1},{TRUE})"]

Check = tuple  # (sheet, cell, value|formula|text, expected[, functions the check needs])


def osa(args: List[str], timeout: float = 20) -> str:
    try:
        p = subprocess.run(["osascript"] + args, capture_output=True, text=True, timeout=timeout)
        return (p.stdout + p.stderr).strip()
    except subprocess.TimeoutExpired:
        return "TIMEOUT"


def excel(script: str, timeout: float = 30) -> str:
    return osa(["-e", script], timeout)


def build(out: Path) -> List[Tuple[str, List[Check]]]:
    books: List[Tuple[str, List[Check]]] = []

    def save(s: Session, name: str, checks: List[Check]) -> None:
        (out / name).write_bytes(base64.b64decode(s.run("file.saveBytes")["data"]))
        books.append((name, checks))

    for nm, checks in (
        ("Sales", [("Sales", "F2", "value", 1368), ("Sales", "F2", "formula", "=[@Units]*[@Price]"), ("Sales", "I2", "text", "$5,390.25"), ("Sales", "J5", "text", "22.0%")]),
        ("Budget", [("Budget", "N4", "value", 21510), ("Budget", "O11", "formula", "=SUM(O4:O10)")]),
        ("Grades", [("Grades", "G2", "value", "C", ["IFS"]), ("Grades", "J2", "value", 76.25, ["XLOOKUP"]), ("Grades", "J5", "value", "Harper")]),
    ):
        s = Session()
        new_sample(s, nm)
        save(s, f"{nm.lower()}.xlsx", checks)

    s = Session()
    s.new_workbook()
    r = s.run
    r("range.setValues", range="A1", values=[["Name", "Amt", "When"], ["a", 1, "2026-01-05"], ["b", 2, "2026-02-05"], ["c", 3, "2026-03-05"]])
    r("cell.set", cell="E1", input="=SEQUENCE(3)")
    r("cell.set", cell="F1", input="=SUMPRODUCT(B2:B4,B2:B4)")
    r("cell.set", cell="F2", input="=TEXT(C2,\"mmm d\")")
    r("selection.set", range="A1:C4")
    r("insert.table")
    r("home.mergeCenter", range="H1:J1")
    r("cell.set", cell="H1", input="Merged")
    r("home.conditionalFormat", range="B2:B4", rule={"type": "dataBar"})
    r("home.conditionalFormat", range="B2:B4", rule={"type": "iconSet", "iconSet": "3Arrows"})
    r("data.validation", range="K1:K5", kind="list", f1="x,y")
    r("insert.hyperlink", cell="L1", target="https://example.com", text="link")
    r("review.newNote", cell="L2", text="note")
    r("selection.set", range="A1:B4")
    r("insert.chart", kind="lineMarkers")
    r("view.freezePanes", cell="B2")
    r("formulas.defineName", name="Amts", refersTo="=Sheet1!$B$2:$B$4")
    r("cell.set", cell="G4", input="=SUM(Amts)")
    r("insert.sparkline", data="B2:B4", location="M2")
    r("home.insertSheet", name="Other", index=1)
    r("cell.set", cell="A1", input="=Sheet1!B2*2")
    r("sheet.activate", sheet="Sheet1")
    save(s, "features.xlsx", [
        ("Sheet1", "E3", "value", 3, ["SEQUENCE"]), ("Sheet1", "F1", "value", 14), ("Sheet1", "F2", "value", "Jan 5"), ("Sheet1", "G4", "value", 6),
        ("Sheet1", "H1", "value", "Merged"), ("Other", "A1", "value", 2),
    ])
    return books


def probe_functions() -> List[str]:
    """Functions the local Excel can't evaluate (typed into a scratch workbook, never saved)."""
    lines = [f'set formula of range "A{i + 1}" of ws to "={f.replace(chr(34), chr(92) + chr(34))}"' for i, f in enumerate(PROBE)]
    reads = [f'set end of out to ((value of range "A{i + 1}" of ws) as text)' for i in range(len(PROBE))]
    out = excel('tell application "Microsoft Excel"\nset wb to make new workbook\nset ws to worksheet 1 of wb\nset out to {}\n' + "\n".join(lines + reads) + "\nclose wb saving no\nreturn out\nend tell", 60)
    vals = [v.strip() for v in out.split(",")]
    return [PROBE[i].split("(")[0] for i, v in enumerate(vals) if v == "missing value"]


def main() -> int:
    if sys.platform != "darwin" or not Path("/Applications/Microsoft Excel.app").exists():
        print("This check needs Microsoft Excel on macOS.")
        return 2
    ap = argparse.ArgumentParser()
    ap.add_argument("--keep", help="folder to keep the files in (default: a temporary folder)")
    ap.add_argument("--record", nargs="?", const=str(RESULTS), help=f"keep the results, with the Excel version and the commit, in this JSON file (default {RESULTS.relative_to(ROOT)})")
    args = ap.parse_args()
    out = Path(args.keep or tempfile.mkdtemp(prefix="reflexgrid-excel-"))
    out.mkdir(parents=True, exist_ok=True)
    was_running = subprocess.run(["pgrep", "-x", "Microsoft Excel"], capture_output=True).returncode == 0
    books = build(out)
    print(f"files in {out}")
    subprocess.run(["open", "-a", "Microsoft Excel"])
    for _ in range(60):
        if excel('tell application "Microsoft Excel" to get version', 10) not in ("TIMEOUT", ""):
            break
        time.sleep(1)
    missing = probe_functions()
    print("functions this Excel lacks:", ", ".join(missing) or "none")
    failures = 0
    compared = matched = skipped_all = opened_books = 0
    diffs: List[Dict[str, str]] = []
    for name, checks in books:
        subprocess.run(["open", "-a", "Microsoft Excel", str(out / name)])
        opened = False
        for _ in range(40):
            got = excel('tell application "Microsoft Excel" to get name of every workbook', 10)
            if name in got:
                opened = True
                break
            time.sleep(1)
        if not opened:
            print(f"FAIL {name}: Excel did not open it (a repair prompt may be showing)")
            failures += 1
            diffs.append({"case": name, "excel": "did not open the file", "reflexgrid": "wrote it"})
            continue
        opened_books += 1
        for check in checks:
            sheet, addr, kind, want = check[:4]
            needs = check[4] if len(check) > 4 else []
            res = osa(["-e", CELL_SCRIPT, name, sheet, addr])
            if "|" not in res:
                print(f"FAIL {name} {sheet}!{addr}: {res}")
                failures += 1
                compared += 1
                diffs.append({"case": f"{name} {sheet}!{addr}", "excel": res, "reflexgrid": str(want)})
                continue
            v, f, t = res.split("|", 2)
            got = {"value": v, "formula": f, "text": t}[kind]
            if isinstance(want, (int, float)) and kind == "value":
                try:
                    ok = abs(float(got) - float(want)) < 1e-6
                except ValueError:
                    ok = False
            else:
                ok = got == str(want)
            uses_missing = any(n in missing for n in needs) or any(m.upper() + "(" in f.upper() for m in missing)
            status = "ok  " if ok else ("skip" if uses_missing else "FAIL")
            failures += status == "FAIL"
            if status == "skip":
                skipped_all += 1
            else:
                compared += 1
                matched += ok
                if not ok:
                    diffs.append({"case": f"{name} {sheet}!{addr} {kind}", "excel": got, "reflexgrid": str(want)})
            print(f"{status} {name} {sheet}!{addr} {kind}: Excel={got!r} expected={want!r}")
        excel(f'tell application "Microsoft Excel"\nset saved of workbook "{name}" to false\nclose workbook "{name}" saving yes\nend tell', 60)
        # Excel's own file, read back and recalculated by gridcore.
        wb, warnings = read_xlsx((out / name).read_bytes())
        cached = {}
        for si, sh in enumerate(wb.sheets):
            for rc, c in sh.cells.items():
                if c.formula is not None:
                    cached[(si, rc)] = c.value
        Calc().recalc_all(wb)
        same = skipped = 0
        for (si, rc), v in cached.items():
            nv = wb.sheets[si].value(rc)
            if v is CellError.NAME and missing:
                skipped += 1
            elif nv == v or (isinstance(v, float) and isinstance(nv, float) and abs(v - nv) <= 1e-9 * max(1.0, abs(v))):
                same += 1
            else:
                failures += 1
                print(f"FAIL {name} {wb.sheets[si].name}!{rc}: Excel cached {v!r}, gridcore computed {nv!r}")
                diffs.append({"case": f"{name} {wb.sheets[si].name}!{a1(*rc)} recalculated", "excel": repr(v), "reflexgrid": repr(nv)})
        compared += same + (len(cached) - same - skipped)
        matched += same
        skipped_all += skipped
        print(f"     {name}: Excel's file read back ({len(warnings)} warnings); {same} formulas match Excel, {skipped} skipped")
    extra = {"workbooks": {"opened": opened_books, "total": len(books)}}
    if skipped_all:
        extra["not_compared"] = {"cases": skipped_all, "why": f"they use functions this Excel lacks ({', '.join(missing)})"}
    record(args.record, "files", "Workbooks ReflexGrid wrote: opened by Excel, their values, formulas and displayed text "
           "read there, saved by Excel, and Excel's files recalculated by ReflexGrid to Excel's results",
           compared, matched, diffs, **extra)
    if not was_running and excel('tell application "Microsoft Excel" to count of workbooks', 10) == "0":
        excel('tell application "Microsoft Excel" to quit', 20)
    print("PASS" if failures == 0 else f"FAIL ({failures})")
    return 0 if failures == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
