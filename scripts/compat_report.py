"""Writes ReflexGrid's Excel compatibility report from a test run.

The unit tests compare ReflexGrid with answers recorded from Microsoft Excel (``tests/*_expected.json``,
captured with ``scripts/excel_vba_oracle_mac.py``); this script groups pytest's JUnit XML report by area
and adds the browser suites' results:

    poetry run pytest --junitxml=reports/unit.xml
    poetry run python scripts/compat_report.py reports/unit.xml --save-env      # a CI job summary
    python scripts/compat_report.py reports/unit.xml --also reports/py311/unit.xml \\
        --e2e reports/e2e --page > docs/compatibility.md                       # the docs site page

``--save-env`` records the Python and system the tests ran on (``env.json`` beside the XML), for a
report written later on another machine. ``--e2e`` is a folder of ``<suite>.json`` files
(``{"suite": "vba_userforms", "result": "success"}``) written by the browser jobs.
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import platform
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

REPO = "https://github.com/milochen0418/reflex_grid"

# (area, test modules, what is checked); {names} are filled from the recorded Excel answers.
AREAS: List[Tuple[str, Tuple[str, ...], str]] = [
    ("VBA language and library", ("test_vba_excel_cases",),
     "Expressions, conversions, strings, dates, `Format`, arrays, `On Error`, classes… against "
     "{excel_cases} answers: {excel_mac} recorded from Microsoft Excel for Mac, {excel_win} that follow "
     "Excel for Windows where Excel for Mac lacks the feature"),
    ("Excel object model from VBA", ("test_vba_object_cases",),
     "`Range`, `Worksheet`, `Application`, tables, filters, names, formats… against {object_cases} "
     "answers recorded from Microsoft Excel"),
    ("UserForms", ("test_vba_forms",),
     "Form designs in `.xlsm` and `.frm`/`.frx` files, the Forms object model, events and their "
     "order, against {form_answers} answers recorded from Microsoft Excel"),
    ("Macros and event macros", ("test_vba_macros",),
     "The `macro.*` commands, VBA functions in formulas, `Worksheet_Change` and other events, "
     "`MsgBox`/`InputBox`, undo, the sandbox"),
    (".xlsm files and VBA projects", ("test_vba_file",),
     "Compound files, MS-OVBA compression, `vbaProject.bin` and `.xlsm` packages, read back with "
     "openpyxl"),
    ("Formulas, functions and number formats", ("test_core", "test_formula", "test_functions_fuzz", "test_numfmt"),
     "Values and coercion, dates, the formula language, every worksheet function with hostile "
     "arguments, number formats as Excel shows them"),
    ("Engine and .xlsx files", ("test_engine", "test_engine_more"),
     "Commands end to end: editing, recalculation, undo, clipboard, structure, formatting, sort and "
     "filter, tables, names, `.xlsx` round trips"),
    ("Editing together", ("test_collab", "test_books"),
     "Several people on one workbook, undo that takes back only your own changes, the shared "
     "workbook store"),
]

RESULT_WORDS = {"success": "Passed", "failure": "**Failed**", "cancelled": "Cancelled", "skipped": "Skipped"}


def recorded_answers() -> Dict[str, int]:
    """How many answers were recorded from Excel, per test file."""
    from tests.vba_excel_cases import build_module

    tests = ROOT / "tests"
    forms = json.loads((tests / "vba_forms_expected.json").read_text("utf-8"))
    all_cases = len(json.loads((tests / "vba_excel_expected.json").read_text("utf-8")))
    mac = len(build_module(windows=False)[1])
    return {
        "excel_cases": all_cases,
        "excel_mac": mac,
        "excel_win": all_cases - mac,
        "object_cases": len(json.loads((tests / "vba_object_expected.json").read_text("utf-8"))),
        "form_answers": sum(len(v) for v in forms.values()),
    }


def read_junit(path: Path) -> Tuple[Dict[str, Dict[str, int]], List[Tuple[str, str]]]:
    """Counts per test module (passed, known, failed, skipped) and the known differences."""
    counts: Dict[str, Dict[str, int]] = defaultdict(lambda: {"passed": 0, "known": 0, "failed": 0, "skipped": 0})
    known: List[Tuple[str, str]] = []
    for case in ET.parse(path).getroot().iter("testcase"):
        parts = (case.get("classname") or "").split(".")
        module = parts[1] if len(parts) > 1 else parts[0]
        outcome = "passed"
        for child in case:
            if child.tag in ("failure", "error"):
                outcome = "failed"
            elif child.tag == "skipped":
                if child.get("type") == "pytest.xfail":
                    outcome = "known"
                    name = re.sub(r"^[^\[]*\[|\]$", "", case.get("name") or "")
                    known.append((name, (child.get("message") or "").strip()))
                else:
                    outcome = "skipped"
        counts[module][outcome] += 1
    return counts, known


def read_env(junit: Path) -> Dict[str, str]:
    env = junit.with_name("env.json")
    if env.exists():
        return json.loads(env.read_text("utf-8"))
    return this_env()


def this_env() -> Dict[str, str]:
    system = platform.system()
    release = Path("/etc/os-release")
    if release.exists():
        m = re.search(r'^PRETTY_NAME="?([^"\n]+)', release.read_text(), re.M)
        system = m.group(1) if m else system
    elif system == "Darwin":
        system = f"macOS {platform.mac_ver()[0]}"
    return {"python": platform.python_version(), "os": system}


def commit() -> str:
    sha = os.getenv("GITHUB_SHA")
    if not sha:
        try:
            sha = subprocess.run(["git", "rev-parse", "HEAD"], cwd=ROOT, capture_output=True, text=True, check=True).stdout.strip()
        except (OSError, subprocess.CalledProcessError):
            sha = ""
    return sha


def run_url() -> Optional[str]:
    run = os.getenv("GITHUB_RUN_ID")
    if not run:
        return None
    return f"{os.getenv('GITHUB_SERVER_URL', 'https://github.com')}/{os.getenv('GITHUB_REPOSITORY', 'milochen0418/reflex_grid')}/actions/runs/{run}"


def suites() -> List[Tuple[str, str]]:
    """The browser suites and what each checks, from testcases/README.md."""
    rows = re.findall(r"^\| `(\w+)` \| (.+?) \|$", (ROOT / "testcases" / "README.md").read_text("utf-8"), re.M)
    return [(name, what) for name, what in rows]


def area_table(counts: Dict[str, Dict[str, int]], answers: Dict[str, int], with_what: bool) -> List[str]:
    head = "| Area | What is checked | Passed | Known differences | Failed |" if with_what else "| Area | Passed | Known differences | Failed |"
    lines = [head, "|" + "---|" * (head.count("|") - 1)]
    mapped = set()
    total = {"passed": 0, "known": 0, "failed": 0, "skipped": 0}
    rows = [(area, modules, what.format(**answers)) for area, modules, what in AREAS]
    other = tuple(sorted(m for m in counts if not any(m in mods for _, mods, _ in AREAS)))
    if other:
        rows.append(("Other", other, "Other unit tests"))
    for area, modules, what in rows:
        c = {k: sum(counts[m][k] for m in modules if m in counts) for k in total}
        mapped.update(modules)
        for k in total:
            total[k] += c[k]
        failed = f"**{c['failed']}**" if c["failed"] else "0"
        if c["skipped"]:
            failed += f" ({c['skipped']} skipped)"
        cells = [area] + ([what] if with_what else []) + [str(c["passed"]), str(c["known"]), failed]
        lines.append("| " + " | ".join(cells) + " |")
    failed = f"**{total['failed']}**" if total["failed"] else "0"
    if total["skipped"]:
        failed += f" ({total['skipped']} skipped)"
    cells = ["**All unit tests**"] + ([""] if with_what else []) + [f"**{total['passed']}**", str(total["known"]), failed]
    lines.append("| " + " | ".join(cells) + " |")
    return lines


def known_list(known: List[Tuple[str, str]]) -> List[str]:
    by_reason: Dict[str, List[str]] = defaultdict(list)
    for name, reason in known:
        by_reason[reason].append(name)
    return [f"- **{', '.join(names)}**: {reason}" for reason, names in by_reason.items()]


def e2e_table(folder: Optional[Path]) -> List[str]:
    results: Dict[str, str] = {}
    if folder and folder.is_dir():
        for f in folder.glob("*.json"):
            data = json.loads(f.read_text("utf-8"))
            results[data["suite"]] = data["result"]
    if not results:
        return []
    lines = ["| Suite | What it checks | Result |", "|---|---|---|"]
    for name, what in suites():
        if name in results:
            lines.append(f"| `{name}` | {what} | {RESULT_WORDS.get(results[name], results[name])} |")
    return lines


def report(junit: Path, also: List[Path], e2e: Optional[Path], page: bool) -> str:
    answers = recorded_answers()
    env = read_env(junit) if junit.exists() else this_env()
    sha = commit()
    url = run_url()
    out: List[str] = []
    if junit.exists():
        counts, known = read_junit(junit)
    else:
        counts, known = {}, []

    if page:
        today = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d")
        out += [
            "---",
            "title: Excel compatibility report",
            'description: "How closely ReflexGrid matches Microsoft Excel in its automated tests: the VBA language, '
            'the Excel object model, UserForms, .xlsm files, formulas and editing together."',
            "---",
            "",
            "ReflexGrid's tests run on GitHub Actions for every change, on a clean Linux machine. Many of them "
            "compare ReflexGrid with answers recorded from Microsoft Excel: `scripts/excel_vba_oracle_mac.py` "
            "runs the same VBA in Excel and saves Excel's answers in the repository "
            f"([`tests/*_expected.json`]({REPO}/tree/main/tests)), so anyone can check them. Excel itself "
            "can't run on the Linux machine; the recorded answers are what it is compared with.",
            "",
            "This page is written by the test run itself (`scripts/compat_report.py`), each time `main` changes.",
            "",
            f"- **Commit**: [{sha[:7]}]({REPO}/commit/{sha})" if sha else "- **Commit**: unknown",
            f"- **Tested**: {today} (UTC), {env['os']}, Python {env['python']}",
        ]
        for extra in also:
            if extra.exists():
                c, _ = read_junit(extra)
                tot = {k: sum(v[k] for v in c.values()) for k in ("passed", "known", "failed")}
                env2 = read_env(extra)
                out.append(f"- **Also tested**: Python {env2['python']}: {tot['passed']} passed, {tot['known']} "
                           f"known differences, {tot['failed']} failed")
            else:
                out.append(f"- **Also tested**: no results from {extra.parent.name}")
        if url:
            out.append(f"- **Test run**: [GitHub Actions]({url})")
        out += ["", "## Unit tests", ""]
    else:
        out += [f"### Excel compatibility: Python {env['python']}, {env['os']}", ""]

    if counts:
        out += area_table(counts, answers, with_what=page)
    else:
        out.append(f"No unit test results ({junit} is missing).")

    if known:
        out += ["", "## Known differences from Excel" if page else "**Known differences from Excel**", ""]
        if page:
            out += ["These cases give a different answer from Excel's on purpose or for now; the tests expect it "
                    "(`KNOWN` in `tests/test_vba_object_cases.py`).", ""]
        out += known_list(known)

    browser = e2e_table(e2e)
    if browser:
        out += ["", "## In the browser" if page else "**In the browser**", ""]
        if page:
            out += ["End-to-end suites drive the real app with Playwright (headless Chromium), "
                    f"one job per suite ([`testcases/`]({REPO}/tree/main/testcases)).", ""]
        out += browser

    if page:
        out += [
            "",
            "## Limits",
            "",
            "What ReflexGrid doesn't do yet (pivot tables, Power Query, controls placed on sheets, macro "
            "recording, a VBA debugger, anything outside the workbook from a macro) is listed in "
            "[Excel macros (VBA)](vba.md#not-supported-and-what-happens-instead) and the "
            f"[README's FAQ]({REPO}#faq).",
        ]
    return "\n".join(out) + "\n"


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("junit", type=Path, help="pytest's JUnit XML report")
    ap.add_argument("--also", type=Path, action="append", default=[], help="another run's report (e.g. another Python)")
    ap.add_argument("--e2e", type=Path, help="folder of the browser suites' <suite>.json results")
    ap.add_argument("--page", action="store_true", help="write the docs site page instead of a CI summary")
    ap.add_argument("--save-env", action="store_true", help="record this Python and system beside the report")
    args = ap.parse_args(argv)
    if args.save_env:
        args.junit.parent.mkdir(parents=True, exist_ok=True)
        args.junit.with_name("env.json").write_text(json.dumps(this_env()), "utf-8")
    sys.stdout.write(report(args.junit, args.also, args.e2e, args.page))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
