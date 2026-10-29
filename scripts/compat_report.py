"""Writes ReflexGrid's Excel compatibility report from a test run and the checks in Microsoft Excel.

The unit tests compare ReflexGrid with answers recorded from Microsoft Excel (``tests/*_expected.json``);
the checks in Excel itself (``excel_vba_oracle_mac.py --all --record``, ``excel_check_mac.py --record``)
keep their results in ``docs/excel-oracle.json``. This script puts them together with the browser
suites' results:

    poetry run pytest --junitxml=reports/unit.xml
    poetry run python scripts/compat_report.py reports/unit.xml --save-env      # a CI job summary
    poetry run python scripts/compat_report.py reports/unit.xml --page           # the page, to read
    python scripts/compat_report.py reports/unit-py3.15/unit.xml --also reports/unit-py3.11/unit.xml \\
        --e2e reports/e2e --site docs                                            # what CI publishes

``--site DIR`` writes the page (``compatibility.md``), the same report as data (``compatibility.json``)
and the JUnit XML reports (``data/``) into the docs site's folder. ``--save-env`` records the Python and
system the tests ran on (``env.json`` beside the XML), for a report written later on another machine.
``--e2e`` is a folder of ``<suite>.json`` files (``{"suite": "vba_userforms", "result": "success"}``)
written by the browser jobs.
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

REPO = "https://github.com/milochen0418/reflex_grid"
SITE = "https://milochen0418.github.io/reflex_grid"
PAGE = f"{SITE}/compatibility.html"
TITLE = "ReflexGrid Excel compatibility report"
AUTHOR = "Milo Chen"
ORACLE = ROOT / "docs" / "excel-oracle.json"
KINDS = ("passed", "known_differences", "failed", "skipped")

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
    ("Bugs fixed in GridCraft", ("test_gridcraft_issues",),
     "Issues that [GridCraft](https://github.com/storytold/gridcraft), whose design gridcore follows, fixed "
     "against Excel 365's answers: functions, number formats, recalculation, `.xlsx` files, fills, filters and links"),
    ("AI through the AI Hub", ("test_ai",),
     "How AI Fill and formula suggestions word their questions and read the answers, and the tools the "
     "hub's assistant reads workbooks with"),
    ("Editing together", ("test_collab", "test_books"),
     "Several people on one workbook, undo that takes back only your own changes, the shared "
     "workbook store"),
    ("Third-party VBA: VBA-JSON's own specs", ("test_third_party_vba",),
     "The spec workbook of [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) (MIT), written for Excel by its "
     "authors: ReflexGrid opens the published `.xlsm` and runs its own suite, one test per spec"),
]

# How the summary names each check in Excel (docs/excel-oracle.json's keys).
SHORT = {
    "vba-language": "VBA language and library cases",
    "vba-objects": "object-model scenarios",
    "userforms": "UserForm descriptions",
    "userform-events": "UserForm event steps",
    "userform-designer": "form-designer descriptions",
    "files": "file and formula checks",
}

RESULT_WORDS = {"success": "Passed", "failure": "**Failed**", "cancelled": "Cancelled", "skipped": "Skipped"}


# ---------------------------------------------------------------- reading


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
    """Counts per test module and the known differences (expected failures, with their reasons)."""
    counts: Dict[str, Dict[str, int]] = defaultdict(lambda: dict.fromkeys(KINDS, 0))
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
                    outcome = "known_differences"
                    name = re.sub(r"^[^\[]*\[|\]$", "", case.get("name") or "")
                    known.append((name, (child.get("message") or "").strip()))
                else:
                    outcome = "skipped"
        counts[module][outcome] += 1
    return counts, known


def this_env() -> Dict[str, str]:
    system = platform.system()
    release = Path("/etc/os-release")
    if release.exists():
        m = re.search(r'^PRETTY_NAME="?([^"\n]+)', release.read_text(), re.M)
        system = m.group(1) if m else system
    elif system == "Darwin":
        system = f"macOS {platform.mac_ver()[0]}"
    return {"python": platform.python_version(), "os": system}


def read_env(junit: Path) -> Dict[str, str]:
    env = junit.with_name("env.json")
    return json.loads(env.read_text("utf-8")) if env.exists() else this_env()


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
    return re.findall(r"^\| `(\w+)` \| (.+?) \|$", (ROOT / "testcases" / "README.md").read_text("utf-8"), re.M)


def git(*args: str) -> Optional[str]:
    try:
        p = subprocess.run(["git", *args], cwd=ROOT, capture_output=True, text=True)
    except OSError:
        return None
    return p.stdout.strip() if p.returncode == 0 else None


def excel_freshness(excel: Dict[str, Any], head: str) -> List[Dict[str, Any]]:
    """For each commit the checks in Excel ran on: how far this report's commit is from it, and which
    of the files the checks depend on (``oracle_record.INPUTS``) changed since."""
    from scripts.oracle_record import INPUTS

    out = []
    for ran in sorted({c["reflexgrid_commit"] for c in excel.get("checks", {}).values()}):
        base = ran.split()[0] if ran else ""
        entry: Dict[str, Any] = {"checks_commit": base, "uncommitted_changes": "uncommitted" in ran, "report_commit": head}
        count = git("rev-list", "--count", f"{base}..{head}") if base and head else None
        changed = git("diff", "--name-only", base, head, "--", *INPUTS) if count is not None else None
        if count is None or changed is None:
            entry["known"] = False
        else:
            entry.update(known=True, commits_since=int(count), changed_since=[f for f in changed.splitlines() if f],
                         compare=f"{REPO}/compare/{base[:12]}...{head[:12]}")
        out.append(entry)
    return out


def excel_totals(excel: Dict[str, Any]) -> Tuple[int, int]:
    checks = excel.get("checks", {}).values()
    return sum(c["matched"] for c in checks), sum(c["compared"] for c in checks)


def collect(junit: Path, also: List[Path], e2e: Optional[Path], oracle: Path) -> Dict[str, Any]:
    """Everything the report says, as data (what ``compatibility.json`` holds)."""
    data: Dict[str, Any] = {
        "name": TITLE,
        "url": PAGE,
        "repository": REPO,
        "license": "MIT",
        "commit": commit(),
        "generated": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d"),
        "test_run": run_url(),
        "unit_tests": None,
        "known_differences": [],
        "browser_suites": [],
        "excel": None,
        "excel_freshness": [],
        "excel_overall": None,
    }
    if junit.exists():
        answers = recorded_answers()
        counts, known = read_junit(junit)
        rows = [(area, modules, what.format(**answers)) for area, modules, what in AREAS]
        other = tuple(sorted(m for m in counts if not any(m in mods for _, mods, _ in AREAS)))
        if other:
            rows.append(("Other", other, "Other unit tests"))
        areas = [{"area": area, "checks": what, **{k: sum(counts[m][k] for m in modules if m in counts) for k in KINDS}}
                 for area, modules, what in rows]
        env = read_env(junit)
        data["unit_tests"] = {"python": env["python"], "os": env["os"], "areas": areas,
                              "total": {k: sum(a[k] for a in areas) for k in KINDS}, "also": []}
        by_reason: Dict[str, List[str]] = defaultdict(list)
        for name, reason in known:
            by_reason[reason].append(name)
        data["known_differences"] = [{"cases": names, "reason": reason} for reason, names in by_reason.items()]
        for extra in also:
            if extra.exists():
                c, _ = read_junit(extra)
                data["unit_tests"]["also"].append({**read_env(extra), **{k: sum(v[k] for v in c.values()) for k in KINDS}})
    results: Dict[str, str] = {}
    if e2e and e2e.is_dir():
        for f in e2e.glob("*.json"):
            r = json.loads(f.read_text("utf-8"))
            results[r["suite"]] = r["result"]
    data["browser_suites"] = [{"suite": name, "checks": what, "result": results[name]} for name, what in suites() if name in results]
    if oracle.exists():
        data["excel"] = json.loads(oracle.read_text("utf-8"))
        data["excel_freshness"] = excel_freshness(data["excel"], data["commit"])
        matched, compared = excel_totals(data["excel"])
        if compared:
            versions = sorted({c["excel"] for c in data["excel"]["checks"].values()})
            data["excel_overall"] = {
                "matched": matched,
                "compared": compared,
                "percent": round(100 * matched / compared, 2),
                "statement": f"ReflexGrid matched {' and '.join(versions)} in {matched} of {compared} published "
                             f"comparison cases ({100 * matched / compared:.1f}%). These cases cover what the checks "
                             "exercise, not every Excel feature or workload.",
            }
    return data


# ---------------------------------------------------------------- writing


def _failed(c: Dict[str, int]) -> str:
    text = f"**{c['failed']}**" if c["failed"] else "0"
    return text + (f" ({c['skipped']} skipped)" if c["skipped"] else "")


def _short(sha: str) -> str:
    return sha[:7] + (" with uncommitted changes" if "uncommitted" in sha else "")


def unit_table(unit: Dict[str, Any], with_what: bool) -> List[str]:
    head = ["Area"] + (["What is checked"] if with_what else []) + ["Passed", "Known differences", "Failed"]
    lines = ["| " + " | ".join(head) + " |", "|" + "---|" * len(head)]
    for a in unit["areas"]:
        cells = [a["area"]] + ([a["checks"]] if with_what else []) + [str(a["passed"]), str(a["known_differences"]), _failed(a)]
        lines.append("| " + " | ".join(cells) + " |")
    t = unit["total"]
    cells = ["**All unit tests**"] + ([""] if with_what else []) + [f"**{t['passed']}**", str(t["known_differences"]), _failed(t)]
    lines.append("| " + " | ".join(cells) + " |")
    return lines


def excel_table(excel: Dict[str, Any], page: bool) -> List[str]:
    lines = ["| Check | Same answer as Excel | Differences | Not compared | Last run |", "|---|---|---|---|---|"]
    for c in excel.get("checks", {}).values():
        missing = c.get("not_compared", {}).get("cases", 0)
        sha = c["reflexgrid_commit"]
        run = f"{c['date']}, [{_short(sha)}]({REPO}/commit/{sha.split()[0]})" if page and sha else f"{c['date']}, {_short(sha)}"
        lines.append(f"| {c['title']} | {c['matched']} of {c['compared']} | {len(c['differences'])} | {missing or ''} | {run} |")
    return lines


def summary_lines(data: Dict[str, Any]) -> List[str]:
    out = []
    unit = data["unit_tests"]
    if unit:
        t = unit["total"]
        same = all(a["passed"] == t["passed"] and a["failed"] == t["failed"] for a in unit["also"])
        pythons = " and ".join([unit["python"]] + [a["python"] for a in unit["also"]])
        each = " on each" if unit["also"] and same else ""
        out.append(f"- **Unit tests** on Python {pythons}: {t['passed']} pass{each}, "
                   f"{t['known_differences']} are known differences from Excel, {t['failed']} fail"
                   f"{'' if same else ' (the Pythons differ: see below)'}.")
    if data["browser_suites"]:
        ok = sum(s["result"] == "success" for s in data["browser_suites"])
        out.append(f"- **In the browser**: {ok} of {len(data['browser_suites'])} end-to-end suites pass, among them "
                   "several people editing one workbook, VBA macros and UserForms.")
    if data["excel"]:
        checks = data["excel"].get("checks", {})
        versions = " and ".join(sorted({c["excel"] for c in checks.values()}))
        parts = [f"{c['matched']} of {c['compared']} {SHORT.get(k, k)}" for k, c in checks.items()]
        out.append(f"- **In {versions}**: " + "; ".join(parts) + " gave the same answer as in ReflexGrid.")
    if data["excel_overall"]:
        out.append(f"- **Overall**: {data['excel_overall']['statement']}")
    stale = [f for f in data["excel_freshness"] if f.get("changed_since") or f["uncommitted_changes"]]
    for f in stale:
        out.append(f"- **Note**: the checks in Excel ran on commit {f['checks_commit'][:7]}, and files they depend on "
                   "changed since, so they may not describe this commit until they run again (see below).")
    return out


def freshness_lines(data: Dict[str, Any], page: bool) -> List[str]:
    """Which commit the checks in Excel ran on, against this report's commit."""
    out = []
    excel = data["excel"]
    for f in data["excel_freshness"]:
        runs = [c for c in excel["checks"].values() if c["reflexgrid_commit"].split()[0] == f["checks_commit"]]
        where = "; ".join(sorted({f"{c['excel']} on {c['os']}, {c['date']}" for c in runs}))
        ran = f"[{f['checks_commit'][:7]}]({REPO}/commit/{f['checks_commit']})" if page else f["checks_commit"][:7]
        head = f"[{f['report_commit'][:7]}]({REPO}/commit/{f['report_commit']})" if page else f["report_commit"][:7]
        text = f"The checks ran on commit {ran}" + (" with uncommitted changes" if f["uncommitted_changes"] else "") + f", in {where}. "
        if not f["known"]:
            text += f"This report is for commit {head}; it couldn't compare the two here (no git history)."
        elif f["commits_since"] == 0:
            text += "That is this report's commit."
        else:
            n = f["commits_since"]
            later = f"{n} commit{'s' if n != 1 else ''} later" + (f" ([changes]({f['compare']}))" if page else "")
            text += f"This report is for commit {head}, {later}. "
            if f["changed_since"]:
                files = f["changed_since"]
                shown = ", ".join(f"`{x}`" for x in files[:5]) + (f" and {len(files) - 5} more" if len(files) > 5 else "")
                text += (f"Since the checks ran, {len(files)} file{'s' if len(files) != 1 else ''} they depend on changed "
                         f"({shown}), so their results describe commit {f['checks_commit'][:7]} and may not hold for "
                         "this one until the checks run again.")
            else:
                text += ("Nothing the checks depend on (the engine, the test cases, the check scripts) changed since, "
                         "so their results hold for this commit too.")
        out.append(text)
    return out


def json_ld(data: Dict[str, Any], downloads: List[Tuple[str, str, str]]) -> List[str]:
    """schema.org Dataset markup, so search engines (Google Dataset Search…) can index the report."""
    ld = {
        "@context": "https://schema.org",
        "@type": "Dataset",
        "name": TITLE,
        "description": ("Automated test results showing how closely ReflexGrid, an open-source online spreadsheet, "
                        "matches Microsoft Excel: the VBA language and library, the Excel object model, UserForms, "
                        ".xlsm files, formulas, end-to-end browser suites, and checks run in Microsoft Excel itself."),
        "url": PAGE,
        "creator": {"@type": "Person", "name": AUTHOR},
        "license": "https://opensource.org/licenses/MIT",
        "isAccessibleForFree": True,
        "dateModified": data["generated"],
        "version": data["commit"][:7] or None,
        "isBasedOn": REPO,
        "keywords": ["Excel compatibility", "VBA", "UserForms", "xlsm", "spreadsheet", "test results"],
        "distribution": [{"@type": "DataDownload", "name": label, "encodingFormat": kind, "contentUrl": f"{SITE}/{path}"}
                         for label, path, kind in downloads],
    }
    return ['<script type="application/ld+json">', json.dumps(ld, indent=1, ensure_ascii=False), "</script>"]


def render(data: Dict[str, Any], page: bool, downloads: Optional[List[Tuple[str, str, str]]] = None) -> str:
    out: List[str] = []
    unit = data["unit_tests"]
    sha = data["commit"]
    if page:
        out += [
            "---",
            "title: Excel compatibility report",
            'description: "How closely ReflexGrid matches Microsoft Excel, in its automated tests and in Excel itself: '
            'the VBA language, the Excel object model, UserForms, .xlsm files, formulas and editing together."',
            "---",
            "",
            "{% raw %}",
            *summary_lines(data),
            "",
            "ReflexGrid's tests run on GitHub Actions for every change, on a clean Linux machine. Many of them compare "
            "ReflexGrid with answers recorded from Microsoft Excel and kept in the repository "
            f"([`tests/*_expected.json`]({REPO}/tree/main/tests)), so anyone can check them. Excel can't run on that "
            "machine, so the checks in Excel itself run on a Mac with Excel, and their latest results are published "
            "with this page. The page is written by the test run (`scripts/compat_report.py`) each time `main` "
            "changes.",
            "",
            f"- **Commit**: [{sha[:7]}]({REPO}/commit/{sha})" if sha else "- **Commit**: unknown",
        ]
        if unit:
            out.append(f"- **Tested**: {data['generated']} (UTC), {unit['os']}, Python {unit['python']}")
            for a in unit["also"]:
                out.append(f"- **Also tested**: Python {a['python']}: {a['passed']} passed, {a['known_differences']} "
                           f"known differences, {a['failed']} failed")
        if data["test_run"]:
            out.append(f"- **Test run**: [GitHub Actions]({data['test_run']})")
        out += ["", "## Unit tests", ""]
    elif unit:
        out += [f"### Excel compatibility: Python {unit['python']}, {unit['os']}", ""]

    out += unit_table(unit, with_what=page) if unit else ["No unit test results."]

    if data["known_differences"]:
        out += ["", "## Known differences from Excel" if page else "**Known differences from Excel**", ""]
        if page:
            out += ["These cases give a different answer from Excel's, on purpose or for now; the tests expect it "
                    "(`KNOWN` in `tests/test_vba_object_cases.py` and `tests/test_third_party_vba.py`).", ""]
        out += [f"- **{', '.join(k['cases'])}**: {k['reason']}" for k in data["known_differences"]]

    if data["browser_suites"]:
        out += ["", "## In the browser" if page else "**In the browser**", ""]
        if page:
            out += ["End-to-end suites drive the real app with Playwright (headless Chromium), one job per suite "
                    f"([`testcases/`]({REPO}/tree/main/testcases)).", ""]
        out += ["| Suite | What it checks | Result |", "|---|---|---|"]
        out += [f"| `{s['suite']}` | {s['checks']} | {RESULT_WORDS.get(s['result'], s['result'])} |" for s in data["browser_suites"]]

    if data["excel"]:
        checks = list(data["excel"].get("checks", {}).values())
        out += ["", "## Checked in Microsoft Excel" if page else "**Checked in Microsoft Excel**", ""]
        if page:
            out += ["These checks run ReflexGrid's output in Microsoft Excel itself: Excel opens the macro-enabled "
                    "workbooks and the files ReflexGrid writes, runs the same VBA, describes the same UserForms, and "
                    "its answers are compared with ReflexGrid's (`scripts/excel_vba_oracle_mac.py --all --record`, "
                    "`scripts/excel_check_mac.py --record`). They need Excel, so they run on a Mac rather than in CI, "
                    "and each row says when and on which commit it last ran.", ""]
        out += [*freshness_lines(data, page), ""]
        out += excel_table(data["excel"], page)
        if page:
            notes = []
            for c in checks:
                name = c["title"].split(":")[0]
                for key, what in (("not_compared", "not compared"), ("not_run", "not run on the Mac")):
                    if c.get(key, {}).get("cases"):
                        notes.append(f"- **{name}**: {c[key]['cases']} {what}, because {c[key]['why']}.")
                if c["differences"]:
                    notes.append(f"- **{name}**, the differences: {', '.join(d['case'] for d in c['differences'])} "
                                 "(the known differences above).")
            out += ["", *notes, "", "Every difference, with Excel's answer and ReflexGrid's, is in "
                    "[excel-oracle.json](excel-oracle.json)."]

    if page:
        if downloads:
            out += ["", "## Download", "", *[f"- [{label}]({path})" for label, path, _ in downloads]]
        out += [
            "",
            "## Limits",
            "",
            "What ReflexGrid doesn't do yet (pivot tables, Power Query, controls placed on sheets, macro recording, a "
            "VBA debugger, anything outside the workbook from a macro) is listed in "
            "[Excel macros (VBA)](vba.md#not-supported-and-what-happens-instead) and the "
            f"[README's FAQ]({REPO}#faq).",
            "",
            "## Cite this report",
            "",
            f"{AUTHOR} and contributors ({data['generated'][:4]}). *{TITLE}*, commit {sha[:7] or 'unknown'}, "
            f"{data['generated']}. {PAGE}",
            "",
            *([f"To quote the result: *{data['excel_overall']['statement']}*", ""] if data["excel_overall"] else []),
            f"For the software itself, see [CITATION.cff]({REPO}/blob/main/CITATION.cff) (GitHub's *Cite this "
            "repository*).",
            "{% endraw %}",
        ]
        if downloads:
            out += ["", *json_ld(data, downloads)]
    return "\n".join(out) + "\n"


def write_site(data: Dict[str, Any], junits: List[Path], folder: Path) -> None:
    """The page, the report as data and the JUnit XML reports, into the docs site's folder."""
    downloads = [("This report as JSON", "compatibility.json", "application/json")]
    if data["excel"]:
        downloads.append(("The checks in Microsoft Excel, with every difference (JSON)", "excel-oracle.json", "application/json"))
    (folder / "data").mkdir(parents=True, exist_ok=True)
    for junit in junits:
        if junit.exists():
            version = ".".join(read_env(junit)["python"].split(".")[:2])
            name = f"data/unit-py{version}.xml"
            shutil.copyfile(junit, folder / name)
            downloads.append((f"Unit test results, Python {version} (JUnit XML)", name, "application/xml"))
    (folder / "compatibility.json").write_text(json.dumps(data, indent=1, ensure_ascii=False) + "\n", "utf-8")
    (folder / "compatibility.md").write_text(render(data, page=True, downloads=downloads), "utf-8")


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("junit", type=Path, help="pytest's JUnit XML report")
    ap.add_argument("--also", type=Path, action="append", default=[], help="another run's report (e.g. another Python)")
    ap.add_argument("--e2e", type=Path, help="folder of the browser suites' <suite>.json results")
    ap.add_argument("--oracle", type=Path, default=ORACLE, help="the checks in Excel (default docs/excel-oracle.json)")
    ap.add_argument("--page", action="store_true", help="print the docs site page instead of a CI summary")
    ap.add_argument("--site", type=Path, help="write the page, compatibility.json and the JUnit XML into this folder")
    ap.add_argument("--save-env", action="store_true", help="record this Python and system beside the report")
    args = ap.parse_args(argv)
    if args.save_env:
        args.junit.parent.mkdir(parents=True, exist_ok=True)
        args.junit.with_name("env.json").write_text(json.dumps(this_env()), "utf-8")
    data = collect(args.junit, args.also, args.e2e, args.oracle)
    if args.site:
        write_site(data, [args.junit, *args.also], args.site)
        print(f"wrote {args.site / 'compatibility.md'}, compatibility.json and data/")
    else:
        sys.stdout.write(render(data, page=args.page))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
