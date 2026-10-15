"""Keeps what the checks against Microsoft Excel found, with the Excel and the ReflexGrid commit they ran on.

``excel_vba_oracle_mac.py --record`` and ``excel_check_mac.py --record`` write their results into
``docs/excel-oracle.json``: it is published with the docs site, and ``compat_report.py`` shows it in the
Excel compatibility report. Each check keeps its own date, Excel version and commit, so one check can be
run again without the others.
"""

from __future__ import annotations

import datetime
import json
import platform
import subprocess
from pathlib import Path
from typing import Any, Dict, List, Optional

ROOT = Path(__file__).resolve().parents[1]
RESULTS = ROOT / "docs" / "excel-oracle.json"
EXCEL_APP = Path("/Applications/Microsoft Excel.app")

ABOUT = ("Checks of ReflexGrid against the Microsoft Excel installed on a Mac (scripts/excel_vba_oracle_mac.py, "
         "scripts/excel_check_mac.py). Each check records the Excel version, the macOS version, the ReflexGrid "
         "commit and the date it ran, how many cases Excel and ReflexGrid answered alike, and every difference.")

# The checks in the order the report lists them.
ORDER = ("vba-language", "vba-objects", "userforms", "userform-events", "userform-designer", "files")

# What the checks' results depend on: when one of these changes after a check ran, its results may not hold
# any more (``compat_report.py`` says so).
INPUTS = ("gridcore", "tests/vba_excel_cases.py", "tests/vba_object_cases.py", "tests/vba_form_cases.py",
          "scripts/excel_vba_oracle_mac.py", "scripts/excel_check_mac.py")


def excel_version() -> str:
    def read(key: str) -> str:
        p = subprocess.run(["defaults", "read", str(EXCEL_APP / "Contents" / "Info"), key], capture_output=True, text=True)
        return p.stdout.strip()

    short, build = read("CFBundleShortVersionString"), read("CFBundleVersion")
    return f"{short} (build {build.split('.')[-1]})" if short and build else short or "unknown"


def commit() -> str:
    """The commit the check ran on, marked when what it depends on (``INPUTS``) had uncommitted changes."""

    def git(*args: str) -> str:
        return subprocess.run(["git", *args], cwd=ROOT, capture_output=True, text=True).stdout.strip()

    sha = git("rev-parse", "HEAD")
    if sha and git("status", "--porcelain", "--", *INPUTS):
        sha += " with uncommitted changes"
    return sha


def record(path: Optional[str], key: str, title: str, compared: int, matched: int,
           differences: List[Dict[str, Any]], **extra: Any) -> None:
    """Puts one check's results into ``path`` (nothing when ``path`` is empty)."""
    if not path:
        return
    p = Path(path)
    data = json.loads(p.read_text("utf-8")) if p.exists() else {}
    checks = dict(data.get("checks", {}))
    checks[key] = {
        "title": title,
        "compared": compared,
        "matched": matched,
        "differences": differences,
        **extra,
        "excel": f"Microsoft Excel for Mac {excel_version()}",
        "os": f"macOS {platform.mac_ver()[0]}",
        "reflexgrid_commit": commit(),
        "date": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d"),
    }
    ordered = {k: checks[k] for k in ORDER if k in checks}
    ordered.update({k: v for k, v in checks.items() if k not in ordered})
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps({"about": ABOUT, "checks": ordered}, indent=1, ensure_ascii=False) + "\n", "utf-8")
    print(f"recorded {key} in {p}")
