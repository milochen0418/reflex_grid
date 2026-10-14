"""The Excel object model cases of ``vba_object_cases.py``: each runs in ReflexGrid without an
internal error, and gives the answers Microsoft Excel gave (``vba_object_expected.json``, captured
with ``scripts/excel_vba_oracle_mac.py --objects --save``; dates as the US locale writes them).

``KNOWN`` lists where ReflexGrid knowingly differs from Excel, with the reason (also in
docs/vba.md)."""

import json
from pathlib import Path

import pytest

from gridcore.engine import Session
from tests.vba_object_cases import build_module

EXPECTED_PATH = Path(__file__).parent / "vba_object_expected.json"
EXPECTED = json.loads(EXPECTED_PATH.read_text("utf-8")) if EXPECTED_PATH.exists() else {}

# Case name -> why ReflexGrid's answer differs from Excel's.
_WIDTH = "Range.Text of General numbers (and ####) depends on the column's width in Excel; ReflexGrid shows up to 11 characters"
_MAC = "Excel for Mac measures columns in its own points (default width 10, 65 pt); ReflexGrid uses Excel for Windows' (8.43, 48 pt)"
KNOWN = {
    "date and time value": _WIDTH,
    "long digit string": _WIDTH,
    "general text of numbers": _WIDTH,
    "positions": _MAC,
    "ColumnWidth": _MAC + "; for columns of different widths ReflexGrid answers Null, as Microsoft documents",
    "RowHeight": "for rows of different heights ReflexGrid answers Null, as Microsoft documents (Excel for Mac answers the first row's)",
    "formula normalised": "ReflexGrid doesn't keep the spaces typed inside a formula (= 1 + 2 becomes =1+2)",
    "SpecialCells": "the same cells, split into different areas",
}


@pytest.fixture(scope="module")
def results():
    code, names, this_workbook = build_module()
    s = Session()
    s.new_workbook()
    s.run("macro.setModule", name="Module1", code=code)
    s.run("macro.setModule", name="ThisWorkbook", code=this_workbook)
    assert s.run("macro.list")["diagnostics"] == []
    r = s.run("macro.run", name="RunAll", timeLimit=600)
    assert r["status"] == "done", r.get("error")
    out = {}
    for i, name in enumerate(names, 1):
        a = s.run("cell.get", cell=f"'Results'!A{i}")["value"] or ""
        b = s.run("cell.get", cell=f"'Results'!B{i}")["value"] or ""
        out[name] = [a, b]
    return out


def test_no_case_hits_an_internal_error(results):
    broken = {n: r[0] for n, r in results.items() if "E51;" in r[0] or "?51" in r[0] or "?51" in r[1]}
    assert not broken, broken


@pytest.mark.skipif(not EXPECTED, reason="Excel's answers not captured yet (scripts/excel_vba_oracle_mac.py --objects --save)")
@pytest.mark.parametrize("case", list(EXPECTED) or ["-"])
def test_case_matches_excel(results, case):
    if case in KNOWN:
        pytest.xfail(KNOWN[case])
    assert [x.rstrip() for x in results[case]] == EXPECTED[case]
