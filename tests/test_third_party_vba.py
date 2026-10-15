"""Third-party VBA in ReflexGrid, checked by its own tests: the spec workbook of VBA-JSON
(``tests/third_party/vba-json``, MIT), written for Microsoft Excel by its authors. ReflexGrid opens the
published ``.xlsm``, runs its VBA-TDD suite, and each spec becomes a test here; these cases come from
outside ReflexGrid, so they don't share the assumptions of its own tests.

``KNOWN`` lists the specs that fail in ReflexGrid and why; one that starts passing must leave the list.
"""

import re
from pathlib import Path

import pytest

from gridcore.engine import Session

WORKBOOK = Path(__file__).parent / "third_party" / "vba-json" / "VBA-JSON-Specs.xlsm"

# Runs the workbook's suite and writes each spec's description, result and failed expectations to a sheet.
DRIVER = """
Public Sub ReflexGridSpecs()
    Dim Suite As SpecSuite, Spec As SpecDefinition, Ex As SpecExpectation, Row As Long, Msg As String
    Set Suite = Specs.Specs()
    Row = 1
    For Each Spec In Suite.SpecsCol
        Msg = ""
        For Each Ex In Spec.FailedExpectations
            Msg = Msg & Ex.FailureMessage & " | "
        Next
        Worksheets("Results").Cells(Row, 1).Value = Spec.Description
        Worksheets("Results").Cells(Row, 2).Value = Spec.ResultName
        Worksheets("Results").Cells(Row, 3).Value = Msg
        Row = Row + 1
    Next
End Sub
"""

KNOWN = {
    "should convert dates to ISO 8601": "VBA-JSON converts local time to UTC with Windows API calls "
    "(GetTimeZoneInformation in kernel32), which ReflexGrid's macro sandbox doesn't allow (error 445)",
}


def _session() -> Session:
    s = Session()
    s.run("file.open", path=str(WORKBOOK))
    return s


def _spec_names():
    code = _session().run("macro.getModule", name="Specs")["code"]
    return re.findall(r'\.It\("((?:[^"]|"")*)"\)', code)


SPECS = _spec_names()


@pytest.fixture(scope="module")
def results():
    s = _session()
    s.run("home.insertSheet", name="Results")
    s.run("macro.addModule", name="ReflexGridSpecs", code=DRIVER)
    r = s.run("macro.run", name="ReflexGridSpecs", timeLimit=120)
    assert r["status"] == "done", r.get("error")
    out = []
    for i in range(1, len(SPECS) + 2):
        desc = s.run("cell.get", cell=f"'Results'!A{i}")["value"]
        if not desc:
            break
        out.append((desc, s.run("cell.get", cell=f"'Results'!B{i}")["value"], s.run("cell.get", cell=f"'Results'!C{i}")["value"] or ""))
    return out


def test_the_workbook_compiles_and_runs_every_spec(results):
    assert _session().run("macro.list")["diagnostics"] == []
    assert [d for d, _, _ in results] == SPECS


@pytest.mark.parametrize("index", range(len(SPECS)), ids=SPECS)
def test_spec_passes(results, index):
    desc, result, failures = results[index]
    if desc in KNOWN:
        if result == "Pass":
            pytest.fail(f"passes now: take {desc!r} out of KNOWN")
        pytest.xfail(KNOWN[desc])
    assert result == "Pass", failures
