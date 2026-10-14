"""ReflexGrid's VBA gives the answers Excel's VBA gave for the cases in ``vba_excel_cases.py``.

The expected answers were captured from Microsoft Excel (``scripts/excel_vba_oracle_mac.py``;
dates as the US locale writes them); cases Excel for Mac can't run (CDec, Scripting.Dictionary)
follow Excel for Windows."""

import json
from pathlib import Path

import pytest

from gridcore.vba.compiler import compile_project
from gridcore.vba.runtime import Execution
from tests.vba_excel_cases import CLASS1, build_module

EXPECTED = json.loads((Path(__file__).parent / "vba_excel_expected.json").read_text("utf-8"))


@pytest.fixture(scope="module")
def results():
    code, names = build_module(windows=True)
    code = code.split("Sub WriteAll()")[0]
    rt = compile_project([("Module1", "standard", code, ()), ("Class1", "class", CLASS1, ())])
    assert not rt.errors, rt.errors
    ex = Execution(rt, time_limit=120)
    res = rt.run(ex, "RunAll")
    return dict(zip(names, [str(x) for x in res.data]))


def test_every_case_has_an_expectation(results):
    assert set(results) == set(EXPECTED)


@pytest.mark.parametrize("case", list(EXPECTED))
def test_case_matches_excel(results, case):
    assert results[case] == EXPECTED[case]
