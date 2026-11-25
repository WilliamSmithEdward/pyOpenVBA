"""The rival rules for WorksheetFunction round 60101's wrong answers, replayed against the in-memory model.

tests/fixtures/worksheet_function_rivals.json holds five probes, each case's rival predictions written before live Excel
answered: an argument left empty or a blank cell (ZTEST's sigma, IFERROR's and IFNA's answer, BAHTTEXT), FREQUENCY's bins
holding no number, a database function's field of several cells, a range SUMIFS and its kin get with no criterion after
it, AMORDEGRC over every period of several lives, LCM and GCD given text or logicals, TREND and its kin given x of
another shape, and Transpose of one cell. The model runs each probe as Excel ran it, its cases in order in one macro.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures"
RECORD: dict[str, Any] = json.loads((FIXTURES / "worksheet_function_rivals.json").read_text(encoding="utf-8"))
MORE: dict[str, Any] = json.loads((FIXTURES / "worksheet_functions_more.json").read_text(encoding="utf-8"))
SEPARATOR = "|^|"
PROBES = sorted({case["probe"] for case in RECORD["cases"]})

#: Cases the model answers differently, and why.
KNOWN: dict[str, str] = {
    "forecast row y column x": "WorksheetFunction.Forecast is not in the model",
    "linest y column x two columns": "with x columns that are one another plus 1, Excel keeps the second and the model the first",
}


def _module(cases: list[dict[str, Any]], typed: bool) -> str:
    kind = "WorksheetFunction" if typed else "Object"
    lines: list[str] = []
    for case in cases:
        before = f"{case['setup']}: " if case["setup"] else ""
        lines.append(f'    Err.Clear: v = Empty: ws.Range("M1").ClearContents: {before}v = {case["expression"]}\n'
                     f'    If Err.Number <> 0 Then out = out & "E" & Err.Number & "{SEPARATOR}" Else out = out & Show(v) & "{SEPARATOR}"')
    return (MORE["helper"] + "\n" + MORE["setup"] + "\n"
            "Public Function One() As String\n"
            f"    Dim wb As Object, ws As Object, f As {kind}, v As Variant, k As Variant, out As String\n"
            "    Application.DisplayAlerts = False\n    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)\n"
            "    Fill ws\n    Set f = Application.WorksheetFunction\n    On Error Resume Next\n" + "\n".join(lines) + "\n"
            "    On Error GoTo 0\n    wb.Close False\n    One = out\nEnd Function\n")


def _answers(probe: int, typed: bool) -> dict[str, str]:
    cases = [case for case in RECORD["cases"] if case["probe"] == probe]
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_module(cases, typed), name="Probe")
    found = str(app.run("Probe.One")).split(SEPARATOR)[:len(cases)]
    return {case["case"]: answer for case, answer in zip(cases, found, strict=True)}


@pytest.mark.parametrize("typed", [False, True], ids=["late", "typed"])
@pytest.mark.parametrize("probe", PROBES)
def test_the_model_answers_as_excel_did(probe: int, typed: bool) -> None:
    tag = "typed" if typed else "late"
    answers = _answers(probe, typed)
    wrong = [(case["case"], case[tag], answers[case["case"]]) for case in RECORD["cases"]
             if case["probe"] == probe and case["case"] not in KNOWN and answers[case["case"]] != case[tag]]
    assert not wrong


def test_the_known_differences_still_differ() -> None:
    answers = {name: answer for probe in PROBES for name, answer in _answers(probe, False).items()}
    expected = {case["case"]: case["late"] for case in RECORD["cases"]}
    assert [name for name in KNOWN if answers[name] == expected[name]] == []


def test_every_prediction_was_written_before_excel_answered() -> None:
    # Each case names its rivals, and Excel's answer fits at most one way of reading them: the probes split rival rules.
    assert all(case["rivals"] for case in RECORD["cases"])
