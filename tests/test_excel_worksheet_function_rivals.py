"""The rival rules for the wrong answers WorksheetFunction rounds found, replayed against the in-memory model.

tests/fixtures/worksheet_function_rivals.json holds seven probes, each case's rival predictions written before live Excel
answered. Round 60101's: an argument left empty or a blank cell (ZTEST's sigma, IFERROR's and IFNA's answer, BAHTTEXT),
FREQUENCY's bins holding no number, a database function's field of several cells, a range SUMIFS and its kin get with no
criterion after it, AMORDEGRC over every period of several lives, LCM and GCD given text or logicals, TREND and its kin
given x of another shape, and Transpose of one cell. Round 60102's: Unichar(True), a constant under a database criteria
label that names no field, and LINEST's switches given a range. Round 60103's: DB's month, Transpose(Empty), and a
lookup's range_lookup given a cell of text. The model runs each probe as Excel ran it, its cases in
order in one macro; where the model refuses a case, case by case, each after every setup before it.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures"
RECORD: dict[str, Any] = json.loads((FIXTURES / "worksheet_function_rivals.json").read_text(encoding="utf-8"))
MORE: dict[str, Any] = json.loads((FIXTURES / "worksheet_functions_more.json").read_text(encoding="utf-8"))
SEPARATOR = "|^|"
PROBES = sorted({case["probe"] for case in RECORD["cases"]})
REFUSED = "refused"

#: Cases the model answers differently, and why.
KNOWN: dict[str, str] = {
    "forecast row y column x": "WorksheetFunction.Forecast is not in the model",
    "linest y column x two columns": "with x columns that are one another plus 1, Excel keeps the second and the model the first",
}
#: Cases the model refuses as not implemented: a range of several cells for a regression function's switch from a macro,
#: which Excel mostly reads by its first cell, yet not always.
REFUSALS: frozenset[str] = frozenset({
    "linest stats two cells", "linest stats zero first", "linest stats one first", "linest const zero first",
    "linest stats blank cells", "logest stats zero first", "trend const zero first", "linest stats text cells",
    "linest stats row 1 across", "linest stats row 1 off column", "linest stats column 1 off row", "linest stats block",
    "linest const column 1 off row", "trend const off row", "growth const off row", "logest stats off row",
})


def _module(cases: list[dict[str, Any]], typed: bool, *, setups: list[str] | None = None) -> str:
    kind = "WorksheetFunction" if typed else "Object"
    lines: list[str] = [f"    {setup}" for setup in setups or []]
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


def _run(source: str) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name="Probe")
    return str(app.run("Probe.One"))


def _answers(probe: int, typed: bool) -> dict[str, str]:
    cases = [case for case in RECORD["cases"] if case["probe"] == probe]
    try:
        found = _run(_module(cases, typed)).split(SEPARATOR)[:len(cases)]
    except VBAUnsupportedError:
        found: list[str] = []
        for at, case in enumerate(cases):
            setups = [one["setup"] for one in cases[:at] if one["setup"] and not one["setup"].startswith('ws.Range("M1")')]
            try:
                found.append(_run(_module([case], typed, setups=setups)).split(SEPARATOR)[0])
            except VBAUnsupportedError:
                found.append(REFUSED)
    return {case["case"]: answer for case, answer in zip(cases, found, strict=True)}


@pytest.mark.parametrize("typed", [False, True], ids=["late", "typed"])
@pytest.mark.parametrize("probe", PROBES)
def test_the_model_answers_as_excel_did(probe: int, typed: bool) -> None:
    tag = "typed" if typed else "late"
    answers = _answers(probe, typed)
    wrong = [(case["case"], case[tag], answers[case["case"]]) for case in RECORD["cases"]
             if case["probe"] == probe and case["case"] not in KNOWN and case["case"] not in REFUSALS
             and answers[case["case"]] != case[tag]]
    assert not wrong
    assert sorted(name for name, answer in answers.items() if answer == REFUSED) == sorted(
        case["case"] for case in RECORD["cases"] if case["probe"] == probe and case["case"] in REFUSALS)


def test_the_known_differences_still_differ() -> None:
    answers = {name: answer for probe in PROBES for name, answer in _answers(probe, False).items()}
    expected = {case["case"]: case["late"] for case in RECORD["cases"]}
    assert [name for name in KNOWN if answers[name] == expected[name]] == []


def test_every_case_names_its_rivals() -> None:
    assert all(case["rivals"] for case in RECORD["cases"])
    assert set(KNOWN) | REFUSALS <= {case["case"] for case in RECORD["cases"]}
