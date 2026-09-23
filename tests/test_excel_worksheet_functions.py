"""Worksheet functions called from VBA, replayed against the in-memory model.

tests/fixtures/worksheet_functions.json is what
scripts/measure_worksheet_functions.py saw in live Excel: each case calls
one function through WorksheetFunction and through Application, with
numbers, text, Booleans, Dates, errors, VBA arrays and ranges of a small
sheet, and records the type and value of each answer or the error. The
model runs the same calls in the same order.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / "fixtures" / "worksheet_functions.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]
BATCH = 40

#: Calls the model answers differently, and why.
KNOWN: dict[str, str] = {}


def _module() -> str:
    batches: list[str] = []
    for start in range(0, len(CASES), BATCH):
        lines = [f"Private Function Batch{start // BATCH}(ws As Object) As String", "Dim out As String, v As Variant",
                 "On Error Resume Next"]
        for case in CASES[start:start + BATCH]:
            call = f"({case['arguments']})" if case["arguments"] else ""
            for target in (f"Application.WorksheetFunction.{case['function']}{call}",
                           f"Application.{case['function']}{call}"):
                lines += ["Err.Clear", "v = Empty", f"v = {target}",
                          'If Err.Number <> 0 Then out = out & "E" & Err.Number & "^" Else out = out & Show(v) & "^"']
            lines.append('out = out & "|"')
        lines += ["On Error GoTo 0", f"Batch{start // BATCH} = out", "End Function"]
        batches.append("\n".join(lines))
    probe = ["Public Function Probe() As String", "Dim ws As Object, out As String",
             "Set ws = ActiveWorkbook.Worksheets(1)", "Fill ws",
             *[f"out = out & Batch{index}(ws)" for index in range(len(batches))], "Probe = out", "End Function"]
    return "\n".join([RECORD["helper"], RECORD["setup"], *batches, *probe]) + "\n"


@pytest.fixture(scope="module")
def answers() -> list[list[str]]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_module(), name="Probe")
    return [part.split("^")[:2] for part in str(app.run("Probe")).split("|")[: len(CASES)]]


def _label(case: dict[str, Any]) -> str:
    return f"{case['function']}({case['arguments']})"


@pytest.mark.parametrize("index", [
    pytest.param(index, marks=pytest.mark.xfail(reason=KNOWN[_label(case)], strict=True))
    if _label(case) in KNOWN else index for index, case in enumerate(CASES)], ids=[_label(case) for case in CASES])
def test_a_worksheet_function_answers_as_excels_does(answers: list[list[str]], index: int) -> None:
    case = CASES[index]
    assert answers[index] == [case["worksheet_function"], case["application"]]
