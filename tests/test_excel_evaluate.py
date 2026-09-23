"""Evaluate, the bracket form and Worksheet.Evaluate, replayed against the in-memory model.

tests/fixtures/evaluate.json is what scripts/measure_evaluate.py saw in
live Excel: each case hands one expression to Application.Evaluate, to
[...] or to Worksheet.Evaluate on a sheet that is not active, and
records the Range, array, error value or value that came back. The model
runs the same code in the same order.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "evaluate.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]
BATCH = 30

#: Cases the model answers differently, and why.
KNOWN: dict[tuple[str, str], str] = {}


def _call(how: str, expression: str) -> str:
    quoted = expression.replace('"', '""')
    if how == "Application":
        return f'Application.Evaluate("{quoted}")'
    if how == "Other":
        return f'wb.Worksheets("Other").Evaluate("{quoted}")'
    return f"[{expression}]"


def _module(cases: list[dict[str, Any]]) -> str:
    batches: list[str] = []
    for start in range(0, len(cases), BATCH):
        lines = [f"Private Function Batch{start // BATCH}(wb As Object) As String", "Dim out As String",
                 "On Error Resume Next"]
        for case in cases[start:start + BATCH]:
            lines += ["Err.Clear", f'out = out & Show({_call(case["how"], case["expression"])}) & "^"',
                      'If Err.Number <> 0 Then out = out & "E" & Err.Number & "^"', 'out = out & "|"']
        lines += ["On Error GoTo 0", f"Batch{start // BATCH} = out", "End Function"]
        batches.append("\n".join(lines))
    probe = ["Public Function Probe() As String", "Dim wb As Object, out As String",
             "Set wb = ActiveWorkbook", "Fill wb",
             *[f"out = out & Batch{index}(wb)" for index in range(len(batches))], "Probe = out", "End Function"]
    return "\n".join([RECORD["helper"], RECORD["setup"], *batches, *probe]) + "\n"


def _key(case: dict[str, Any]) -> tuple[str, str]:
    return case["how"], case["expression"]


@pytest.fixture(scope="module")
def answers() -> dict[tuple[str, str], str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_module(CASES), name="Probe")
    got = [part.split("^")[0] for part in str(app.run("Probe")).split("|")[: len(CASES)]]
    return {_key(case): answer for case, answer in zip(CASES, got, strict=True)}


def _param(index: int, case: dict[str, Any]) -> Any:
    key = _key(case)
    marks = [pytest.mark.xfail(reason=KNOWN[key], strict=True)] if key in KNOWN else []
    return pytest.param(index, marks=marks, id=f"{case['how']}-{case['expression'][:40]}")


@pytest.mark.parametrize("index", [_param(index, case) for index, case in enumerate(CASES)])
def test_evaluate_answers_as_excels_does(answers: dict[tuple[str, str], str], index: int) -> None:
    case = CASES[index]
    assert answers[_key(case)] == case["answer"]
