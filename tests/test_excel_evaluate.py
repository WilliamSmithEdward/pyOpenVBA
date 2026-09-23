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
KNOWN = {
    ("Application", "ROW()"): "ROW() with no argument is one number to the engine, a one-item array to Evaluate",
    ("Application", "COLUMN()"): "COLUMN() with no argument is one number to the engine, a one-item array to Evaluate",
    ("Other", "ROW()"): "ROW() with no argument is one number to the engine, a one-item array to Evaluate",
    ("Application", 'IF(A1:A3>1,"big","small")'): "the engine does not run IF once per item of a block",
    ("Application", "UPPER(B1:B3)"): "the engine does not run a function once per item of a block given for a value",
}
#: Cases that report themselves unsupported, and why.
UNSUPPORTED = {
    ("Application", "SUM(A1:A5 A2:B3)"): "an intersection inside a formula",
    ("Application", "XLOOKUP(3,A1:A5,B1:B5)"): "XLOOKUP, which Evaluate hands back as a Range",
}


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


#: The cases the model runs: an unsupported one would stop the rest of its procedure.
RUN = [case for case in CASES if _key(case) not in UNSUPPORTED]


@pytest.fixture(scope="module")
def answers() -> dict[tuple[str, str], str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_module(RUN), name="Probe")
    got = [part.split("^")[0] for part in str(app.run("Probe")).split("|")[: len(RUN)]]
    return {_key(case): answer for case, answer in zip(RUN, got, strict=True)}


def _param(index: int, case: dict[str, Any]) -> Any:
    key = _key(case)
    marks = [pytest.mark.xfail(reason=KNOWN[key], strict=True)] if key in KNOWN else []
    return pytest.param(index, marks=marks, id=f"{case['how']}-{case['expression'][:40]}")


@pytest.mark.parametrize("index", [_param(index, case) for index, case in enumerate(RUN)])
def test_evaluate_answers_as_excels_does(answers: dict[tuple[str, str], str], index: int) -> None:
    case = RUN[index]
    assert answers[_key(case)] == case["answer"]


@pytest.mark.parametrize("key", sorted(UNSUPPORTED))
def test_what_evaluate_cannot_work_out_says_so(key: tuple[str, str]) -> None:
    from pyopenvba.exceptions import VBAUnsupportedError

    case = next(one for one in CASES if _key(one) == key)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_module([case]), name="Probe")
    with pytest.raises(VBAUnsupportedError, match=UNSUPPORTED[key].split(",")[0].split(" ")[0]):
        app.run("Probe")
