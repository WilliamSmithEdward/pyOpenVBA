"""Formulas that work with references, replayed against the in-memory model.

tests/fixtures/reference_forms.json is what
scripts/measure_reference_forms.py saw in live Excel: the range operator
between references, intersections, unions, and INDEX, OFFSET, INDIRECT,
CHOOSE and IF where a reference is wanted, each written to a cell with
Range.Formula, with what the cell holds and how Excel spells it back.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / "fixtures" / "reference_forms.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]

ITEM = '''Private Function Item(v As Variant) As String
    If IsError(v) Then
        Item = "Error:" & CStr(CLng(v))
    ElseIf IsEmpty(v) Then
        Item = "Empty"
    Else
        Item = TypeName(v) & ":" & CStr(v)
    End If
End Function
'''


@pytest.fixture(scope="module")
def answers() -> list[list[str]]:
    lines = ["Public Function Probe() As String", "Dim ws As Object, out As String, r As Long",
             "Set ws = ActiveWorkbook.Worksheets(1)", "Fill ws"]
    for row, case in enumerate(CASES, 1):
        lines.append(f'ws.Range("E{row}").Formula = "{case["formula"].replace(chr(34), chr(34) * 2)}"')
    lines += [f"For r = 1 To {len(CASES)}",
              'out = out & ws.Cells(r, 5).Formula & "^" & Item(ws.Cells(r, 5).Value) & "|"', "Next r",
              "Probe = out", "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("\n".join([ITEM, RECORD["setup"], *lines]) + "\n", name="Probe")
    return [part.split("^") for part in str(app.run("Probe")).split("|")[: len(CASES)]]


@pytest.mark.parametrize("index", range(len(CASES)), ids=[case["formula"] for case in CASES])
def test_the_cell_holds_what_excels_does(answers: list[list[str]], index: int) -> None:
    case = CASES[index]
    assert answers[index] == [case["spelled"], case["value"]]
