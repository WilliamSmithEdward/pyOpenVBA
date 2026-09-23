"""SUM, AVERAGE, the variances and their kin, against Excel's own doubles, bit for bit.

tests/fixtures/variance.json is what scripts/measure_variance.py read from
live Excel: for each set of numbers, typed into A1 down, the eight bytes
each function's cell holds. The model's answer has to be the same double.
"""

from __future__ import annotations

import json
import struct
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "variance.json").read_text(encoding="utf-8"))
FORMULAS: dict[str, str] = RECORD["formulas"]
SETS: list[dict[str, Any]] = RECORD["sets"]

#: Answers the model does not give as Excel does yet.
MISSED: set[tuple[int, str]] = set()


def _answers(values: list[float]) -> dict[str, float]:
    app = ExcelApplication()
    app.add_workbook()
    code = ["Public Sub Fill()", "Dim ws As Object", "Set ws = ActiveWorkbook.Worksheets(1)"]
    code += [f"ws.Cells({row}, 1).Value = {value!r}" for row, value in enumerate(values, 1)]
    for column, name in enumerate(FORMULAS, 3):
        formula = FORMULAS[name].format(r=f"A1:A{len(values)}")
        code.append(f'ws.Cells(1, {column}).Formula = "={formula}"')
    app.add_module("\n".join([*code, "End Sub"]) + "\n", name="Fill")
    app.run("Fill")
    out: dict[str, float] = {}
    for column, name in enumerate(FORMULAS, 3):
        value = app.evaluate(f"ActiveWorkbook.Worksheets(1).Cells(1, {column}).Value")
        assert isinstance(value, float), (name, value)
        out[name] = value
    return out


def _cases() -> list[Any]:
    cases: list[Any] = []
    for index, one in enumerate(SETS):
        for name in FORMULAS:
            if one["answers"][name].startswith("E"):
                continue
            marks = [pytest.mark.xfail(reason="Excel's variance rule for data under a thousandth", strict=True)] \
                if (index, name) in MISSED else []
            cases.append(pytest.param(index, name, marks=marks, id=f"{index}-{name}"))
    return cases


_CACHE: dict[int, dict[str, float]] = {}


@pytest.mark.parametrize(("index", "name"), _cases())
def test_the_answer_is_excels_double(index: int, name: str) -> None:
    if index not in _CACHE:
        _CACHE[index] = _answers(SETS[index]["stored"])
    got = struct.pack(">d", _CACHE[index][name]).hex().upper()
    assert got == SETS[index]["answers"][name]


def test_the_numbers_are_stored_as_typed() -> None:
    for one in SETS:
        assert one["stored"] == [float(value) for value in one["typed"]]
