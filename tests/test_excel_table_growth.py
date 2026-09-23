"""A table growing on its own, replayed against the in-memory model.

tests/fixtures/tables/growth.json is what scripts/measure_table_growth.py
saw in live Excel: Table1 over A1:C4, formulas in column H naming parts
of it by address, and one thing written next to it -- a value, a
formula, a block, a copy -- after which the table's range and those
formulas were read back.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / "fixtures" / "tables" / "growth.json").read_text(encoding="utf-8"))

#: Cases the model answers differently, and why.
KNOWN: dict[str, str] = {
    "copy_below": "a copy pasted under a table puts a table row in first, the cells under it moving down, which "
                  "the model does not do",
    "fill_down": "FillDown into the row under a table puts a table row in first, which the model does not do",
}


@pytest.mark.parametrize("case", [
    pytest.param(case, id=case["name"],
                 marks=[pytest.mark.xfail(reason=KNOWN[case["name"]], strict=True)] if case["name"] in KNOWN else [])
    for case in RECORD["cases"]])
def test_a_table_grows_as_excels_does(case: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    lines = [RECORD["reader"], "Public Function Probe() As String", "Dim ws As Object, t As Object",
             "Set ws = ActiveWorkbook.Worksheets(1)", RECORD["setup"]]
    lines += [f'ws.Cells({row}, 8).Formula = "{formula}"' for row, formula in enumerate(RECORD["references"], 1)]
    lines += ["On Error Resume Next", case["writing"], 'Probe = Err.Number & "~;~" & Read(ws)', "End Function"]
    app.add_module("\n".join(lines) + "\n", name="Probe")
    assert str(app.run("Probe")) == case["answer"]
