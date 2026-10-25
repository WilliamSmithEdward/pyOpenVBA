"""Range.PivotCell, PivotTable, PivotField and PivotItem over every cell of a pivot table, replayed against the
in-memory model.

tests/fixtures/pivot_cells/ holds what scripts/measure_pivot_cells.py saw in
live Excel: for each workbook of pivot tables Excel saved for the other
fixtures, every cell of each table's TableRange2, its PivotCell's type,
table, field, item, data field, row and column items, custom subtotal
function and range, and the range's own field, item and table, an error's
number where one is raised; with the table's whole TableRange1 and the cell
just right of the table. The model opens the same workbooks and runs the
macro Excel ran.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures"
RECORDS = sorted((FIXTURES / "pivot_cells").glob("*.json"))


def dumped(record: dict[str, Any], path: Path) -> dict[str, list[str]]:
    """Each table's cells as the macro Excel ran dumps them in the model, from the workbook at ``path``."""
    app = ExcelApplication()
    app.add_module(record["dump"], name="PivotCells")
    text = str(app.run("Dump", str(path)))
    return {key: cells.split("~") for key, cells in (one.split("^", 1) for one in text.split("`") if "^" in one)}


def same(found: dict[str, list[str]], record: dict[str, Any]) -> None:
    # Excel lists a sheet's tables after an open in an order that turns from one open to the next (see
    # tests/test_excel_pivot_shared_caches.py), so the tables are matched by name.
    assert sorted(found) == sorted(record["tables"])
    for key, cells in record["tables"].items():
        assert found[key] == cells, key


@pytest.mark.parametrize("path", RECORDS, ids=lambda path: path.stem)
def test_every_cell_answers_as_excels_does(path: Path) -> None:
    """Values, items, subtotals and grand totals in each layout, the header's field names and buttons, the page
    fields, and the cells beside the table, which are no table's."""
    record: dict[str, Any] = json.loads(path.read_text(encoding="utf-8"))
    same(dumped(record, FIXTURES / record["file"]), record)


def test_the_models_own_layouts_answer_as_excels_do(tmp_path: Path) -> None:
    """Every layout of tests/fixtures/pivot_layouts the model makes, saved: the model's own definitions of the
    tables give each cell Excel's answer."""
    from test_excel_pivot_layouts import UNSUPPORTED, built

    built(tmp_path).run("Build")
    record: dict[str, Any] = json.loads((FIXTURES / "pivot_cells" / "pivot_layouts.json").read_text(encoding="utf-8"))
    record["tables"] = {key: cells for key, cells in record["tables"].items()
                        if key.partition("!")[0] not in UNSUPPORTED}
    same(dumped(record, tmp_path / "pivot_layouts.xlsx"), record)


def test_a_cell_outside_every_pivot_table_has_no_pivot_cell() -> None:
    """Error 1004 from Range.PivotCell, PivotTable, PivotField and PivotItem where no table is."""
    app = ExcelApplication()
    app.add_module(
        "Public Function Run() As String\nDim out As String, name As Variant, x As Object\nOn Error Resume Next\n"
        'For Each name In Array("PivotCell", "PivotTable", "PivotField", "PivotItem")\nErr.Clear\n'
        'Set x = CallByName(Range("C3"), name, VbGet)\nout = out & name & "=" & Err.Number & ";"\nNext\n'
        "Run = out\nEnd Function\n", name="Outside")
    assert app.run("Run") == "PivotCell=1004;PivotTable=1004;PivotField=1004;PivotItem=1004;"
