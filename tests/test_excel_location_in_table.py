"""Range.LocationInTable over every cell of a pivot table, replayed against the in-memory model.

tests/fixtures/location_in_table/ holds what scripts/measure_location_in_table.py saw in live Excel: for each
workbook of pivot tables the pivot fixtures saved, every cell of each table's TableRange2, the table's whole
TableRange1 and the cell just right of the table, its LocationInTable and its PivotCell's type, an error's number
where one is raised. The model opens the same workbooks and runs the macro Excel ran.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures"
RECORDS = sorted((FIXTURES / "location_in_table").glob("*.json"))


@pytest.mark.parametrize("path", RECORDS, ids=lambda path: path.stem)
def test_every_cell_answers_as_excels_does(path: Path) -> None:
    """Values, items, subtotals, grand totals, headers, buttons, page fields and blanks, the whole table, and the
    cells beside it."""
    record: dict[str, Any] = json.loads(path.read_text(encoding="utf-8"))
    app = ExcelApplication()
    app.add_module(record["dump"], name="Locations")
    text = str(app.run("Dump", str(FIXTURES / record["file"])))
    found = {key: cells.split("~") for key, cells in (one.split("^", 1) for one in text.split("`") if "^" in one)}
    # Excel lists a sheet's tables after an open in an order that turns from one open to the next, so the tables
    # are matched by name (see tests/test_excel_pivot_cells.py).
    assert sorted(found) == sorted(record["tables"])
    for key, cells in record["tables"].items():
        assert found[key] == cells, key
