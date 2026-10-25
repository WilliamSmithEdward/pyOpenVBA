"""A pivot table's standard deviations and variances, replayed against the in-memory model to the last bit.

tests/fixtures/pivot_variance/ holds what scripts/measure_pivot_variance.py
saw in live Excel: one pivot table showing each group's StdDev, StdDevP, Var
and VarP over groups of values chosen to set the ways of working a variance
out apart, large values a little apart, decimals, tiny values, values all the
same and whole numbers, with the grand total of them all, and the workbook
Excel saved, whose cells hold each value to the bit.
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication
from test_excel_pivot_layouts import sheet_cells

FIXTURES = Path(__file__).parent / "fixtures" / "pivot_variance"
RECORD: dict[str, Any] = json.loads((FIXTURES / "pivot_variance.json").read_text(encoding="utf-8"))


def test_every_deviation_and_variance_is_excels_to_the_bit(tmp_path: Path) -> None:
    """A sample's from the sum of squares less the square of the sum over the count, a whole population's from the
    count times the sum of squares less the square of the sum: 100000001, 100000002 and 100000003 come to 0, as in
    Excel, where two passes would keep their spread."""
    rows = ", ".join(f'"{group}|{value}"' for group, values in RECORD["groups"].items() for value in values)
    # The macro Excel ran names its file under a Windows folder; here it goes under this system's.
    code = RECORD["build"].replace("{rows}", rows).replace("{folder}\\", str(tmp_path) + os.sep)
    app = ExcelApplication()
    app.add_module(code, name="Builder")
    assert app.run("Build") == "done"
    assert sheet_cells(tmp_path / "pivot_variance.xlsx", {"data"}) == sheet_cells(FIXTURES / "pivot_variance.xlsx",
                                                                                 {"data"})
