"""The chart members the checklist had as UNASSESSED, replayed as Excel answered them
(scripts/measure_unassessed_charts.py)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

MEASURED = json.loads((Path(__file__).parent / "fixtures/unassessed_charts.json").read_text(encoding="utf-8"))
#: Measured, and left reporting themselves: filtering a category, 3-D chart types and a secondary axis, which the model
#: does not have, and InvertIfNegative through a legend key on a chart that is no column or bar.
REPORTED = {"category_filtered", "walls", "axis_secondary", "key_invert_set", "key_invert_series"}


def _run(row: dict[str, str]) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MEASURED["probe"].format(action=row["action"]), name="Measure")
    return app.run("Probe")


@pytest.mark.parametrize("row", [row for row in MEASURED["rows"] if row["name"] not in REPORTED],
                         ids=lambda row: row["name"])
def test_unassessed_charts(row: dict[str, str]) -> None:
    assert _run(row) == row["result"]


@pytest.mark.parametrize("row", [row for row in MEASURED["rows"] if row["name"] in REPORTED], ids=lambda row: row["name"])
def test_what_was_left_out_reports_itself(row: dict[str, str]) -> None:
    with pytest.raises(VBAUnsupportedError):
        _run(row)
