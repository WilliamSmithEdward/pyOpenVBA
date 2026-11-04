"""The effect members the checklist had as UNASSESSED, replayed as Excel answered them
(scripts/measure_unassessed_effects.py)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

MEASURED = json.loads((Path(__file__).parent / "fixtures/unassessed_effects.json").read_text(encoding="utf-8"))
#: Measured, and left reporting themselves: each preset shadow sets its own blur and offsets and only two were read, and
#: Shape.ThreeD is a feature of its own.
REPORTED = {"shadow_type", "shadow_type_2", "shadow_type_bad", "three_shape"}


def _run(row: dict[str, str]) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MEASURED["probe"].format(action=row["action"]), name="Measure")
    return app.run("Probe")


@pytest.mark.parametrize("row", [row for row in MEASURED["rows"] if row["name"] not in REPORTED],
                         ids=lambda row: row["name"])
def test_unassessed_effects(row: dict[str, str]) -> None:
    assert _run(row) == row["result"]


@pytest.mark.parametrize("row", [row for row in MEASURED["rows"] if row["name"] in REPORTED], ids=lambda row: row["name"])
def test_what_was_left_out_reports_itself(row: dict[str, str]) -> None:
    with pytest.raises(VBAUnsupportedError):
        _run(row)
