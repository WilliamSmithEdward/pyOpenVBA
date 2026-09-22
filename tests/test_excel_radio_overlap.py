"""Measured overlapping and nested Forms group-box behavior."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORDS = json.loads((Path(__file__).parent / "fixtures/shapes/radio_overlap.json").read_text())


@pytest.mark.parametrize("record", RECORDS,
                         ids=[f'{r["layout"]}-{r["reverse"]}-{r["operation"]}' for r in RECORDS])
def test_measured_radio_overlap(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim sh As Object, box1 As Object, box2 As Object\n'
                   + str(record["body"]) + 'End Function\n', name="Probe")
    assert app.run("Report") == record["reported"]
