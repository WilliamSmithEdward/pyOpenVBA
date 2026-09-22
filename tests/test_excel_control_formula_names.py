"""Replay reference-formula control bindings measured in Excel."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORDS = json.loads((Path(__file__).parent / "fixtures/shapes/control_formula_names.json").read_text())


@pytest.mark.parametrize("record", RECORDS, ids=[r["name"] for r in RECORDS])
def test_measured_formula_control_names(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim sh As Object, n As Long\n'
                   + str(record["body"]) + 'End Function\n', name="Probe")
    assert app.run("Report") == record["reported"]
