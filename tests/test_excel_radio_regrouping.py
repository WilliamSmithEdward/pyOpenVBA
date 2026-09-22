"""Measured group-box creation and deletion around existing radios."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORDS = json.loads((Path(__file__).parent / "fixtures/shapes/radio_regrouping.json").read_text())


@pytest.mark.parametrize("record", RECORDS,
                         ids=[f'{r["operation"]}-{r["chosen"]}-{r["outside"]}' for r in RECORDS])
def test_measured_radio_regrouping(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim a As Object, b As Object, c As Object, box As Object\n'
                   + str(record["body"]) + 'End Function\n', name="Probe")
    assert app.run("Report") == record["reported"]
