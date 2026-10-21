"""Exact native coercion bodies; only version/build environment header excluded."""

import hashlib
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._outline import Outline
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parent / "fixtures/worksheet_outlines"


@pytest.mark.parametrize("name", ["worksheet-outline-coercions.json", "worksheet-outline-coercion-strings.json"])
def test_native_outline_coercion_matrix(name: str) -> None:
    record = json.loads((ROOT / name).read_bytes())
    assert record["completed"] and record["session_closed_utc"]
    assert record["status"]["ok"] and not record["status"]["dialogs"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(record["source"], name="Coercion")
    actual = app.run("Probe")
    assert isinstance(actual, str)
    native_header, _, native_body = record["raw_result"].partition("\n")
    model_header, _, model_body = actual.partition("\n")
    assert native_header.startswith("version=16.0|build=") and model_header.startswith("version=16.0|build=")
    assert model_body == native_body
    provenance = json.loads((ROOT / "provenance.json").read_bytes())
    assert hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == provenance["files"][name]


@pytest.mark.parametrize("property_name,values", [("SummaryColumn", [0, 1, -1, 2]), ("SummaryRow", [-1, 2, 3])])
def test_unmeasured_integer_summary_values_are_atomic(property_name: str, values: list[int]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.workbook.sheets_[0]
    outline = Outline(sheet)
    before = (outline.SummaryRow(), outline.SummaryColumn(), app.workbook.saved)
    for value in values:
        with pytest.raises(VBAUnsupportedError, match="unmeasured integer"):
            outline.vba_set(property_name, value)
        assert (outline.SummaryRow(), outline.SummaryColumn(), app.workbook.saved) == before


def test_summary_column_integer_three_uses_original_native_raw_proof() -> None:
    record = json.loads((ROOT / "initial.json").read_bytes())
    case = next(case for case in record["controls"] if case["name"] == "summary-sides")
    assert 'PutP(ActiveSheet.Outline,"SummaryColumn",3)' in case["source"]
    assert ";badcol=ok:-4152(3)" in case["raw_result"]
    app = ExcelApplication()
    app.add_workbook()
    outline = Outline(app.workbook.sheets_[0])
    outline.vba_set("SummaryColumn", -4131)
    assert outline.SummaryColumn() == -4131
    outline.vba_set("SummaryColumn", 3)
    assert outline.SummaryColumn() == -4152
