"""Structural edits measured against native Excel."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORDS = json.loads((Path(__file__).parent / "fixtures/range_edit.json").read_text())


@pytest.mark.parametrize("record", RECORDS, ids=[r["name"] for r in RECORDS])
def test_native_structural_edit(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Report() As String\nDim cell As Object, n As Long\n'
                   + str(record["body"]) + 'End Function', name="Probe")
    assert app.run("Report") == record["reported"]


@pytest.mark.parametrize("name", ["insert_row", "delete_row", "insert_column", "delete_columns"])
def test_structural_edit_round_trip(tmp_path: Path, name: str) -> None:
    record = next(r for r in RECORDS if r["name"] == name)
    app = ExcelApplication()
    app.add_workbook()
    body = record["body"]
    app.add_module('Function Build() As String\nDim cell As Object, n As Long\n'
                   + body.replace("Report", "Build") + 'End Function', name="Builder")
    expected = app.run("Build")
    path = tmp_path / "edited.xlsm"
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False)
    report = body[body.index('Report = CStr(n)'):]
    reopened.add_module('Function Report() As String\nDim cell As Object, n As Long\n'
                        + 'Worksheets("Sheet1").Activate\n' + report + 'End Function', name="Probe")
    assert reopened.run("Report") == expected


def test_qualified_deleted_reference_calculates_and_survives_another_edit() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Function Report() As String
Worksheets.Add.Name = "Other"
Worksheets("Sheet1").Range("A1").Value = 7
Range("B2").Formula = "=Sheet1!A1"
Worksheets("Sheet1").Rows("1:1").Delete
Worksheets("Sheet1").Rows("2:2").Insert
Report = CStr(Range("B2").Value)
End Function''', name="Probe")
    assert app.run("Report") == "Error 2023"


@pytest.mark.parametrize("axis, edge", [("Rows", "A1048576"), ("Columns", "XFD1")])
def test_insertion_at_capacity_is_atomic(axis: str, edge: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'''Function Report() As String
Dim n As Long
Range("{edge}").Value = "kept"
Range("B2").Formula = "=A1"
On Error Resume Next
{axis}(1).Insert
n = Err.Number
On Error GoTo 0
Report = CStr(n) & "|" & Range("{edge}").Value & "|" & Range("B2").Formula
End Function''', name="Probe")
    assert app.run("Report") == "1004|kept|=A1"
