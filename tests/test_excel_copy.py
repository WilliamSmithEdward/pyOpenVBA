"""Copy overlap, geometry and blank-cell behavior measured in Excel."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORDS = json.loads((Path(__file__).parent / "fixtures/range_copy.json").read_text())


@pytest.mark.parametrize("record", RECORDS, ids=[r["name"] for r in RECORDS])
def test_measured_copy(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Report() As String\nDim cell As Object, n As Long\n'
                   + str(record["body"]) + 'End Function', name="Probe")
    assert app.run("Report") == record["reported"]


@pytest.mark.parametrize("name", ["overlap_down", "tile", "blank_tail", "formula", "format"])
def test_copy_round_trip(tmp_path: Path, name: str) -> None:
    record = next(r for r in RECORDS if r["name"] == name)
    app = ExcelApplication()
    app.add_workbook()
    body = record["body"]
    app.add_module('Function Build() As String\nDim cell As Object, n As Long\n'
                   + body.replace("Report", "Build") + 'End Function', name="Builder")
    expected = app.run("Build")
    path = tmp_path / "copied.xlsm"
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False)
    report = body[body.index('Report = CStr(n)'):]
    reopened.add_module('Function Report() As String\nDim cell As Object, n As Long\n'
                        + report + 'End Function', name="Probe")
    assert reopened.run("Report") == expected


def test_bold_blank_copy_and_removal_round_trip(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Sub Build()
Range("A1").Font.Bold = True
Range("A1").Copy Range("B2")
End Sub''', name="Builder")
    app.run("Build")
    path = tmp_path / "bold.xlsm"
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False)
    report = '''Function Report() As String
Report = CStr(Range("A1").Font.Bold) & "|" & CStr(Range("B2").Font.Bold)
End Function'''
    reopened.add_module(report + '\nSub Unbold()\nRange("B2").Font.Bold = False\nEnd Sub', name="Probe")
    assert reopened.run("Report") == "True|True"
    reopened.run("Unbold")
    reopened.save(path)
    final = ExcelApplication.open(path, with_vba=False)
    final.add_module(report, name="Probe")
    assert final.run("Report") == "True|False"
