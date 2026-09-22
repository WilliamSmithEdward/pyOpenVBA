"""Merged-range semantics measured in native Excel."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORDS = json.loads((Path(__file__).parent / "fixtures/range_merge.json").read_text())


@pytest.mark.parametrize("record", RECORDS, ids=[r["name"] for r in RECORDS])
def test_measured_merge(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim cell As Object, n As Long, answer As String\n'
                   + str(record["body"]) + 'End Function\n', name="Probe")
    assert app.run("Report") == record["reported"]


@pytest.mark.parametrize("across", [False, True])
def test_merge_and_unmerge_round_trip(tmp_path: Path, across: bool) -> None:
    import zipfile
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'''Sub Build()
Range("B2").Value = "kept"
Range("B2:D4").Merge Across:={str(across)}
End Sub''', name="Builder")
    app.run("Build")
    path = tmp_path / "merged.xlsm"
    app.save(path)
    original = zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml")
    reopened = ExcelApplication.open(path, with_vba=False)
    reopened.add_module('''Public Function Report() As String
Report = Range("C3").MergeArea.Address
End Function
Public Sub UnmergeIt()
Range("B2:D4").UnMerge
End Sub''', name="Probe")
    assert reopened.run("Report") == ("$B$3:$D$3" if across else "$B$2:$D$4")
    reopened.sheet(1).set_value("F1", "unrelated")
    reopened.save(path)
    xml = zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml")
    assert original.split(b"<mergeCells", 1)[1].split(b"</mergeCells>", 1)[0] == xml.split(b"<mergeCells", 1)[1].split(b"</mergeCells>", 1)[0]
    reopened.run("UnmergeIt")
    reopened.save(path)
    assert b"<mergeCells" not in zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml")
    final = ExcelApplication.open(path, with_vba=False)
    assert final.sheet(1).value("B2") == "kept"


def test_for_each_includes_blank_cells_outside_used_range() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Public Function Counted() As Long
Dim cell As Object
Range("A1").Value = 1
For Each cell In Range("A1:C3")
Counted = Counted + 1
Next cell
End Function''', name="Probe")
    assert app.run("Counted") == 9
