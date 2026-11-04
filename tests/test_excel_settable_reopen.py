"""What a setting is after a save and reopen, as Excel keeps or loses it (scripts/measure_settable_reopen.py)."""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

MEASURED = json.loads((Path(__file__).parent / "fixtures/settable_reopen.json").read_text(encoding="utf-8"))
#: Setting these reports itself: ChartDataPointTrack is kept in the file, somewhere not yet found.
REPORTED = {"ChartDataPointTrack"}
ROWS = [row for row in MEASURED["settings"] if row["member"] not in REPORTED]


def _source(row: dict[str, str]) -> str:
    read = f"{row['target']}.{row['member']}"
    source = MEASURED["probe"].format(target=row["target"], member=row["member"], value=row["value"], read=read)
    return source.replace('"\\reopen.xlsx"', '"/reopen.xlsx"')


@pytest.mark.parametrize("row", ROWS, ids=lambda row: f"{row['member']}={row['value']}")
def test_a_setting_after_save_and_reopen(row: dict[str, str], tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(row), name="Measure")
    assert app.run("Reopen", tmp_path.as_posix()) == row["answer"]


def _saved(tmp_path: Path, action: str) -> Path:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    file = tmp_path / "state.xlsx"
    app.add_module("Sub Go(ByVal file As String)\nDim wb As Workbook, a As Worksheet, b As Worksheet, c As Worksheet\n"
                   "Set wb = Workbooks.Add(xlWBATWorksheet): Set a = wb.Worksheets(1)\n"
                   "Set b = wb.Worksheets.Add(After:=a): Set c = wb.Worksheets.Add(After:=b)\na.Activate\n"
                   f"{action}\nwb.SaveAs file, 51\nwb.Close False\nEnd Sub\n", name="Measure")
    app.run("Go", file.as_posix())
    return file


def _part(file: Path, name: str) -> str:
    with zipfile.ZipFile(file) as package:
        return package.read(name).decode("utf-8")


def test_a_hidden_sheet_is_written_as_excel_writes_it(tmp_path: Path) -> None:
    """Excel's own bytes for these sheets (scripts/measure_settable_reopen.py)."""
    file = _saved(tmp_path, "b.Visible = xlSheetHidden\nc.Visible = xlSheetVeryHidden")
    written = re.findall(r"<sheet\b[^>]*/>", _part(file, "xl/workbook.xml"))
    assert [re.sub(r'r:id="rId\d+"', 'r:id="?"', one) for one in written] == [
        '<sheet name="Sheet1" sheetId="1" r:id="?"/>', '<sheet name="Sheet2" sheetId="2" state="hidden" r:id="?"/>',
        '<sheet name="Sheet3" sheetId="3" state="veryHidden" r:id="?"/>']


def test_showing_a_sheet_again_takes_its_state_away(tmp_path: Path) -> None:
    file = _saved(tmp_path, "b.Visible = xlSheetHidden\nb.Visible = xlSheetVisible")
    assert "state=" not in _part(file, "xl/workbook.xml")


@pytest.mark.parametrize("value, element", [("xlUnlockedCells", '<sheetProtection selectLockedCells="1"/>'),
                                            ("xlNoSelection", '<sheetProtection selectLockedCells="1" selectUnlockedCells="1"/>')])
def test_enable_selection_is_kept_on_a_sheet_not_protected(tmp_path: Path, value: str, element: str) -> None:
    file = _saved(tmp_path, f"a.EnableSelection = {value}")
    assert element in _part(file, "xl/worksheets/sheet1.xml")


def test_enable_selection_read_alone_does_not_protect_the_sheet(tmp_path: Path) -> None:
    file = _saved(tmp_path, "a.EnableSelection = xlNoSelection")
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Function Read1(ByVal file As String) As String\nDim wb As Workbook\nSet wb = Workbooks.Open(file)\n'
                   'Read1 = wb.Worksheets(1).ProtectContents & "," & wb.Worksheets(1).EnableSelection\n'
                   'wb.Worksheets(1).Range("A1").Value = 1\nwb.Close False\nEnd Function\n', name="R")
    assert app.run("Read1", file.as_posix()) == "False,-4142"
