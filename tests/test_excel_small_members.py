"""Small members everyday macros read, replayed from live Excel (scripts/measure_small_members.py,
tests/fixtures/small_members.json): Creator across Excel's objects, Range.InsertIndent, Workbook.HasVBProject and
CheckCompatibility, Name.Index, and Names' Parent and Application."""
from __future__ import annotations

import json
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/small_members.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(RECORD['source'], name='SmallProbe')
    return made


def test_creator_of_every_object(app: ExcelApplication) -> None:
    assert str(app.run('SmallProbe.Creators')).split('|') == list(RECORD['creators'].values())


def test_a_sheet_name_in_a_book_named_for_its_sheet_finds_the_sheets_own() -> None:
    """Workbooks.Add(xlWBATWorksheet) names its book Sheet1, as its sheet; Names("Sheet1!Local") is the sheet's."""
    made = ExcelApplication()
    made.add_workbook()
    made.add_module('Function P() As String\n'
                    '    Dim wb As Workbook\n'
                    '    Set wb = Workbooks.Add(xlWBATWorksheet)\n'
                    '    wb.Worksheets(1).Names.Add "Local", "=4"\n'
                    '    P = wb.Name & " " & wb.Names("Sheet1!Local").Name\n'
                    'End Function\n', name='Fresh')
    assert made.run('Fresh.P') == 'Sheet1 Sheet1!Local'


def test_each_case_as_excel_answered(app: ExcelApplication) -> None:
    assert str(app.run('SmallProbe.Cases')).split('|') == [answer for _, _, answer in RECORD['cases']]


def test_check_compatibility_is_saved(tmp_path: Path) -> None:
    """Saved as checkCompatibility="1" in the workbookPr, as Excel saves it, and read back."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub Save(ByVal path As String, ByVal on As Boolean)\nActiveWorkbook.CheckCompatibility = on\n'
                   'ActiveWorkbook.SaveAs path, xlOpenXMLWorkbook\nEnd Sub\n', name='Saving')
    for name, setting in (('On.xlsx', True), ('Off.xlsx', False)):
        app.run('Saving.Save', str(tmp_path / name), setting)
        with zipfile.ZipFile(tmp_path / name) as package:
            workbook = package.read('xl/workbook.xml').decode('utf-8')
        assert ('checkCompatibility="1"' in workbook) is setting
        assert ExcelApplication.open(tmp_path / name).workbook.CheckCompatibility() is setting
