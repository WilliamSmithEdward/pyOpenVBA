"""Range.SparklineGroups replayed from live Excel (scripts/measure_sparklines.py, tests/fixtures/sparklines.json):
Add's groups, what groups and sparklines read, Modify and its kin, Clear, Delete and Range.Clear, the extension
each save writes, and the groups a saved file gives back."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/sparklines.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('sparklines')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='SparklineProbe')
    return str(app.run('SparklineProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][3]


def _extension(sheet: str) -> str:
    """The worksheet's extLst, revision ids left out."""
    found = re.search(r'<extLst>.*</extLst>', sheet, re.DOTALL)
    return re.sub(r' xr2:uid="[^"]*"', '', found.group(0) if found else '')


@pytest.mark.parametrize('saved', list(RECORD['parts']))
def test_each_save_writes_the_extension_excel_wrote(run: tuple[list[str], Path], saved: str) -> None:
    with zipfile.ZipFile(run[1] / saved) as package:
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    assert _extension(sheet) == re.sub(r' xr2:uid="[^"]*"', '', RECORD['parts'][saved])


@pytest.mark.parametrize('saved', ['two.xlsx', 'three.xlsx', 'styled.xlsx', 'elsewhere.xlsx'])
def test_a_saved_file_gives_its_groups_back(run: tuple[list[str], Path], saved: str) -> None:
    """Reopened, a file's groups read as they did before the save, and an unchanged save keeps the extension."""
    app = ExcelApplication.open(run[1] / saved)
    app.add_module('Function P() As String\n    Dim g As Object, s As String\n'
                   '    For Each g In Worksheets(1).Cells.SparklineGroups\n'
                   '        s = s & "|" & g.Type & " " & g.SourceData & " " & g.Location.Address & " " & g.LineWeight & '
                   '" " & g.SeriesColor.Color\n    Next\n    P = Mid(s, 2)\nEnd Function\n', name='Reading')
    expected = {'two.xlsx': '2 A2:D2 $E$3 0.75 9592887|1 A1:D1 $E$1 0.75 9592887',
                'three.xlsx': '2 A1:D3 $E$1:$E$3 0.75 9592887', 'styled.xlsx': '1 A1:D1 $E$1 1.5 255',
                'elsewhere.xlsx': '1 Sheet2!A1:D1 $E$1 0.75 9592887'}
    assert app.run('Reading.P') == expected[saved]


def test_a_row_put_in_on_a_sheet_with_sparklines_is_refused() -> None:
    """The model moves no sparkline."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub S()\n    ActiveSheet.Range("E1").SparklineGroups.Add xlSparkLine, "A1:D1"\n'
                   '    ActiveSheet.Rows(1).Insert\nEnd Sub\n', name='Edit')
    with pytest.raises(VBAUnsupportedError, match='sparkline'):
        app.run('Edit.S')
