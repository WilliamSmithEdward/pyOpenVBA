"""A sparkline group's DateRange and ModifyDateRange, and RightToLeftPlotOrder, replayed from live Excel
(scripts/measure_sparkline_dates.py, tests/fixtures/sparkline_dates.json): what each reads and sets, the group each
save writes, and the date range a saved file gives back."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/sparkline_dates.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('sparkline_dates')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='DatesProbe')
    return str(app.run('DatesProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][3]


@pytest.mark.parametrize('saved', list(RECORD['groups']))
def test_each_save_writes_the_group_excel_wrote(run: tuple[list[str], Path], saved: str) -> None:
    with zipfile.ZipFile(run[1] / saved) as package:
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<x14:sparklineGroup\b.*?</x14:sparklineGroup>', sheet, re.DOTALL)
    assert re.sub(r' xr2:uid="[^"]*"', '', found.group(0) if found else '') == \
        re.sub(r' xr2:uid="[^"]*"', '', RECORD['groups'][saved])


@pytest.mark.parametrize('saved', ['dated.xlsx', 'cell.xlsx', 'absolute.xlsx'])
def test_a_saved_date_range_reads_back(run: tuple[list[str], Path], saved: str) -> None:
    app = ExcelApplication.open(run[1] / saved)
    app.add_module('Function P() As String\n    With Worksheets(1).Range("E1").SparklineGroups(1)\n'
                   '        P = .DateRange & " " & .Axes.Horizontal.IsDateAxis\n    End With\nEnd Function\n',
                   name='Reading')
    expected = {'dated.xlsx': 'A6:D6 True', 'cell.xlsx': 'A6 True', 'absolute.xlsx': '$A$6:$D$6 True'}
    assert app.run('Reading.P') == expected[saved]
