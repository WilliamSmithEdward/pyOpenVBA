"""ChartObject, Chart, ChartGroup, ChartTitle and Legend members that read simply, and ChartObject's RoundedCorners
and Shadow, replayed from live Excel (scripts/measure_chart_object_reads.py, tests/fixtures/chart_object_reads.json):
what each case reads, the chart part its workbook saves, and what the saved chart reads once opened again."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_object_reads.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
#: What a saved chart reads once opened again: RoundedCorners and Shadow.
REOPEN = ('Public Function Reopen(ByVal path As String) As String\n    Dim wb As Workbook, co As Object\n'
          '    Set wb = Workbooks.Open(path)\n    Set co = wb.Worksheets(1).ChartObjects(1)\n'
          '    Reopen = co.RoundedCorners & "," & co.Shadow\n    wb.Close False\nEnd Function\n')


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(SOURCE, name='ObjectReadsProbe')
    made.add_module(REOPEN, name='ObjectReadsReopen')
    return made


@pytest.fixture(scope='module')
def run(app: ExcelApplication, tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_object_reads')
    return str(app.run('ObjectReadsProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    if 'ColumnWidth' in RECORD['cases'][index][0]:
        pytest.skip('a chart object does not yet move with the columns it is placed over')
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_saved_chart_reads_back_its_look(app: ExcelApplication, run: tuple[list[str], Path], index: int) -> None:
    read = RECORD['cases'][index][1].removeprefix('ok ').split(',')
    assert str(app.run('ObjectReadsReopen.Reopen', str(run[1] / f'c{index}.xlsx'))) == f'{read[3]},{read[4]}'
