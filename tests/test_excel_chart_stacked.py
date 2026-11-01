"""Stacked and 100% stacked column, bar, line and area charts, replayed from live Excel
(scripts/measure_chart_stacked.py, tests/fixtures/chart_stacked.json): what each case reads and the chart part its
workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_stacked.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
#: What a saved chart reads once opened again: its type and its group's overlap.
REOPEN = ('Public Function Reopen(ByVal path As String) As String\n    Dim wb As Workbook, ch As Object\n'
          '    Set wb = Workbooks.Open(path)\n    Set ch = wb.Worksheets(1).ChartObjects(1).Chart\n'
          '    Reopen = ch.ChartType & "," & ch.ChartGroups(1).Overlap\n'
          '    wb.Close False\nEnd Function\n')


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(SOURCE, name='StackedProbe')
    made.add_module(REOPEN, name='StackedReopen')
    return made


@pytest.fixture(scope='module')
def run(app: ExcelApplication, tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_stacked')
    return str(app.run('StackedProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_saved_chart_reads_back_its_type(app: ExcelApplication, run: tuple[list[str], Path],
                                              index: int) -> None:
    read = RECORD['cases'][index][1].removeprefix('ok ').split(',')
    assert str(app.run('StackedReopen.Reopen', str(run[1] / f'c{index}.xlsx'))) == f'{read[0]},{read[3]}'


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
