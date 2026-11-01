"""Series.MarkerForegroundColor, MarkerBackgroundColor and their ColorIndex forms, replayed from live Excel
(scripts/measure_chart_marker_colours.py, tests/fixtures/chart_marker_colours.json): what each case answers, the
chart part its workbook saves, and what the saved series reads once the workbook is opened again."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_marker_colours.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
READ = ('sr.MarkerForegroundColor & "," & sr.MarkerBackgroundColor & "," & sr.MarkerForegroundColorIndex & "," & '
        'sr.MarkerBackgroundColorIndex & "," & sr.MarkerStyle')
REOPEN = ('Public Function Reopen(ByVal path As String) As String\n    Dim wb As Workbook, sr As Object\n'
          '    Set wb = Workbooks.Open(path)\n    Set sr = wb.Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1)\n'
          f'    Reopen = {READ}\n    wb.Close False\nEnd Function\n')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[ExcelApplication, list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_marker_colours')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE + REOPEN, name='MarkerProbe')
    return app, str(app.run('MarkerProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert run[1][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert chart_part(run[2] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', [index for index, case in enumerate(RECORD['cases'])
                                   if case[1].startswith('ok ') and case[1].count(',') == 4])
def test_each_saved_series_reads_back(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert run[0].run('MarkerProbe.Reopen', str(run[2] / f'c{index}.xlsx')) == RECORD['cases'][index][1][3:]
