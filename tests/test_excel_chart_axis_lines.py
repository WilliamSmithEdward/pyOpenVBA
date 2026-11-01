"""Axis tick marks, tick label position, gridlines and crossings replayed from live Excel
(scripts/measure_chart_axis_lines.py, tests/fixtures/chart_axis_lines.json): what each case answers, the chart part
its workbook saves, and what the saved chart's axis reads once the workbook is opened again."""
from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_axis_lines.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
READ = ('"|" & a.MajorTickMark & "," & a.MinorTickMark & "," & a.TickLabelPosition & "," & a.HasMajorGridlines & '
        '"," & a.HasMinorGridlines & "," & a.Crosses')
REOPEN = ('Public Function Reopen(ByVal path As String, ByVal which As Long) As String\n'
          '    Dim wb As Workbook, a As Object\n    Set wb = Workbooks.Open(path)\n'
          f'    Set a = wb.Worksheets(1).ChartObjects(1).Chart.Axes(which)\n    Reopen = {READ}\n'
          '    wb.Close False\nEnd Function\n')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[ExcelApplication, list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_axis_lines')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE + REOPEN, name='AxisProbe')
    return app, str(app.run('AxisProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert run[1][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert chart_part(run[2] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_saved_axis_reads_back(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    case, answer, _ = RECORD['cases'][index]
    if 'ChartType = xlLine' in case:
        pytest.skip('the case reads the axis before the chart changed type; the saved chart is a line chart')
    which = int(re.match(r'\S+ (\d):', case).group(1))  # type: ignore[union-attr]
    # The axis as it stands after the case: the second read, or the first when the change failed.
    state = answer.split('|')[2 if answer.startswith('ok') else 1]
    expected = '|' + ','.join(state.split(',')[:6])
    assert run[0].run('AxisProbe.Reopen', str(run[2] / f'c{index}.xlsx'), which) == expected
