"""An axis's Gridlines and TickLabels, replayed from live Excel (scripts/measure_chart_axis_parts.py,
tests/fixtures/chart_axis_parts.json): what each case reads and the chart part its workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_axis_parts.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
#: What a saved chart's value axis reads once opened again: its gridlines, label format and link.
REOPEN = ('Public Function Reopen(ByVal path As String) As String\n    Dim wb As Workbook, vx As Object\n'
          '    Set wb = Workbooks.Open(path)\n    Set vx = wb.Worksheets(1).ChartObjects(1).Chart.Axes(xlValue)\n'
          '    Reopen = vx.HasMajorGridlines & "," & vx.HasMinorGridlines & "," & vx.TickLabels.NumberFormat & "," & '
          'vx.TickLabels.NumberFormatLinked\n    wb.Close False\nEnd Function\n')
COMPLETED = [index for index, (_, answer, _) in enumerate(RECORD['cases']) if answer.startswith('ok ')]


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(SOURCE, name='AxisPartsProbe')
    made.add_module(REOPEN, name='AxisPartsReopen')
    return made


@pytest.fixture(scope='module')
def run(app: ExcelApplication, tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_axis_parts')
    return str(app.run('AxisPartsProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', COMPLETED)
def test_each_saved_chart_reads_back_what_it_read(app: ExcelApplication, run: tuple[list[str], Path],
                                                  index: int) -> None:
    reads = RECORD['cases'][index][1][3:].split(',')
    expected = ','.join([reads[0], reads[1], reads[4], reads[5]])
    assert str(app.run('AxisPartsReopen.Reopen', str(run[1] / f'c{index}.xlsx'))) == expected
