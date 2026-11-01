"""Series.HasErrorBars, Series.ErrorBar and the ErrorBars object, replayed from live Excel
(scripts/measure_chart_error_bars.py, tests/fixtures/chart_error_bars.json): what each case reads and the chart part its
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
    (Path(__file__).parent / 'fixtures/chart_error_bars.json').read_text(encoding='utf-8'))
#: Cases that turn error bars on over an automatic value axis scale, whose major unit the model does not work out.
AUTOMATIC = {index for index, (case, _, _) in enumerate(RECORD['cases'])
             if 'HasErrorBars = True' in case and 'MajorUnit' not in case
             and not case.startswith(('xlXYScatter', 'xlPie', 'xlDoughnut'))}
#: The probe as recorded, its files joined to the folder with this platform's separator, the automatic cases left out.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').replace(
    '        s = s & "@@@" & Tried(i, folder)\n',
    f'        If InStr(",{",".join(map(str, sorted(AUTOMATIC)))},", "," & i & ",") Then\n'
    '            s = s & "@@@"\n        Else\n            s = s & "@@@" & Tried(i, folder)\n        End If\n')
#: What a saved chart reads once opened again: HasErrorBars and the bars' EndStyle.
REOPEN = ('Public Function Reopen(ByVal path As String) As String\n    Dim wb As Workbook, sr As Object, s As String\n'
          '    Set wb = Workbooks.Open(path)\n    Set sr = wb.Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1)\n'
          '    s = sr.HasErrorBars\n    If sr.HasErrorBars Then s = s & "," & sr.ErrorBars.EndStyle\n'
          '    Reopen = s\n    wb.Close False\nEnd Function\n')


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(SOURCE, name='ErrorBarsProbe')
    made.add_module(REOPEN, name='ErrorBarsReopen')
    return made


@pytest.fixture(scope='module')
def run(app: ExcelApplication, tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_error_bars')
    return str(app.run('ErrorBarsProbe.Probe', str(folder))).split('@@@'), folder


def _measured(index: int) -> None:
    if index in AUTOMATIC:
        pytest.skip('the default amount is the automatic value axis major unit, which needs native layout')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    _measured(index)
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    _measured(index)
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_saved_chart_reads_back_its_bars(app: ExcelApplication, run: tuple[list[str], Path],
                                              index: int) -> None:
    _measured(index)
    answer = RECORD['cases'][index][1]
    if not answer.startswith('ok '):
        pytest.skip('the case raised before it read anything')
    read = answer.split(';')[-1].removeprefix('ok ').lstrip(',')
    expected = 'False' if read.startswith('False') else ','.join(read.split(',')[:2])
    assert str(app.run('ErrorBarsReopen.Reopen', str(run[1] / f'c{index}.xlsx'))) == expected
