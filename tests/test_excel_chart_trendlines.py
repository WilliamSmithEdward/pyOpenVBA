"""Series.Trendlines, Trendlines.Add and the Trendline object, replayed from live Excel
(scripts/measure_chart_trendlines.py, tests/fixtures/chart_trendlines.json): what each case reads, the chart part its
workbook saves, and what the saved chart reads once opened again."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_trendlines.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
#: What a saved chart reads once opened again: the first series' trendline count and each one's type and name.
REOPEN = ('Public Function Reopen(ByVal path As String) As String\n'
          '    Dim wb As Workbook, sr As Object, s As String, i As Long\n'
          '    Set wb = Workbooks.Open(path)\n    Set sr = wb.Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1)\n'
          '    s = sr.Trendlines.Count\n'
          '    For i = 1 To sr.Trendlines.Count: s = s & "|" & sr.Trendlines(i).Type & "," & sr.Trendlines(i).Name: Next\n'
          '    Reopen = s\n    wb.Close False\nEnd Function\n')


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(SOURCE, name='TrendlinesProbe')
    made.add_module(REOPEN, name='TrendlinesReopen')
    return made


@pytest.fixture(scope='module')
def run(app: ExcelApplication, tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_trendlines')
    return str(app.run('TrendlinesProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_saved_chart_reads_back_its_trendlines(app: ExcelApplication, run: tuple[list[str], Path],
                                                    index: int) -> None:
    answer = RECORD['cases'][index][1]
    if not answer.startswith('ok '):
        pytest.skip('the case raised before it read anything')
    count, *lines = answer.removeprefix('ok ').split(';')[-1].split('|')
    expected = count + ''.join(f'|{fields[1]},{fields[6]}' for fields in (line.split(',') for line in lines))
    assert str(app.run('TrendlinesReopen.Reopen', str(run[1] / f'c{index}.xlsx'))) == expected
