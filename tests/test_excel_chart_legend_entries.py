"""Legend.LegendEntries, LegendEntry.Delete and Legend.Clear, replayed from live Excel
(scripts/measure_chart_legend_entries.py, tests/fixtures/chart_legend_entries.json): what each case answers, the chart
part its workbook saves, and how many entries the saved legend shows once the workbook is opened again."""
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
    (Path(__file__).parent / 'fixtures/chart_legend_entries.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
REOPEN = ('Public Function Reopen(ByVal path As String) As String\n    Dim wb As Workbook, ch As Object\n'
          '    Set wb = Workbooks.Open(path)\n    Set ch = wb.Worksheets(1).ChartObjects(1).Chart\n'
          '    Reopen = ch.HasLegend: If ch.HasLegend Then Reopen = ch.Legend.LegendEntries.Count\n    wb.Close False\nEnd Function\n')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[ExcelApplication, list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_legend_entries')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE + REOPEN, name='EntriesProbe')
    return app, str(app.run('EntriesProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert run[1][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert chart_part(run[2] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_saved_chart_reads_back(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    part = RECORD['cases'][index][2]
    if '<c:legend>' not in part:
        expected = 'False'
    else:
        # A pie's legend shows its first series' points, any other chart its series, less the entries deleted.
        points = re.search(r'<c:val>.*?<c:ptCount val="(\d+)"/>', part, re.S)
        total = int(points.group(1)) if '<c:pieChart>' in part and points else part.count('<c:ser>')
        expected = str(total - part.count('<c:legendEntry>'))
    assert run[0].run('EntriesProbe.Reopen', str(run[2] / f'c{index}.xlsx')) == expected
