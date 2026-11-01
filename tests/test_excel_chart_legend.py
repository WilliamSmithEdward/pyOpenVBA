"""Chart.Legend replayed from live Excel (scripts/measure_chart_legend.py, tests/fixtures/chart_legend.json): what each
case answers, the chart part its workbook saves, and what the saved legend reads once the workbook is opened again."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_legend.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
READ = ('"|" & ch.HasLegend & "," & ch.Legend.Position & "," & ch.Legend.IncludeInLayout & "," & ch.Legend.Shadow & '
        '"," & TypeName(ch.Legend.Parent)')
REOPEN = ('Public Function Reopen(ByVal path As String) As String\n    Dim wb As Workbook, ch As Object\n'
          '    Set wb = Workbooks.Open(path)\n    Set ch = wb.Worksheets(1).ChartObjects(1).Chart\n'
          f'    If ch.HasLegend Then Reopen = {READ} Else Reopen = "|False"\n    wb.Close False\nEnd Function\n')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[ExcelApplication, list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_legend')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE + REOPEN, name='LegendProbe')
    return app, str(app.run('LegendProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert run[1][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert chart_part(run[2] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_saved_legend_reads_back(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    _, answer, part = RECORD['cases'][index]
    expected = '|' + answer.split('|')[-1] if '<c:legend>' in part else '|False'
    assert run[0].run('LegendProbe.Reopen', str(run[2] / f'c{index}.xlsx')) == expected
