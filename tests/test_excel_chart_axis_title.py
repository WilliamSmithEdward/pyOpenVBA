"""Axis title text, links and formatting replayed from live Excel (scripts/measure_chart_axis_title.py,
tests/fixtures/chart_axis_title.json): what each case answers, the chart part its workbook saves, and what the saved
axis title reads once the workbook is opened again."""
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
    (Path(__file__).parent / 'fixtures/chart_axis_title.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
READ = ('"|" & a.HasTitle & "," & t.Text & "," & t.Caption & "," & t.Formula & "," & t.FormulaR1C1 & "," & '
        't.FormulaLocal & "," & t.FormulaR1C1Local & "," & t.Orientation & "," & t.HorizontalAlignment & "," & '
        't.VerticalAlignment & "," & t.ReadingOrder & "," & t.Position & "," & t.IncludeInLayout & "," & t.Shadow')
REOPEN = ('Public Function Reopen(ByVal path As String, ByVal which As Long) As String\n'
          '    Dim wb As Workbook, a As Object, t As Object\n    Set wb = Workbooks.Open(path)\n'
          '    Set a = wb.Worksheets(1).ChartObjects(1).Chart.Axes(which)\n'
          '    If a.HasTitle Then\n        Set t = a.AxisTitle\n'
          f'        Reopen = {READ}\n    Else\n        Reopen = "|False"\n    End If\n'
          '    wb.Close False\nEnd Function\n')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[ExcelApplication, list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_axis_title')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE + REOPEN, name='AxisTitleProbe')
    return app, str(app.run('AxisTitleProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert run[1][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert chart_part(run[2] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_saved_axis_title_reads_back(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    case, answer, _ = RECORD['cases'][index]
    which = int(re.match(r'(\d):', case).group(1))  # type: ignore[union-attr]
    # The title as it stands after the case: the last read, or none once it was deleted.
    expected = '|False' if 'Delete' in case else '|' + answer.split('|')[-1]
    assert run[0].run('AxisTitleProbe.Reopen', str(run[2] / f'c{index}.xlsx'), which) == expected
