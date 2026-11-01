"""Series.HasDataLabels and HasLeaderLines, replayed from live Excel
(scripts/measure_chart_data_labels.py, tests/fixtures/chart_data_labels.json): what each case answers, the chart
part its workbook saves, and whether the saved series shows labels once the workbook is opened again."""
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
    (Path(__file__).parent / 'fixtures/chart_data_labels.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
REOPEN = ('Public Function Reopen(ByVal path As String) As String\n    Dim wb As Workbook, ch As Object\n'
          '    Set wb = Workbooks.Open(path)\n    Set ch = wb.Worksheets(1).ChartObjects(1).Chart\n'
          '    Reopen = ch.SeriesCollection(1).HasDataLabels\n    wb.Close False\nEnd Function\n')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[ExcelApplication, list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_data_labels')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE + REOPEN, name='LabelsProbe')
    return app, str(app.run('LabelsProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert run[1][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert chart_part(run[2] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_saved_chart_reads_back(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    part = RECORD['cases'][index][2]
    first = re.search(r'<c:ser>.*?</c:ser>', part, re.S)
    assert first is not None
    expected = str('<c:showVal val="1"/>' in first.group(0))
    assert run[0].run('LabelsProbe.Reopen', str(run[2] / f'c{index}.xlsx')) == expected
