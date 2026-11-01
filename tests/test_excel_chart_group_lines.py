"""ChartGroup.HasDropLines, HasHiLoLines, HasUpDownBars and HasSeriesLines, replayed from live Excel
(scripts/measure_chart_group_lines.py, tests/fixtures/chart_group_lines.json): what each case reads and the chart part its
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
    (Path(__file__).parent / 'fixtures/chart_group_lines.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
#: What a saved chart's group reads once opened again: the four flags.
REOPEN = ('Public Function Reopen(ByVal path As String) As String\n    Dim wb As Workbook, cg As Object\n'
          '    Set wb = Workbooks.Open(path)\n    Set cg = wb.Worksheets(1).ChartObjects(1).Chart.ChartGroups(1)\n'
          '    Reopen = cg.HasDropLines & "," & cg.HasHiLoLines & "," & cg.HasUpDownBars & "," & cg.HasSeriesLines\n'
          '    wb.Close False\nEnd Function\n')
#: The element each flag is written as.
ELEMENTS = ('<c:dropLines/>', '<c:hiLowLines/>', '<c:upDownBars>', '<c:serLines/>')


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(SOURCE, name='GroupLinesProbe')
    made.add_module(REOPEN, name='GroupLinesReopen')
    return made


@pytest.fixture(scope='module')
def run(app: ExcelApplication, tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_group_lines')
    return str(app.run('GroupLinesProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_saved_chart_reads_back_what_it_wrote(app: ExcelApplication, run: tuple[list[str], Path],
                                                  index: int) -> None:
    # Opened again, a group has the lines its part holds: a kind that does not draw one wrote nothing of it.
    part = RECORD['cases'][index][2]
    expected = ','.join(str(element in part) for element in ELEMENTS)
    assert str(app.run('GroupLinesReopen.Reopen', str(run[1] / f'c{index}.xlsx'))) == expected


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
