"""Chart.ChartGroups and a ChartGroup's GapWidth, Overlap, VaryByCategories, FirstSliceAngle and DoughnutHoleSize,
replayed from live Excel (scripts/measure_chart_groups.py, tests/fixtures/chart_groups.json): what each case answers,
the chart part its workbook saves, and the group's gap width and first slice angle once the workbook is opened
again."""
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
    (Path(__file__).parent / 'fixtures/chart_groups.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
REOPEN = ('Public Function Reopen(ByVal path As String) As String\n    Dim wb As Workbook, ch As Object\n'
          '    Set wb = Workbooks.Open(path)\n    Set ch = wb.Worksheets(1).ChartObjects(1).Chart\n'
          '    Reopen = ch.ChartGroups(1).GapWidth & "," & ch.ChartGroups(1).FirstSliceAngle\n'
          '    wb.Close False\nEnd Function\n')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[ExcelApplication, list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_groups')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE + REOPEN, name='GroupsProbe')
    return app, str(app.run('GroupsProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert run[1][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert chart_part(run[2] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_saved_group_reads_back(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    part = RECORD['cases'][index][2]
    gap = re.search(r'<c:gapWidth val="(\d+)"/>', part)
    first = re.search(r'<c:firstSliceAng val="(\d+)"/>', part)
    expected = f'{gap.group(1) if gap else 150},{first.group(1) if first else 0}'
    assert run[0].run('GroupsProbe.Reopen', str(run[2] / f'c{index}.xlsx')) == expected
