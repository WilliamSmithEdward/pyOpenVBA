"""A sparkline group's Points, Axes, DisplayBlanksAs, DisplayHidden, PlotBy and colours, replayed from live Excel
(scripts/measure_sparkline_points.py, tests/fixtures/sparkline_points.json): what each reads and sets, and the
group each save writes."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/sparkline_points.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('sparkline_points')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='PointsProbe')
    return str(app.run('PointsProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][3]


@pytest.mark.parametrize('index', list(RECORD['groups']))
def test_each_save_writes_the_group_excel_wrote(run: tuple[list[str], Path], index: str) -> None:
    with zipfile.ZipFile(run[1] / f'c{index}.xlsx') as package:
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<x14:sparklineGroup\b.*?</x14:sparklineGroup>', sheet, re.DOTALL)
    assert re.sub(r' xr2:uid="[^"]*"', '', found.group(0) if found else '') == \
        re.sub(r' xr2:uid="[^"]*"', '', RECORD['groups'][index])
