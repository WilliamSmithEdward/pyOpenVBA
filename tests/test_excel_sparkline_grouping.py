"""SparklineGroups.ClearGroups, Group and Ungroup replayed from live Excel (scripts/measure_sparkline_grouping.py,
tests/fixtures/sparkline_grouping.json): the groups left, newest first, and the extension each save writes."""
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
    (Path(__file__).parent / 'fixtures/sparkline_grouping.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('sparkline_grouping')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='GroupingProbe')
    return str(app.run('GroupingProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][3]


@pytest.mark.parametrize('saved', list(RECORD['parts']))
def test_each_save_writes_the_groups_excel_wrote(run: tuple[list[str], Path], saved: str) -> None:
    with zipfile.ZipFile(run[1] / saved) as package:
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<extLst>.*</extLst>', sheet, re.DOTALL)
    assert re.sub(r' xr2:uid="[^"]*"', '', found.group(0) if found else '') == \
        re.sub(r' xr2:uid="[^"]*"', '', RECORD['parts'][saved])
