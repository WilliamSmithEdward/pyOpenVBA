"""Workbook.EnableConnections, ConvertComments, CanCheckIn and ExclusiveAccess replayed from live Excel
(scripts/measure_workbook_small_methods.py, tests/fixtures/workbook_small_methods.json): what each answers."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/workbook_small_methods.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('workbook_small_methods')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='SmallProbe')
    return str(app.run('SmallProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]
