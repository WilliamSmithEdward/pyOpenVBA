"""CodeModule.InsertLines and AddFromString with text that goes on with " _", replayed from live Excel
(scripts/measure_vbe_continuation.py, tests/fixtures/vbe_continuation.json): CountOfLines, and the module the
workbook then saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.excel import ExcelFile

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/vbe_continuation.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('vbe_continuation')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='ContinuationProbe')
    return str(app.run('ContinuationProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_count_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_module_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert ExcelFile(run[1] / f'c{index}.xlsm').get_module('Module1') == RECORD['cases'][index][2]
