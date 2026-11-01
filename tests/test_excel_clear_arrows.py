"""Worksheet.ClearArrows and ClearCircles with nothing drawn, replayed from live Excel
(scripts/measure_clear_arrows.py, tests/fixtures/clear_arrows.json): nothing changes and the workbook stays saved,
but ClearArrows on a protected sheet is error 1004."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/clear_arrows.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def answers(tmp_path_factory: pytest.TempPathFactory) -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='ArrowsProbe')
    return str(app.run('ArrowsProbe.Probe', str(tmp_path_factory.mktemp('clear_arrows')))).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]
