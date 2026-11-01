"""Worksheet.EnableFormatConditionsCalculation replayed from live Excel
(scripts/measure_format_conditions_calculation.py, tests/fixtures/format_conditions_calculation.json): what it reads
after each setting and once reopened, and the sheetPr the saved sheet holds."""
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
    (Path(__file__).parent / 'fixtures/format_conditions_calculation.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('format_conditions_calculation')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='CalculationProbe')
    return str(app.run('CalculationProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_nothing_is_saved(run: tuple[list[str], Path], index: int) -> None:
    with zipfile.ZipFile(run[1] / f'c{index}.xlsx') as package:
        xml = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<sheetPr\b.*?(?:/>|</sheetPr>)', xml, re.S)
    assert (found.group() if found else '') == RECORD['cases'][index][2]
