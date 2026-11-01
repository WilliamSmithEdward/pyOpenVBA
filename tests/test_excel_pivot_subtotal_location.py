"""PivotTable.SubtotalLocation, replayed from live Excel (scripts/measure_pivot_subtotal_location.py,
tests/fixtures/pivot_subtotal_location.json): each row field's LayoutSubtotalLocation and every cell of the table
after each case, the errors, and each pivot field's subtotalTop as the saved definition holds it."""
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
    (Path(__file__).parent / 'fixtures/pivot_subtotal_location.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def subtotal_tops(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        name = next(one for one in package.namelist() if re.fullmatch(r'xl/pivotTables/pivotTable\d+\.xml', one))
        xml = package.read(name).decode('utf-8')
    return [' '.join(re.findall(r'\b(?:axis|subtotalTop|outline|compact)="[^"]*"', field))
            for field in re.findall(r'<pivotField\b[^>]*>', xml)]


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('pivot_subtotal_location')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='SubtotalProbe')
    return str(app.run('SubtotalProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_definition_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert subtotal_tops(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
