"""Shape.Visible and ShapeRange.Visible replayed from live Excel (scripts/measure_shape_visible.py,
tests/fixtures/shape_visible.json): what both shapes read after each setting and once the workbook is opened again,
and the cNvPr each saved shape holds."""
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
    (Path(__file__).parent / 'fixtures/shape_visible.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_visible')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='VisibleProbe')
    return str(app.run('VisibleProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_shape_identity_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    with zipfile.ZipFile(run[1] / f'c{index}.xlsx') as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    assert re.findall(r'<xdr:cNvPr\b[^>]*>', xml) == RECORD['cases'][index][2]
