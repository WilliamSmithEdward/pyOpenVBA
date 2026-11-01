"""Shape settings changed on a shape a workbook opened with, replayed from live Excel
(scripts/measure_loaded_shape_properties.py, tests/fixtures/loaded_shape/): what the shape reads after the change,
and the anchor the saved drawing holds, patched into the markup the shape came with."""
from __future__ import annotations

import json
import os
import re
import shutil
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FOLDER = Path(__file__).parent / 'fixtures/loaded_shape'
RECORD: dict[str, Any] = json.loads((FOLDER / 'cases.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def anchor(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.search(r'<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>', xml, re.S)
    assert found is not None
    return found.group()


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('loaded_shape')
    for base in ('plain', 'set'):
        shutil.copyfile(FOLDER / f'{base}.xlsx', folder / f'{base}.xlsx')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='LoadedProbe')
    return str(app.run('LoadedProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_anchor_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert anchor(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][3]
