"""ShapeRange's IncrementLeft, IncrementTop, IncrementRotation, Flip, ZOrder, Decorative and AutoShapeType, replayed
from live Excel (scripts/measure_shape_range_more.py, tests/fixtures/shape_range_more.json): every shape's
settings after each case, the errors, and each saved shape's preset geometry."""
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
    (Path(__file__).parent / 'fixtures/shape_range_more.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def geometries(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    return re.findall(r'<a:prstGeom\b.*?</a:prstGeom>', xml, re.S)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_range_more')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='RangeProbe')
    return str(app.run('RangeProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_geometry_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert geometries(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
