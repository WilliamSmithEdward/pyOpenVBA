"""Shape.Duplicate, ShapeRange.Duplicate and the rotation a flip mirrors, replayed from live Excel
(scripts/measure_shape_duplicate.py, tests/fixtures/shape_duplicate.json): what each case answers, every shape's
name, id and box, and each saved anchor with its creation id and cell corners cut out."""
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
    (Path(__file__).parent / 'fixtures/shape_duplicate.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def anchors(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.findall(r'<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>', xml, re.S)
    return [re.sub(r'<xdr:from>.*?</xdr:to>', '',
                   re.sub(r'(<a16:creationId\b[^>]*\bid=")\{[0-9A-F-]{36}\}', r'\1{ID}', anchor), flags=re.S)
            for anchor in found]


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_duplicate')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='DuplicateProbe')
    return str(app.run('DuplicateProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_drawing_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert anchors(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
