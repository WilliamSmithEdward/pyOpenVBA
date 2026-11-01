"""Shape.Flip, HorizontalFlip, VerticalFlip, ZOrder and ZOrderPosition replayed from live Excel
(scripts/measure_shape_flip_order.py, tests/fixtures/shape_flip_order.json): what the shapes read after each change,
and the shapes' order and transforms in the saved drawing, also once the workbook is opened again and saved."""
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
    (Path(__file__).parent / 'fixtures/shape_flip_order.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def drawing(path: Path) -> list[list[str]]:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found: list[list[str]] = []
    for anchor in re.findall(r'<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>', xml, re.S):
        name = re.search(r'<xdr:cNvPr\b[^>]*\bname="([^"]*)"', anchor)
        xfrm = re.search(r'<a:xfrm\b[^>]*>', anchor)
        assert name is not None and xfrm is not None
        found.append([name.group(1), xfrm.group()])
    return found


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_flip_order')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='FlipProbe')
    return str(app.run('FlipProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_drawing_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert drawing(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_a_reopened_drawing_keeps_its_order_and_flips(run: tuple[list[str], Path], index: int) -> None:
    reopened = ExcelApplication.open(run[1] / f'c{index}.xlsx', with_vba=False)
    again = run[1] / f'again{index}.xlsx'
    reopened.save(again)
    assert drawing(again) == RECORD['cases'][index][2]
