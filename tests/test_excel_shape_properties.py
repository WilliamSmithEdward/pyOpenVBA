"""Shape alternative text, title, decorative flag, rotation, increments and locks replayed from live Excel
(scripts/measure_shape_properties.py, tests/fixtures/shape_properties.json): what each case answers, the shape markup
its workbook saves, and what the saved shape reads once the workbook is opened again."""
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
    (Path(__file__).parent / 'fixtures/shape_properties.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
READ = ('"|" & p.AlternativeText & "," & p.Title & "," & p.Decorative & "," & p.Rotation & "," & p.Left & "," & '
        'p.Top & "," & p.LockAspectRatio & "," & p.Locked')
REOPEN = ('Public Function Reopen(ByVal path As String) As String\n    Dim wb As Workbook, p As Object\n'
          '    Set wb = Workbooks.Open(path)\n    Set p = wb.Worksheets(1).Shapes(1)\n'
          f'    Reopen = {READ}\n    wb.Close False\nEnd Function\n')


def shape_markup(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.search(r'<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>', xml, re.S)
    assert found is not None
    return re.sub(r'(<a16:creationId\b[^>]*\bid=")\{[0-9A-F-]{36}\}', r'\1{ID}', found.group())


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[ExcelApplication, list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_properties')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE + REOPEN, name='ShapeProbe')
    return app, str(app.run('ShapeProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert run[1][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_shape_as_excel_saved_it(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert shape_markup(run[2] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_saved_shape_reads_back(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    answer = RECORD['cases'][index][1]
    assert run[0].run('ShapeProbe.Reopen', str(run[2] / f'c{index}.xlsx')) == '|' + answer.split('|')[-1]
