"""FillFormat's Patterned and Pattern, replayed from live Excel (scripts/measure_shape_fill_patterns.py,
tests/fixtures/shape_fill_patterns.json): what each case reads and the spPr and style its saved shape has."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/shape_fill_patterns.json').read_text(encoding='utf-8'))
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
INDEXES = range(len(RECORD['cases']))
#: Combinations no case measured, which report themselves rather than guess.
UNMEASURED = ['sh.Line.Visible = msoFalse: sh.Fill.Patterned 22',
              'sh.Fill.Patterned 22: sh.Fill.Visible = msoFalse: sh.Fill.BackColor.RGB = 255',
              'sh.Fill.Patterned 22: sh.Fill.BackColor.Brightness = 0.4']


def drawn(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.search(r'<xdr:spPr\b.*?</xdr:spPr>(?:<xdr:style>.*?</xdr:style>)?', xml, re.DOTALL)
    return found.group(0) if found else ''


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_fill_patterns')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE, name='FillPatternProbe')
    return str(app.run('FillPatternProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', INDEXES, ids=[RECORD['cases'][index][0] for index in INDEXES])
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', INDEXES, ids=[RECORD['cases'][index][0] for index in INDEXES])
def test_each_shape_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    # The xfrm's place and the geometry are other tests'; the fill, line and style are this one's.
    strip = re.compile(r'<a:xfrm\b.*?</a:xfrm>|<a:prstGeom\b.*?</a:prstGeom>|<xdr:spPr[^>]*>')
    case = RECORD['cases'][index]
    if case[2]:
        assert strip.sub('', drawn(run[1] / f'c{index}.xlsx')) == strip.sub('', case[2])


@pytest.mark.parametrize('statements', UNMEASURED)
def test_an_unmeasured_combination_is_refused(statements: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim sh As Object\n'
                   '    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40)\n'
                   f'    {statements}\nEnd Sub\n', name='Unmeasured')
    with pytest.raises(VBAUnsupportedError):
        app.run('Unmeasured.Go')
