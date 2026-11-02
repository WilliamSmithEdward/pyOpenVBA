"""Shape.Fill and Shape.Line, replayed from live Excel (scripts/measure_shape_fill_line.py,
tests/fixtures/shape_fill_line.json): what each case reads and the spPr and style its saved shape has."""
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
    (Path(__file__).parent / 'fixtures/shape_fill_line.json').read_text(encoding='utf-8'))
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
KEPT = range(len(RECORD['cases']))
#: Combinations no case measured, which report themselves rather than guess.
UNMEASURED = ['sh.Line.Visible = msoFalse: sh.Fill.Visible = msoFalse',
              'sh.Line.Visible = msoFalse: sh.Fill.Transparency = 0.5',
              'sh.Line.Visible = msoFalse: sh.Line.Transparency = 0.5',
              'sh.Line.Transparency = 0.5: sh.Line.Visible = msoFalse',
              'sh.Line.Transparency = 0.5: sh.Line.DashStyle = msoLineDash',
              'sh.Line.DashStyle = msoLineDash: sh.Line.Transparency = 0.5',
              'sh.Line.DashStyle = msoLineRoundDot: sh.Line.Visible = msoFalse',
              'sh.Line.DashStyle = msoLineRoundDot: sh.Fill.Visible = msoFalse',
              'sh.Fill.Visible = msoFalse: sh.Line.DashStyle = msoLineRoundDot',
              'sh.Fill.Visible = msoFalse: sh.Fill.Solid',
              'sh.Line.Visible = msoFalse: sh.Fill.ForeColor.RGB = 255: sh.Line.Weight = 3',
              'sh.Line.Visible = msoFalse: sh.Fill.ForeColor.RGB = 255: sh.Fill.ForeColor.RGB = 0']


@pytest.mark.parametrize('statements', UNMEASURED)
def test_an_unmeasured_combination_is_refused(statements: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim sh As Object\n'
                   '    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40)\n'
                   f'    {statements}\nEnd Sub\n', name='Unmeasured')
    with pytest.raises(VBAUnsupportedError):
        app.run('Unmeasured.Go')


def drawn(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.search(r'<xdr:spPr\b.*?</xdr:spPr>(?:<xdr:style>.*?</xdr:style>)?', xml, re.DOTALL)
    return found.group(0) if found else ''


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_fill_line')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE, name='FillLineProbe')
    return str(app.run('FillLineProbe.Probe', str(folder))).split('@@@'), folder


def test_a_duplicate_keeps_the_look_and_changes_its_own() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Function Go() As String\n    Dim sh As Object, d As Object\n'
                   '    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40)\n'
                   '    sh.Fill.ForeColor.RGB = 255: sh.TextFrame.MarginLeft = 3\n    Set d = sh.Duplicate\n'
                   '    Go = d.Fill.ForeColor.RGB & " " & d.TextFrame.MarginLeft\n'
                   '    d.Fill.ForeColor.RGB = 65280: d.TextFrame.MarginLeft = 5\n'
                   '    Go = Go & " " & sh.Fill.ForeColor.RGB & " " & sh.TextFrame.MarginLeft\nEnd Function\n',
                   name='Copied')
    assert app.run('Copied.Go') == '255 3 255 3'


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_shape_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    # The xfrm's place and the geometry are other tests'; the fill, line and style are this one's.
    strip = re.compile(r'<a:xfrm\b.*?</a:xfrm>|<a:prstGeom\b.*?</a:prstGeom>|<xdr:spPr[^>]*>')
    assert strip.sub('', drawn(run[1] / f'c{index}.xlsx')) == strip.sub('', RECORD['cases'][index][2])
