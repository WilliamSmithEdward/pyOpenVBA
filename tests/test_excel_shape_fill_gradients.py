"""FillFormat's gradients, replayed from live Excel (scripts/measure_shape_fill_gradients.py,
tests/fixtures/shape_fill_gradients.json): what each case reads and the spPr and style its saved shape has."""
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
    (Path(__file__).parent / 'fixtures/shape_fill_gradients.json').read_text(encoding='utf-8'))
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
INDEXES = range(len(RECORD['cases']))
#: Combinations no case measured, which report themselves rather than guess.
UNMEASURED = ['sh.Fill.TwoColorGradient 1, 1: sh.Fill.Transparency = 0.5: sh.Fill.Solid',
              'sh.Fill.TwoColorGradient 1, 1: sh.Fill.Transparency = 0.5: sh.Fill.Patterned 22',
              'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(1).Transparency = 0.5: '
              'sh.Fill.GradientStops(1).Color.RGB = 255',
              'sh.Fill.TwoColorGradient 1, 1: sh.Fill.Visible = msoFalse: sh.Fill.BackColor.RGB = 255',
              'sh.Fill.TwoColorGradient 1, 1: sh.Fill.ForeColor.SchemeColor = 3',
              'sh.Line.Visible = msoFalse: sh.Fill.TwoColorGradient 1, 1: sh.Fill.ForeColor.RGB = 255',
              'sh.Fill.TwoColorGradient 1, 1: sh.Fill.BackColor.Brightness = 0.4']


def drawn(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.search(r'<xdr:spPr\b.*?</xdr:spPr>(?:<xdr:style>.*?</xdr:style>)?', xml, re.DOTALL)
    return found.group(0) if found else ''


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_fill_gradients')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE, name='FillGradientProbe')
    return str(app.run('FillGradientProbe.Probe', str(folder))).split('@@@'), folder


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


def test_a_duplicate_keeps_the_gradient_and_changes_its_own() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Function Go() As String\n    Dim sh As Object, d As Object\n'
                   '    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40)\n'
                   '    sh.Fill.TwoColorGradient 1, 1\n    Set d = sh.Duplicate\n'
                   '    Go = d.Fill.GradientStops.Count\n'
                   '    d.Fill.GradientStops.Insert 255, 0.5\n'
                   '    Go = Go & " " & d.Fill.GradientStops.Count & " " & sh.Fill.GradientStops.Count\nEnd Function\n',
                   name='Copied')
    assert app.run('Copied.Go') == '2 3 2'


def test_for_each_walks_the_stops() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Function Go() As String\n    Dim sh As Object, st As Object\n'
                   '    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40)\n'
                   '    sh.Fill.TwoColorGradient 1, 3\n'
                   '    For Each st In sh.Fill.GradientStops: Go = Go & st.Position & " ": Next\nEnd Function\n',
                   name='Walked')
    assert app.run('Walked.Go') == '0 0.5 1 '
