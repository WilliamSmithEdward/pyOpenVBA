"""LineFormat's arrowheads, back colour, pattern and InsetPen, and ColorFormat's Brightness, TintAndShade, SchemeColor
and ObjectThemeColor, replayed from live Excel (scripts/measure_shape_line_ends.py,
tests/fixtures/shape_line_ends.json): what each case reads and the spPr and style its saved shape has.

Excel reads SchemeColor as the palette's nearest colour by a measure no fit has found, so those reads are left out of
the replay, and a colour with no exact match in the palette reports itself."""
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
    (Path(__file__).parent / 'fixtures/shape_line_ends.json').read_text(encoding='utf-8'))
SCHEME_READ = re.compile(r'    Err\.Clear\n    v = ""\n    v = sh\.\w+\.\w+\.SchemeColor\n.*?\n    s = s & "," & v\n')
SOURCE = SCHEME_READ.sub('', RECORD['source']).replace('folder & "\\', f'folder & "{os.sep}')
KEPT_READS = [index + 1 for index, read in enumerate(RECORD['reads']) if not read.endswith('.SchemeColor')]
INDEXES = range(len(RECORD['cases']))
#: Combinations no case measured, which report themselves rather than guess.
UNMEASURED = ['sh.Line.DashStyle = msoLineDash: sh.Line.Pattern = 22',
              'sh.Line.Transparency = 0.5: sh.Line.Pattern = 22',
              'sh.Line.Visible = msoFalse: sh.Line.Pattern = 22',
              'sh.Line.Pattern = 22: sh.Line.Visible = msoFalse: sh.Line.ForeColor.RGB = 255',
              'sh.Line.Pattern = 22: sh.Line.ForeColor.Brightness = 0.4',
              'sh.Line.InsetPen = msoTrue: sh.Line.Visible = msoFalse',
              'sh.Line.InsetPen = msoTrue: sh.Line.Transparency = 0.5',
              'sh.Line.Visible = msoFalse: sh.Line.InsetPen = msoTrue',
              'sh.Line.BackColor.RGB = 255: sh.Line.Visible = msoFalse',
              'sh.Line.ForeColor.RGB = 255: sh.Line.BackColor.RGB = 255',
              'sh.Line.ForeColor.Brightness = 0.4: sh.Line.ForeColor.TintAndShade = 0.2',
              'sh.Fill.ForeColor.Brightness = 0.4: sh.Fill.ForeColor.TintAndShade = 0.2: sh.Fill.ForeColor.Brightness = 0.1',
              'sh.Fill.ForeColor.SchemeColor = 3: sh.Fill.ForeColor.TintAndShade = 0.4',
              'sh.Line.Visible = msoFalse: sh.Fill.ForeColor.ObjectThemeColor = 6',
              'sh.Line.Visible = msoFalse: sh.Fill.ForeColor.Brightness = 0.4',
              'Dim v: v = sh.Fill.ForeColor.SchemeColor',
              'sh.Fill.ForeColor.SchemeColor = 3: sh.Fill.ForeColor.Brightness = 0.4: sh.Fill.ForeColor.Brightness = 0.1']
#: Through a ShapeRange, none of these is measured.
RANGE_UNMEASURED = ['r.Line.EndArrowheadStyle = 2', 'r.Line.Pattern = 22', 'r.Line.InsetPen = msoTrue',
                    'Dim v: Set v = r.Line.BackColor', 'r.Fill.ForeColor.Brightness = 0.4',
                    'r.Line.ForeColor.TintAndShade = 0.4', 'r.Fill.ForeColor.ObjectThemeColor = 6',
                    'r.Line.ForeColor.SchemeColor = 3']


def kept(answer: str) -> str:
    """An answer without its SchemeColor reads."""
    parts = answer.split(',')
    return ','.join([parts[0], *(parts[index] for index in KEPT_READS)])


def drawn(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.search(r'<xdr:spPr\b.*?</xdr:spPr>(?:<xdr:style>.*?</xdr:style>)?', xml, re.DOTALL)
    return found.group(0) if found else ''


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_line_ends')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE, name='LineEndsProbe')
    return str(app.run('LineEndsProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', INDEXES, ids=[RECORD['cases'][index][0] for index in INDEXES])
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == kept(RECORD['cases'][index][1])


@pytest.mark.parametrize('index', INDEXES, ids=[RECORD['cases'][index][0] for index in INDEXES])
def test_each_shape_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    # The xfrm's place and the geometry are other tests'; the fill, line and style are this one's.
    strip = re.compile(r'<a:xfrm\b.*?</a:xfrm>|<a:prstGeom\b.*?</a:prstGeom>|<xdr:spPr[^>]*>')
    case = RECORD['cases'][index]
    if not case[2]:
        return
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


@pytest.mark.parametrize('statements', RANGE_UNMEASURED)
def test_a_range_refuses_what_only_one_shape_has_measured(statements: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim r As Object\n'
                   '    ActiveSheet.Shapes.AddLine(10, 10, 110, 50).Name = "a"\n'
                   '    ActiveSheet.Shapes.AddLine(20, 20, 120, 60).Name = "b"\n'
                   '    Set r = ActiveSheet.Shapes.Range(Array("a", "b"))\n'
                   f'    {statements}\nEnd Sub\n', name='Ranged')
    with pytest.raises(VBAUnsupportedError):
        app.run('Ranged.Go')


def test_a_duplicate_keeps_the_new_look_and_changes_its_own() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Function Go() As String\n    Dim sh As Object, d As Object\n'
                   '    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40)\n'
                   '    sh.Line.Pattern = 22: sh.Fill.ForeColor.Brightness = 0.4\n    Set d = sh.Duplicate\n'
                   '    Go = d.Line.Pattern & " " & d.Fill.ForeColor.Brightness\n'
                   '    d.Line.Pattern = 48: d.Fill.ForeColor.Brightness = -0.25\n'
                   '    Go = Go & " " & sh.Line.Pattern & " " & sh.Fill.ForeColor.Brightness\nEnd Function\n',
                   name='Copied')
    assert app.run('Copied.Go') == '22 0.4 22 0.4'


def test_scheme_color_reads_what_was_set_and_an_exact_palette_match() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Function Go() As String\n    Dim sh As Object\n'
                   '    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40)\n'
                   '    Go = sh.Fill.BackColor.SchemeColor\n'
                   '    sh.Fill.ForeColor.RGB = 255: Go = Go & " " & sh.Fill.ForeColor.SchemeColor\n'
                   '    sh.Fill.ForeColor.SchemeColor = 80: Go = Go & " " & sh.Fill.ForeColor.SchemeColor\n'
                   'End Function\n', name='Schemed')
    assert app.run('Schemed.Go') == '9 10 80'


def test_windows_own_scheme_colours_report_themselves() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40).Fill.ForeColor.SchemeColor = 70\n'
                   'End Sub\n', name='Windowsy')
    with pytest.raises(VBAUnsupportedError):
        app.run('Windowsy.Go')
