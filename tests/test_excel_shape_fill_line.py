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

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/shape_fill_line.json').read_text(encoding='utf-8'))
#: Settings that make Excel write out the shape's whole style, which the model refuses: the replay runs those cases
#: with no statements and leaves their answers out.
REFUSED = ('Visible = msoFalse', 'Line.Transparency = 0.5')
SOURCE = '\n'.join(line for line in RECORD['source'].split('\n') if not any(part in line for part in REFUSED)
                   ).replace('folder & "\\', f'folder & "{os.sep}')
KEPT = [index for index, (case, _, _) in enumerate(RECORD['cases']) if not any(part in case for part in REFUSED)]


def drawn(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.search(r'<xdr:spPr\b.*?</xdr:spPr>(?:<xdr:style>.*?</xdr:style>)?', xml, re.DOTALL)
    return found.group(0) if found else ''


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_fill_line')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='FillLineProbe')
    return str(app.run('FillLineProbe.Probe', str(folder))).split('@@@'), folder


def test_a_duplicate_keeps_the_look_and_changes_its_own() -> None:
    app = ExcelApplication()
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
