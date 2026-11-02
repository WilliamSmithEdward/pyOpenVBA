"""ShapeRange.Fill and ShapeRange.Line, replayed from live Excel (scripts/measure_shape_range_fill_line.py,
tests/fixtures/shape_range_fill_line.json): what each case reads and the spPr each saved shape has."""
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
    (Path(__file__).parent / 'fixtures/shape_range_fill_line.json').read_text(encoding='utf-8'))
#: One-off follow-ups to a range's hiding, each too strange to generalise from one case, which the model refuses: a
#: Transparency writes a black scrgbClr, a colour is ignored and drops the hidden copy, and a range that starts with
#: a line writes its own hidden noFill. The replay runs them with no statements and leaves their answers out.
REFUSED = ['sr.Fill.Visible = msoFalse: b.Fill.Transparency = 0.5',
           'sr.Fill.Visible = msoFalse: sr.Fill.ForeColor.RGB = 255']
LINE_FIRST = 'lineb: sr.Fill.Visible = msoFalse'
KEPT = [index for index, (case, _, _) in enumerate(RECORD['cases'])
        if case != LINE_FIRST and not any(case.endswith(statements) for statements in REFUSED)]


def _source() -> str:
    """The probe as recorded, with the refused statements taken out and this platform's separator in its paths."""
    source = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
    for statements in REFUSED:
        source = source.replace(f'        {statements}\n', '')
    range_line = '        Set sr = ws.Shapes.Range(Array(d.Name, a.Name))\n'
    place = next(index for index, (case, _, _) in enumerate(RECORD['cases']) if case == LINE_FIRST)
    return source.replace(f'    Case {place}\n{range_line}        sr.Fill.Visible = msoFalse\n',
                          f'    Case {place}\n{range_line}')


def drawn(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.findall(r'<xdr:spPr\b.*?</xdr:spPr>', xml, re.DOTALL)
    return '|'.join(re.sub(r'<a:xfrm\b.*?</a:xfrm>|<a:prstGeom\b.*?</a:prstGeom>|</?xdr:spPr[^>]*>', '', one)
                    for one in found)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_range_fill_line')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_source(), name='RangeFillProbe')
    return str(app.run('RangeFillProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_shape_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert drawn(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('statements', [*REFUSED, 'Set sr = ws.Shapes.Range(Array(d.Name, a.Name)): '
                                                  'sr.Fill.Visible = msoFalse'])
def test_a_follow_up_too_strange_to_generalise_is_refused(statements: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ws As Object, a As Object, b As Object, d As Object, sr As Object\n'
                   '    Set ws = ActiveSheet\n    Set a = ws.Shapes.AddShape(1, 10, 10, 100, 40)\n'
                   '    Set b = ws.Shapes.AddShape(1, 10, 60, 100, 40)\n'
                   '    Set d = ws.Shapes.AddLine(10, 160, 110, 200)\n'
                   f'    Set sr = ws.Shapes.Range(Array(a.Name, b.Name))\n    {statements}\nEnd Sub\n', name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
