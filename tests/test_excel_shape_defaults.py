"""Shape.SetShapesDefaultProperties, replayed from live Excel (scripts/measure_shape_columns.py,
tests/fixtures/shape_columns.json)."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_columns.json').read_text(encoding='utf-8'))
#: A shape's Characters.Font and a line's TextFrame are other gaps.
READS = [read for read in RECORD['reads'] if read.startswith('c.') and 'TextFrame' not in read]
PLACES = [RECORD['reads'].index(read) + 1 for read in READS]
CASES = [case for case in RECORD['cases'] if 'SetShapes' in case[0] and 'TextFrame' not in case[0]]
STRIP = re.compile(r'<a:xfrm\b.*?</a:xfrm>')


@pytest.mark.parametrize('case', CASES, ids=[case[0][42:] for case in CASES])
def test_each_case_as_excel_answered(case: list[Any], tmp_path: Path) -> None:
    shape, statements = case[0].split(': ', 1)
    lines = ['Public Function Go() As String', '    Dim wb As Object, ws As Object, sh As Object, c As Object, s As String, v As String',
             '    Set wb = ActiveWorkbook: Set ws = ActiveSheet', f'    Set sh = {shape}', f'    {statements}', '    s = "ok"', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "," & v']
    lines += ['    Go = s', 'End Function']
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('\n'.join(lines) + '\n', name='Defaults')
    parts = case[1].split(',')
    assert app.run('Defaults.Go') == ','.join([parts[0], *(parts[place] for place in PLACES)])
    path = tmp_path / 'b.xlsx'
    app.save(path)
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
        theme = package.read('xl/theme/theme1.xml').decode('utf-8')
    assert re.search(r'<a:objectDefaults>.*?</a:objectDefaults>', theme, re.DOTALL).group(0) == case[2][1]  # type: ignore[union-attr]
    found = re.findall(r'<xdr:spPr\b.*?</xdr:spPr>(?:<xdr:style>.*?</xdr:style>)?(?:<xdr:txBody>.*?</xdr:txBody>)?', xml, re.DOTALL)
    assert [STRIP.sub('', one) for one in found] == [STRIP.sub('', one) for one in case[2][0]]


@pytest.mark.parametrize('statements', ['sh.TextFrame.MarginLeft = 3', 'sh.TextFrame.Characters.Text = "Hi"'])
def test_text_settings_report_themselves(statements: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim sh As Object\n    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40)\n'
                   f'    {statements}\n    sh.SetShapesDefaultProperties\nEnd Sub\n', name='Text')
    with pytest.raises(VBAUnsupportedError):
        app.run('Text.Go')
