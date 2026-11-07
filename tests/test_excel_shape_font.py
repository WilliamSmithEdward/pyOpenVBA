"""TextFrame.Characters.Font on all of a shape's text, replayed from live Excel (scripts/measure_shape_font.py,
tests/fixtures/shape_font.json)."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_font.json').read_text(encoding='utf-8'))
#: A tint on a text box's automatic colour, TextFrame2's Font and text of two lines are not built here.
CASES = [case for case in RECORD['cases'] if not (
    'TextFrame2' in case[0] or 'vbLf' in case[0] or ('AddTextbox' in case[0] and 'TintAndShade' in case[0]))]
IDS = [f'{index}: {case[0]}' for index, case in enumerate(CASES)]


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_case_as_excel_answered(case: list[Any], tmp_path: Path) -> None:
    shape, statements = case[0].split(': ', 1)
    lines = ['Public Function Go() As String', '    Dim ws As Object, sh As Object, s As String, v As String', '    Set ws = ActiveSheet',
             f'    Set sh = {shape}', '    On Error GoTo Failed', f'    {statements}', '    s = "ok"', 'Measured:', '    On Error Resume Next']
    for read in RECORD['reads']:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    Go = s', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function']
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('\n'.join(lines) + '\n', name='Fonts')
    assert app.run('Fonts.Go') == case[1]
    path = tmp_path / 'b.xlsx'
    app.save(path)
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    assert re.findall(r'<xdr:txBody>.*?</xdr:txBody>', xml, re.DOTALL)[0] == case[2]


@pytest.mark.parametrize('statements', ['sh.TextFrame.Characters.Font.Name = "Verdana"'])
def test_what_was_not_measured_reports_itself(statements: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim sh As Object\n    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40)\n'
                   f'    sh.TextFrame.Characters.Text = "Hi"\n    {statements}\nEnd Sub\n', name='Unmeasured')
    with pytest.raises(VBAUnsupportedError):
        app.run('Unmeasured.Go')
