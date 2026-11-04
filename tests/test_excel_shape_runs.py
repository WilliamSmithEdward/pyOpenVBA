"""Characters(start, length) and TextRange2's ranges on a shape's text, replayed from live Excel (scripts/measure_shape_runs.py,
tests/fixtures/shape_runs.json)."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_runs.json').read_text(encoding='utf-8'))
#: Each of these reports itself: its rules are not worked out.
UNBUILT = ('Font.Fill', 'vbLf & "ef": s2 = sh.TextFrame2.TextRange.Words', 'Characters(3, 0).Font', 'Characters(1, 0).Font',
           # A Characters' Parent is the old drawing object, a Rectangle here, which the model does not have.
           'TypeName(sh.TextFrame.Characters(1, 5).Parent)')
READS = RECORD['reads']
CASES = [case for case in RECORD['cases'] if not any(one in case[0] for one in UNBUILT)]
IDS = [f'{index}: {case[0]}' for index, case in enumerate(CASES)]


def _replayed(case: list[Any], tmp_path: Path) -> tuple[Any, str]:
    shape, statements = case[0].split(': ', 1)
    lines = ['Public Function Go() As String', '    Dim ws As Object, sh As Object, o As Object, s As String, v As String, s2 As Variant', '    Set ws = ActiveSheet',
             f'    Set sh = {shape}', '    On Error GoTo Failed', f'    {statements}', '    s = "ok"', 'Measured:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "~" & Replace(v, vbLf, "<LF>")']
    lines += ['    Go = s', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function']
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('\n'.join(lines) + '\n', name='Runs')
    answer = app.run('Runs.Go')
    path = tmp_path / 'b.xlsx'
    app.save(path)
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    return answer, re.findall(r'<xdr:txBody>.*?</xdr:txBody>', xml, re.DOTALL)[0]


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_case_as_excel_answered(case: list[Any], tmp_path: Path) -> None:
    answer, body = _replayed(case, tmp_path)
    assert dict(zip(['status', *READS], answer.split('~'), strict=True)) == dict(zip(['status', *READS], case[1].split('~'), strict=True))
    assert body == case[2]


@pytest.mark.parametrize('statements', ['sh.TextFrame.Characters.Text = "Hello World": sh.TextFrame2.TextRange.Characters(3, 0).Font.Bold = msoTrue',
                                        'sh.TextFrame.Characters.Text = "ab" & vbLf & "cd": x = sh.TextFrame2.TextRange.Words.Count'])
def test_what_is_not_worked_out_reports_itself(statements: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Public Sub Go()\n    Dim sh As Object, x As Variant\n    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40)\n    {statements}\nEnd Sub\n', name='Unbuilt')
    with pytest.raises(VBAUnsupportedError):
        app.run('Unbuilt.Go')
