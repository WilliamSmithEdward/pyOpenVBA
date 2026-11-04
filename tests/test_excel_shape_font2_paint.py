"""TextRange2.Font.Fill and Line, replayed from live Excel (scripts/measure_shape_font2_paint.py,
tests/fixtures/shape_font2_paint.json)."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_font2_paint.json').read_text(encoding='utf-8'))
#: Each of these reports itself: gradients, patterns, textures and the legacy palette on the letters, and their shadow.
UNBUILT = ('TwoColorGradient', 'Patterned', 'OneColorGradient', 'PresetGradient', 'PresetTextured', 'Line.Pattern =', 'SchemeColor', 'Shadow.Visible')
#: SchemeColor reads the nearest palette colour by a rule not worked out.
READS = [read for read in RECORD['reads'] if 'SchemeColor' not in read]
PLACES = [RECORD['reads'].index(read) + 1 for read in READS]
CASES = [case for case in RECORD['cases'] if not any(one in case[0] for one in UNBUILT)]
IDS = [f'{index}: {case[0]}' for index, case in enumerate(CASES)]


def _replayed(case: list[Any], tmp_path: Path) -> tuple[Any, str]:
    shape, statements = case[0].split(': ', 1)
    lines = ['Public Function Go() As String', '    Dim ws As Object, sh As Object, s As String, v As String', '    Set ws = ActiveSheet',
             f'    Set sh = {shape}', '    On Error GoTo Failed', f'    {statements}', '    s = "ok"', 'Measured:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "~" & v']
    lines += ['    Go = s', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function']
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('\n'.join(lines) + '\n', name='Paint')
    answer = app.run('Paint.Go')
    path = tmp_path / 'b.xlsx'
    app.save(path)
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    return answer, re.findall(r'<xdr:txBody>.*?</xdr:txBody>', xml, re.DOTALL)[0]


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_case_as_excel_answered(case: list[Any], tmp_path: Path) -> None:
    answer, body = _replayed(case, tmp_path)
    parts = case[1].split('~')
    assert dict(zip(['status', *READS], answer.split('~'), strict=True)) == dict(zip(['status', *READS], [parts[0], *(parts[place] for place in PLACES)], strict=True))
    assert body == case[2]


@pytest.mark.parametrize('statement', ['Fill.Patterned 5', 'Fill.TwoColorGradient 1, 1', 'Line.Pattern = 5', 'Line.DashStyle = 3', 'Fill.ForeColor.SchemeColor = 3'])
def test_what_is_not_worked_out_reports_itself(statement: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim sh As Object\n    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40)\n'
                   f'    sh.TextFrame2.TextRange.Font.{statement}\nEnd Sub\n', name='Unbuilt')
    with pytest.raises(VBAUnsupportedError):
        app.run('Unbuilt.Go')
