"""Shape.Callout, replayed from live Excel (scripts/measure_shape_columns.py, tests/fixtures/shape_columns.json)."""
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
READS = [read for read in RECORD['reads'] if read.startswith('sh.Callout.') or read in ('sh.AutoShapeType', 'sh.Type')]
PLACES = [RECORD['reads'].index(read) + 1 for read in READS]
CASES = [case for case in RECORD['cases'] if 'Callout' in case[0] and 'Column' not in case[0] and 'SetShapes' not in case[0]
         or re.match(r'ws\.Shapes\.AddShape\((?:1[012]\d|61), 10, 10, 100, 40\): ', case[0]) and 'SetShapes' not in case[0]]
IDS = [f'{index}: {case[0]}' for index, case in enumerate(CASES)]


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_case_as_excel_answered(case: list[Any], tmp_path: Path) -> None:
    shape, statements = case[0].split(': ', 1)
    lines = ['Public Function Go() As String', '    Dim ws As Object, sh As Object, s As String, v As String', '    Set ws = ActiveSheet',
             f'    Set sh = {shape}', '    On Error GoTo Failed', f'    {statements}', '    s = "ok"', 'Measured:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "," & v']
    lines += ['    Go = s', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function']
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('\n'.join(lines) + '\n', name='Callouts')
    parts = case[1].split(',')
    assert app.run('Callouts.Go') == ','.join([parts[0], *(parts[place] for place in PLACES)])
    path = tmp_path / 'b.xlsx'
    app.save(path)
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    geometry = re.compile(r'<a:prstGeom.*?</a:prstGeom>')
    assert geometry.findall(xml)[0] == geometry.findall(case[2][0][0])[0]


def test_what_was_not_measured_reports_itself() -> None:
    for statements in ('sh.Callout.CustomLength 5', 'sh.Callout.Angle = 2'):
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module('Public Sub Go()\n    Dim sh As Object\n    Set sh = ActiveSheet.Shapes.AddCallout(4, 10, 10, 100, 40)\n'
                       f'    {statements}\nEnd Sub\n', name='Unmeasured')
        with pytest.raises(VBAUnsupportedError):
            app.run('Unmeasured.Go')
