"""TextFrame2.Column, replayed from live Excel (scripts/measure_shape_columns.py, tests/fixtures/shape_columns.json).
The fixture holds Shape.Callout and SetShapesDefaultProperties too, which are not built yet."""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_columns.json').read_text(encoding='utf-8'))
CASES = [case for case in RECORD['cases'] if 'Column' in case[0] or case[0] in ('ws.Shapes.AddShape(1, 10, 10, 100, 40): ',
                                                                                'ws.Shapes.AddTextbox(1, 10, 10, 100, 40): ',
                                                                                'ws.Shapes.AddLine(10, 10, 110, 50): ')]


@pytest.mark.parametrize('case', CASES, ids=[case[0] for case in CASES])
def test_each_case_as_excel_answered(case: list[Any], tmp_path: Path) -> None:
    shape, statements = case[0].split(': ', 1)
    source = ('Public Function Go() As String\n    Dim ws As Object, sh As Object, s As String\n    Set ws = ActiveSheet\n'
              f'    Set sh = {shape}\n    On Error GoTo Failed\n    {statements}\n    s = "ok"\nMeasured:\n    On Error Resume Next\n'
              '    Dim v As String\n    Err.Clear\n    v = sh.TextFrame2.Column.Number\n    If Err.Number Then v = "E" & Err.Number\n    s = s & "," & v\n'
              '    Err.Clear\n    v = "": v = sh.TextFrame2.Column.Spacing\n    If Err.Number Then v = "E" & Err.Number\n    s = s & "," & v\n'
              '    Go = s\n    Exit Function\nFailed:\n    s = "Err " & Err.Number\n    Resume Measured\nEnd Function\n')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Columns')
    assert app.run('Columns.Go') == ','.join(case[1].split(',')[:3])
    path = tmp_path / 'b.xlsx'
    app.save(path)
    import zipfile

    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    want = re.findall(r'<a:bodyPr[^>]*?/?>', case[2][0][0] if case[2][0] else '')
    assert re.findall(r'<a:bodyPr[^>]*?/?>', xml) == want
