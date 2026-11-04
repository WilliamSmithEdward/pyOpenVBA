"""A Let into a variable that holds an object, against live Excel (scripts/measure_vba_default_let.py)."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError

ROOT = Path(__file__).resolve().parents[1]
ANSWERS: dict[str, str] = json.loads((ROOT / 'tests/fixtures/vba_default_let.json').read_text(encoding='utf-8'))
BODY = """Function Probe() As String
    Dim ws As Worksheet, h As @@T@@, s As String
    Set ws = Worksheets.Add
    ws.Range("A1").Value = 7
    Set h = @@E@@
    On Error Resume Next
    h = @@V@@
    s = Err.Number
    Err.Clear
    s = s & "|" & ws.Range("B2").Formula & "|" & TypeName(h)
    Probe = s
End Function
"""
CASES = {
    'range': ('Range', 'ws.Range("B2")'), 'range_many': ('Range', 'ws.Range("B2:C3")'), 'range_multi': ('Range', 'ws.Range("B2,D4")'),
    'sheet': ('Worksheet', 'ws'), 'book': ('Workbook', 'ActiveWorkbook'), 'app': ('Application', 'Application'), 'font': ('Font', 'ws.Range("B2").Font'), 'interior': ('Interior', 'ws.Range("B2").Interior'),
    'name': ('Name', 'ActiveWorkbook.Names.Add("Nm", "=1")'), 'shape': ('Shape', 'ws.Shapes.AddShape(1, 1, 1, 9, 9)'), 'comment': ('Comment', 'ws.Range("A9").AddComment("x")'), 'hyperlink': ('Hyperlink', 'ws.Hyperlinks.Add(ws.Range("A8"), "http://x.test/")'), 'list': ('ListObject', 'ws.ListObjects.Add(1, ws.Range("F1:G2"), , 1)'), 'style': ('Style', 'ActiveWorkbook.Styles(1)'), 'sheets': ('Sheets', 'Worksheets'), 'names': ('Names', 'ActiveWorkbook.Names'), 'areas': ('Areas', 'ws.Range("B2,D4").Areas'), 'pagesetup': ('PageSetup', 'ws.PageSetup'), 'validation': ('Validation', 'ws.Range("B2").Validation'),
}
VALUES = {'number': '1', 'text': '"x"', 'nothing_text': 'Empty', 'null': 'Null', 'range': 'ws.Range("A1")', 'sheet': 'ws'}
#: Collection is VBA's own class, which the table does not hold: its typed refusal is not modelled.
ASKED = sorted(key for key in ANSWERS if key.split()[0] in CASES)


@pytest.mark.parametrize('key', ASKED)
def test_a_let_goes_to_the_default_member_of_what_the_variable_holds(key: str) -> None:
    name, value, binding = key.split()
    declared, expression = CASES[name]
    source = BODY.replace('@@T@@', declared if binding == 'typed' else 'Object').replace('@@E@@', expression).replace('@@V@@', VALUES[value])
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Harness')
    if ANSWERS[key].startswith('compile: '):
        with pytest.raises(VBACompileError, match=ANSWERS[key].removeprefix('compile: ')[:20]):
            app.run('Harness.Probe')
        return
    assert app.run('Harness.Probe') == ANSWERS[key]
