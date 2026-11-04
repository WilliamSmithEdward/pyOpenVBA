"""A Let into a variable that holds an object, by the object's class and the value, typed and through an Object.

    python scripts/measure_vba_default_let.py

writes tests/fixtures/vba_default_let.json, keeps what it holds and asks only the rest.
"""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

CASES = {
    'range': ('Range', 'ws.Range("B2")', 'r2'), 'range_many': ('Range', 'ws.Range("B2:C3")', 'r2'), 'range_multi': ('Range', 'ws.Range("B2,D4")', 'r2'),
    'sheet': ('Worksheet', 'ws', ''), 'book': ('Workbook', 'ActiveWorkbook', ''), 'app': ('Application', 'Application', ''), 'font': ('Font', 'ws.Range("B2").Font', ''), 'interior': ('Interior', 'ws.Range("B2").Interior', ''),
    'name': ('Name', 'ActiveWorkbook.Names.Add("Nm", "=1")', ''), 'shape': ('Shape', 'ws.Shapes.AddShape(1, 1, 1, 9, 9)', ''), 'comment': ('Comment', 'ws.Range("A9").AddComment("x")', ''), 'hyperlink': ('Hyperlink', 'ws.Hyperlinks.Add(ws.Range("A8"), "http://x.test/")', ''), 'list': ('ListObject', 'ws.ListObjects.Add(1, ws.Range("F1:G2"), , 1)', ''), 'style': ('Style', 'ActiveWorkbook.Styles(1)', ''), 'collection': ('Collection', 'New Collection', ''), 'sheets': ('Sheets', 'Worksheets', ''), 'names': ('Names', 'ActiveWorkbook.Names', ''), 'areas': ('Areas', 'ws.Range("B2,D4").Areas', ''), 'pagesetup': ('PageSetup', 'ws.PageSetup', ''), 'validation': ('Validation', 'ws.Range("B2").Validation', ''),
}
VALUES = {'number': '1', 'text': '"x"', 'nothing_text': 'Empty', 'null': 'Null', 'range': 'ws.Range("A1")', 'sheet': 'ws'}
BODY = '''Function Probe() As String
    Dim ws As Worksheet, h As @@T@@, s As String
    Set ws = Worksheets.Add
    ws.Range("A1").Value = 7
    Set h = @@E@@
    On Error Resume Next
    h = @@V@@
    s = Err.Number
    Err.Clear
    s = s & "|" & ws.Range("B2").Formula & "|" & TypeName(h)
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Probe = s
End Function
'''
out = {}
target = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_default_let.json'
if target.exists():
    out = json.loads(target.read_text(encoding='utf-8'))
excel = None
try:
    for name, (declared, expression, _) in CASES.items():
        for value, text in VALUES.items():
            if name not in ('range', 'sheet', 'font') and value not in ('number', 'text'):
                continue
            for binding, typed in (('typed', declared), ('late', 'Object')):
                key = f'{name} {value} {binding}'
                if key in out:
                    continue
                if excel is None:
                    excel = ExcelSession(HarnessConfig(lock_wait_s=1200.0))
                    excel.__enter__()
                    excel.new_document()
                result = excel.run_vba(BODY.replace('@@T@@', typed).replace('@@E@@', expression).replace('@@V@@', text), 'Probe', timeout=30.0, line_numbers=False)
                refused = [d for d in result.dialogs if d.classification == 'compile-error']
                out[key] = 'compile: ' + refused[0].message[16:60] if refused else (result.value if result.ok else 'stopped ' + str(result.message)[:40])
                print(key, out[key], flush=True)
                if not result.ok:
                    excel.__exit__(None, None, None)
                    excel = None
finally:
    if excel is not None:
        excel.__exit__(None, None, None)
    target.write_text(json.dumps(out, indent=0), encoding='utf-8')
