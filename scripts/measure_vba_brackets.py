"""A member name in brackets, typed and through an Object.

VBA binds a name it finds on a typed class. Otherwise, and always when bound late, it calls the object's DISPID_EVALUATE with the
text: Application, Worksheet, Chart and DialogSheet take that, and any other object is 438. A class that takes no other name
refuses at compile time. Nothing here is a member the rule was fitted to on Worksheet alone: Range and Sheets split it.

    python scripts/measure_vba_brackets.py

writes tests/fixtures/vba_brackets.json.
"""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

LINES = ['v = ws.[B2]', 'v = ws.[B2].Value', 'With ws\nv = .[B2].Value\nEnd With', 'v = ws.[1+1]', 'v = Application.[1+1]', 'v = ws.[_CodeName]', 'v = r.[Value]', 'v = ws.[Nope]', 'v = ws.Nope', 'v = ws.[_Default]', 'v = ws._Default', 'v = ws.[_NewEnum]', 'v = ch.[Nope]', 'v = ss.[Nope]', 'v = ws.[Name]', 'v = ws.[Index]', 'ws.[Index] = 1', 'v = r.[Offset](1, 2, 3)', 'r.[Count] = 1', 'v = r.[Nope]', 'v = r.[_Default]', 'v = r._Default']
BODY = '''Function Probe() As String
    Dim ws As @W@, r As @R@, ch As @C@, ss As @S@, v As Variant
    Set ws = Worksheets.Add
    Set r = ws.Range("B2"): Set ss = Worksheets
    r.Value = 5
    On Error Resume Next
    @L@
    Probe = Err.Number & "|" & TypeName(v)
    On Error GoTo 0
    Application.DisplayAlerts = False
    Dim d As Object
    Set d = Worksheets(1)
    d.Delete
    Application.DisplayAlerts = True
End Function
'''
target = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_brackets.json'
out = json.loads(target.read_text(encoding='utf-8')) if target.exists() else {}
for line in LINES:
    for binding, names in (('typed', ('Worksheet', 'Range', 'Chart', 'Sheets')), ('late', ('Object',) * 4)):
        if f'{line} {binding}' in out:
            continue
        source = BODY.replace('@W@', names[0]).replace('@R@', names[1]).replace('@C@', names[2]).replace('@S@', names[3]).replace('@L@', line)
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(source, 'Probe', timeout=30.0, line_numbers=False)
        refused = [d for d in result.dialogs if d.classification == 'compile-error']
        out[f'{line} {binding}'] = 'compile: ' + refused[0].message[16:70] if refused else (result.value if result.ok else 'stopped ' + str(result.message)[:50])
        print(f'{line} {binding}'.ljust(36), out[f'{line} {binding}'], flush=True)
        target.write_text(json.dumps(out, indent=0) + '\n', encoding='utf-8')
