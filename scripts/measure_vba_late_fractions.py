"""What a Long property does with a fraction given through an Object, read back.

Two rules were live. Cut: the whole part, as Application.MaxIterations = 1.5 gave 1. Half down: the nearest whole number with a half
going down, as Chart.DisplayBlanksAs = 2.6 gave 3. Each predicts differently at 2.6 and 1.7, and VBA's own rounding, which a typed
call does, differs at 1.5 and 3.5. Settings a macro changes are put back.

    python scripts/measure_vba_late_fractions.py

writes tests/fixtures/vba_late_fractions.json.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_late_fractions.json'
VALUES = ['1.5', '1.7', '2.5', '2.6', '3.2', '3.5', '-1.5']
SOURCE = '''Private Function Tried(o As Object, ByVal member As String, ByVal value As Double) As String
    On Error Resume Next
    CallByName o, member, VbLet, 1
    Err.Clear
    CallByName o, member, VbLet, value
    Tried = Err.Number & ":" & CallByName(o, member, VbGet)
End Function

Private Function Direct() As String
    Dim o As Object, s As String, keep As Variant
    Dim ch As Object, co As Object, ws As Object
    Set ws = Worksheets.Add
    ws.Range("A1:A3").Value = 1
    Set co = ws.ChartObjects.Add(1, 1, 200, 100)
    Set ch = co.Chart
    Set o = Application
    keep = o.MaxIterations
    On Error Resume Next
@@LINES@@
    o.MaxIterations = keep
    On Error GoTo 0
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Direct = s
End Function

Public Function Probe() As String
    Probe = Direct()
End Function
'''


def line(holder: str, member: str, value: str) -> str:
    return (f'    {holder}.{member} = 1: Err.Clear\n    {holder}.{member} = {value}\n'
            f'    s = s & "{holder}.{member} {value}=" & Err.Number & ":" & {holder}.{member} & ";": Err.Clear\n')


def main() -> None:
    lines = ''.join(line(holder, member, value) for holder, member in (('o', 'MaxIterations'), ('ch', 'DisplayBlanksAs'), ('ch', 'ChartStyle'), ('ws', 'Tab.ColorIndex')) for value in VALUES)
    source = SOURCE.replace('@@LINES@@', lines)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=120.0, line_numbers=False)
    assert result.ok and not result.dialogs, (result.message, result.dialogs)
    answers = dict(one.split('=', 1) for one in str(result.value).split(';') if one)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers}, indent=0) + '\n', encoding='utf-8')
    for key, value in answers.items():
        print(key.ljust(30), value)


if __name__ == '__main__':
    main()
