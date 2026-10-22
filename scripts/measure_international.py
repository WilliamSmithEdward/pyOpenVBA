"""Application.International and the separator settings, as this machine's Excel reads them.

The model works in the en-US locale this machine runs, so these readings
are its answers: every International index, a few that are not one, and
DecimalSeparator, ThousandsSeparator and UseSystemSeparators with their
types. Nothing is set, since the settings are the user's.

    python scripts/measure_international.py

writes tests/fixtures/international.json, which tests/test_excel_international.py
replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/international.json'

BODY = '''Dim s As String, i As Long, v As Variant
For i = -1 To 60
On Error Resume Next
v = Application.International(i)
If Err.Number <> 0 Then
s = s & "|" & i & "=e" & Err.Number
Err.Clear
ElseIf IsNull(v) Then
s = s & "|" & i & "=Null"
Else
s = s & "|" & i & "=" & TypeName(v) & ":" & CStr(v)
End If
On Error GoTo 0
Next
s = s & "|" & TypeName(Application.DecimalSeparator) & ":" & Application.DecimalSeparator
s = s & "|" & TypeName(Application.ThousandsSeparator) & ":" & Application.ThousandsSeparator
s = s & "|" & TypeName(Application.UseSystemSeparators) & ":" & Application.UseSystemSeparators
Probe = Mid(s, 2)'''

if __name__ == '__main__':
    record({'international': 'Function Probe() As String\n' + BODY + '\nEnd Function\n'}, OUT)
