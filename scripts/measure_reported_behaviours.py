"""The VBA from issue reports #249, #253, #254, #258, #259, #260 and #261, run in live Excel, so each answer the reports
quote is a measured one the model is held to.

Predictions written before measuring, from the reports: a Variant number is less than a Variant string, "900" < 1000
False; a Function As Long() returns its array; a ByVal array argument is a copy; &HFF01& is the Long 65281; the Mid
statement replaces in place; Range("A1:C3").Value is 3 by 3 on a sheet holding only A1; a formula entered before
Calculation goes manual reads as a Double.

    python scripts/measure_reported_behaviours.py

writes tests/fixtures/reported_behaviours.json, which tests/test_vba_reported_behaviours.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/reported_behaviours.json'
SOURCE = '''Public Function Pair() As Long()
    Dim r(1 To 2) As Long
    r(1) = 3: r(2) = 4
    Pair = r
End Function

Private Sub SetFirst(ByVal arr As Variant)
    arr(0) = 99
End Sub

Private Sub SetFormula(ByVal ws As Worksheet)
    ws.Range("A1").Formula = "=1+2"
End Sub

Private Function ReadUnderManual(ByVal ws As Worksheet) As String
    Dim was As Long
    was = Application.Calculation
    Application.Calculation = xlCalculationManual
    ReadUnderManual = TypeName(ws.Range("A1").Value) & ":" & ws.Range("A1").Value
    Application.Calculation = was
End Function

Private Function Tried(ByVal which As Long, ByVal ws As Worksheet) As String
    Dim a As Variant, b As Variant, c As Variant, v As Variant, s As String
    On Error GoTo Failed
    Select Case which
    Case 0
        a = "unpaid": b = 1000: c = "900"
        Tried = (a < b) & " " & (b < a) & " " & (c < b) & " " & (a > b) & " " & (c = b)
    Case 1
        v = Pair()
        Tried = TypeName(v) & " " & LBound(v) & "-" & UBound(v) & " " & v(1) & v(2)
    Case 2
        v = Array(1, 2)
        SetFirst v
        Tried = CStr(v(0))
    Case 3
        Tried = CStr(&HFF01&) & " " & CStr(&HFF01) & " " & CStr(&HFFFF&) & " " & CStr(&H8000&) & " " & _
                TypeName(&HFF01&) & " " & TypeName(&HFF01) & " " & CStr(&O177777&) & " " & CStr(&O177777)
    Case 4
        s = "abc"
        Mid$(s, 2, 1) = "X"
        Tried = s
        s = "abcdef"
        Mid(s, 2) = "XY"
        Tried = Tried & " " & s
        s = "abc"
        Mid(s, 3, 5) = "XYZ"
        Tried = Tried & " " & s
    Case 5
        ws.Range("A1").Value = 1
        v = ws.Range("A1:C3").Value
        Tried = LBound(v, 1) & "-" & UBound(v, 1) & "x" & LBound(v, 2) & "-" & UBound(v, 2) & " " & _
                TypeName(v(3, 3))
    Case 6
        SetFormula ws
        Tried = ReadUnderManual(ws)
    End Select
    Exit Function
Failed:
    Tried = "Err " & Err.Number & " " & Err.Description
End Function

Public Function Probe() As String
    Dim s As String, i As Long, wb As Workbook
    For i = 0 To 6
        Set wb = Workbooks.Add(xlWBATWorksheet)
        s = s & "|" & Tried(i, wb.Worksheets(1))
        wb.Close False
    Next
    Probe = Mid(s, 2)
End Function
'''
CASES = ['#249 variant comparisons', '#253 function As Long()', '#254 ByVal array', '#258 hex and octal literals',
         '#259 Mid statement', '#260 Value of a block past the used range', '#261 formula read under manual']


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(SOURCE, 'Probe', timeout=120.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('|')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE,
              'cases': dict(zip(CASES, answers, strict=True))}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases'].items():
        print(case, '->', answer)


if __name__ == '__main__':
    main()
