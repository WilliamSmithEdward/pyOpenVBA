"""Measure Range.Consolidate: what it writes by position and by labels, with each function, and what the
destination sheet's ConsolidationFunction, ConsolidationOptions and ConsolidationSources read afterwards.

Each case runs on a new workbook of three sheets: Sheet1 and Sheet2 hold the sources, Sheet3 takes the result. The
answer is Sheet3's A1:F8 formulas and number formats, then its consolidation settings, or the error.

    python scripts/measure_consolidate.py

writes tests/fixtures/consolidate.json, which tests/test_excel_consolidate.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/consolidate.json'
BOTH = 'Array("Sheet1!R1C1:R4C3", "Sheet2!R1C1:R3C3")'
DATA = 'Array("Sheet1!R2C2:R4C3", "Sheet2!R2C2:R3C3")'
FUNCTIONS = ['xlSum', 'xlCount', 'xlAverage', 'xlMax', 'xlMin', 'xlProduct', 'xlCountNums', 'xlStDev', 'xlStDevP',
             'xlVar', 'xlVarP']
#: Each case: statements on the workbook (wb, s1, s2, s3), run after the sources are written.
CASES: list[list[str]] = [
    [f's3.Range("A1").Consolidate {DATA}, xlSum'],
    [f's3.Range("A1").Consolidate {DATA}'],
    [f's3.Range("B2").Consolidate {DATA}, xlSum'],
    [f's3.Range("A1:A1").Consolidate {BOTH}, xlSum, True, True'],
    [f's3.Range("A1").Consolidate {BOTH}, xlSum, True, False'],
    [f's3.Range("A1").Consolidate {BOTH}, xlSum, False, True'],
    [f's3.Range("A1").Consolidate {BOTH}, xlSum, False, False'],
    [f's3.Range("B2").Consolidate {BOTH}, xlSum, True, True'],
    [f's3.Range("A1:B2").Consolidate {DATA}, xlSum'],
    [f's3.Range("A1:D5").Consolidate {DATA}, xlSum'],
    [f's3.Range("A1:B2").Consolidate {BOTH}, xlSum, True, True'],
    ['s3.Range("A1").Consolidate "Sheet1!R2C2:R4C3", xlSum'],
    ['s3.Range("A1").Consolidate Array("Sheet1!B2:C4"), xlSum'],
    ['s3.Range("A1").Consolidate Array("Sheet1!R2C2:R4C3", "Sheet1!R2C2:R4C3"), xlSum'],
    ['s3.Range("A1").Consolidate Array("[Book9]Sheet1!R2C2:R4C3"), xlSum'],
    ['s3.Range("A1").Consolidate Array("Sheet2!R1C5:R2C6"), xlSum'],
    ['s3.Range("A1").Consolidate Array("Sheet2!R1C5:R2C6", "Sheet2!R1C5:R2C6"), xlCount'],
    ['s3.Range("A1").Consolidate Array("Sheet2!R1C5:R2C6"), xlCountNums'],
    ['s3.Range("A1").Consolidate Array("Sheet2!R1C5:R2C6"), xlAverage'],
    ['s3.Range("A1").Consolidate'],
    [f's3.Range("A1").Consolidate {DATA}, xlSum', 's3.Range("A1:F8").ClearContents', 's3.Range("A1").Consolidate'],
    [f's3.Range("A1").Consolidate {DATA}, xlMax', 's3.Range("A1:F8").ClearContents',
     's3.Range("A1").Consolidate Function:=xlMin'],
    [f's3.Range("A1").Consolidate {BOTH}, xlSum, True, True, True'],
    [f's3.Range("A1").Consolidate {DATA}, xlSum, CreateLinks:=False'],
    [f's3.Range("A1").Consolidate {DATA}, 99'],
    [f's3.Range("A1").Consolidate {DATA}, xlUnknown'],
    ['s1.Range("A2").Value = "pear"', f's3.Range("A1").Consolidate {BOTH}, xlSum, True, True'],
    ['s1.Range("B1").Value = Empty', f's3.Range("A1").Consolidate {BOTH}, xlSum, True, True'],
    ['s1.Range("A3").Value = 7', f's3.Range("A1").Consolidate {BOTH}, xlSum, True, True'],
    ['s1.Range("A3").Value = Empty', f's3.Range("A1").Consolidate {BOTH}, xlSum, True, True'],
    ['s1.Range("B2:C4").NumberFormat = "0.00"', f's3.Range("A1").Consolidate {DATA}, xlSum'],
    ['s1.Range("B2:C4").NumberFormat = "0.00"', f's3.Range("A1").Consolidate {BOTH}, xlSum, True, True'],
    ['s3.Range("A1:F8").NumberFormat = "0.0"', f's3.Range("A1").Consolidate {DATA}, xlSum'],
    ['s3.Range("F8").Value = "keep"', 's3.Range("C3").Value = "over"', f's3.Range("A1").Consolidate {DATA}, xlSum'],
    ['s1.Range("B2").Formula = "=10+1"', f's3.Range("A1").Consolidate {DATA}, xlSum'],
    ['s1.Range("B2").Value = "text"', f's3.Range("A1").Consolidate {DATA}, xlSum'],
    ['s1.Range("B2").Value = CVErr(xlErrDiv0)', f's3.Range("A1").Consolidate {DATA}, xlSum'],
    ['s1.Range("B2").Value = True', f's3.Range("A1").Consolidate {DATA}, xlSum'],
    ['s1.Range("B2").Value = "5"', f's3.Range("A1").Consolidate {DATA}, xlSum'],
    ['s1.Range("B2").Value = 0.1', 's2.Range("B2").Value = 0.2', f's3.Range("A1").Consolidate {DATA}, xlSum'],
    ['s1.Range("B2").Value = DateSerial(2024, 1, 2)', f's3.Range("A1").Consolidate {DATA}, xlSum'],
    ['s3.Range("A1").Consolidate Array("Sheet1!R1C1:R4C3", "Sheet2!R1C1:R3C3"), xlSum, True, True',
     's3.Range("A1:F8").ClearContents', 's3.Range("A1").Consolidate Array("Sheet2!R1C1:R3C3"), xlCount'],
    [f's1.Range("A1").Consolidate {DATA}, xlSum'],
    ['s3.Range("A1").Consolidate Array("Sheet1!R1C1:R4C3"), xlSum, True, True'],
    ['s3.Range("A1").Consolidate Array("Sheet1!R1C1:R1C3"), xlSum, True, False'],
    ['s3.Range("A1").Consolidate Array("Sheet1!R1C1"), xlSum'],
    ['s3.Range("A1").Consolidate Array("Sheet1!R1C1:R4C3", "Sheet2!R1C1:R3C3"), xlSum, True, True',
     's3.Range("H1").Consolidate Array("Sheet1!R2C2:R3C2"), xlMax'],
    ['s3.Range("A1").Consolidate Array("Sheet1!R1C1:R4C3", "Sheet2!R1C1:R3C3"), xlSum, 1, 0'],
    ['s3.Range("A1").Consolidate Array("Sheet1!R1C1:R4C3", "Sheet2!R1C1:R3C3"), xlSum, "True", "False"'],
    ['s3.Range("A1").Consolidate Array(), xlSum'],
    ['s3.Range("A1").Consolidate Array("Sheet1!R2C2:R4C3,Sheet2!R2C2:R3C3"), xlSum'],
    ['s3.Range("A1").Consolidate Array("NoSheet!R2C2:R4C3"), xlSum'],
    ['wb.Names.Add "Src", "=Sheet1!$B$2:$C$4"', 's3.Range("A1").Consolidate Array("Src"), xlSum'],
    ['s3.Range("A1").Consolidate Array("Sheet1!C2:C3"), xlSum'],
    ['s3.Range("A1").Consolidate Array("Sheet1!R2C2:R4C3"), xlSum', 's3.Range("A1").Consolidate Array("Sheet2!R2C2:R3C3")'],
    ['s3.Range("A1").Consolidate Array("Sheet2!R1C5:R2C6"), xlCount'],
    ['s3.Range("B1").Value = "old"', 's3.Range("A2").Value = "old2"',
     's3.Range("A1").Consolidate Array("Sheet2!R1C5:R2C6"), xlSum'],
    ['s3.Range("B5").Value = "old"', 's3.Range("A1").Value = "corner"', f's3.Range("A1").Consolidate {BOTH}, xlSum, True, True'],
    ['s1.Range("A1:C4").NumberFormat = "0.00"', f's3.Range("A1").Consolidate {BOTH}, xlSum, True, True'],
    ['s2.Range("B2:C3").NumberFormat = "0.0"', f's3.Range("A1").Consolidate {DATA}, xlSum'],
    ['s1.Range("B2").ClearContents', 's1.Range("B2").NumberFormat = "0.000"', 's2.Range("B2").NumberFormat = "0.0"',
     f's3.Range("A1").Consolidate {DATA}, xlSum'],
    ['s1.Range("B2").Value = "text"', 's1.Range("B2").NumberFormat = "0.000"', 's2.Range("B2").NumberFormat = "0.0"',
     f's3.Range("A1").Consolidate {DATA}, xlSum'],
    ['s1.Range("B2").Value = "text"', f's3.Range("A1").Consolidate {DATA}, xlCount'],
    ['s1.Range("C1").Value = 2024', f's3.Range("A1").Consolidate {BOTH}, xlSum, True, True'],
    ['s1.Range("B3").Value = "n/a"', f's3.Range("A1").Consolidate {BOTH}, xlSum, True, True'],
    ['s1.Range("B3").Value = "n/a"', f's3.Range("A1").Consolidate {BOTH}, xlCount, True, True'],
]
for name in FUNCTIONS:
    CASES.append([f's3.Range("A1").Consolidate {BOTH}, {name}, True, True'])


def _probe() -> str:
    lines = ['Private Function Settings(ByVal ws As Worksheet) As String',
             '    Dim s As String, v As Variant, i As Long', '    On Error Resume Next',
             '    s = "F=" & ws.ConsolidationFunction', '    If Err.Number <> 0 Then s = "F=E" & Err.Number: Err.Clear',
             '    v = ws.ConsolidationOptions',
             '    If Err.Number <> 0 Then',
             '        s = s & " O=E" & Err.Number: Err.Clear', '    ElseIf IsArray(v) Then',
             '        s = s & " O=" & TypeName(v) & "(" & LBound(v) & ")"',
             '        For i = LBound(v) To UBound(v): s = s & "," & TypeName(v(i)) & ":" & CStr(v(i)): Next',
             '    Else', '        s = s & " O=" & TypeName(v)', '    End If',
             '    v = ws.ConsolidationSources',
             '    If Err.Number <> 0 Then',
             '        s = s & " S=E" & Err.Number: Err.Clear', '    ElseIf IsArray(v) Then',
             '        s = s & " S=" & TypeName(v) & "(" & LBound(v) & ")"',
             '        For i = LBound(v) To UBound(v): s = s & "," & CStr(v(i)): Next',
             '    Else', '        s = s & " S=" & TypeName(v)', '    End If',
             '    Settings = s', 'End Function', '',
             'Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, s1 As Worksheet, s2 As Worksheet, s3 As Worksheet, c As Range, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set s1 = wb.Worksheets(1)', '    s1.Name = "Sheet1"',
             '    Set s2 = wb.Worksheets.Add(After:=s1)', '    s2.Name = "Sheet2"',
             '    Set s3 = wb.Worksheets.Add(After:=s2)', '    s3.Name = "Sheet3"',
             '    s1.Range("A1:C4").Value = Array(Array("", "Q1", "Q2"), Array("Apple", 1, 2), Array("Pear", 3, 4), '
             'Array("Fig", 5, 6))',
             '    s1.Range("A1:C1").Value = Array("", "Q1", "Q2")', '    s1.Range("A2:C2").Value = Array("Apple", 1, 2)',
             '    s1.Range("A3:C3").Value = Array("Pear", 3, 4)', '    s1.Range("A4:C4").Value = Array("Fig", 5, 6)',
             '    s2.Range("A1:C1").Value = Array("", "q2", "Q3")', '    s2.Range("A2:C2").Value = Array("pear", 10, 20)',
             '    s2.Range("A3:C3").Value = Array("Kiwi", 30, 40)',
             '    s2.Range("E1").Value = 1', '    s2.Range("F1").Value = "x"', '    s2.Range("F2").Value = 2.5',
             '    On Error GoTo Failed', '    Select Case which']
    for index, statements in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
    lines += ['    End Select', '    For Each c In s3.Range("A1:F8")',
              '        If Not IsEmpty(c.Formula) And c.Formula <> "" Then s = s & ";" & c.Address(False, False) & "=" '
              '& c.Formula & IIf(c.NumberFormat = "General", "", "{" & c.NumberFormat & "}")',
              '    Next', '    Tried = Mid(s, 2) & " | " & Settings(s3) & " | " & Settings(s1)', 'Done:',
              '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description & " | " & Settings(s3)', '    Resume Done',
              'End Function', '',
              'Public Function Probe(ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _run(source: str, first: int, last: int) -> list[str] | None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (first, last), timeout=300.0)
    if result.outcome == 'runner-error' or result.dialogs:
        return None
    assert result.ok, result
    return str(result.value).split('@@@')


def main() -> None:
    source = _probe()
    answers: list[str] = []
    for first in range(0, len(CASES), 30):
        last = min(first + 29, len(CASES) - 1)
        found = _run(source, first, last)
        if found is None:
            found = [(_run(source, index, index) or ['Excel failed'])[0] for index in range(first, last + 1)]
        answers += found
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, answer] for statements, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for index, (statements, answer) in enumerate(record['cases']):
        print(index, statements[-1][:70], '->', answer)


if __name__ == '__main__':
    main()
