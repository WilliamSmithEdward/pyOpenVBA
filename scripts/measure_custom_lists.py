"""Application custom lists: CustomListCount, GetCustomListContents, GetCustomListNum, AddCustomList and
DeleteCustomList, and the lists at work in AutoFill and in Range.Sort's OrderCustom.

The lists are Excel's own, kept for the user, so the probe records how many there were and deletes every list it
added before it returns, whatever happened.

    python scripts/measure_custom_lists.py

writes tests/fixtures/custom_lists.json, which tests/test_excel_custom_lists.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/custom_lists.json'

#: Each step: a label, statements run first, and the expression recorded. ``n`` is the count the probe began with.
STEPS = [
    ('count', [], 'CStr(Application.CustomListCount)'),
    ('list 1', [], 'Shown(Application.GetCustomListContents(1))'),
    ('list 2', [], 'Shown(Application.GetCustomListContents(2))'),
    ('list 3', [], 'Shown(Application.GetCustomListContents(3))'),
    ('list 4', [], 'Shown(Application.GetCustomListContents(4))'),
    ('num days', [], 'CStr(Application.GetCustomListNum(Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")))'),
    ('num lower', [], 'CStr(Application.GetCustomListNum(Array("sun", "mon", "tue", "wed", "thu", "fri", "sat")))'),
    ('num part', [], 'CStr(Application.GetCustomListNum(Array("Sun", "Mon")))'),
    ('num none', [], 'CStr(Application.GetCustomListNum(Array("x", "y")))'),
    ('add', ['Application.AddCustomList Array("Low", "Medium", "High")'], 'CStr(Application.CustomListCount - n)'),
    ('num added', [], 'CStr(Application.GetCustomListNum(Array("Low", "Medium", "High")) - n)'),
    ('added', [], 'Shown(Application.GetCustomListContents(n + 1))'),
    ('add again', ['Application.AddCustomList Array("Low", "Medium", "High")'],
     'CStr(Application.CustomListCount - n)'),
    ('add lower', ['Application.AddCustomList Array("low", "medium", "high")'],
     'CStr(Application.CustomListCount - n)'),
    ('add odd', ['Application.AddCustomList Array(" Alpha ", "Beta,Gamma", 3, "", "Delta")'],
     'CStr(Application.CustomListCount - n) & " " & Shown(Application.GetCustomListContents(Application.CustomListCount))'),
    ('add range', ['ws.Range("A1:A3").Value = Application.Transpose(Array("x1", "y1", "z1"))',
                   'Application.AddCustomList ws.Range("A1:A3")'],
     'Shown(Application.GetCustomListContents(Application.CustomListCount))'),
    ('add block', ['ws.Range("B1:D2").Value = Array("b1", "c1", "d1")', 'ws.Range("B2:D2").Value = Array("b2", "c2", "d2")',
                   'Application.AddCustomList ws.Range("B1:D2")'],
     'Shown(Application.GetCustomListContents(Application.CustomListCount))'),
    ('add block rows', ['Application.AddCustomList ws.Range("B1:D2"), True'],
     'Shown(Application.GetCustomListContents(Application.CustomListCount))'),
    ('add one', ['Application.AddCustomList Array("Only")'], 'CStr(Application.CustomListCount - n)'),
    ('add number', ['ws.Range("E1:E2").Value = Application.Transpose(Array(1, 2))',
                    'Application.AddCustomList ws.Range("E1:E2")'],
     'CStr(Application.CustomListCount - n) & " " & Shown(Application.GetCustomListContents(Application.CustomListCount))'),
    ('delete built in', ['Application.DeleteCustomList 1'], 'CStr(Application.CustomListCount - n)'),
    ('delete past', ['Application.DeleteCustomList Application.CustomListCount + 1'], '"done"'),
    ('contents 0', [], 'Shown(Application.GetCustomListContents(0))'),
    ('contents past', [], 'Shown(Application.GetCustomListContents(Application.CustomListCount + 1))'),
    ('fill', ['ws.Range("F1").Value = "Medium"', 'ws.Range("F1").AutoFill ws.Range("F1:F5")'], 'Column(ws, "F1:F5")'),
    ('fill lower', ['ws.Range("G1").Value = "medium"', 'ws.Range("G1").AutoFill ws.Range("G1:G4")'],
     'Column(ws, "G1:G4")'),
    ('fill two', ['ws.Range("H1").Value = "Low"', 'ws.Range("H2").Value = "High"', 'ws.Range("H1:H2").AutoFill ws.Range("H1:H6")'],
     'Column(ws, "H1:H6")'),
    ('sort', ['ws.Range("I1:I7").Value = Application.Transpose(Array("High", "low", "Zeta", "Medium", 5, "high", "Apple"))',
              'ws.Range("I1:I7").Sort Key1:=ws.Range("I1"), Order1:=xlAscending, Header:=xlNo, OrderCustom:=n + 2'],
     'Column(ws, "I1:I7")'),
    ('sort down', ['ws.Range("J1:J7").Value = Application.Transpose(Array("High", "low", "Zeta", "Medium", 5, "high", "Apple"))',
                   'ws.Range("J1:J7").Sort Key1:=ws.Range("J1"), Order1:=xlDescending, Header:=xlNo, OrderCustom:=n + 2'],
     'Column(ws, "J1:J7")'),
    ('sort days', ['ws.Range("K1:K5").Value = Application.Transpose(Array("Wed", "Sunday", "Mon", "sat", "Tue"))',
                   'ws.Range("K1:K5").Sort Key1:=ws.Range("K1"), Order1:=xlAscending, Header:=xlNo, OrderCustom:=2'],
     'Column(ws, "K1:K5")'),
    ('sort past', ['ws.Range("L1:L2").Value = Application.Transpose(Array("b", "a"))',
                   'ws.Range("L1:L2").Sort Key1:=ws.Range("L1"), Header:=xlNo, OrderCustom:=Application.CustomListCount + 2'],
     'Column(ws, "L1:L2")'),
    ('count now', [], 'CStr(Application.CustomListCount - n)'),
    ('odd items', ['Application.AddCustomList Array("p1", 3, "p2")'],
     'CStr(Application.CustomListCount - n) & " " & Shown(Application.GetCustomListContents(Application.CustomListCount))'),
    ('empty item', ['Application.AddCustomList Array("q1", "", "q2")'],
     'CStr(Application.CustomListCount - n) & " " & Shown(Application.GetCustomListContents(Application.CustomListCount))'),
    ('numbers only', ['Application.AddCustomList Array(1, 2)'], 'CStr(Application.CustomListCount - n)'),
    ('block cols', ['ws.Range("M1:N3").Value = Application.Transpose(Application.Transpose(Array("m1", "m2", "m3")))',
                    'ws.Range("M1").Value = "m1"', 'ws.Range("M2").Value = "m2"', 'ws.Range("M3").Value = "m3"',
                    'ws.Range("N1").Value = "n1"', 'ws.Range("N2").Value = "n2"', 'ws.Range("N3").Value = "n3"',
                    'Application.AddCustomList ws.Range("M1:N3")'],
     'CStr(Application.CustomListCount - n) & " " & Shown(Application.GetCustomListContents(Application.CustomListCount - 1)) '
     '& " / " & Shown(Application.GetCustomListContents(Application.CustomListCount))'),
    ('block cols false', ['ws.Range("O1").Value = "o1"', 'ws.Range("O2").Value = "o2"', 'ws.Range("P1").Value = "p1x"',
                          'ws.Range("P2").Value = "p2x"', 'Application.AddCustomList ws.Range("O1:P2"), False'],
     'CStr(Application.CustomListCount - n) & " " & Shown(Application.GetCustomListContents(Application.CustomListCount - 1)) '
     '& " / " & Shown(Application.GetCustomListContents(Application.CustomListCount))'),
    ('block rows true', ['ws.Range("Q1").Value = "q1x"', 'ws.Range("Q2").Value = "q2x"', 'ws.Range("R1").Value = "r1x"',
                         'ws.Range("R2").Value = "r2x"', 'Application.AddCustomList ws.Range("Q1:R2"), True'],
     'CStr(Application.CustomListCount - n) & " " & Shown(Application.GetCustomListContents(Application.CustomListCount - 1)) '
     '& " / " & Shown(Application.GetCustomListContents(Application.CustomListCount))'),
    ('range again', ['Application.AddCustomList ws.Range("A1:A3")'], 'CStr(Application.CustomListCount - n)'),
    ('range blank', ['ws.Range("S1").Value = "s1"', 'ws.Range("S3").Value = "s3"', 'Application.AddCustomList ws.Range("S1:S3")'],
     'CStr(Application.CustomListCount - n) & " " & Shown(Application.GetCustomListContents(Application.CustomListCount))'),
    ('num range', [], 'CStr(Application.GetCustomListNum(ws.Range("A1:A3")) - n)'),
    ('num other case', [], 'CStr(Application.GetCustomListNum(Array("LOW", "medium", "HIGH")) - n)'),
    ('fill upper', ['ws.Range("T1").Value = "MEDIUM"', 'ws.Range("T1").AutoFill ws.Range("T1:T3")'], 'Column(ws, "T1:T3")'),
    ('fill mixed', ['ws.Range("U1").Value = "mEDium"', 'ws.Range("U1").AutoFill ws.Range("U1:U3")'], 'Column(ws, "U1:U3")'),
    ('fill up', ['ws.Range("V5").Value = "Low"', 'ws.Range("V5").AutoFill ws.Range("V2:V5")'], 'Column(ws, "V2:V5")'),
    ('fill copy', ['ws.Range("W1").Value = "Low"', 'ws.Range("W1").AutoFill ws.Range("W1:W3"), xlFillCopy'],
     'Column(ws, "W1:W3")'),
    ('fill skip', ['ws.Range("X1").Value = "Low"', 'ws.Range("X2").Value = "Medium"', 'ws.Range("X3").Value = "Low"',
                   'ws.Range("X1:X3").AutoFill ws.Range("X1:X6")'], 'Column(ws, "X1:X6")'),
    ('fill spaced', ['ws.Range("Y1").Value = " Alpha "', 'ws.Range("Y1").AutoFill ws.Range("Y1:Y4")'],
     'Column(ws, "Y1:Y4")'),
    ('delete added', ['Application.DeleteCustomList n + 1'],
     'CStr(Application.CustomListCount - n) & " " & Shown(Application.GetCustomListContents(n + 1))'),
]

HELPERS = '''Private Function Shown(ByVal items As Variant) As String
    Dim i As Long, s As String
    For i = LBound(items) To UBound(items)
        s = s & "|" & TypeName(items(i)) & ":" & items(i)
    Next
    Shown = LBound(items) & "-" & UBound(items) & " " & Mid(s, 2)
End Function

Private Function Column(ByVal ws As Worksheet, ByVal address As String) As String
    Dim c As Range, s As String
    For Each c In ws.Range(address)
        s = s & "|" & c.Formula
    Next
    Column = Mid(s, 2)
End Function
'''


def _probe() -> str:
    lines = [HELPERS, 'Public Function Probe() As String',
             '    Dim wb As Workbook, ws As Worksheet, n As Long, s As String, v As String, i As Long',
             '    n = Application.CustomListCount', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    On Error Resume Next']
    for label, statements, read in STEPS:
        lines += ['    Err.Clear', '    v = ""']
        lines += [f'    {statement}' for statement in statements]
        lines += [f'    If Err.Number = 0 Then v = {read}',
                  '    If Err.Number <> 0 Then v = "Err " & Err.Number & " " & Err.Description',
                  f'    s = s & "@@@{label}=" & v']
    lines += ['    For i = Application.CustomListCount To n + 1 Step -1', '        Application.DeleteCustomList i',
              '    Next', '    s = s & "@@@restored=" & (Application.CustomListCount = n)', '    wb.Close False',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = dict(one.split('=', 1) for one in str(result.value).split('@@@'))
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'answers': answers}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for label, answer in answers.items():
        print(label, '->', answer)


if __name__ == '__main__':
    main()
