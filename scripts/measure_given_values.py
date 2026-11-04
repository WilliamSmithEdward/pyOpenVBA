"""What Excel does with a Currency, a Date, a Single and ranges given to each settable property of Range, Font, Interior, Border,
Borders, Style and Characters that the checklist calls VERIFIED: the error, and what the property then reads.

    python scripts/measure_given_values.py [seconds]

writes tests/fixtures/given_values.json member by member, and can be run again until "left 0". A member that does not
answer in time is recorded as such. Typed as the class it is reached by, and As Object.
"""
from __future__ import annotations

import csv
import json
import sys
import time
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/given_values.json'
KINDS = ['x = CCur(1.5)', 'x = CCur(-2)', 'x = CCur(1000)', 'x = CCur(0)', 'x = #3/1/2019#', 'x = #12:30:00 PM#', 'x = CSng(0.25)', 'x = CDec(2)',
         'Set x = ws.Range("E1")', 'Set x = ws.Range("G1")', 'Set x = ws.Range("E1:E2")', 'Set x = ws.Range("E1,F1")', 'Set x = ws.Range("G1,E1")', 'Set x = ws.Range("E1:E2,G1")', 'Set x = ws.Range("H1,E1")']
#: How each class is reached from a sheet: the object, and its type.
REACH = {'Range': 'ws.Range("B2")', 'Font': 'ws.Range("B2").Font', 'Interior': 'ws.Range("B2").Interior', 'Border': 'ws.Range("B2").Borders(xlEdgeBottom)', 'Borders': 'ws.Range("B2").Borders',
         'Style': 'wb.Styles(sn)', 'Characters': 'ws.Range("B2").Characters(1, 2)'}
#: Changes what follows, or asks for a window or a printer; the rest of a Range's members are measured.
LEFT_OUT = {('Range', 'Value'), ('Range', 'Value2'), ('Range', 'Formula')}
HELPER = '''Private Function Note(e As Long, v As Variant) As String
    If e <> 0 Then
        Note = "e" & e
    ElseIf IsNull(v) Then
        Note = "Null"
    ElseIf IsEmpty(v) Then
        Note = "Empty"
    ElseIf IsObject(v) Then
        Note = "Object " & TypeName(v)
    ElseIf IsArray(v) Then
        Note = TypeName(v)
    Else
        Note = TypeName(v) & " " & CStr(v)
    End If
End Function
'''


def members() -> list[tuple[str, str]]:
    rows = csv.DictReader((ROOT / 'docs/excel_checklist.csv').open(encoding='utf-8', newline=''))
    return sorted((row['object'], row['member']) for row in rows
                  if row['object'] in REACH and row['kind'] == 'property' and row['status'] == 'VERIFIED' and (row['object'], row['member']) not in LEFT_OUT)


def source(owner: str, member: str, typed: bool) -> str:
    kind = owner if typed else 'Object'
    lines = [HELPER, 'Public Function Probe() As String', f'    Dim wb As Object, ws As Object, o As {kind}, x As Variant, v As Variant, s As String, e As Long, sn As String, n As Long',
             '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    On Error Resume Next']
    for value in KINDS:
        lines += ['    n = n + 1: sn = "Given" & n: wb.Styles.Add sn: Err.Clear', '    Set ws = wb.Worksheets.Add',
                  '    ws.Range("B2").Value = "abc": ws.Range("E1").Value = 1.5: ws.Range("E2").Value = 2: ws.Range("F1").Value = 3: ws.Range("G1").Value = "abc"',
                  f'    Set o = {REACH[owner]}', '    x = Empty: Err.Clear', f'    {value}', f'    o.{member} = x', '    e = Err.Number: Err.Clear',
                  f'    v = Empty: v = o.{member}', '    s = s & "@@" & Note(e, Empty) & "|" & Note(Err.Number, v): Err.Clear']
    lines += ['    wb.Close False', '    Probe = Mid(s, 3)', 'End Function', '']
    return '\n'.join(lines)


def main() -> None:
    budget = float(sys.argv[1]) if len(sys.argv) > 1 else 500.0
    record = json.loads(OUT.read_text(encoding='utf-8')) if OUT.exists() else {'kinds': KINDS, 'answers': {}, 'unanswered': []}
    every = members()
    started = time.monotonic()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        for owner, member in every:
            key = f'{owner}.{member}'
            if key in record['answers'] or key in record['unanswered']:
                continue
            if time.monotonic() - started > budget:
                break
            late = excel.run_vba(source(owner, member, False), 'Probe', (), timeout=40.0)
            typed = excel.run_vba(source(owner, member, True), 'Probe', (), timeout=40.0) if late.ok and not late.dialogs else late
            if not (late.ok and typed.ok) or late.dialogs or typed.dialogs:
                record['unanswered'].append(key)
                print('unanswered', key, late.outcome, typed.outcome, str(typed.message)[:70], flush=True)
            else:
                record['answers'][key] = {'late': str(late.value).split('@@'), 'typed': str(typed.value).split('@@')}
            OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    left = [one for one in every if f'{one[0]}.{one[1]}' not in record['answers'] and f'{one[0]}.{one[1]}' not in record['unanswered']]
    print('measured', len(record['answers']), 'unanswered', len(record['unanswered']), 'left', len(left), flush=True)


if __name__ == '__main__':
    main()
