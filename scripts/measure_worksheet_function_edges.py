"""What tests/fixtures/worksheet_function_types.json found wrong, measured with its neighbours: Text's format codes letter by
letter, and the arguments the engine read otherwise, as worksheet formulas and through WorksheetFunction.

    python scripts/measure_worksheet_function_edges.py

writes tests/fixtures/worksheet_function_edges.json, which tests/test_excel_worksheet_function_edges.py replays.
"""
from __future__ import annotations

import json
import string
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/worksheet_function_edges.json'
LETTERS = [*string.ascii_lowercase, *string.ascii_uppercase]
CODES = [one * count for one in LETTERS for count in (1, 2, 3, 4)]
#: A date with a time, a whole date, and a plain number.
TEXTS = [f'=TEXT({value},"{code}")' for value in ('45000.75', '1234.5') for code in CODES]
FORMULAS = [
    '=TEXT(1234.5,TRUE)', '=TEXT(1234.5,FALSE)', '=TEXT(TRUE,"0")', '=TEXT("x","0")', '=TEXT(1234.5,1)',
    '=NUMBERVALUE(TRUE)', '=NUMBERVALUE(FALSE)', '=NUMBERVALUE("1,5",",",".")',
    '=DECIMAL("x",16)', '=DECIMAL("X",16)', '=DECIMAL("x1",16)', '=DECIMAL("0x1F",16)', '=DECIMAL("0X1F",16)', '=DECIMAL("x",10)', '=DECIMAL("0x1",10)', '=DECIMAL("1F",16)',
    '=DECIMAL("y",16)', '=DECIMAL("",16)',
    '=VDB(2,300,3650,0,1)', '=VDB(300,2,3650,0,1)', '=VDB(2400,-1,3650,0,1)', '=VDB(-1,300,3650,0,1)', '=VDB(0,300,3650,0,1)', '=VDB(2,2,3650,0,1)', '=VDB(2,300,10,0,5)',
    '=SYD(30000,-1,10,1)', '=SYD(1,300,10,1)', '=SYD(-1,300,10,1)', '=SYD(30000,0,10,1)',
    '=DDB(30000,-1,10,1)', '=DDB(1,300,10,1)', '=DDB(-1,300,10,1)', '=DB(30000,-1,10,1)', '=DB(1,300,10,1)', '=DB(-1,300,10,1)',
    '=SLN(30000,-1,10)', '=SLN(1,300,10)', '=SLN(-1,300,10)',
    '=GAMMALN(1)', '=GAMMALN(2)', '=GAMMALN(3)', '=GAMMALN(0.5)', '=GAMMALN(1.5)', '=GAMMALN.PRECISE(2)',
    '=ZTEST({3,6,7,8,6,5,4,2,1,9},-1,1)', '=ZTEST({3,6,7,8,6,5,4,2,1,9},0,1)', '=ZTEST({3,6,7,8,6,5,4,2,1,9},1,1)', '=ZTEST({3,6,7,8,6,5,4,2,1,9},4,1)', '=ZTEST({3,6,7,8,6,5,4,2,1,9},100,1)',
    '=MIRR({-120000,39000,30000,21000,37000,46000},-1,0.12)', '=MIRR({-120000,39000,-30000,21000,37000,46000},-1,0.12)', '=MIRR({-120000,39000,30000,21000,37000,46000},0.1,-1)',
    '=PV(-1,240,500,,0)', '=PV(-1,240,500,,1)', '=FV(-1,240,500)', '=PMT(-1,240,500)',
    '=POISSON(0,0,TRUE)', '=POISSON(1,0,TRUE)', '=POISSON(0,0,FALSE)', '=POISSON(1,0,FALSE)',
    '=HYPGEOMDIST(1,0,8,20)', '=RATE(48,0,8000)', '=MUNIT(1)', '=ROWS(MUNIT(1))',
]
#: Through WorksheetFunction, where a cell cannot tell.
CALLS = ['AscW(f.Unichar(2))', 'AscW(f.Unichar(1))', 'AscW(f.Unichar(31))', 'AscW(f.Unichar(65))', 'TypeName(f.MUnit(1))', 'TypeName(f.Unichar(2))',
         'f.Subtotal(9, ws.Range("A1:A3"))', 'f.Subtotal(1, ws.Range("A1:A3"))', 'f.Subtotal(109, ws.Range("A1:A3"))', 'f.Subtotal(9, ws.Range("A1:A3"), ws.Range("A1:A2"))',
         'f.Subtotal(99, ws.Range("A1:A3"))', 'f.Subtotal(0, ws.Range("A1:A3"))', 'f.Subtotal("9", ws.Range("A1:A3"))', 'f.Subtotal(9.9, ws.Range("A1:A3"))', 'f.Subtotal(True, ws.Range("A1:A3"))']


def source(kind: str) -> str:
    lines = ['Private Function Note(e As Long, v As Variant) As String', '    If e <> 0 Then', '        Note = "e" & e', '    ElseIf IsArray(v) Then', '        Note = TypeName(v)', '    ElseIf IsError(v) Then',
             '        Note = CStr(v)', '    Else', '        Note = TypeName(v) & " " & CStr(v)', '    End If', 'End Function', '']
    parts = []
    every = [f'ws.Evaluate("{formula.replace(chr(34), chr(34) * 2)}")' for formula in [*TEXTS, *FORMULAS]] + CALLS
    for number, at in enumerate(range(0, len(every), 150)):
        lines += [f'Private Function Part{number}(f As {kind}, ws As Object) As String', '    Dim v As Variant, s As String', '    On Error Resume Next']
        lines += [f'    v = Empty: Err.Clear: v = {call}: s = s & "@@" & Note(Err.Number, v): Err.Clear' for call in every[at:at + 150]]
        lines += [f'    Part{number} = Mid(s, 3)', 'End Function', '']
        parts.append(f'Part{number}(f, ws)')
    lines += ['Public Function Probe() As String', f'    Dim wb As Object, ws As Object, f As {kind}, s As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
              '    ws.Range("A1").Value = 1: ws.Range("A2").Value = 2: ws.Range("A3").Value = 3', '    Set f = Application.WorksheetFunction',
              f'    s = {" & chr(64) & chr(64) & ".join(parts)}', '    wb.Close False', '    Probe = s', 'End Function', '']
    return '\n'.join(lines)


def main() -> None:
    every = [*TEXTS, *FORMULAS, *CALLS]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        typed = excel.run_vba(source('WorksheetFunction'), 'Probe', (), timeout=120.0)
        assert typed.ok and not typed.dialogs, typed
    answers = str(typed.value).split('@@')
    assert len(answers) == len(every), (len(answers), len(every))
    OUT.write_text(json.dumps({'source': source('WorksheetFunction'), 'cases': [list(pair) for pair in zip(every, answers, strict=True)]}, indent=0) + '\n', encoding='utf-8')
    print(len(every), 'cases')


if __name__ == '__main__':
    main()
