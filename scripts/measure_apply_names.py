"""Measure Range.ApplyNames: which references in formulas it turns into the names that refer to them.

Each case runs on a new workbook: it defines names, writes formulas, calls ApplyNames and reads the formulas back.

    python scripts/measure_apply_names.py

writes tests/fixtures/apply_names.json, which tests/test_excel_apply_names.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/apply_names.json'
#: Each case: names to define (name, refers-to), formulas by cell, the ApplyNames call on ws, the cells read back.
CASES: list[tuple[list[tuple[str, str]], list[tuple[str, str]], str, str]] = [
    ([('Rate', '=Sheet1!$A$1')], [('C1', '=A1*2'), ('C2', '=$A$1*2'), ('C3', '=A$1+1'), ('C4', '=B1+A1')],
     'ws.Range("C1:C4").ApplyNames', 'C1:C4'),
    ([('Rate', '=Sheet1!$A$1')], [('C1', '=A1*2'), ('C2', '=$A$1*2'), ('C3', '=A$1+1')],
     'ws.Range("C1:C3").ApplyNames IgnoreRelativeAbsolute:=False', 'C1:C3'),
    ([('Rate', '=Sheet1!$A$1'), ('Tax', '=Sheet1!$A$2')], [('C1', '=A1*A2'), ('C2', '=A2-A1')],
     'ws.Range("C1:C2").ApplyNames Names:="Tax"', 'C1:C2'),
    ([('Rate', '=Sheet1!$A$1'), ('Tax', '=Sheet1!$A$2')], [('C1', '=A1*A2')],
     'ws.Range("C1").ApplyNames Names:=Array("Rate", "Tax")', 'C1'),
    ([('Data', '=Sheet1!$A$1:$A$5')], [('C1', '=SUM(A1:A5)'), ('C2', '=SUM(A1:A4)'), ('C3', '=SUM($A$1:$A$5)')],
     'ws.Range("C1:C3").ApplyNames', 'C1:C3'),
    ([('Rate', '=Sheet1!$A$1')], [('C1', '=A1*2'), ('D1', '=A1*3')], 'ws.Range("C1").ApplyNames', 'C1:D1'),
    ([('Rate', '=Sheet1!$A$1')], [('C1', '=Sheet1!A1*2'), ('C2', '="A1"&A1')], 'ws.Range("C1:C2").ApplyNames',
     'C1:C2'),
    ([('Rate', '=Sheet1!$A$1')], [('C1', '=A1*2')], 'ws.Range("C1").ApplyNames Names:="NoSuch"', 'C1'),
    ([], [('C1', '=A1*2')], 'ws.Range("C1").ApplyNames', 'C1'),
    ([('Rate', '=Sheet1!$A$1')], [('C1', '5')], 'ws.Range("C1").ApplyNames', 'C1'),
    ([('Rate', '=Sheet1!$A$1'), ('Rate2', '=Sheet1!$A$1')], [('C1', '=A1*2')], 'ws.Range("C1").ApplyNames', 'C1'),
    ([('Jan', '=Sheet1!$B$1:$B$4'), ('Sales', '=Sheet1!$A$2:$C$2')], [('E1', '=B2*2')],
     'ws.Range("E1").ApplyNames', 'E1'),
    ([('Jan', '=Sheet1!$B$1:$B$4'), ('Sales', '=Sheet1!$A$2:$C$2')], [('E1', '=B2*2')],
     'ws.Range("E1").ApplyNames UseRowColumnNames:=False', 'E1'),
    ([('Rate', '=Sheet1!$A$1')], [('C1', '=A1*2')], 'ws.Range("C1").ApplyNames Order:=xlColumnThenRow', 'C1'),
    ([('Rate', '=Sheet1!$A$1')], [('C1', '=SUM(A1,A1)')], 'ws.Range("C1").ApplyNames', 'C1'),
    ([('Rate', '=Sheet1!$A$1')], [('C1', '=A1:A1')], 'ws.Range("C1").ApplyNames', 'C1'),
    ([('Rate', '=Sheet1!$A$1')], [('C1', '{=A1*2}')], 'ws.Range("C1").ApplyNames', 'C1'),
    ([('Rate', '=Sheet1!$A$1')], [('C1', '=A1*2')], 'ws.Range("C1").ApplyNames AppendLast:=True', 'C1'),
    ([('Low', '=Sheet1!$A$1'), ('low2', '=Sheet1!$A$2')], [('C1', '=A1+A2')], 'ws.Range("C1").ApplyNames', 'C1'),
    ([('Rate', '=0.5')], [('C1', '=0.5*2')], 'ws.Range("C1").ApplyNames', 'C1'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, c As Range, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    ws.Name = "Sheet1"',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (names, formulas, call, read) in enumerate(CASES):
        lines.append(f'    Case {index}')
        for name, refers in names:
            lines.append(f'        wb.Names.Add Name:="{name}", RefersTo:="{refers}"')
        for cell, formula in formulas:
            if formula.startswith('{'):
                lines.append(f'        ws.Range("{cell}").FormulaArray = "{formula[1:-1]}"')
            else:
                lines.append(f'        ws.Range("{cell}").Formula = "{formula.replace(chr(34), chr(34) * 2)}"')
        lines.append(f'        {call}')
        lines.append(f'        For Each c In ws.Range("{read}")')
        lines.append('            s = s & "|" & c.Formula')
        lines.append('        Next')
        lines.append('        Tried = Mid(s, 2)')
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


#: How Excel spells a word a function and a defined name share, and one that names nothing, which ApplyNames'
#: answers rest on: written through Range.Formula, read back.
SPELLING = '''Public Function Spelling() As String
    Dim wb As Workbook, ws As Worksheet, s As String, c As Range
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    wb.Names.Add Name:="Rate", RefersTo:="=" & ws.Name & "!$A$1"
    wb.Names.Add Name:="sumx", RefersTo:="=1"
    ws.Range("C1").Formula = "=Rate*2"
    ws.Range("C2").Formula = "=rate*2"
    ws.Range("C3").Formula = "=RATE*2"
    ws.Range("C4").Formula = "=Sum*2"
    ws.Range("C5").Formula = "=rate(1,2,3)"
    ws.Range("C6").Formula = "=SUMX+1"
    ws.Range("C7").Formula = "=Abs*2"
    For Each c In ws.Range("C1:C7")
        s = s & "|" & c.Formula
    Next
    wb.Close False
    Spelling = Mid(s, 2)
End Function
'''


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        # First, since the last case's name for a constant leaves Excel's automation failing after it.
        spelled = excel.run_vba(SPELLING, 'Spelling', (), timeout=60.0)
        assert spelled.ok, spelled
        result = excel.run_vba(source, 'Probe', (0, len(CASES) - 1), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[names, formulas, call, read, answer]
                        for (names, formulas, call, read), answer in zip(CASES, answers, strict=True)],
              'names_spelling': {'source': SPELLING, 'answer': spelled.value}}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for names, formulas, call, _read, answer in record['cases']:
        print(call[22:60], '|', [f for _, f in formulas], '->', answer)


if __name__ == '__main__':
    main()
