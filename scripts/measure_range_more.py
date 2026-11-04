"""The rest of a Range's settable properties, held to the checklist's VERIFIED standard (docs/host_completeness.md): each given
each kind of value on the ranges it applies to, fresh and already changed, late and early bound; what the range then
reads, and the same after a save and reopen.

    python scripts/measure_range_more.py

writes tests/fixtures/range_more.json, which tests/test_excel_range_more.py replays. Its "early" holds the answers that
differ with c typed As Range. Refuses to run off 96 DPI, since StandardHeight and the rows' sizes follow the display.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from measure_format_properties import _runner
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_more.json'
BAD = ['Null', 'Empty', 'True', '"x"', '""']
SWITCHES = ['True', 'False', '1', '0', '-1', '2', '"True"', '"x"', 'Null', 'Empty', '""']
FORMULAS = ['"=1+1"', '"=SUM(1,2)"', '"=A1"', '"=R1C1"', '"=RC[-1]"', '"=NoSuchName"', '"=SUM("', '"text"', '"5"', '5', '1.5', 'True', 'Null',
            'Empty', '""', 'CVErr(xlErrNA)', '"=SEQUENCE(2)"']
FORMATS = ['"General"', '"0.00"', '"@"', '""', 'True', 'Null', 'Empty', '5', '"bad["', 'String(256, "0")']
NAMES = ['"Named"', '"=B2"', '"1bad"', '"has space"', '""', '5', 'True', 'Null', 'Empty', '"Sheet1!Local"', 'String(255, "a")', 'String(256, "a")']
LEVELS = ['1', '2', '8', '9', '0', '-1', '1.5', '"2"', *BAD]
BREAKS = ['xlPageBreakNone', 'xlPageBreakManual', 'xlPageBreakAutomatic', '0', '1', '99', '"1"', *BAD]
C2 = 'ws.Range("C2")'
B2C3 = 'ws.Range("B2:C3")'
ROW = 'ws.Rows(2)'
COLUMN = 'ws.Columns(2)'
#: (member, the range, what it has first, values, what to read after).
MEMBERS = [
    ('Formula2', C2, FORMULAS, ['Formula2', 'Value', 'HasFormula', 'HasSpill']),
    ('FormulaLocal', C2, FORMULAS, ['FormulaLocal', 'Formula', 'Value', 'HasFormula']),
    ('FormulaR1C1', C2, FORMULAS, ['FormulaR1C1', 'Formula', 'Value', 'HasFormula']),
    ('FormulaR1C1Local', C2, FORMULAS, ['FormulaR1C1Local', 'Formula', 'Value', 'HasFormula']),
    ('FormulaArray', C2, FORMULAS, ['Formula', 'Value', 'HasArray']),
    ('FormulaArray', 'ws.Range("C2:C3")', FORMULAS, ['Formula', 'Value', 'HasArray']),
    ('NumberFormatLocal', C2, FORMATS, ['NumberFormatLocal', 'NumberFormat']),
    ('MergeCells', B2C3, SWITCHES, ['MergeCells', 'Address']),
    ('MergeCells', C2, SWITCHES, ['MergeCells', 'Address']),
    ('Name', C2, NAMES, ['Name']),
    ('ID', C2, ['"anid"', '""', '5', 'True', 'Null', 'Empty'], ['ID']),
    ('Hidden', ROW, SWITCHES, ['Hidden']),
    ('Hidden', COLUMN, SWITCHES, ['Hidden']),
    ('Hidden', C2, ['True', 'False'], ['Hidden']),
    ('OutlineLevel', ROW, LEVELS, ['OutlineLevel']),
    ('OutlineLevel', COLUMN, LEVELS, ['OutlineLevel']),
    ('OutlineLevel', C2, ['1', '2'], ['OutlineLevel']),
    ('PageBreak', ROW, BREAKS, ['PageBreak']),
    ('PageBreak', COLUMN, BREAKS, ['PageBreak']),
    ('PageBreak', C2, ['xlPageBreakManual', 'xlPageBreakNone'], ['PageBreak']),
    ('UseStandardHeight', ROW, SWITCHES, ['UseStandardHeight']),
    ('UseStandardHeight', C2, ['True', 'False', 'Null'], ['UseStandardHeight']),
    ('UseStandardWidth', COLUMN, SWITCHES, ['UseStandardWidth']),
    ('UseStandardWidth', C2, ['True', 'False', 'Null'], ['UseStandardWidth']),
]
FIRST = {'Hidden': ['', 'c.Hidden = True'], 'MergeCells': ['', 'c.MergeCells = True'], 'OutlineLevel': ['', 'c.OutlineLevel = 3'],
         'PageBreak': ['', 'c.PageBreak = xlPageBreakManual'], 'UseStandardHeight': ['', 'c.RowHeight = 40'],
         'UseStandardWidth': ['', 'c.ColumnWidth = 30'], 'Name': ['', 'c.Name = "Before"'], 'ID': ['', 'c.ID = "before"'],
         'NumberFormatLocal': ['', 'c.NumberFormat = "0.00"']}


def cases() -> list[tuple[str, str, str, str, list[str]]]:
    return [(member, target, first, value, reads) for member, target, values, reads in MEMBERS
            for first in FIRST.get(member, ['', 'c.Formula = "=2+2"']) for value in values]


def source(typed: bool = False) -> str:
    kind = 'Range' if typed else 'Object'
    body = ('Private wb As Object\n\nPublic Sub Setup()\n    Set wb = Workbooks.Add(xlWBATWorksheet)\nEnd Sub\n'
            '\nPublic Sub Reopen(ByVal folder As String)\n    Dim path As String\n'
            '    path = folder & Application.PathSeparator & "more.xlsx"\n'
            '    wb.SaveAs path, xlOpenXMLWorkbook\n    wb.Close False\n    Set wb = Workbooks.Open(path)\nEnd Sub\n'
            '\nPublic Sub Finish()\n    wb.Close False\nEnd Sub\n'
            '\nPrivate Function Show(v As Variant) As String\n    If IsNull(v) Then\n        Show = "Null"\n'
            '    ElseIf IsEmpty(v) Then\n        Show = "Empty"\n    ElseIf IsObject(v) Then\n        Show = "Object " & TypeName(v)\n'
            '    ElseIf IsArray(v) Then\n        Show = TypeName(v)\n    ElseIf IsError(v) Then\n        Show = "Error " & CStr(v)\n'
            '    ElseIf Len(CStr(v)) > 60 Then\n        Show = TypeName(v) & " " & Len(CStr(v))\n'
            '    Else\n        Show = TypeName(v) & " " & CStr(v)\n    End If\nEnd Function\n')
    for number, (member, target, first, value, reads) in enumerate(cases()):
        read = ''.join(f'    v = Empty\n    v = c.{one}\n    If Err.Number <> 0 Then\n        s = s & "|e" & Err.Number\n'
                       f'        Err.Clear\n    Else\n        s = s & "|" & Show(v)\n    End If\n' for one in reads)
        body += (f'\nPrivate Function R{number}(ws As Object) As String\n    Dim c As Object, s As String, v As Variant\n'
                 f'    On Error Resume Next\n    Set c = {target}\n{read}'
                 f'    R{number} = Replace(s, "\'" & ws.Name & "\'!", "<sheet>!")\nEnd Function\n'
                 f'\nPublic Function S{number}() As String\n    Dim ws As Object, c As {kind}, e As Long\n'
                 f'    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))\n'
                 f'    ws.Range("B2").Value = 1\n    ws.Range("A1").Value = 3\n    Set c = {target}\n'
                 + '    On Error Resume Next\n' + (f'    {first}\n    Err.Clear\n' if first else '')
                 + f'    c.{member} = {value}\n    e = Err.Number\n    Err.Clear\n'
                 f'    S{number} = Replace("e" & e & R{number}(ws), wb.Name, "<book>")\nEnd Function\n'
                 f'\nPublic Function A{number}() As String\n    A{number} = Replace(R{number}(wb.Worksheets({number + 2})), wb.Name, "<book>")\nEnd Function\n')
    return body


def main() -> None:
    every = cases()
    count = len(every)
    runner = _runner([f'S{number}' for number in range(count)], 'First') + _runner([f'A{number}' for number in range(count)], 'Second')
    runner += ('\nPublic Function Probe(ByVal folder As String) As String\n    Dim out As String\n    Setup\n'
               '    out = First()\n    Reopen folder\n    out = out & "@@" & Second()\n    Finish\n    Probe = out\nEnd Function\n')
    only = _runner([f'S{number}' for number in range(count)], 'First') + (
        '\nPublic Function Probe() As String\n    Setup\n    Probe = First()\n    Finish\nEnd Function\n')
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        standard = excel.run_vba('Public Function Probe() As Double\n    Probe = ActiveSheet.StandardHeight\nEnd Function\n', 'Probe', (), timeout=30.0)
        assert float(str(standard.value)) == 15.0, f'not at 96 DPI: StandardHeight {standard.value}'
        result = excel.run_vba(source() + runner, 'Probe', (folder,), timeout=500.0)
        assert result.ok and not result.dialogs, result
        bound = excel.run_vba(source(typed=True) + only, 'Probe', (), timeout=500.0)
        assert bound.ok and not bound.dialogs, bound
    # A format's text can end in @, so a case begins at the @@ before what a case begins with.
    answers = re.split(r'@@(?=e-?\d|\|)', str(result.value))
    early = re.split(r'@@(?=e-?\d|\|)', str(bound.value))
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source(),
              'early': {str(number): early[number] for number in range(count) if early[number] != answers[number]},
              'cases': [[every[number][0], every[number][1], every[number][2], every[number][3], answers[number], answers[count + number]]
                        for number in range(count)]}
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    print(count, 'cases', len(record['early']), 'differ when typed')


if __name__ == '__main__':
    main()
