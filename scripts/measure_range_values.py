"""Range.Value, Value2, Formula, NumberFormat and Style, held to the checklist's VERIFIED standard
(docs/host_completeness.md): each given each kind of value, on one cell and on two, fresh and already formatted, with
Range.Text read in each state. What the cells then read, and the same after a save and reopen.

Each answer is the error from the assignment, then what the first cell reads, as TypeName and value.

    python scripts/measure_range_values.py

writes tests/fixtures/range_values.json, which tests/test_excel_range_values.py replays. Its "early" holds the answers
that differ with c typed As Range.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from measure_format_properties import _runner
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_values.json'
READS = ['Value', 'Value2', 'Formula', 'NumberFormat', 'Text', 'HasFormula', 'Style']
VALUES = ['5', '1.5', '-1', '0', '1E+308', '123456789012345678', 'True', 'False', 'Null', 'Empty', '""', '"text"', '"5"', '"1.5"',
          '"=1+1"', '"\'x"', '"\'5"', '"TRUE"', '"1/2/2020"', '"1,000"', '"50%"', '"$5"', '" 5 "', '"#N/A"', '"=SUM("',
          '#1/2/2020#', '#1/2/2020 3:04:05 PM#', 'CCur(1.5)', 'CDec(1.5)', 'CVErr(xlErrNA)', 'Array(1, 2)',
          'String(255, "a")', 'String(32767, "a")', 'String(32768, "a")', 'Application']
FORMULAS = ['"=1+1"', '"=SUM(1,2)"', '"=A1"', '"=NoSuchName"', '"=#REF!"', '"=SUM("', '"=1+"', '"text"', '"5"', '"\'x"', '5', '1.5',
            'True', 'False', 'Null', 'Empty', '""', 'CVErr(xlErrNA)', '#1/2/2020#', 'Array("=1", "=2")', '"=SEQUENCE(2)"',
            '"=" & String(8192, "1")', '"=" & String(8193, "1")', 'Application']
FORMATS = ['"General"', '"0.00"', '"@"', '"yyyy-mm-dd"', '"0.00;[Red]-0.00"', '"""x"""', '"[Red]0"', '"[h]:mm"', '"#,##0"', '"0%"',
           '"m/d/yyyy"', '"bad["', '"0.0.0"', '""', '" "', '5', '1.5', 'True', 'False', 'Null', 'Empty', 'String(255, "0")',
           'String(256, "0")', 'Application']
STYLES = ['"Normal"', '"Good"', '"Currency"', '"normal"', '"No Such Style"', '""', '5', 'True', 'Null', 'Empty', 'wb.Styles(1)',
          'wb.Styles("Good")', 'Application']
MEMBERS = {'Value': VALUES, 'Value2': VALUES, 'Formula': FORMULAS, 'NumberFormat': FORMATS, 'Style': STYLES}
FIRST = ['', 'c.Cells(1, 1).Value = 7: c.NumberFormat = "0.00": c.Font.Bold = True']
TARGETS = {'one': 'ws.Range("B2")', 'two': 'ws.Range("B2:C2")'}
#: A read of Text is decided by the column's width where a number does not fit, which follows the display's DPI, so
#: the values that overflow it are read with their column wide enough to hold them.
WIDE = 'ws.Columns("B:C").ColumnWidth = 60'


def cases() -> list[tuple[str, str, str, str]]:
    return [(member, target, first, value) for member, values in MEMBERS.items() for target in TARGETS
            for first in FIRST for value in values]


def _state() -> str:
    reads = ''.join(f'    v = Empty\n    v = c.Cells(1, 1).{read}\n    If Err.Number <> 0 Then\n        s = s & "|e" & Err.Number\n'
                    f'        Err.Clear\n    Else\n        s = s & "|" & Show(v)\n    End If\n' for read in READS)
    return ('\nPrivate Function State(c As Object) As String\n    Dim s As String, v As Variant\n    On Error Resume Next\n'
            + reads + '    State = s\nEnd Function\n')


def source(typed: bool = False) -> str:
    kind = 'Range' if typed else 'Object'
    body = ('Private wb As Object\n\nPublic Sub Setup()\n    Set wb = Workbooks.Add(xlWBATWorksheet)\nEnd Sub\n'
            '\nPublic Sub Reopen(ByVal folder As String)\n    Dim path As String\n'
            '    path = folder & Application.PathSeparator & "values.xlsx"\n'
            '    wb.SaveAs path, xlOpenXMLWorkbook\n    wb.Close False\n    Set wb = Workbooks.Open(path)\nEnd Sub\n'
            '\nPublic Sub Finish()\n    wb.Close False\nEnd Sub\n'
            '\nPrivate Function Show(v As Variant) As String\n    If IsNull(v) Then\n        Show = "Null"\n'
            '    ElseIf IsEmpty(v) Then\n        Show = "Empty"\n    ElseIf IsObject(v) Then\n        Show = "Object " & TypeName(v)\n'
            '    ElseIf IsArray(v) Then\n        Show = TypeName(v)\n    ElseIf IsError(v) Then\n        Show = "Error " & CStr(v)\n'
            '    ElseIf Len(CStr(v)) > 60 Then\n        Show = TypeName(v) & " " & Len(CStr(v))\n'
            '    Else\n        Show = TypeName(v) & " " & CStr(v)\n    End If\nEnd Function\n')
    body += _state()
    for number, (member, target, first, value) in enumerate(cases()):
        body += (f'\nPublic Function S{number}() As String\n    Dim ws As Object, c As {kind}, e As Long\n'
                 f'    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))\n    {WIDE}\n'
                 f'    Set c = {TARGETS[target]}\n' + (f'    {first}\n' if first else '')
                 + f'    On Error Resume Next\n    c.{member} = {value}\n    e = Err.Number\n    Err.Clear\n'
                 f'    S{number} = "e" & e & State(c)\nEnd Function\n'
                 f'\nPublic Function A{number}() As String\n'
                 f'    A{number} = State(wb.Worksheets({number + 2}).Range("B2"))\nEnd Function\n')
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
        result = excel.run_vba(source() + runner, 'Probe', (folder,), timeout=500.0)
        assert result.ok and not result.dialogs, result
        bound = excel.run_vba(source(typed=True) + only, 'Probe', (), timeout=500.0)
        assert bound.ok and not bound.dialogs, bound
    answers = str(result.value).split('@@')
    early = str(bound.value).split('@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source(), 'reads': READS,
              'early': {str(number): early[number] for number in range(count) if early[number] != answers[number]},
              'cases': [[*every[number], answers[number], answers[count + number]] for number in range(count)]}
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    print(count, 'cases', len(record['early']), 'differ when typed')


if __name__ == '__main__':
    main()
