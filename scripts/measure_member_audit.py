"""Settable properties of validation, comments, hyperlinks, tables, sorts and windows, held to the checklist's VERIFIED
standard (docs/host_completeness.md): each given each kind of value on a fresh object and on one already changed, what
the object then reads, and the same after a save and reopen.

Each case builds its object on a sheet of its own. Each answer is the error from the assignment, then every property of
that object, as TypeName and value.

    python scripts/measure_member_audit.py

writes tests/fixtures/member_audit.json, which tests/test_excel_member_audit.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/member_audit.json'
BAD = ['Null', 'Empty', 'True', '"x"', '""']
SWITCHES = ['True', 'False', '1', '0', '-1', '2', '"True"', '"x"', 'Null', 'Empty', '""']
TEXTS = ['"words"', '"é"', '5', '1.5', 'True', 'Null', 'Empty', '""', 'String(32, "a")', 'String(33, "a")',
         'String(225, "a")', 'String(226, "a")', 'String(254, "a")', 'String(255, "a")', 'String(256, "a")',
         'String(2084, "a")', 'String(2085, "a")']
#: Each family: its setup statements on sheet ws, the object, what is read back, and what a changed one has first.
FAMILIES = {
    'Validation': ('ws.Range("B2").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "a,b"',
                   'ws.Range("B2").Validation',
                   ['AlertStyle', 'ErrorMessage', 'ErrorTitle', 'Formula1', 'IgnoreBlank', 'InCellDropdown',
                    'InputMessage', 'InputTitle', 'ShowError', 'ShowInput', 'Type'],
                   ['', 'o.ErrorMessage = "e": o.ErrorTitle = "t": o.InputMessage = "i": o.InputTitle = "n": '
                        'o.IgnoreBlank = False: o.InCellDropdown = False: o.ShowError = False: o.ShowInput = False']),
    'Comment': ('ws.Range("B2").AddComment "note"', 'ws.Range("B2").Comment', ['Visible', 'Text'],
                ['', 'o.Visible = True']),
    'Hyperlink': ('ws.Hyperlinks.Add ws.Range("B2"), "http://example.com/"', 'ws.Hyperlinks(1)',
                  ['Address', 'SubAddress', 'ScreenTip', 'TextToDisplay', 'Name'],
                  ['', 'o.SubAddress = "s": o.ScreenTip = "t": o.TextToDisplay = "d"']),
    'ListColumn': ('ws.Range("B2:C2").Value = Array("x", "y"): ws.Range("B3:C3").Value = Array(1, 2): '
                   'ws.ListObjects.Add xlSrcRange, ws.Range("B2:C3"), , xlYes: ws.ListObjects(1).ShowTotals = True',
                   'ws.ListObjects(1).ListColumns(1)', ['Name', 'TotalsCalculation', 'Index'],
                   ['', 'o.Name = "Named": o.TotalsCalculation = xlTotalsCalculationSum']),
    'ListObject': ('ws.Range("B2:C2").Value = Array("x", "y"): ws.Range("B3:C3").Value = Array(1, 2): '
                   'ws.ListObjects.Add xlSrcRange, ws.Range("B2:C3"), , xlYes', 'ws.ListObjects(1)',
                   ['ShowTotals', 'ShowHeaders', 'Name'], ['', 'o.ShowTotals = True']),
    'Sort': ('ws.Range("B2:B3").Value = 1: ws.Sort.SetRange ws.Range("B2:B3")', 'ws.Sort', ['MatchCase', 'Header'],
             ['', 'o.MatchCase = True']),
}
VALUES = {
    'Validation': {'ErrorMessage': TEXTS, 'ErrorTitle': TEXTS, 'InputMessage': TEXTS, 'InputTitle': TEXTS,
                   'IgnoreBlank': SWITCHES, 'InCellDropdown': SWITCHES, 'ShowError': SWITCHES, 'ShowInput': SWITCHES},
    'Comment': {'Visible': SWITCHES},
    'Hyperlink': {'SubAddress': TEXTS, 'ScreenTip': TEXTS, 'TextToDisplay': TEXTS},
    'ListColumn': {'Name': TEXTS, 'TotalsCalculation': ['xlTotalsCalculationNone', 'xlTotalsCalculationSum',
                                                        'xlTotalsCalculationCustom', '11', '-1', '99', '"1"', *BAD]},
    'ListObject': {'ShowTotals': SWITCHES},
    'Sort': {'MatchCase': SWITCHES},
}


def cases() -> list[tuple[str, str, str, str]]:
    """Each (family, what the object has first, property, value)."""
    return [(family, first, prop, value) for family, values in VALUES.items() for prop, tried in values.items()
            for first in FAMILIES[family][3] for value in tried]


def _state(family: str) -> str:
    _, target, props, _ = FAMILIES[family]
    reads = ''.join(f'    v = Empty\n    v = o.{prop}\n    If Err.Number <> 0 Then\n        s = s & "|e" & Err.Number\n'
                    f'        Err.Clear\n    Else\n        s = s & "|" & Show(v)\n    End If\n' for prop in props)
    return (f'\nPrivate Function State{family}(ws As Object) As String\n    Dim s As String, v As Variant, o As Object\n'
            f'    On Error Resume Next\n    Set o = {target}\n{reads}    State{family} = s\nEnd Function\n')


def source() -> str:
    """Setup makes the book; S(i) makes case i on a sheet of its own; Reopen saves; A(i) reads the case again."""
    every = cases()
    body = ('Private wb As Object\n\nPublic Sub Setup()\n    Set wb = Workbooks.Add(xlWBATWorksheet)\nEnd Sub\n'
            '\nPublic Sub Reopen(ByVal folder As String)\n    Dim path As String\n'
            '    path = folder & Application.PathSeparator & "audit.xlsx"\n'
            '    wb.SaveAs path, xlOpenXMLWorkbook\n    wb.Close False\n    Set wb = Workbooks.Open(path)\nEnd Sub\n'
            '\nPublic Sub Finish()\n    wb.Close False\nEnd Sub\n'
            '\nPrivate Function Show(v As Variant) As String\n    If IsNull(v) Then\n        Show = "Null"\n'
            '    ElseIf IsEmpty(v) Then\n        Show = "Empty"\n    ElseIf Len(CStr(v)) > 40 Then\n'
            '        Show = TypeName(v) & " " & Len(CStr(v))\n    Else\n        Show = TypeName(v) & " " & CStr(v)\n'
            '    End If\nEnd Function\n')
    body += ''.join(_state(family) for family in FAMILIES)
    for number, (family, first, prop, value) in enumerate(every):
        setup, target, _, _ = FAMILIES[family]
        body += (f'\nPublic Function S{number}() As String\n    Dim ws As Object, o As Object, e As Long\n'
                 f'    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))\n'
                 f'    {setup}\n    Set o = {target}\n'
                 + (f'    {first}\n' if first else '')
                 + f'    On Error Resume Next\n    o.{prop} = {value}\n    e = Err.Number\n    Err.Clear\n'
                 f'    S{number} = "e" & e & State{family}(ws)\nEnd Function\n'
                 f'\nPublic Function A{number}() As String\n    A{number} = State{family}(wb.Worksheets({number + 2}))\n'
                 f'End Function\n')
    return body


def _runner(calls: list[str], name: str) -> str:
    """A function answering each of the calls joined by @@, split up since one procedure cannot hold them all."""
    parts = [calls[at:at + 300] for at in range(0, len(calls), 300)]
    body = ''
    for number, part in enumerate(parts):
        lines = ''.join(f'    out = out & "@@" & {call}()\n' for call in part)
        body += (f'\nPrivate Function {name}{number}() As String\n    Dim out As String\n{lines}'
                 f'    {name}{number} = Mid(out, 3)\nEnd Function\n')
    joined = ' & "@@" & '.join(f'{name}{number}()' for number in range(len(parts)))
    return body + f'\nPrivate Function {name}() As String\n    {name} = {joined}\nEnd Function\n'


def main() -> None:
    every = cases()
    probe = source()
    runner = _runner([f'S{number}' for number in range(len(every))], 'First')
    runner += _runner([f'A{number}' for number in range(len(every))], 'Second')
    runner += ('\nPublic Function Probe(ByVal folder As String) As String\n    Dim out As String\n    Setup\n'
               '    out = First()\n    Reopen folder\n    out = out & "@@" & Second()\n    Finish\n'
               '    Probe = out\nEnd Function\n')
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        result = excel.run_vba(probe + runner, 'Probe', (folder,), timeout=500.0)
        assert result.ok and not result.dialogs, result
    answers = str(result.value).split('@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': probe,
              'states': {family: FAMILIES[family][2] for family in FAMILIES},
              'cases': [[*case, answers[number], answers[len(every) + number]] for number, case in enumerate(every)]}
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    print(len(every), 'cases')


if __name__ == '__main__':
    main()
