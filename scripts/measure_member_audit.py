"""Settable properties of validation, comments, hyperlinks, tables, sorts and windows, held to the checklist's VERIFIED
standard (docs/host_completeness.md): each given each kind of value on a fresh object and on one already changed, what
the object then reads, and the same after a save and reopen.

Each case builds its object on a sheet of its own. Each answer is the error from the assignment, then every property of
that object, as TypeName and value.

    python scripts/measure_member_audit.py

writes tests/fixtures/member_audit.json, which tests/test_excel_member_audit.py replays. Its "early" holds the answers
that differ with o typed, which the model follows.
"""
from __future__ import annotations

import json
import re
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
_WINDOW_FRESH = ('o.Caption = "Cap": o.DisplayFormulas = False: o.DisplayGridlines = True: o.DisplayHeadings = True: '
                 'o.DisplayHorizontalScrollBar = True: o.DisplayVerticalScrollBar = True: '
                 'o.DisplayWorkbookTabs = True: o.DisplayZeros = True: o.TabRatio = 0.6')
_WINDOW_CHANGED = ('o.Caption = "Other": o.DisplayFormulas = True: o.DisplayGridlines = False: o.DisplayHeadings = False: '
                   'o.DisplayHorizontalScrollBar = False: o.DisplayVerticalScrollBar = False: '
                   'o.DisplayWorkbookTabs = False: o.DisplayZeros = False: o.TabRatio = 0.3')
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
    # A window's settings outlast a sheet, so each case sets them all first, and a read activates the case's sheet.
    'Window': ('ws.Activate', 'ActiveWindow',
               ['Caption', 'DisplayFormulas', 'DisplayGridlines', 'DisplayHeadings', 'DisplayHorizontalScrollBar',
                'DisplayVerticalScrollBar', 'DisplayWorkbookTabs', 'DisplayZeros', 'TabRatio'],
               [_WINDOW_FRESH, _WINDOW_CHANGED]),
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
    'Window': {'Caption': ['"words"', '5', 'True', 'Null', 'Empty', '""', 'String(200, "a")', 'String(201, "a")',
                           'String(255, "a")'],
               'DisplayFormulas': SWITCHES, 'DisplayGridlines': SWITCHES, 'DisplayHeadings': SWITCHES,
               'DisplayHorizontalScrollBar': SWITCHES, 'DisplayVerticalScrollBar': SWITCHES,
               'DisplayWorkbookTabs': SWITCHES, 'DisplayZeros': SWITCHES,
               'TabRatio': ['0', '0.25', '0.5', '1', '1.5', '-0.1', '"0.5"', *BAD]},
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
            f'    On Error Resume Next\n    ws.Activate\n    Set o = {target}\n{reads}'
            f'    State{family} = Replace(s, wb.Name, "<book>")\nEnd Function\n')


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


def _typed(probe: str) -> str:
    """The probe with each case's o typed by its family, which is how a macro usually holds it."""
    parts = re.split(r'(?=\nPublic Function S\d+\(\))', probe)
    for at, part in enumerate(parts):
        found = re.search(r'State(\w+)\(ws\)', part)
        if part.startswith('\nPublic Function S') and found:
            parts[at] = part.replace('o As Object', f'o As {found.group(1)}', 1)
    return ''.join(parts)


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
        bound = excel.run_vba(_typed(probe) + _runner([f'S{number}' for number in range(len(every))], 'First')
                              + '\nPublic Function Probe() As String\n    Setup\n    Probe = First()\n    Finish\nEnd Function\n',
                              'Probe', (), timeout=500.0)
        assert bound.ok and not bound.dialogs, bound
    answers = str(result.value).split('@@')
    early = str(bound.value).split('@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': probe,
              'early': {str(number): early[number] for number in range(len(every)) if early[number] != answers[number]},
              'states': {family: FAMILIES[family][2] for family in FAMILIES},
              'cases': [[*case, answers[number], answers[len(every) + number]] for number, case in enumerate(every)]}
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    print(len(every), 'cases')


if __name__ == '__main__':
    main()
