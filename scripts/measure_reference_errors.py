"""What a function that takes only a range does with an error or an array given there.

pyOfficeEditor measured that such a parameter takes an array one item at a
time, every item #VALUE!; pyOpenVBA measured that ROW and COLUMN pass a
#REF! through. These probes give the rest of those functions an error, as
a #REF! the formula holds and as a value LET hands them, and an array.
Each formula is written to its own cell of a fresh sheet, with Formula
where it holds #REF! and Formula2 where it uses LET, and read back as its
value or error number.

    python scripts/measure_reference_errors.py

writes tests/fixtures/formula/reference_errors.json, which
tests/test_formula_reference_errors.py replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/formula/reference_errors.json'

#: Each function, with the error or array standing in its range parameter as {x}.
FUNCTIONS = [
    'ROW({x})', 'COLUMN({x})', 'ROWS({x})', 'COLUMNS({x})', 'COUNTIF({x},1)', 'COUNTIFS({x},1)',
    'SUMIF({x},1)', 'SUMIFS(A1:A2,{x},1)', 'AVERAGEIF({x},1)', 'MAXIFS(A1:A2,{x},1)', 'COUNTBLANK({x})',
    'RANK(1,{x})', 'OFFSET({x},0,0)', 'CELL("row",{x})', 'AREAS({x})', 'ISFORMULA({x})', 'FORMULATEXT({x})',
    'SUBTOTAL(9,{x})', 'DSUM({x},1,A1:A2)', 'SHEET({x})', 'SHEETS({x})', 'ISREF({x})',
]
#: What stands in the range parameter: a #REF! written in the formula, then values LET hands over.
STANDINS = {
    'ref': ('Formula', '#REF!', ''),
    'div0': ('Formula2', 'x', 'LET(x,1/0,'),
    'na': ('Formula2', 'x', 'LET(x,NA(),'),
    'array': ('Formula2', 'x', 'LET(x,{1,2;3,4},'),
}


def case(standin: str, template: str) -> tuple[str, str]:
    """The property a formula is written with, and the formula."""
    prop, token, wrapper = STANDINS[standin]
    body = template.replace('{x}', token)
    return prop, '=' + (wrapper + body + ')' if wrapper else body)


def macro(standin: str) -> str:
    # Each formula sits in its own row, far enough apart that an array result spills without blocking another.
    lines = ['Function Probe() As String', 'Dim ws As Worksheet, s As String, v As Variant',
             'Set ws = Worksheets.Add', 'ws.Range("A1").Value = 1', 'ws.Range("A2").Value = 2',
             'On Error Resume Next']
    for index, template in enumerate(FUNCTIONS):
        prop, formula = case(standin, template)
        cell = f'ws.Cells({index * 3 + 1}, 3)'
        quoted = formula.replace('"', '""')
        lines += ['Err.Clear', f'{cell}.{prop} = "{quoted}"',
                  'If Err.Number <> 0 Then',
                  '  s = s & "|e" & Err.Number',
                  'Else',
                  f'  v = {cell}.Value',
                  f'  s = s & "|" & IIf(IsError(v), "", TypeName(v) & ":") & CStr(v)',
                  f'  If {cell}.HasSpill Then s = s & ":" & {cell}.SpillingToRange.Address(False, False) _',
                  f'      & ":" & CStr({cell}.SpillingToRange.Cells(2).Value)',
                  'End If']
    lines += ['Application.DisplayAlerts = False', 'ws.Delete', 'Application.DisplayAlerts = True',
              'Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


#: Legacy formulas handing an array to a range-only parameter, written with Formula and read back with Formula2:
#: where Excel puts the @ says it takes the call's answer for an array.
READBACK = ['=LET(x,{1,2;3,4},COUNTIF(x,1))', '=LET(x,{1,2;3,4},ROW(x))', '=LET(x,{1,2;3,4},COUNTBLANK(x))',
            '=LET(x,{1,2;3,4},SUMIF(x,1))', '=LET(x,A1:A2,COUNTIF(x,1))', '=LET(x,{1,2;3,4},ROWS(x))',
            '=COUNTIF(LET(x,{1,2},x),1)', '=LET(x,1,COUNTIF(x,1))', '=LET(x,{1,2;3,4},SUBTOTAL(9,x))']


def readback() -> str:
    lines = ['Function Probe() As String', 'Dim ws As Worksheet, s As String', 'Set ws = Worksheets.Add',
             'ws.Range("A1").Value = 1', 'ws.Range("A2").Value = 2', 'On Error Resume Next']
    for index, formula in enumerate(READBACK):
        cell = f'ws.Cells({index * 3 + 1}, 3)'
        quoted = formula.replace('"', '""')
        lines += ['Err.Clear', f'{cell}.Formula = "{quoted}"',
                  'If Err.Number <> 0 Then',
                  '  s = s & "|e" & Err.Number',
                  'Else',
                  f'  s = s & "|" & {cell}.Formula2 & "~" & CStr({cell}.Value) & "~" & {cell}.HasSpill',
                  'End If']
    lines += ['Application.DisplayAlerts = False', 'ws.Delete', 'Application.DisplayAlerts = True',
              'Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


if __name__ == '__main__':
    record({**{standin: macro(standin) for standin in STANDINS}, 'readback': readback()}, OUT)
