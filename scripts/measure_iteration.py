"""Circular references, with and without iterative calculation.

Each probe adds a sheet, sets Application.Iteration, MaxIterations and
MaxChange as it needs, writes formulas that feed themselves and records
what the cells hold (CStr, fifteen digits), what CircularReference names
and what the settings read back; then it deletes the sheet and puts the
settings back to Excel's defaults, off, 100 and 0.001.

    python scripts/measure_iteration.py

writes tests/fixtures/iteration.json, which tests/test_excel_iteration.py
replays, with the calcPr of workbooks Excel saved with each setting and
the settings one of them brings when it is opened.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/iteration.json'

#: Each probe's body, ws being its sheet and s what it returns.
BODIES = {
    'defaults': 's = Application.Iteration & "|" & Application.MaxIterations & "|" & Application.MaxChange',
    'set': '''Application.Iteration = True: Application.MaxIterations = 50: Application.MaxChange = 0.01
s = Application.Iteration & "|" & Application.MaxIterations & "|" & Application.MaxChange''',
    'iteration_from_a_number': 'Application.Iteration = 2: s = CStr(Application.Iteration)',
    'iteration_from_text': 'Application.Iteration = "True": s = CStr(Application.Iteration)',
    'iteration_from_other_text': '''On Error Resume Next
Application.Iteration = "abc": s = Err.Number & ":" & Application.Iteration
On Error GoTo 0''',
    'max_iterations_bounds': '''Dim v As Variant
On Error Resume Next
For Each v In Array(0, 1, 32767, 32768, -5, 2.5, 3.5, 1E+20)
    Err.Clear: Application.MaxIterations = v
    s = s & "|" & Err.Number & ":" & Application.MaxIterations
Next
On Error GoTo 0''',
    'max_change_bounds': '''Dim v As Variant
On Error Resume Next
For Each v In Array(0, -1, 1E+300, 1E-300, 0.5)
    Err.Clear: Application.MaxChange = v
    s = s & "|" & Err.Number & ":" & Application.MaxChange
Next
Err.Clear: Application.MaxChange = "abc": s = s & "|" & Err.Number & ":" & Application.MaxChange
On Error GoTo 0''',
    'off_self': '''ws.Range("A1").Formula = "=A1+1"
s = CStr(ws.Range("A1").Value) & "|" & ws.CircularReference.Address
Application.Calculate: s = s & "|" & CStr(ws.Range("A1").Value)''',
    'off_none': 's = TypeName(ws.CircularReference)',
    'off_pair': '''ws.Range("A1").Formula = "=B1+1": ws.Range("B1").Formula = "=A1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value) & "|" & ws.CircularReference.Address''',
    'off_pair_reversed': '''ws.Range("B1").Formula = "=A1+1": ws.Range("A1").Formula = "=B1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value) & "|" & ws.CircularReference.Address''',
    'off_dependent': '''ws.Range("A1").Formula = "=A1+1": ws.Range("C1").Formula = "=A1*2+3"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("C1").Value) & "|" & ws.CircularReference.Address''',
    'off_value_kept': '''ws.Range("A1").Value = 5: ws.Range("B1").Formula = "=A1*2"
ws.Range("A1").Formula = "=B1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value)''',
    'on_self': '''Application.Iteration = True
ws.Range("A1").Formula = "=A1+1"
s = CStr(ws.Range("A1").Value)
Application.Calculate: s = s & "|" & CStr(ws.Range("A1").Value)
Application.CalculateFull: s = s & "|" & CStr(ws.Range("A1").Value)
s = s & "|" & TypeName(ws.CircularReference)''',
    'on_self_one_iteration': '''Application.Iteration = True: Application.MaxIterations = 1
ws.Range("A1").Formula = "=A1+1"
s = CStr(ws.Range("A1").Value)
Application.Calculate: s = s & "|" & CStr(ws.Range("A1").Value)''',
    'on_converging': '''Application.Iteration = True
ws.Range("A1").Formula = "=(A1+10)/2"
s = CStr(ws.Range("A1").Value)
Application.Calculate: s = s & "|" & CStr(ws.Range("A1").Value)''',
    'on_converging_coarse': '''Application.Iteration = True: Application.MaxChange = 0.1
ws.Range("A1").Formula = "=(A1+10)/2"
s = CStr(ws.Range("A1").Value)''',
    'on_converging_exact_change': '''Application.Iteration = True: Application.MaxChange = 0.625
ws.Range("A1").Formula = "=(A1+10)/2"
s = CStr(ws.Range("A1").Value)''',
    'on_pair': '''Application.Iteration = True
ws.Range("A1").Formula = "=B1+1": ws.Range("B1").Formula = "=A1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value)
Application.Calculate: s = s & "|" & CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value)''',
    'on_interest': '''Application.Iteration = True
ws.Range("B1").Formula = "=1000+B2": ws.Range("B2").Formula = "=B1*0.1"
s = CStr(ws.Range("B1").Value) & "|" & CStr(ws.Range("B2").Value)''',
    'on_interest_reversed': '''Application.Iteration = True
ws.Range("B2").Formula = "=B1*0.1": ws.Range("B1").Formula = "=1000+B2"
s = CStr(ws.Range("B1").Value) & "|" & CStr(ws.Range("B2").Value)''',
    'on_dependent': '''Application.Iteration = True
ws.Range("A1").Formula = "=(A1+10)/2": ws.Range("C1").Formula = "=A1*2"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("C1").Value)''',
    'on_three': '''Application.Iteration = True
ws.Range("A1").Formula = "=C1+1": ws.Range("B1").Formula = "=A1+1": ws.Range("C1").Formula = "=B1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value) & "|" & CStr(ws.Range("C1").Value)''',
    'on_after_off': '''ws.Range("A1").Formula = "=A1+1"
s = CStr(ws.Range("A1").Value)
Application.Iteration = True: s = s & "|" & CStr(ws.Range("A1").Value)
Application.Calculate: s = s & "|" & CStr(ws.Range("A1").Value)
Application.Iteration = False: s = s & "|" & CStr(ws.Range("A1").Value)
Application.Calculate: s = s & "|" & CStr(ws.Range("A1").Value)''',
    'on_a_value_changes': '''Application.Iteration = True
ws.Range("A1").Value = 1: ws.Range("B1").Formula = "=(B1+A1*10)/2"
s = CStr(ws.Range("B1").Value)
ws.Range("A1").Value = 2: s = s & "|" & CStr(ws.Range("B1").Value)''',
    'on_error_in_loop': '''Application.Iteration = True
ws.Range("A1").Formula = "=1/(A1-A1)"
s = CStr(ws.Range("A1").Value)''',
    'on_text_in_loop': '''Application.Iteration = True
ws.Range("A1").Formula = "=A1&""x"""
s = CStr(ws.Range("A1").Value)''',
    'on_sum_range': '''Application.Iteration = True
ws.Range("A1").Value = 1: ws.Range("A2").Value = 2: ws.Range("A3").Formula = "=SUM(A1:A3)"
s = CStr(ws.Range("A3").Value)''',
    'on_manual': '''Application.Iteration = True: Application.Calculation = xlCalculationManual
ws.Range("A1").Formula = "=A1+1"
s = CStr(ws.Range("A1").Value)
ws.Calculate: s = s & "|" & CStr(ws.Range("A1").Value)
Application.Calculate: s = s & "|" & CStr(ws.Range("A1").Value)
Application.Calculation = xlCalculationAutomatic: s = s & "|" & CStr(ws.Range("A1").Value)''',
    'off_dependent_recalculated': '''ws.Range("A1").Formula = "=A1+1": ws.Range("C1").Formula = "=A1*2+3"
Application.Calculate: s = CStr(ws.Range("C1").Value)
Application.CalculateFull: s = s & "|" & CStr(ws.Range("C1").Value) & "|" & ws.CircularReference.Address''',
    'off_broken': '''ws.Range("A1").Formula = "=A1+1": ws.Range("A1").Value = 5
s = CStr(ws.Range("A1").Value) & "|" & TypeName(ws.CircularReference)''',
    'off_both_at_once': '''ws.Range("A1:A2").Formula = "=A2+1"
ws.Range("A2").Formula = "=A1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("A2").Value) & "|" & ws.CircularReference.Address''',
    'off_across_sheets': '''Dim other As Worksheet
Set other = Worksheets.Add(After:=ws)
ws.Range("A1").Formula = "='" & other.Name & "'!A1+1": other.Range("A1").Formula = "='" & ws.Name & "'!A1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(other.Range("A1").Value)
s = s & "|" & IIf(ws.CircularReference.Parent.Name = ws.Name, "ws", "other") & ws.CircularReference.Address
s = s & "|" & IIf(other.CircularReference.Parent.Name = ws.Name, "ws", "other") & other.CircularReference.Address
Application.DisplayAlerts = False: other.Delete: Application.DisplayAlerts = True''',
    'off_two_loops': '''ws.Range("A1").Formula = "=A1+1": ws.Range("B1").Formula = "=B1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value) & "|" & ws.CircularReference.Address
ws.Range("B1").Value = 1: s = s & "|" & ws.CircularReference.Address''',
    'off_precedent_edit': '''ws.Range("A1").Formula = "=A1+D1": ws.Range("C1").Formula = "=A1+1"
ws.Range("D1").Value = 5
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("C1").Value)''',
    'off_dependent_after_broken': '''ws.Range("A1").Formula = "=A1+1": ws.Range("C1").Formula = "=A1*2+3"
ws.Range("A1").Value = 1
s = CStr(ws.Range("C1").Value) & "|" & TypeName(ws.CircularReference)''',
    'off_stale_member': '''ws.Range("A1").Formula = "=A2+1": ws.Range("A2").Formula = "=A3+1": ws.Range("A3").Formula = "=A1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("A2").Value) & "|" & CStr(ws.Range("A3").Value)''',
    'off_upstream_formula': '''ws.Range("D1").Formula = "=2*3": ws.Range("A1").Formula = "=A1+D1"
ws.Range("D1").Formula = "=2*4"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("D1").Value)''',
    'on_across_sheets_closing_on_other': '''Dim other As Worksheet
Application.Iteration = True
Set other = Worksheets.Add(After:=ws)
ws.Range("A1").Formula = "=1000+'" & other.Name & "'!A1": other.Range("A1").Formula = "='" & ws.Name & "'!A1*0.1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(other.Range("A1").Value)
Application.DisplayAlerts = False: other.Delete: Application.DisplayAlerts = True''',
    'on_across_sheets_closing_on_other_before': '''Dim other As Worksheet
Application.Iteration = True
Set other = Worksheets.Add(Before:=ws)
ws.Range("A1").Formula = "=1000+'" & other.Name & "'!A1": other.Range("A1").Formula = "='" & ws.Name & "'!A1*0.1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(other.Range("A1").Value)
Application.DisplayAlerts = False: other.Delete: Application.DisplayAlerts = True''',
    'on_text_beside_number': '''Application.Iteration = True
ws.Range("A1").Formula = "=A1+1": ws.Range("B1").Formula = "=B1&""x"""
s = CStr(ws.Range("A1").Value) & "|" & Len(ws.Range("B1").Value)''',
    'on_logical_loop': '''Application.Iteration = True
ws.Range("A1").Formula = "=NOT(A1)"
s = CStr(ws.Range("A1").Value)''',
    'on_logical_beside_number': '''Application.Iteration = True
ws.Range("A1").Formula = "=A1+1": ws.Range("B1").Formula = "=NOT(B1)"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value)''',
    'on_unrelated_edit': '''Application.Iteration = True
ws.Range("A1").Formula = "=A1+1"
ws.Range("D5").Value = 1: s = CStr(ws.Range("A1").Value)
ws.Range("D6").Formula = "=1+1": s = s & "|" & CStr(ws.Range("A1").Value)''',
    'on_precedent_edit': '''Application.Iteration = True
ws.Range("A1").Formula = "=A1+D1"
s = CStr(ws.Range("A1").Value)
ws.Range("D1").Value = 1: s = s & "|" & CStr(ws.Range("A1").Value)''',
    'on_dependent_edit': '''Application.Iteration = True
ws.Range("A1").Formula = "=A1+1": ws.Range("C1").Formula = "=A1*2"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("C1").Value)
ws.Range("C1").Formula = "=A1*3": s = s & "|" & CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("C1").Value)''',
    'on_upstream_formula': '''Application.Iteration = True
ws.Range("D1").Formula = "=2*5": ws.Range("A1").Formula = "=(A1+D1)/2"
s = CStr(ws.Range("A1").Value)''',
    'on_downstream_converging': '''Application.Iteration = True
ws.Range("A1").Formula = "=A1*0+5": ws.Range("A2").Formula = "=(A2+A1)/2"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("A2").Value)''',
    'on_cycle_against_position': '''Application.Iteration = True
ws.Range("A3").Formula = "=A1+1": ws.Range("A2").Formula = "=A3+1": ws.Range("A1").Formula = "=A2+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("A2").Value) & "|" & CStr(ws.Range("A3").Value)''',
    'on_cycle_with_position': '''Application.Iteration = True
ws.Range("A1").Formula = "=A2+1": ws.Range("A2").Formula = "=A3+1": ws.Range("A3").Formula = "=A1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("A2").Value) & "|" & CStr(ws.Range("A3").Value)''',
    'on_across_sheets': '''Dim other As Worksheet
Application.Iteration = True
Set other = Worksheets.Add(After:=ws)
other.Range("A1").Formula = "='" & ws.Name & "'!A1*0.1": ws.Range("A1").Formula = "=1000+'" & other.Name & "'!A1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(other.Range("A1").Value)
Application.DisplayAlerts = False: other.Delete: Application.DisplayAlerts = True''',
    'on_across_sheets_reversed': '''Dim other As Worksheet
Application.Iteration = True
Set other = Worksheets.Add(Before:=ws)
other.Range("A1").Formula = "='" & ws.Name & "'!A1*0.1": ws.Range("A1").Formula = "=1000+'" & other.Name & "'!A1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(other.Range("A1").Value)
Application.DisplayAlerts = False: other.Delete: Application.DisplayAlerts = True''',
    'on_zero_change': '''Application.Iteration = True: Application.MaxChange = 0
ws.Range("A1").Formula = "=(A1+10)/2": ws.Range("B1").Formula = "=B1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value)''',
    'on_negative_change': '''Application.Iteration = True: Application.MaxChange = -1
ws.Range("A1").Formula = "=A1*0+5": ws.Range("B1").Formula = "=B1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value)''',
    'on_settles': '''Application.Iteration = True
ws.Range("A1").Formula = "=A1*0+5": ws.Range("B1").Formula = "=A1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value)''',
    'on_two_loops': '''Application.Iteration = True
ws.Range("A1").Formula = "=(A1+10)/2": ws.Range("B1").Formula = "=B1+1"
s = CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value)
Application.Calculate: s = s & "|" & CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value)''',
    'on_ordinary_formula': '''Application.Iteration = True
ws.Range("A1").Value = 3: ws.Range("B1").Formula = "=A1*2"
s = CStr(ws.Range("B1").Value)''',
}
#: Probes whose CircularReference Excel answers differently run to run, measured this many times each.
REPEATED = ('off_pair_reversed', 'off_both_at_once', 'off_two_loops', 'off_pair')
ROUNDS = 6
#: Workbooks saved with each setting, whose calcPr the file keeps; the first is opened again, beside the workbook
#: the probe runs in and then in an Excel of its own.
SAVED = {
    'iterating': 'Application.Iteration = True: Application.MaxIterations = 50: Application.MaxChange = 0.01',
    'not_iterating': 'Application.Iteration = False: Application.MaxIterations = 50: Application.MaxChange = 0.01',
    'defaults': 'Application.Iteration = False: Application.MaxIterations = 100: Application.MaxChange = 0.001',
}


def file_probe(folder: Path) -> str:
    lines = ['Function Probe() As String', 'Dim wb As Workbook, s As String', 'Application.DisplayAlerts = False']
    for name, setting in SAVED.items():
        path = str(folder / f'{name}.xlsx').replace('"', '""')
        lines += ['Set wb = Workbooks.Add(xlWBATWorksheet)', 'wb.Worksheets(1).Range("A1").Value = 1', setting,
                  f'wb.SaveAs Filename:="{path}", FileFormat:=xlOpenXMLWorkbook', 'wb.Close False']
    iterating = str(folder / 'iterating.xlsx').replace('"', '""')
    lines += ['Application.Iteration = False: Application.MaxIterations = 100: Application.MaxChange = 0.001',
              f'Set wb = Workbooks.Open("{iterating}")',
              's = Application.Iteration & "|" & Application.MaxIterations & "|" & Application.MaxChange',
              'wb.Close False',
              's = s & "|" & Application.Iteration & "|" & Application.MaxIterations & "|" & Application.MaxChange',
              'Application.Iteration = False: Application.MaxIterations = 100: Application.MaxChange = 0.001',
              # The first workbook a fresh Excel opens: an instance of its own, quit before the probe ends.
              'Dim fresh As Object', 'Set fresh = CreateObject("Excel.Application")', 'On Error Resume Next',
              f'fresh.Workbooks.Open "{iterating}"',
              's = s & "|" & fresh.Iteration & "|" & fresh.MaxIterations & "|" & fresh.MaxChange',
              'fresh.Workbooks(1).Close False', 'fresh.Quit', 'On Error GoTo 0', 'Set fresh = Nothing',
              'Application.DisplayAlerts = True', 'Probe = s', 'End Function']
    return '\n'.join(lines) + '\n'


def macro(body: str) -> str:
    return f'''Function Probe() As String
Dim ws As Worksheet, s As String
Set ws = Worksheets.Add
{body}
Application.Iteration = False: Application.MaxIterations = 100: Application.MaxChange = 0.001
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = False
ws.Delete
Application.DisplayAlerts = True
Probe = s
End Function
'''


def main() -> None:
    record({name: macro(body) for name, body in BODIES.items()}, OUT)
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        answers: dict[str, list[str]] = {}
        for name in REPEATED:
            for _ in range(ROUNDS):
                again = excel.run_vba(macro(BODIES[name]), 'Probe', timeout=60.0, line_numbers=False)
                assert again.ok and not again.dialogs, again
                answers.setdefault(name, []).append(str(again.value))
        saved = excel.run_vba(file_probe(Path(folder)), 'Probe', timeout=120.0, line_numbers=False)
        assert saved.ok and not saved.dialogs, saved
        files = {}
        for name in SAVED:
            with zipfile.ZipFile(Path(folder) / f'{name}.xlsx') as package:
                workbook = package.read('xl/workbook.xml').decode('utf-8')
            found = re.search(r'<calcPr\b[^>]*/?>', workbook)
            files[name] = found.group(0) if found else ''
    fixture = json.loads(OUT.read_text(encoding='utf-8'))
    fixture['repeated'] = answers
    fixture['file'] = {'settings': SAVED, 'calcPr': files, 'opened': saved.value}
    OUT.write_text(json.dumps(fixture, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(fixture['file'], indent=2))


if __name__ == '__main__':
    main()
