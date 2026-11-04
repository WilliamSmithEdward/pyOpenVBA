"""The worksheet functions macros call most, held to the checklist's VERIFIED standard (docs/host_completeness.md): each with
each kind of value for each argument, called on a WorksheetFunction typed and on one held As Object.

    python scripts/measure_worksheet_function_args.py

writes tests/fixtures/worksheet_function_args.json, which tests/test_excel_worksheet_function_args.py replays both ways.
A call with an argument left out does not compile typed, so only late binding answers it.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/worksheet_function_args.json'
KINDS = ['Null', 'Empty', 'True', 'False', '"x"', '""', '"2"', '1.5', '0', '-1', '99', '1E+300', 'ws.Range("A1")', 'ws.Range("A1:A3")', 'ws.Range("Z9")',
         'Array(1, 2)', 'CVErr(xlErrNA)', 'Nothing', 'ws']
TABLE = 'ws.Range("A1:C3")'
COLUMN = 'ws.Range("A1:A3")'
#: A function and its calls: ARG stands for each kind of value. Calls with an argument left out are late bound alone.
CALLS: dict[str, list[str]] = {
    'VLookup': ['ARG, %s, 2, False' % TABLE, '2, ARG, 2, False', '2, %s, ARG, False' % TABLE, '2, %s, 2, ARG' % TABLE, '2, %s, 2' % TABLE, '"b", %s, 1, True' % 'ws.Range("B1:B3")'],
    'HLookup': ['ARG, %s, 2, False' % TABLE, '1, ARG, 2, False', '1, %s, ARG, False' % TABLE, '1, %s, 2, ARG' % TABLE],
    'Match': ['ARG, %s, 0' % COLUMN, '2, ARG, 0', '2, %s, ARG' % COLUMN, '2, %s' % COLUMN],
    'Index': ['ARG, 1, 1', '%s, ARG, 1' % TABLE, '%s, 1, ARG' % TABLE, '%s, 2' % TABLE, '%s, 0, 2' % TABLE, '%s, 2, 0' % TABLE, '%s, 1, 1, ARG' % TABLE],
    'Sum': ['ARG', '1, ARG', '%s, ARG' % COLUMN],
    'CountIf': ['ARG, 2', '%s, ARG' % COLUMN],
    'SumIf': ['ARG, ">1"', '%s, ARG' % COLUMN, '%s, ">1", ARG' % COLUMN, '%s, ">1"' % COLUMN],
    'Max': ['ARG', '1, ARG', '%s, ARG' % COLUMN],
    'Min': ['ARG', '1, ARG', '%s, ARG' % COLUMN],
    'Round': ['ARG, 1', '1.25, ARG'],
    'Count': ['ARG', '1, ARG'],
    'CountA': ['ARG', '1, ARG'],
    'Average': ['ARG', '1, ARG', '%s, ARG' % COLUMN],
}
HELPER = '''Private Function Note(e As Long, v As Variant) As String
    If e <> 0 Then
        Note = "e" & e
    ElseIf IsObject(v) Then
        Note = "Object " & TypeName(v)
    ElseIf IsNull(v) Then
        Note = "Null"
    ElseIf IsEmpty(v) Then
        Note = "Empty"
    ElseIf IsArray(v) Then
        Note = TypeName(v)
    ElseIf IsError(v) Then
        Note = CStr(v)
    Else
        Note = TypeName(v) & " " & CStr(v)
    End If
End Function
'''
BOOK = ('Private Sub Fill(ws As Object)\n    ws.Range("A1").Value = 1: ws.Range("A2").Value = 2: ws.Range("A3").Value = 3\n'
        '    ws.Range("B1").Value = "a": ws.Range("B2").Value = "b": ws.Range("B3").Value = "c"\n'
        '    ws.Range("C1").Value = 10: ws.Range("C2").Value = 20: ws.Range("C3").Value = 30\nEnd Sub\n')


def _held(kind: str) -> str:
    return f'Set x = {kind}' if kind == 'Nothing' or kind.startswith('ws') else f'x = {kind}'


def cases() -> list[tuple[str, str, bool]]:
    """Each (function, its call, whether it compiles typed)."""
    found: list[tuple[str, str, bool]] = []
    for name, calls in CALLS.items():
        found.append((name, f'v = f.{name}()', False))
        for call in calls:
            if 'ARG' in call:
                # Each kind reaches the call in a Variant, as a macro holds it: literal text or Nothing for an argument typed
                # Double or Range does not compile.
                found += [(name, f'{_held(kind)}: v = f.{name}({call.replace("ARG", "x")})', True) for kind in KINDS]
            else:
                found.append((name, f'v = f.{name}({call})', True))
    return found


def source(typed: bool) -> str:
    """A probe answering every case joined by @@; what does not compile typed answers "no compile" there."""
    every = cases()
    kind = 'WorksheetFunction' if typed else 'Object'
    body = HELPER + BOOK
    parts = [every[at:at + 150] for at in range(0, len(every), 150)]
    for number, part in enumerate(parts):
        lines = [f'Private Function Part{number}(f As {kind}, ws As Object) As String', '    Dim v As Variant, x As Variant, e As Long, s As String', '    On Error Resume Next']
        for _, call, compiles in part:
            if typed and not compiles:
                lines.append('    s = s & "@@no compile"')
            else:
                lines += ['    v = Empty: Err.Clear', f'    {call}', '    e = Err.Number: Err.Clear', '    s = s & "@@" & Note(e, v)']
        lines += [f'    Part{number} = Mid(s, 3)', 'End Function', '']
        body += '\n'.join(lines) + '\n'
    joined = ' & "@@" & '.join(f'Part{number}(f, ws)' for number in range(len(parts)))
    return body + (f'Public Function Probe() As String\n    Dim wb As Object, ws As Object, f As {kind}\n    Set wb = Workbooks.Add(xlWBATWorksheet)\n'
                   f'    Set ws = wb.Worksheets(1)\n    Fill ws\n    Set f = Application.WorksheetFunction\n    Probe = {joined}\n    wb.Close False\nEnd Function\n')


def main() -> None:
    every = cases()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        late = excel.run_vba(source(False), 'Probe', (), timeout=280.0)
        assert late.ok and not late.dialogs, late
        typed = excel.run_vba(source(True), 'Probe', (), timeout=280.0)
        assert typed.ok and not typed.dialogs, typed
    answers, bound = str(late.value).split('@@'), str(typed.value).split('@@')
    assert len(answers) == len(bound) == len(every), (len(answers), len(bound), len(every))
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'late_source': source(False), 'typed_source': source(True),
              'cases': [[name, call, one, two] for (name, call, _), one, two in zip(every, answers, bound, strict=True)]}
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    print(len(every), 'cases', sum(case[2] != case[3] and case[3] != 'no compile' for case in record['cases']), 'differ by binding')


if __name__ == '__main__':
    main()
