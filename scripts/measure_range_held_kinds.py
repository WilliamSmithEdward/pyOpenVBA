"""Range's VERIFIED members given the two kinds of value the held-out worksheet function tests found the first fixtures lacked:
a Currency and a range of separate areas, with a Date and a Single beside them, late and early bound.

    python scripts/measure_range_held_kinds.py predict    writes the model's answers
    python scripts/measure_range_held_kinds.py measure    writes Excel's beside them, in tests/fixtures/range_held_kinds.json
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/range_held_kinds.json'
VALUES = ['x = CCur(1.5)', 'x = CCur(-2)', 'x = CCur(1000)', 'Set x = ws.Range("E1,G1")', 'Set x = ws.Range("E1:E2,G1:G2")', 'x = #3/1/2019#', 'x = CSng(0.25)']
#: The member and what to read after, on one cell and on a block.
SETTERS = {'Value': ['Value', 'Formula'], 'Value2': ['Value2', 'Formula'], 'Formula': ['Formula', 'Value'], 'FormulaLocal': ['FormulaLocal', 'Value'],
           'NumberFormat': ['NumberFormat'], 'NumberFormatLocal': ['NumberFormatLocal'], 'Style': ['Style'], 'ColumnWidth': ['ColumnWidth'], 'HorizontalAlignment': ['HorizontalAlignment'],
           'VerticalAlignment': ['VerticalAlignment'], 'WrapText': ['WrapText'], 'MergeCells': ['MergeCells'], 'Font.Bold': ['Font.Bold'], 'Font.Size': ['Font.Size'], 'Font.Name': ['Font.Name'],
           'Font.Color': ['Font.Color'], 'Interior.Color': ['Interior.Color'], 'Interior.Pattern': ['Interior.Pattern']}
CALLS = {'Offset': 'r = c.Offset(x).Address', 'OffsetColumn': 'r = c.Offset(0, x).Address', 'Resize': 'r = c.Resize(x).Address', 'ResizeColumn': 'r = c.Resize(, x).Address', 'Merge': 'c.Merge x',
         'FindNext': 'r = c.FindNext(x).Address', 'FindPrevious': 'r = c.FindPrevious(x).Address', 'Find': 'r = c.Find(x).Address', 'Replace': 'r = c.Replace(x, "z")', 'ReplaceWith': 'r = c.Replace("abc", x)'}
HELPER = '''Private Function Note(e As Long, v As Variant) As String
    If e <> 0 Then
        Note = "e" & e
    ElseIf IsNull(v) Then
        Note = "Null"
    ElseIf IsEmpty(v) Then
        Note = "Empty"
    ElseIf IsArray(v) Then
        Note = TypeName(v)
    Else
        Note = TypeName(v) & " " & CStr(v)
    End If
End Function
'''


def cases() -> list[tuple[str, str, str, str, list[str]]]:
    """Each (member, range, value, statement, reads)."""
    found: list[tuple[str, str, str, str, list[str]]] = []
    for target in ('B2', 'B2:B3'):
        for member, reads in SETTERS.items():
            found += [(member, target, value, f'c.{member} = x', reads) for value in VALUES]
    for name, call in CALLS.items():
        found += [(name, 'B2:C4', value, call, ['Value']) for value in VALUES]
    return found


def source(typed: bool) -> str:
    """One function a case, so the model can report a gap on one alone, and a Probe of them for Excel."""
    kind = 'Range' if typed else 'Object'
    lines = ['Private wb As Object', '', HELPER, 'Public Sub Setup()', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet)', 'End Sub', '']
    every = cases()
    for number, (_member, target, value, statement, reads) in enumerate(every):
        lines += [f'Public Function S{number}() As String', f'    Dim ws As Object, c As {kind}, x As Variant, r As Variant, v As Variant, s As String, e As Long', '    Set ws = wb.Worksheets.Add',
                  '    ws.Range("B2").Value = "abc": ws.Range("C3").Value = 7: ws.Range("E1").Value = 1.5: ws.Range("G1").Value = "abc"', f'    Set c = ws.Range("{target}")',
                  '    On Error Resume Next', f'    {value}', f'    {statement}', '    e = Err.Number: Err.Clear', '    s = Note(e, r)']
        lines += [f'    v = Empty: v = c.{read}: s = s & "|" & Note(Err.Number, v): Err.Clear' for read in reads]
        lines += [f'    S{number} = s', 'End Function', '']
    parts = [range(at, min(at + 150, len(every))) for at in range(0, len(every), 150)]
    for number, part in enumerate(parts):
        lines += [f'Private Function Part{number}() As String', '    Dim s As String', *[f'    s = s & "@@" & S{one}()' for one in part], f'    Part{number} = Mid(s, 3)', 'End Function', '']
    joined = ' & "@@" & '.join(f'Part{number}()' for number in range(len(parts)))
    lines += ['Public Function Probe() As String', '    Setup', f'    Probe = {joined}', '    wb.Close False', 'End Function', '']
    return '\n'.join(lines)


def predict() -> None:
    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    answers: dict[str, list[str]] = {}
    for name, typed in (('late', False), ('typed', True)):
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(source(typed), name='Held')
        app.run('Held.Setup')
        found: list[str] = []
        for number in range(len(cases())):
            try:
                found.append(str(app.run(f'Held.S{number}')))
            except VBAUnsupportedError:
                found.append('gap')
        answers[name] = found
    OUT.write_text(json.dumps({'late_source': source(False), 'typed_source': source(True), 'cases': [list(case[:4]) for case in cases()], 'predicted': answers}, indent=0) + '\n', encoding='utf-8')
    print({key: len(value) for key, value in answers.items()}, sum(one == 'gap' for value in answers.values() for one in value), 'gaps')


def measure() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record = json.loads(OUT.read_text(encoding='utf-8'))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        late = excel.run_vba(source(False), 'Probe', (), timeout=200.0)
        assert late.ok and not late.dialogs, late
        typed = excel.run_vba(source(True), 'Probe', (), timeout=200.0)
        assert typed.ok and not typed.dialogs, typed
    record['late'] = str(late.value).split('@@')
    record['typed'] = str(typed.value).split('@@')
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    wrong: dict[str, int] = {}
    for tag in ('late', 'typed'):
        for case, want, got in zip(record['cases'], record[tag], record['predicted'][tag], strict=True):
            if want != got:
                wrong[case[0]] = wrong.get(case[0], 0) + 1
                print(tag, case[0], case[1], case[2][:30], 'want', want[:44], 'got', got[:44])
    print(wrong)


if __name__ == '__main__':
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
