"""What a range of separate areas reads as: its value, formula and text, in a Variant, a conversion and an operator, with the
range held As Object and As Range.

    python scripts/measure_range_area_reads.py

writes tests/fixtures/range_area_reads.json, which tests/test_excel_range_area_reads.py replays.
"""
from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_area_reads.json'
TARGETS = ['E1,G1', 'E1:E2,G1', 'G1,E1', 'H1,E1', 'E1:F2,G1:G2', 'E1,J1', 'E1,H1', 'E1,E2', 'E1,J1:J2', 'H1,I1', 'H1:I1,E1', 'J1:J2,E1:E2', 'E1:E2,J1:J2', 'E1:E2', 'E1']
READS = ['v = o', 'v = o.Value', 'v = o.Value2', 'v = o.Formula', 'v = o.FormulaR1C1', 'v = o.FormulaLocal', 'v = o.FormulaR1C1Local', 'v = o.Formula2', 'v = o.Text', 'v = CStr(o)', 'v = CDbl(o)',
         'v = o + 1', 'v = 1 + o', 'v = o & "x"', 'v = (o)', 'v = TypeName(o)', 'v = IsArray(o)', 'v = IsObject(o)', 'v = IsNull(o.Text)']
HELPER = '''Private Function Note(e As Long, v As Variant) As String
    If e <> 0 Then
        Note = "e" & e
    ElseIf IsNull(v) Then
        Note = "Null"
    ElseIf IsEmpty(v) Then
        Note = "Empty"
    ElseIf IsObject(v) Then
        Note = "Object " & TypeName(v)
    ElseIf IsArray(v) Then
        Note = TypeName(v) & " " & UBound(v, 1) & "x" & UBound(v, 2)
    Else
        Note = TypeName(v) & " " & CStr(v)
    End If
End Function
'''


def source(typed: bool) -> str:
    kind = 'Range' if typed else 'Object'
    lines = [HELPER, 'Public Function Probe() As String', f'    Dim wb As Object, ws As Object, o As {kind}, v As Variant, s As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("E1").Value = 1.5: ws.Range("E2").Value = 2: ws.Range("F1").Value = 3: ws.Range("F2").Value = 4: ws.Range("G1").Value = "abc": ws.Range("G2").Value = 6',
             '    ws.Range("J1").Value = 1.5: ws.Range("J2").Value = 2', '    On Error Resume Next']
    for target in TARGETS:
        lines.append(f'    Set o = ws.Range("{target}")')
        lines += [f'    v = Empty: Err.Clear: {read}: s = s & "@@" & Note(Err.Number, v): Err.Clear' for read in READS]
    lines += ['    wb.Close False', '    Probe = Mid(s, 3)', 'End Function', '']
    return '\n'.join(lines)


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=600.0, exclusive=False)) as excel:
        excel.new_document()
        late = excel.run_vba(source(False), 'Probe', (), timeout=60.0)
        assert late.ok and not late.dialogs, late
        typed = excel.run_vba(source(True), 'Probe', (), timeout=60.0)
        assert typed.ok and not typed.dialogs, typed
    every = [(target, read) for target in TARGETS for read in READS]
    one, two = str(late.value).split('@@'), str(typed.value).split('@@')
    assert len(one) == len(two) == len(every), (len(one), len(two), len(every))
    OUT.write_text(json.dumps({'late_source': source(False), 'typed_source': source(True), 'cases': [[*case, a, b] for case, a, b in zip(every, one, two, strict=True)]}, indent=0) + '\n', encoding='utf-8')
    print(len(every), 'cases', sum(a != b for a, b in zip(one, two)), 'differ by binding')


if __name__ == '__main__':
    main()
