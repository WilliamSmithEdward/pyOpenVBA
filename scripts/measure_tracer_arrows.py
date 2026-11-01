"""Range.ShowPrecedents, ShowDependents and ShowErrors, and Worksheet.CircleInvalid: what each answers, and the errors
each raises. The arrows and circles they draw are not saved and no member reads them back, so the answers are all a
macro sees.

Each case makes a workbook with a small web of formulas (A1 a value, B1 =A1*2, C1 =B1+A1, D1 =1/0, E1 =D1+1, F1 a
validated cell holding an invalid value), runs the case's statements and records what they answer.

    python scripts/measure_tracer_arrows.py

writes tests/fixtures/tracer_arrows.json, which tests/test_excel_tracer_arrows.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/tracer_arrows.json'
CASES = [
    'v = TypeName(ws.Range("C1").ShowPrecedents)',
    'ws.Range("C1").ShowPrecedents: v = "done"',
    'ws.Range("C1").ShowPrecedents Remove:=True: v = "done"',
    'ws.Range("A1").ShowPrecedents: v = "done"',
    'ws.Range("A1").ShowDependents: v = "done"',
    'ws.Range("C1").ShowDependents: v = "done"',
    'ws.Range("G9").ShowPrecedents: v = "done"',
    'ws.Range("G9").ShowDependents: v = "done"',
    'ws.Range("A1:C1").ShowPrecedents: v = "done"',
    'ws.Range("A1").ShowDependents Remove:=True: v = "done"',
    'ws.Range("E1").ShowErrors: v = "done"',
    'v = TypeName(ws.Range("E1").ShowErrors)',
    'ws.Range("C1").ShowErrors: v = "done"',
    'ws.Range("D1").ShowErrors: v = "done"',
    'ws.Range("C1").ShowPrecedents: ws.ClearArrows: v = "done"',
    'ws.CircleInvalid: v = "done"',
    'v = TypeName(ws.CircleInvalid)',
    'ws.CircleInvalid: ws.ClearCircles: v = "done"',
    'ws.Protect: ws.Range("C1").ShowPrecedents: v = "done"',
    'ws.Protect: ws.CircleInvalid: v = "done"',
    'v = CStr(ws.Range("C1").ShowPrecedents)',
    'v = CStr(ws.Range("A1").ShowDependents)',
    'v = CStr(ws.Range("G9").ShowPrecedents)',
    'v = CStr(ws.Range("C1").ShowPrecedents(True))',
    'v = ws.Range("E1").ShowErrors.Address',
    'v = ws.Range("D1").ShowErrors.Address',
    'v = ws.Range("C1").ShowErrors.Address',
    'v = ws.Range("G9").ShowErrors.Address',
    'v = ws.Range("D1:E1").ShowErrors.Address',
    'ws.Protect: ws.Range("A1").ShowDependents: v = "done"',
    'ws.Protect: ws.Range("E1").ShowErrors: v = "done"',
    'ws.Range("C1").ShowDependents: ws.Range("E1").ShowErrors: v = "done"',
    'ws.Range("G1").Formula = "=NA()": ws.Range("H1").Formula = "=D1+G1": v = ws.Range("H1").ShowErrors.Address',
    'ws.Range("G1").Formula = "=NA()": ws.Range("H1").Formula = "=G1+D1": v = ws.Range("H1").ShowErrors.Address',
    'ws.Range("G1").Formula = "=NA()": ws.Range("H1").Formula = "=SUM(D1:G1)": v = ws.Range("H1").ShowErrors.Address',
    'ws.Range("H1").Formula = "=E1*2": v = ws.Range("H1").ShowErrors.Address',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String', '    Dim wb As Workbook, ws As Object, v As Variant',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1").Value = 3: ws.Range("B1").Formula = "=A1*2": ws.Range("C1").Formula = "=B1+A1"',
             '    ws.Range("D1").Formula = "=1/0": ws.Range("E1").Formula = "=D1+1"',
             '    ws.Range("F1").Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, "1", "5"',
             '    ws.Range("F1").Value = 9', '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', '    Tried = "ok " & v', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Probe() As String', '    Dim i As Long, s As String',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (), timeout=200.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer] for case, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
