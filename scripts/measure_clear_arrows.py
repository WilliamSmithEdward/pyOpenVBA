"""Worksheet.ClearArrows and ClearCircles on a sheet with no tracer arrows or validation circles: whether they answer,
what they leave in Shapes, on a protected sheet, and whether the workbook is left unsaved.

    python scripts/measure_clear_arrows.py

writes tests/fixtures/clear_arrows.json, which tests/test_excel_clear_arrows.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/clear_arrows.json'
#: Each case: statements on a new workbook's sheet ws, saved first so Saved starts True.
CASES = [
    ['ws.ClearArrows'],
    ['ws.ClearCircles'],
    ['ws.Protect', 'ws.ClearArrows'],
    ['ws.Protect', 'ws.ClearCircles'],
    ['ws.Range("A1").Formula = "=B1"', 'ws.ClearArrows'],
    ['ws.Range("A1").Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, "1", "5"',
     'ws.Range("A1").Value = 9', 'ws.ClearCircles'],
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, v As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed', '    Select Case which']
    for index, statements in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements[:-1]]
        lines += [f'        wb.SaveAs folder & "\\c{index}.xlsx", xlOpenXMLWorkbook', f'        {statements[-1]}']
    lines += ['    End Select', '    Tried = TypeName(v) & " " & ws.Shapes.Count & " " & wb.Saved', 'Done:',
              '    wb.Close False', '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Done',
              'End Function', '', 'Public Function Probe(ByVal folder As String) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i, folder)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    import tempfile

    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=300.0)
        assert result.ok and not result.dialogs, result
    answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, answer] for statements, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, answer in record['cases']:
        print(statements, '->', answer)


if __name__ == '__main__':
    main()
