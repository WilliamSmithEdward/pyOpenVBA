"""ActiveChart on Application, Workbook and Window: which statements make an embedded chart active, which end it,
and what Selection, ActiveCell and ActiveSheet read meanwhile.

Each case makes a workbook of two sheets, with a column chart co of A1:B4 on the first and a second chart co2 below
it, runs its statements, and reads READS, each as its value (TypeName for an object) or "E" and the error number.

    python scripts/measure_active_chart.py

writes tests/fixtures/active_chart.json, which tests/test_excel_active_chart.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/active_chart.json'
CASES = [
    '',
    'co.Activate',
    'co.Chart.Activate',
    'co.Select',
    'co.Chart.Select',
    'co.Activate: ws.Range("C3").Select',
    'co.Activate: ws.Range("C3").Activate',
    'co.Activate: ws2.Activate',
    'co.Activate: ws2.Activate: ws.Activate',
    'co.Activate: co.Delete',
    'co.Activate: co2.Activate',
    'co.Activate: ActiveChart.ChartType = xlLine',
    'co.Activate: ActiveCell.Value = 5',
    'co.Activate: Set wb2 = Workbooks.Add: s0 = TypeName(ActiveChart) & "/" & TypeName(wb.ActiveChart): '
    'wb2.Close False',
    'co.Activate: Set wb2 = Workbooks.Add: wb2.Close False',
    'co.Activate: ActiveWindow.ActiveCell.Select',
    'co.Activate: ws.Shapes(2).Select',
    'co.Activate: Application.Goto ws.Range("D4")',
    'co.Activate: ws.Range("C3").Value = 1',
    'co.Activate: co.Chart.ChartArea.Select',
    'co.Activate: ActiveWindow.RangeSelection.Select',
]
READS = [
    'TypeName(ActiveChart)', 'ActiveChart.Parent.Name', 'TypeName(wb.ActiveChart)',
    'TypeName(ActiveWindow.ActiveChart)',
    'TypeName(Selection)', 'ActiveCell.Address', 'ActiveWindow.RangeSelection.Address', 'ActiveSheet.Name',
    'co.Chart.ChartType', 'ws.Range("A1").Value', 'ws.ChartObjects.Count', 's0',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, wb2 As Workbook, ws As Worksheet, ws2 As Worksheet, co As Object, co2 As Object',
             '    Dim s As String, s0 As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    Set ws2 = wb.Worksheets.Add(After:=ws): ws.Activate',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("B2").Select',
             '    Set co = ws.ChartObjects.Add(150, 10, 200, 120): co.Chart.SetSourceData ws.Range("A1:B4")',
             '    co.Chart.ChartType = xlColumnClustered',
             '    Set co2 = ws.ChartObjects.Add(150, 150, 200, 120): co2.Chart.SetSourceData ws.Range("A1:B4")',
             '    On Error GoTo Failed', '    Select Case which']
    for index, statements in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {statements}'] if statements else [])
    lines += ['    End Select', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', f'    s = s & "," & {read}', '    If Err.Number Then s = s & "E" & Err.Number']
    lines += ['    On Error GoTo 0', '    Tried = "ok " & Mid(s, 2)', 'Closed:', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Closed', 'End Function',
              '', 'Public Function Probe() As String', '    Dim i As Long, s As String',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=200.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[statements, answer] for statements, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, answer in record['cases']:
        print(statements, '->', answer)


if __name__ == '__main__':
    main()
