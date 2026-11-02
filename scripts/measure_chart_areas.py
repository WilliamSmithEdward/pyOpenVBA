"""A chart's ChartArea and PlotArea, and the ChartObjects collection: what they read, what selecting them leaves in
Selection and ActiveChart, and how the chart area follows its chart object.

Each case makes a workbook with a column chart co of A1:B4 at 150, 10, 200 by 120 and a second chart co2 below it,
runs its statements, and reads READS with ca and pa as co's chart area and plot area and cos as the sheet's chart
objects, each as its value (TypeName for an object) or "E" and the error number.

    python scripts/measure_chart_areas.py

writes tests/fixtures/chart_areas.json, which tests/test_excel_chart_areas.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_areas.json'
CASES = [
    '',
    'co.Activate: ca.Select',
    'ca.Select',
    'co.Activate: pa.Select',
    'pa.Select',
    'co.Width = 300: co.Height = 150',
    'co.RoundedCorners = True: co.Shadow = True',
    'co.Activate: co.Chart.ChartType = xlLine',
    'co.Activate: co.Chart.HasTitle = True',
    'cos.Select',
    'cos.Select: ws.Range("A1").Select',
    'co.Activate: ca.Select: pa.Select',
    'co.Activate: pa.Select: ws.Range("C3").Select',
    'co.Activate: ca.Select: co2.Activate',
    'co.Chart.HasLegend = False',
    'co.Left = 0.3: co.Top = 0.6: co.Width = 100.4: co.Height = 50.2',
]
READS = [
    'TypeName(Selection)', 'TypeName(ActiveChart)', 'ca.Name', 'TypeName(ca.Parent)', 'TypeName(ca.Application)',
    'ca.Left', 'ca.Top', 'ca.Width', 'ca.Height', 'ca.Shadow', 'ca.RoundedCorners', 'pa.Name', 'TypeName(pa.Parent)',
    'pa.Position', 'TypeName(pa.Application)', 'TypeName(cos.Application)', 'TypeName(cos.Parent)',
    'TypeName(cos.ShapeRange)', 'cos.ShapeRange.Count', 'cos.ProtectChartObject', 'co.Width', 'co.Height',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, co As Object, co2 As Object, ca As Object, pa As Object',
             '    Dim cos As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("B2").Select',
             '    Set co = ws.ChartObjects.Add(150, 10, 200, 120): co.Chart.SetSourceData ws.Range("A1:B4")',
             '    co.Chart.ChartType = xlColumnClustered',
             '    Set co2 = ws.ChartObjects.Add(150, 150, 200, 120): co2.Chart.SetSourceData ws.Range("A1:B4")',
             '    Set ca = co.Chart.ChartArea: Set pa = co.Chart.PlotArea: Set cos = ws.ChartObjects',
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
