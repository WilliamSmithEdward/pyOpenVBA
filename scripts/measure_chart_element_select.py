"""Selecting a point or a series of a chart: what Selection and ActiveChart answer afterwards, and what moves them on.

Each case makes a workbook with two sheets, the first active, and on it a chart of categories in A2:A4 and two series in
B and C, of the chart type TYPES names. It runs the case's statements, then reads READS, each as its value or "E" and an
error number.

    python scripts/measure_chart_element_select.py

writes tests/fixtures/chart_element_select.json, which tests/test_excel_chart_element_select.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone

from pyvbaharness import ExcelSession, HarnessConfig

from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_element_select.json'
TYPES = {'column': 'xlColumnClustered', 'line': 'xlLine', 'markers': 'xlLineMarkers', 'scatter': 'xlXYScatter', 'pie': 'xlPie', 'area': 'xlArea', 'bar': 'xlBarClustered', 'doughnut': 'xlDoughnut'}
P, Q, S, T = 'ch.SeriesCollection(1).Points(2)', 'ch.SeriesCollection(1).Points(1)', 'ch.SeriesCollection(1)', 'ch.SeriesCollection(2)'
ON_EVERY_KIND = [f'{P}.Select', f'{Q}.Select', f'{T}.Points(3).Select', f'{S}.Select', f'{T}.Select', f'ch.Parent.Activate: {P}.Select', f'ch.Parent.Activate: {S}.Select', f'{S}.Select: {P}.Select', f'{P}.Select: {S}.Select']
ON_COLUMN = [
    f'{P}.Select: ws.Range("B2").Select', f'{P}.Select: ws.Range("B2").Activate', f'{S}.Select: ws.Range("B2").Select', f'{P}.Select: ch.Parent.Activate', f'{P}.Select: ch.ChartArea.Select',
    f'{P}.Select: ch.PlotArea.Select', f'{P}.Select: {P}.Delete', f'{P}.Select: {T}.Delete', f'{S}.Select: {S}.Delete', f'{P}.Select: ch.Parent.Delete', f'{S}.Select: ch.Parent.Delete',
    f'{P}.Select: ws2.Activate', f'{P}.Select: ws2.Activate: ws.Activate', f'{S}.Select: ws2.Activate: ws.Activate', f'ws2.Activate: {P}.Select', f'ws2.Activate: {S}.Select',
    f'{P}.Select: Selection.MarkerSize = 12', f'{P}.Select: Selection.Format.Fill.ForeColor.RGB = 255', f'{S}.Select: Selection.Format.Fill.ForeColor.RGB = 255',
    f'{P}.Select: ch.ChartType = xlLine', f'{P}.Select: ws.Range("D1").Value = 5', f'{P}.Select: ws.Range("B2").Value = 5', f'{P}.Select: ActiveWindow.ScrollRow = 2',
    f'{S}.Select: ch.SeriesCollection.NewSeries', f'{P}.Select: ch.SetSourceData ws.Range("A1:B4")', f'{S}.Select: ch.SetSourceData ws.Range("A1:B4")',
]
CASES = [(kind, statement) for kind in TYPES for statement in ON_EVERY_KIND] + [('column', statement) for statement in ON_COLUMN]
READS = ['TypeName(Selection)', 'Selection.Name', 'Selection.Parent.Name', 'TypeName(ActiveChart)', 'ActiveChart.Parent.Name', 'ActiveSheet.Name', 'TypeName(ActiveCell)', 'ActiveCell.Address',
         'ws.Parent.Windows(1).RangeSelection.Address', 'TypeName(ActiveWindow.Selection)', 'ws.ChartObjects.Count', 'ch.SeriesCollection.Count']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, ws2 As Worksheet, ch As Object, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ws2 = wb.Worksheets.Add(After:=ws)', '    ws.Activate',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart', '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statement) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {TYPES[kind]}', f'        {statement}']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    wb.Close False', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i)', '    Next', '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 30):
            last = min(first + 29, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (first, last), timeout=280.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{kind}: {statement}', answer] for (kind, statement), answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case.replace(P, 'P').replace(Q, 'Q').replace(S, 'S').replace(T, 'T'), '=>', answer)


if __name__ == '__main__':
    main()
