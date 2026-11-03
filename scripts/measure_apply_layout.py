"""Chart.ApplyLayout: each layout on each kind of chart, over charts already changed, and with a ChartType: what the chart
reads afterwards, and the chart part Excel saves.

Each case makes a workbook with a chart of categories in A2:A4 and two series in B and C, of the chart type TYPES names,
runs the case's statements, reads READS, each as its value or "E" and an error number, and saves the workbook. The chart
part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_apply_layout.py

writes tests/fixtures/apply_layout.json, which tests/test_excel_apply_layout.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/apply_layout.json'
TYPES = {'column': 'xlColumnClustered', 'bar': 'xlBarClustered', 'line': 'xlLine', 'markers': 'xlLineMarkers', 'scatter': 'xlXYScatter', 'area': 'xlArea', 'pie': 'xlPie', 'doughnut': 'xlDoughnut'}
S1, S2 = 'ch.SeriesCollection(1)', 'ch.SeriesCollection(2)'
OVER = ['ch.Axes(xlValue).HasMinorGridlines = True', 'ch.Axes(xlValue).HasMajorGridlines = False', 'ch.Axes(xlCategory).HasMajorGridlines = True', 'ch.HasDataTable = True', f'{S1}.HasDataLabels = True',
        f'{S1}.HasDataLabels = True: {S1}.DataLabels.ShowCategoryName = True', 'ch.Axes(xlValue).HasTitle = True: ch.Axes(xlValue).AxisTitle.Text = "V"', 'ch.Axes(xlCategory).HasTitle = True', 'ch.HasTitle = True: ch.ChartTitle.Text = "T"',
        'ch.HasTitle = True', 'ch.HasLegend = False', 'ch.Legend.Position = xlLegendPositionLeft', 'ch.HasAxis(xlValue) = False', 'ch.HasAxis(xlCategory) = False', 'ch.ChartGroups(1).GapWidth = 33', 'ch.ChartGroups(1).Overlap = 10',
        'ch.Axes(xlValue).MajorTickMark = xlCross', 'ch.Axes(xlCategory).TickLabelPosition = xlLow', f'{S1}.Format.Fill.ForeColor.RGB = 255', 'ch.ChartStyle = 5', 'ch.ApplyLayout 2', 'ch.ApplyLayout 3', 'ch.ApplyLayout 5', 'ch.ApplyLayout 6', 'ch.ApplyLayout 7']
CASES = [(kind, f'ch.ApplyLayout {layout}') for kind in TYPES for layout in range(0, 14)]
CASES += [('column', f'{before}: ch.ApplyLayout {layout}') for before in OVER for layout in (1, 4, 11)]
CASES += [('column', f'ch.ApplyLayout {layout}, {kind}') for kind in ('xlLine', 'xlPie', 'xlColumnClustered', 'xlBarClustered') for layout in (1, 3, 9)]
CASES += [(kind, '') for kind in TYPES]
READS = [
    'ch.HasTitle', 'ch.ChartTitle.Text', 'ch.ChartTitle.IncludeInLayout', 'ch.HasLegend', 'ch.Legend.Position', 'ch.Legend.IncludeInLayout', 'ch.HasAxis(xlCategory)', 'ch.HasAxis(xlValue)',
    'ch.Axes(xlValue).HasMajorGridlines', 'ch.Axes(xlValue).HasMinorGridlines', 'ch.Axes(xlCategory).HasMajorGridlines', 'ch.Axes(xlCategory).HasMinorGridlines', 'ch.Axes(xlValue).HasTitle', 'ch.Axes(xlCategory).HasTitle',
    'ch.Axes(xlValue).AxisTitle.Text', 'ch.Axes(xlCategory).AxisTitle.Text', 'ch.Axes(xlValue).MajorTickMark', 'ch.Axes(xlCategory).MajorTickMark', 'ch.Axes(xlValue).TickLabelPosition', 'ch.Axes(xlCategory).TickLabelPosition',
    'ch.ChartGroups(1).GapWidth', 'ch.ChartGroups(1).Overlap', 'ch.HasDataTable', f'{S1}.HasDataLabels', f'{S2}.HasDataLabels', f'{S1}.DataLabels.ShowValue', f'{S1}.DataLabels.ShowCategoryName', f'{S1}.DataLabels.ShowSeriesName', f'{S1}.DataLabels.ShowPercentage',
    f'{S1}.DataLabels.Position', f'{S1}.HasLeaderLines', f'{S1}.Points(3).HasDataLabel', f'{S1}.Points(2).HasDataLabel', f'{S1}.Points(3).DataLabel.ShowSeriesName', f'{S1}.Points(3).DataLabel.ShowValue', f'{S1}.Points(3).DataLabel.Text', 'ch.ChartType', 'ch.ChartStyle']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart', '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statement) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {TYPES[kind]}']
        lines += [f'        {statement}'] if statement else []
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 30):
            last = min(first + 29, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
        parts = [chart_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{kind}: {statement}', answer, part] for (kind, statement), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
