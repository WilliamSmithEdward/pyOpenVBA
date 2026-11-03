"""Series.IsFiltered: filtering a series out on each kind of chart, what its members and the chart's answer then, and the
chart part Excel saves.

Each case makes a workbook with a chart of categories in A2:A4 and two or three series in B to D, of the chart type TYPES
names, runs the case's statements, reads READS, each as its value or "E" and an error number, and saves the workbook. The
chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_series_filter.py

writes tests/fixtures/series_filter.json, which tests/test_excel_series_filter.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/series_filter.json'
TYPES = {'column': 'xlColumnClustered', 'bar': 'xlBarClustered', 'line': 'xlLine', 'markers': 'xlLineMarkers', 'scatter': 'xlXYScatter', 'area': 'xlArea', 'pie': 'xlPie', 'doughnut': 'xlDoughnut'}
F, S = 'ch.FullSeriesCollection', 'ch.SeriesCollection'
F1, F2, F3 = f'{F}(1)', f'{F}(2)', f'{F}(3)'
ON = '.IsFiltered = True'
OFF = '.IsFiltered = False'
ON_EVERY_KIND = [f'{F1}{ON}', f'{F2}{ON}', f'{F3}{ON}', f'{F1}{ON}: {F2}{ON}', f'{F1}{ON}: {F2}{ON}: {F3}{ON}', f'{F1}{ON}: {F1}{OFF}', f'{F2}{ON}: {F1}{ON}: {F2}{OFF}']
ON_COLUMN = [
    f'{F1}{ON}: {F1}{ON}', f'{S}(1){ON}: {S}(1){ON}', f'{F1}{ON}: {F1}.Name = "z"', f'{F1}{ON}: {F1}.Values = Array(7, 8, 9)', f'{F1}{ON}: {F1}.Format.Fill.ForeColor.RGB = 255', f'{F1}{ON}: {F1}.HasDataLabels = True',
    f'{F1}.Format.Fill.ForeColor.RGB = 255: {F1}{ON}', f'{F1}.HasDataLabels = True: {F1}{ON}', f'{F1}.Points(2).Format.Fill.ForeColor.RGB = 255: {F1}{ON}', f'{F1}{ON}: {F1}.Select', f'{F1}.Select: {F1}{ON}', f'{F1}{ON}: {F1}.Delete',
    f'{F1}{ON}: {S}(1).Delete', f'{F1}{ON}: {F2}.Delete', f'{F1}{ON}: {S}.NewSeries', f'{F1}{ON}: ch.SetSourceData ws.Range("A1:B4")', f'{F1}{ON}: ch.ChartType = xlLine', f'{F1}{ON}: ch.HasLegend = False', f'{F1}{ON}: {S}(1).PlotOrder = 2',
    f'{F1}{ON}: {F2}.PlotOrder = 2', f'{S}(1).PlotOrder = 2: {F1}{ON}', f'{F1}{ON}: ch.ChartTitle.Text = "T"', f'{F1}{ON}: {F2}{ON}: {F3}{ON}: ch.HasTitle = True', f'{F1}{ON}: ch.SetElement msoElementDataLabelShow', f'{F1}{ON}: ch.ApplyLayout 2',
    f'{F1}{ON}: ws.Range("B2").Value = 9', f'{F1}{ON}: {F1}.Points(1).HasDataLabel = True', f'{F1}{ON}: ws.Range("B:B").Delete', f'{F1}{ON}: ws.Rows(2).Insert',
]
CASES = [(kind, statement) for kind in TYPES for statement in ON_EVERY_KIND] + [('column', statement) for statement in ON_COLUMN] + [(kind, '') for kind in TYPES]
READS = [f'{S}.Count', f'{F}.Count', f'{S}(1).Name', f'{S}(2).Name', f'{F1}.Name', f'{F2}.Name', f'{F3}.Name', f'{F1}.IsFiltered', f'{F2}.IsFiltered', f'{F3}.IsFiltered', f'{S}(1).IsFiltered', f'{S}(1).PlotOrder', f'{S}(2).PlotOrder', f'{F1}.PlotOrder', f'{F2}.PlotOrder', f'{F3}.PlotOrder',
         f'{F1}.Formula', f'{S}(1).Formula', f'{F1}.Format.Fill.ForeColor.RGB', f'{F1}.HasDataLabels', f'{F1}.Values(1)', f'{F1}.Points.Count', f'TypeName({F1}.Parent)', f'{S}(1).Format.Fill.ForeColor.RGB', 'ch.HasLegend', 'ch.Legend.LegendEntries.Count', 'ch.HasTitle', 'ch.ChartTitle.Text',
         'TypeName(Selection)', 'ch.ChartType', 'ch.PlotBy']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart', '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))', '    ws.Range("D1:D4").Value = Application.Transpose(Array("x", 7, 8, 9))',
             '    ch.SetSourceData ws.Range("A1:D4")', '    On Error GoTo Failed', '    Select Case which']
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
        print(case.replace(F, 'F').replace(S, 'S'), '=>', answer)


if __name__ == '__main__':
    main()
