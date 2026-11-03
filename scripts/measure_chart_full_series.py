"""Chart.FullSeriesCollection, Series.IsFiltered, and Select on the other elements of a chart: what each answers, what
Selection answers afterwards, and the chart part Excel saves.

Each case makes a workbook with a chart of categories in A2:A4 and two series in B and C, of the chart type TYPES names,
runs the case's statements, reads the case's own expression and READS, each as its value or "E" and an error number, and
saves the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_full_series.py

writes tests/fixtures/chart_full_series.json, which tests/test_excel_chart_full_series.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_full_series.json'
TYPES = {'column': 'xlColumnClustered', 'line': 'xlLine', 'pie': 'xlPie', 'scatter': 'xlXYScatter'}
F, S = 'ch.FullSeriesCollection', 'ch.SeriesCollection'
LOOKS = [
    f'TypeName({F})', f'{F}.Count', f'TypeName({F}(1))', f'{F}(1).Name', f'{F}(2).Name', f'{F}("w").Name', f'{F}(3).Name', f'{F}(0).Name', f'{F}("nope").Name', f'{F}.Item(2).Name',
    f'TypeName({F}.Parent)', f'{F}(1).IsFiltered', f'{F}(2).IsFiltered', f'{S}(1).IsFiltered', f'{F}(1).PlotOrder', f'{F}(2).Formula', f'{F}(1) Is {S}(1)', f'TypeName({F}.Application)',
]
FILTER = [
    f'{F}(1).IsFiltered = True', f'{F}(2).IsFiltered = True', f'{F}(1).IsFiltered = False', f'{F}(1).IsFiltered = True: {F}(1).IsFiltered = False',
    f'{S}(1).IsFiltered = True', f'{F}(1).IsFiltered = True: {F}(2).IsFiltered = True', f'{F}(1).IsFiltered = True: {S}(1).Name = "z"', f'{F}(1).IsFiltered = True: {F}(1).Name = "z"',
    f'{F}(1).Select', f'{F}(2).Points(2).Select', f'{F}(1).Format.Fill.ForeColor.RGB = 255', f'{F}(1).Delete', f'{F}.NewSeries', f'{F}.Add ws.Range("B1:B4")', f'{F}(1).Name = "z"',
]
SELECTS = [
    'ch.Legend.Select', 'ch.ChartTitle.Select', 'ch.Axes(xlValue).Select', 'ch.Axes(xlCategory).Select', 'ch.Axes(xlValue).MajorGridlines.Select', 'ch.Axes(xlValue).AxisTitle.Select',
    f'{S}(1).DataLabels.Select', f'{S}(1).Points(1).DataLabel.Select', 'ch.Legend.LegendEntries(1).Select', f'{S}(1).Trendlines(1).Select', f'{S}(1).ErrorBars.Select', 'ch.ChartArea.Select', 'ch.PlotArea.Select',
]
BEFORE = {'ch.ChartTitle.Select': 'ch.HasTitle = True', 'ch.Axes(xlValue).AxisTitle.Select': 'ch.Axes(xlValue).HasTitle = True', f'{S}(1).DataLabels.Select': f'{S}(1).HasDataLabels = True',
          f'{S}(1).Points(1).DataLabel.Select': f'{S}(1).HasDataLabels = True', f'{S}(1).Trendlines(1).Select': f'{S}(1).Trendlines.Add', f'{S}(1).ErrorBars.Select': f'{S}(1).ErrorBar xlY, xlBoth, xlPercent, 5',
          'ch.Axes(xlValue).MajorGridlines.Select': 'ch.Axes(xlValue).HasMajorGridlines = True', 'ch.Legend.Select': 'ch.HasLegend = True', 'ch.Legend.LegendEntries(1).Select': 'ch.HasLegend = True'}
ABSENT = {'ch.Legend.Select': 'ch.HasLegend = False', 'ch.ChartTitle.Select': 'ch.HasTitle = False', 'ch.Axes(xlValue).AxisTitle.Select': '', f'{S}(1).DataLabels.Select': '', f'{S}(1).Points(1).DataLabel.Select': ''}
CASES: list[tuple[str, str, str]] = [('column', '', look) for look in LOOKS]
CASES += [(kind, statement, f'{F}(1).Name') for kind in ('column', 'line', 'pie') for statement in FILTER]
CASES += [(kind, f'{BEFORE[item]}: {item}' if item in BEFORE else item, '') for kind in ('column', 'scatter') for item in SELECTS]
CASES += [('pie', f'{BEFORE[item]}: {item}' if item in BEFORE else item, '') for item in SELECTS if 'Axes' not in item and 'Trendlines' not in item and 'ErrorBars' not in item]
CASES += [('column', f'{before}: {item}' if before else item, '') for item, before in ABSENT.items()]
CASES += [('column', f'ws2.Activate: {item}', '') for item in ('ch.Legend.Select', 'ch.Axes(xlValue).Select')]
CASES += [('column', f'ch.Legend.Select: ws.Range("A1").Select', ''), ('column', 'ch.Legend.Select: ch.HasLegend = False', ''), ('column', 'ch.Axes(xlValue).Select: ch.HasAxis(xlValue) = False', ''),
          ('column', 'ch.HasTitle = True: ch.ChartTitle.Select: ch.HasTitle = False', ''), ('column', 'ch.Legend.Select: Selection.Position = xlLegendPositionTop', ''),
          ('column', 'ch.Axes(xlValue).Select: Selection.MinimumScale = 1', ''), ('column', 'ch.HasTitle = True: ch.ChartTitle.Select: Selection.Text = "T"', '')]
READS = ['TypeName(Selection)', 'Selection.Name', 'TypeName(ActiveChart)', f'{S}.Count', f'{F}.Count', f'{S}(1).Name', f'{F}(1).Name', f'{F}(1).IsFiltered', f'{F}(2).IsFiltered', f'{S}(1).PlotOrder', f'{F}(2).PlotOrder']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ws2 As Worksheet, ch As Object, s As String, v As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ws2 = wb.Worksheets.Add(After:=ws)', '    ws.Activate',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart', '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statement, _) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {TYPES[kind]}']
        lines += [f'        {statement}'] if statement else []
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = ""', '    Select Case which']
    for index, (_, _, look) in enumerate(CASES):
        if look:
            lines += [f'    Case {index}', f'        v = {look}']
    lines += ['    End Select', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "=" & v']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    ws2.Delete', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function',
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
              'cases': [[f'{kind}: {statement} ? {look}', answer, part] for (kind, statement, look), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
