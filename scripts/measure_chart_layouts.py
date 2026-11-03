"""Chart.ChartStyle, ChartColor, ApplyLayout and ChartWizard: what each answers, what the chart reads afterwards, and the
chart part Excel saves.

Each case makes a workbook with a chart of categories in A2:A4 and two series in B and C, of the chart type TYPES names,
runs the case's statements, reads READS, each as its value or "E" and an error number, and saves the workbook. The chart
part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_layouts.py

writes tests/fixtures/chart_layouts.json, which tests/test_excel_chart_layouts.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import normalized

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_layouts.json'
TYPES = {'column': 'xlColumnClustered', 'line': 'xlLineMarkers', 'pie': 'xlPie', 'scatter': 'xlXYScatter'}
STYLES = [f'ch.ChartStyle = {style}' for style in (1, 2, 10, 48, 49, 0, 201, 227, 300, 352, 353, -1)] + ['ch.ClearToMatchStyle', 'ch.ChartStyle = 5: ch.ClearToMatchStyle']
COLOURS = [f'ch.ChartColor = {colour}' for colour in (1, 10, 14, 26, 27, 0)]
LAYOUTS = [f'ch.ApplyLayout {layout}' for layout in (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 0)] + ['ch.ApplyLayout 1, xlLine', 'ch.HasTitle = True: ch.ChartTitle.Text = "T": ch.ApplyLayout 1', 'ch.HasLegend = False: ch.ApplyLayout 3']
WIZARD = ['ch.ChartWizard', 'ch.ChartWizard ws.Range("A1:B4")', 'ch.ChartWizard Gallery:=xlLine', 'ch.ChartWizard Gallery:=xlColumn, Format:=1', 'ch.ChartWizard Gallery:=xlBar', 'ch.ChartWizard Gallery:=xlPie',
          'ch.ChartWizard PlotBy:=xlRows', 'ch.ChartWizard PlotBy:=xlColumns', 'ch.ChartWizard HasLegend:=False', 'ch.ChartWizard HasLegend:=True', 'ch.ChartWizard Title:="T"', 'ch.ChartWizard CategoryTitle:="C"', 'ch.ChartWizard ValueTitle:="V"',
          'ch.ChartWizard ExtraTitle:="E"', 'ch.ChartWizard Source:=ws.Range("A1:C3"), Gallery:=xlColumn, Title:="T", CategoryTitle:="C", ValueTitle:="V", HasLegend:=False', 'ch.ChartWizard CategoryLabels:=0', 'ch.ChartWizard SeriesLabels:=0',
          'ch.ChartWizard Source:=ws.Range("B2:C4"), CategoryLabels:=0, SeriesLabels:=0', 'ch.ChartWizard Source:=ws.Range("A1:C4"), CategoryLabels:=1, SeriesLabels:=1', 'ch.ChartWizard Gallery:=999', 'ch.ChartWizard Title:=""', 'ch.HasTitle = True: ch.ChartWizard Title:=""']
CASES = [(kind, statement) for kind in TYPES for statement in STYLES + COLOURS + LAYOUTS]
CASES += [(kind, statement) for kind in ('column', 'line') for statement in WIZARD]
CHANGES = ['ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = 255', 'ch.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = 255', 'ch.ChartArea.Format.Fill.ForeColor.RGB = 255',
           'ch.HasTitle = True: ch.ChartTitle.Font.Bold = True', 'ch.Legend.Font.Bold = True', 'ch.SeriesCollection(1).HasDataLabels = True: ch.SeriesCollection(1).DataLabels.Font.Bold = True',
           'ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = 255']
CASES += [('column', f'{change}: {after}') for change in CHANGES for after in ('ch.ChartStyle = 5', 'ch.ClearToMatchStyle')]
CASES += [('column', statement) for statement in ('ch.ChartStyle = 5: ch.ChartType = xlLine', 'ch.ChartStyle = 5: ch.ChartStyle = 2', 'ch.ChartStyle = 5: ch.SetSourceData ws.Range("A1:B4")', 'ch.ChartStyle = 5: ch.Parent.Duplicate')]
CASES += [(kind, '') for kind in TYPES]
READS = ['ch.ChartStyle', 'ch.ChartColor', 'ch.HasTitle', 'ch.ChartTitle.Text', 'ch.HasLegend', 'ch.Legend.Position', 'ch.HasAxis(xlCategory)', 'ch.HasAxis(xlValue)', 'ch.Axes(xlValue).HasMajorGridlines', 'ch.Axes(xlValue).HasTitle', 'ch.Axes(xlValue).AxisTitle.Text',
         'ch.Axes(xlCategory).HasTitle', 'ch.Axes(xlCategory).AxisTitle.Text', 'ch.SeriesCollection(1).HasDataLabels', 'ch.HasDataTable', 'ch.ChartType', 'ch.SeriesCollection.Count', 'ch.SeriesCollection(1).Formula', 'ch.PlotBy']


def _part(path: Path) -> str:
    """The chart part, after the names of any style parts that came with it."""
    with zipfile.ZipFile(path) as package:
        names = [entry for entry in package.namelist() if entry.startswith('xl/charts/')]
        chart = next(entry for entry in names if re.fullmatch(r'xl/charts/chart\d+\.xml', entry))
        extra = ' '.join(entry.rsplit('/', 1)[1] for entry in names if entry != chart)
        return (f'EXTRA {extra} ' if extra else '') + normalized(package.read(chart).decode('utf-8'))


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
        parts = [_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{kind}: {statement}', answer, part] for (kind, statement), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
