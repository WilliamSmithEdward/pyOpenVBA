"""Shapes.AddChart2 and Shapes.AddChart, as recorded macros make charts: what each answers, what the chart reads, and each chart part
Excel saves, its style and colours among them.

Each case makes a workbook with categories in A2:A5 and series in B and C, selects a cell or none in the data, runs the case's
statements, reads READS, each as its value or "E" and an error number, and saves the workbook.

    python scripts/measure_add_chart2.py

writes tests/fixtures/add_chart2.json, which tests/test_excel_add_chart2.py replays.
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

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/add_chart2.json'
A = 'ws.Shapes.AddChart2'
AWAY = 'ws.Range("H20").Select: '
CASES = [
    # What it answers, where it goes.
    f'{AWAY}Set n = {A}', f'{AWAY}Set n = {A}()', f'{AWAY}r = TypeName({A}())', f'{AWAY}Set n = {A}(201)', f'{AWAY}Set n = {A}(-1)', f'{AWAY}Set n = {A}(201, xlColumnClustered)', f'{AWAY}Set n = {A}(201, xlColumnClustered, 10, 20)', f'{AWAY}Set n = {A}(201, xlColumnClustered, 10, 20, 300, 150)', f'{AWAY}Set n = {A}(Style:=201, XlChartType:=xlLine, Left:=10, Top:=20, Width:=300, Height:=150)', f'{AWAY}Set n = {A}(201, xlColumnClustered, 10, 20, 300, 150, True)', f'{AWAY}Set n = {A}(201, xlColumnClustered, 10, 20, 300, 150, False)', f'{AWAY}Set n = {A}(201, xlColumnClustered, NewLayout:=True)', f'{AWAY}Set n = {A}(201, xlColumnClustered, NewLayout:=False)',
    # Each kind at its usual style.
    *[f'{AWAY}Set n = {A}({style}, {kind})' for style, kind in ((201, 'xlBarClustered'), (227, 'xlLine'), (227, 'xlLineMarkers'), (251, 'xlPie'), (251, 'xlDoughnut'), (240, 'xlXYScatter'), (276, 'xlArea'), (297, 'xlColumnStacked'), (297, 'xlBarStacked'), (-1, 'xlLine'), (-1, 'xlPie'), (-1, 'xlXYScatter'), (-1, 'xlColumnClustered'), (201, 'xlLine'), (202, 'xlColumnClustered'), (205, 'xlColumnClustered'))],
    # Picking up the data round the selection.
    'ws.Range("B2").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("A1:B3").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("A1:B5").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("A1:B5").Select: Set n = ws.Shapes.AddChart2(251, xlPie)', 'ws.Range("B1:B5").Select: Set n = ws.Shapes.AddChart2(227, xlLine)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(240, xlXYScatter)',
    # The recorded macro's next steps.
    f'{AWAY}Set n = {A}(201, xlColumnClustered): n.Chart.SetSourceData ws.Range("A1:C5")', f'{AWAY}Set n = {A}(201, xlColumnClustered): n.Chart.SetSourceData ws.Range("A1:B5")', f'{AWAY}Set n = {A}(201, xlColumnClustered): n.Chart.SetSourceData ws.Range("A1:C5"): n.Chart.ChartTitle.Text = "T"', f'{AWAY}Set n = {A}(201, xlColumnClustered): n.Chart.SetSourceData ws.Range("A1:C5"): n.Chart.HasTitle = False', f'{AWAY}Set n = {A}(201, xlColumnClustered): n.Chart.SetSourceData ws.Range("A1:C5"): n.Chart.HasLegend = True', f'{AWAY}Set n = {A}(201, xlColumnClustered): n.Chart.SetSourceData ws.Range("A1:C5"): n.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = 255', f'{AWAY}Set n = {A}(201, xlColumnClustered): n.Chart.SetSourceData ws.Range("A1:C5"): n.Chart.ChartType = xlLine', f'{AWAY}Set n = {A}(201, xlColumnClustered): n.Chart.SetSourceData ws.Range("A1:C5"): n.Chart.Axes(xlValue).MajorGridlines.Delete', f'{AWAY}Set n = {A}(201, xlColumnClustered): n.Chart.SeriesCollection.NewSeries', f'{AWAY}Set n = {A}(201, xlColumnClustered): n.Name = "Mine"', f'{AWAY}Set n = {A}(201, xlColumnClustered): n.Select', f'{AWAY}{A}(201, xlColumnClustered).Select', f'{AWAY}{A}(201, xlColumnClustered).Select: ActiveChart.SetSourceData ws.Range("A1:C5")', f'{AWAY}{A}(201, xlColumnClustered).Select: r = TypeName(ActiveChart) & TypeName(Selection)', f'{AWAY}Set n = {A}(201, xlColumnClustered): r = n.Chart.ChartStyle & " " & n.Chart.ChartColor', f'{AWAY}Set n = {A}(227, xlLine): r = n.Chart.ChartStyle & " " & n.Chart.ChartColor', f'{AWAY}Set n = {A}(-1, xlColumnClustered): r = n.Chart.ChartStyle & " " & n.Chart.ChartColor',
    # What it refuses.
    f'{AWAY}Set n = {A}(5, xlColumnClustered)', f'{AWAY}Set n = {A}(0, xlColumnClustered)', f'{AWAY}Set n = {A}(9999, xlColumnClustered)', f'{AWAY}Set n = {A}(201, 9999)', f'{AWAY}Set n = {A}("a")', f'{AWAY}Set n = {A}(201, xlColumnClustered, -5, -5)', f'{AWAY}Set n = {A}(201, xlColumnClustered, 1, 1, 0, 0)',
    # Each kind with its data.
    *[f'ws.Range("A1:{last}5").Select: Set n = {A}({style}, {kind})' for style, kind in ((201, 'xlColumnClustered'), (201, 'xlBarClustered'), (227, 'xlLine'), (227, 'xlLineMarkers'), (251, 'xlPie'), (251, 'xlDoughnut'), (240, 'xlXYScatter'), (276, 'xlArea'), (297, 'xlColumnStacked'), (297, 'xlBarStacked')) for last in 'BC'],
    # The older one.
    f'{AWAY}Set n = ws.Shapes.AddChart', f'{AWAY}Set n = ws.Shapes.AddChart()', f'{AWAY}Set n = ws.Shapes.AddChart(xlLine)', f'{AWAY}Set n = ws.Shapes.AddChart(xlColumnClustered, 10, 20, 300, 150)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart()', f'{AWAY}r = TypeName(ws.Shapes.AddChart())', f'{AWAY}Set n = ws.Shapes.AddChart(xlColumnClustered, 10, 20, 300, 150, True)',
    # What a styled chart reads.
    *[f'ws.Range("A1:C5").Select: Set n = {A}(201, xlColumnClustered): r = {read}' for read in ('n.Chart.ChartTitle.Font.Size', 'n.Chart.ChartTitle.Font.Bold', 'n.Chart.ChartTitle.Font.Color', 'n.Chart.Axes(xlValue).TickLabels.Font.Size', 'n.Chart.Axes(xlValue).TickLabels.Font.Color', 'n.Chart.Axes(xlCategory).MajorTickMark', 'n.Chart.Axes(xlValue).MajorTickMark', 'n.Chart.Axes(xlValue).HasMajorGridlines', 'n.Chart.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB', 'n.Chart.Axes(xlValue).MajorGridlines.Format.Line.Weight', 'n.Chart.Axes(xlValue).Format.Line.Visible', 'n.Chart.Axes(xlCategory).Format.Line.ForeColor.RGB', 'n.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB', 'n.Chart.SeriesCollection(2).Format.Fill.ForeColor.ObjectThemeColor', 'n.Chart.SeriesCollection(1).Format.Line.Visible', 'n.Chart.ChartGroups(1).GapWidth', 'n.Chart.ChartGroups(1).Overlap', 'n.Chart.ChartArea.Format.Fill.ForeColor.RGB', 'n.Chart.ChartArea.Format.Line.ForeColor.RGB', 'n.Chart.ChartArea.Format.Line.Weight', 'n.Chart.ChartArea.Font.Size', 'n.Chart.PlotArea.Format.Fill.Visible', 'n.Chart.Legend.Position', 'n.Chart.Legend.Font.Size', 'n.Chart.ChartArea.RoundedCorners', 'n.RoundedCorners', 'n.Chart.ChartObjects(1).RoundedCorners')],
    *[f'ws.Range("A1:C5").Select: Set n = {A}({style}, {kind}): r = {read}' for style, kind, read in ((227, 'xlLine', 'n.Chart.SeriesCollection(1).MarkerStyle'), (227, 'xlLine', 'n.Chart.SeriesCollection(1).Format.Line.Weight'), (227, 'xlLine', 'n.Chart.SeriesCollection(1).Format.Line.ForeColor.RGB'), (227, 'xlLineMarkers', 'n.Chart.ChartType'), (240, 'xlXYScatter', 'n.Chart.SeriesCollection(1).MarkerStyle'), (240, 'xlXYScatter', 'n.Chart.SeriesCollection(1).MarkerSize'), (251, 'xlPie', 'n.Chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB'), (251, 'xlPie', 'n.Chart.SeriesCollection(1).Points.Count'), (251, 'xlDoughnut', 'n.Chart.ChartGroups(1).DoughnutHoleSize'), (251, 'xlPie', 'n.Chart.ChartGroups(1).FirstSliceAngle'))],
    # What editing a styled chart adds.
    *[f'ws.Range("A1:C5").Select: Set n = {A}(201, xlColumnClustered): {edit}' for edit in ('n.Chart.HasLegend = False', 'n.Chart.Legend.Position = xlLegendPositionRight', 'n.Chart.SeriesCollection(1).HasDataLabels = True', 'n.Chart.ApplyDataLabels', 'n.Chart.SeriesCollection(1).Trendlines.Add', 'n.Chart.SeriesCollection(1).ErrorBar xlY, xlBoth, xlStandardError', 'n.Chart.Axes(xlValue).HasTitle = True', 'n.Chart.Axes(xlCategory).HasTitle = True', 'n.Chart.Axes(xlCategory).HasMajorGridlines = True', 'n.Chart.Axes(xlValue).HasMinorGridlines = True', 'n.Chart.HasDataTable = True', 'n.Chart.ChartType = xlPie', 'n.Chart.ChartType = xlXYScatter', 'n.Chart.ChartType = xlArea', 'n.Chart.ChartType = xlBarClustered', 'n.Chart.ChartStyle = 227', 'n.Chart.ChartStyle = 5', 'n.Chart.ClearToMatchStyle', 'n.Chart.SetElement msoElementLegendBottom', 'n.Chart.ApplyLayout 3', 'n.Chart.SeriesCollection(1).Delete', 'n.Chart.SeriesCollection.NewSeries', 'n.Chart.SetSourceData ws.Range("A1:B5")', 'n.Chart.ChartTitle.Text = "T": n.Chart.ChartTitle.Font.Size = 20', 'n.Chart.ChartGroups(1).GapWidth = 50', 'n.Chart.Axes(xlValue).MajorTickMark = xlOutside', 'n.Chart.Axes(xlValue).TickLabels.Font.Size = 12', 'n.Chart.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = 255', 'n.Chart.ChartArea.Format.Fill.ForeColor.RGB = 255', 'n.Chart.PlotArea.Format.Fill.ForeColor.RGB = 255', 'n.Chart.Legend.Format.Fill.ForeColor.RGB = 255')],
    # Where else the data can come from.
    f'{AWAY}Set n = ws.Range("H20:I22").Parent.Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("H20:I22").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("A1:A5,C1:C5").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("A1:C5").Select: Set n = wb.Worksheets.Add(After:=ws).Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("A1:C5").Select: wb.Worksheets.Add After:=ws: Set n = ws.Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("B2:B5").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("A2:B5").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("A1").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered, NewLayout:=False)', 'ws.Range("A1:B5").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered, NewLayout:=False)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(5, xlColumnClustered)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(5, xlLine, NewLayout:=False)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(48, xlColumnClustered)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(49, xlColumnClustered)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(200, xlColumnClustered)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(216, xlBarClustered)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(-1, xlBarClustered)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(-1, xlColumnStacked)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(-1, xlArea)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(-1, xlLineMarkers)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(332, xlLineMarkers)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(-1, xlDoughnut)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(-1, xlBubble)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(-1, xlRadar)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(-1, xl3DColumn)',
    # Numbers across, more series than colours, more than one chart, and copies.
    'ws.Range("B1:C5").Select: Set n = ws.Shapes.AddChart2(240, xlXYScatter)', 'ws.Range("B2:C5").Select: Set n = ws.Shapes.AddChart2(240, xlXYScatter)', 'ws.Range("D1:J5").Value = 3: ws.Range("A1:J5").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("D1:J5").Value = 3: ws.Range("A1:J5").Select: Set n = ws.Shapes.AddChart2(227, xlLine)', 'ws.Range("D1:E5").Value = 3: ws.Range("A1:E2").Select: Set n = ws.Shapes.AddChart2(251, xlPie)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered): Set n = ws.Shapes.AddChart2(227, xlLine)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered): n.Delete: Set n = ws.Shapes.AddChart2(227, xlLine)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered): Set n = n.Duplicate', f'{AWAY}Set n = ws.Shapes.AddChart2(201, xlColumnClustered): ws.Protect DrawingObjects:=True: Set n = ws.Shapes.AddChart2(201, xlColumnClustered)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart2(201, xlColumnClustered, Left:=1.2, Top:=2.3, Width:=100.4, Height:=50.6)', f'{AWAY}Set n = ws.Shapes.AddChart2(201, xlColumnClustered, 1, 1, -5, -5)',
    'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart(xlLine)', 'ws.Range("A1:B5").Select: Set n = ws.Shapes.AddChart(xlPie)', 'ws.Range("A1:B5").Select: Set n = ws.Shapes.AddChart(xlColumnClustered)', 'ws.Range("A1:C5").Select: Set n = ws.Shapes.AddChart(xlBarClustered): n.Chart.HasTitle = True',
]
READS = ['n.Name', 'n.Type', 'n.HasChart', 'n.Left', 'n.Top', 'n.Width', 'n.Height', 'n.ZOrderPosition', 'n.Chart.ChartType', 'n.Chart.HasTitle', 'n.Chart.HasLegend', 'n.Chart.SeriesCollection.Count', 'n.Chart.SeriesCollection(1).Formula', 'n.Chart.ChartTitle.Text', 'n.Chart.Parent.Name', 'ws.Shapes.Count', 'ws.ChartObjects.Count', 'TypeName(ActiveChart)', 'TypeName(Selection)']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Object, n As Object, s As String, v As String, r As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:A5").Value = Application.Transpose(Array("k", "a", "b", "c", "d"))', '    ws.Range("B1:B5").Value = Application.Transpose(Array("v", 1, 2, 3, 4))', '    ws.Range("C1:C5").Value = Application.Transpose(Array("w", 5, 6, 7, 8))', '    On Error GoTo Failed', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def chart_parts(path: Path) -> dict[str, str]:
    """Every part under xl/charts, the chart's own with its random ids made comparable."""
    found: dict[str, str] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name.startswith('xl/charts/'):
                text = package.read(name).decode('utf-8')
                found[name] = normalized(text) if re.fullmatch(r'xl/charts/chart\d+\.xml', name) else text
    return found


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 25):
            last = min(first + 24, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
        parts = [chart_parts(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    styles: dict[str, str] = {}
    for part in parts:
        for name in list(part):
            if re.fullmatch(r'xl/charts/(style|colors)\d+\.xml', name):
                text = part[name]
                key = f'{"style" if "style" in name else "colors"}{len(styles)}'
                for other, held in styles.items():
                    if held == text:
                        key = other
                styles[key] = text
                part[name] = key
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source, 'styles': styles,
              'cases': [[case, answer, part] for case, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        print(case, '=>', answer, {name[10:]: len(text) for name, text in part.items()})


if __name__ == '__main__':
    main()
