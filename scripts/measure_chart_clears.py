"""Clearing a chart's parts, its legend entries, and what a chart objects collection can do. What each reads, and the chart part
Excel saves.

Each case makes a workbook with categories in A2:A4 and two series in B and C, and a column chart of them. It runs the
case's statements, which set r to what they want recorded, reads READS, each as its value or "E" and an error number, and
saves the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it, or as "" for a workbook left
with other than one chart.

    python scripts/measure_chart_clears.py

writes tests/fixtures/chart_clears.json, which tests/test_excel_chart_clears.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_clears.json'
S1 = 'ch.SeriesCollection(1)'
AREA_LOOK = 'ch.ChartArea.Format.Fill.ForeColor.RGB = 255: ch.ChartArea.Format.Line.ForeColor.RGB = 65280'
PLOT_LOOK = 'ch.PlotArea.Format.Fill.ForeColor.RGB = 255: ch.PlotArea.Format.Line.ForeColor.RGB = 65280'
EVERYTHING = f'{AREA_LOOK}: {PLOT_LOOK}: ch.HasTitle = True: ch.ChartTitle.Text = "t": ch.Legend.Format.Fill.ForeColor.RGB = 255: {S1}.Format.Fill.ForeColor.RGB = 255: ch.Axes(xlValue).Format.Line.ForeColor.RGB = 255: {S1}.HasDataLabels = True'
AREA_READ = 'r = ch.ChartArea.Format.Fill.ForeColor.RGB & " " & ch.ChartArea.Format.Line.ForeColor.RGB & " " & ch.ChartArea.Format.Fill.Visible & " " & ch.ChartArea.Format.Line.Visible'
PLOT_READ = 'r = ch.PlotArea.Format.Fill.ForeColor.RGB & " " & ch.PlotArea.Format.Line.ForeColor.RGB & " " & ch.PlotArea.Format.Fill.Visible & " " & ch.PlotArea.Format.Line.Visible'
E1, KEY = 'ch.Legend.LegendEntries(1)', 'ch.Legend.LegendEntries(1).LegendKey'
CASES = [
    # The chart area.
    'r = ch.ChartArea.ClearFormats', 'r = ch.ChartArea.Clear', 'r = ch.ChartArea.ClearContents', f'{AREA_LOOK}: ch.ChartArea.ClearFormats: {AREA_READ}', f'{AREA_LOOK}: ch.ChartArea.Clear: {AREA_READ}', f'{AREA_LOOK}: ch.ChartArea.ClearContents: {AREA_READ}',
    f'{EVERYTHING}: ch.ChartArea.ClearFormats: {AREA_READ}', f'{EVERYTHING}: ch.ChartArea.Clear: {AREA_READ}', f'{EVERYTHING}: ch.ChartArea.ClearContents: {AREA_READ}', 'ch.ChartArea.Clear: ch.SetSourceData ws.Range("A1:C4"): r = ch.SeriesCollection.Count', 'ch.ChartArea.ClearContents: ch.SetSourceData ws.Range("A1:C4"): r = ch.SeriesCollection.Count', 'ch.ChartArea.Clear: r = ch.HasLegend & " " & ch.HasTitle', 'ch.HasTitle = True: ch.ChartArea.Clear: r = ch.HasLegend & " " & ch.HasTitle', 'ch.HasTitle = True: ch.ChartArea.ClearContents: r = ch.HasLegend & " " & ch.HasTitle', 'ch.HasTitle = True: ch.ChartArea.ClearFormats: r = ch.HasLegend & " " & ch.HasTitle',
    # The plot area and a trendline.
    'r = ch.PlotArea.ClearFormats', f'{PLOT_LOOK}: ch.PlotArea.ClearFormats: {PLOT_READ}', f'{EVERYTHING}: ch.PlotArea.ClearFormats: {PLOT_READ}', f'{S1}.Trendlines.Add: r = {S1}.Trendlines(1).ClearFormats', f'{S1}.Trendlines.Add: {S1}.Trendlines(1).Border.Color = 255: {S1}.Trendlines(1).ClearFormats: r = {S1}.Trendlines(1).Border.Color', f'{S1}.Trendlines.Add: {S1}.Trendlines(1).Format.Line.Weight = 3: {S1}.Trendlines(1).ClearFormats: r = {S1}.Trendlines(1).Format.Line.Weight',
    'r = ch.Legend.Clear', 'ch.Legend.Format.Fill.ForeColor.RGB = 255: ch.Legend.Clear: r = ch.HasLegend', 'ch.HasTitle = True: r = ch.ChartTitle.Clear',
    # A legend entry.
    f'r = TypeName({E1})', f'r = {E1}.Index', 'r = ch.Legend.LegendEntries(2).Index', f'r = TypeName({E1}.Font)', f'r = {E1}.Font.Size & " " & {E1}.Font.Bold & " " & {E1}.Font.Color & " " & {E1}.Font.Name', f'{E1}.Font.Size = 14: r = {E1}.Font.Size & " " & ch.Legend.LegendEntries(2).Font.Size & " " & ch.Legend.Font.Size', f'{E1}.Font.Bold = True: r = {E1}.Font.Bold & " " & ch.Legend.Font.Bold', f'{E1}.Font.Color = 255: r = {E1}.Font.Color', f'ch.Legend.Font.Size = 14: r = {E1}.Font.Size', f'ch.Legend.Font.Size = 14: {E1}.Font.Size = 9: r = {E1}.Font.Size & " " & ch.Legend.Font.Size',
    f'r = TypeName({E1}.Format)', f'r = {E1}.Format.Line.Visible', f'r = {E1}.Format.Fill.Visible', f'r = TypeName({KEY})', f'r = TypeName({KEY}.Format)', f'r = {KEY}.Format.Fill.ForeColor.RGB', f'r = {KEY}.Format.Fill.Visible', f'r = {KEY}.Format.Line.Visible', f'{KEY}.Format.Fill.ForeColor.RGB = 255: r = {KEY}.Format.Fill.ForeColor.RGB & " " & {S1}.Format.Fill.ForeColor.RGB', f'{S1}.Format.Fill.ForeColor.RGB = 255: r = {KEY}.Format.Fill.ForeColor.RGB', f'{KEY}.Format.Line.ForeColor.RGB = 255: r = {KEY}.Format.Line.ForeColor.RGB & " " & {S1}.Format.Line.ForeColor.RGB', f'r = {KEY}.MarkerStyle', f'r = {KEY}.Smooth', f'r = {KEY}.InvertIfNegative', f'{KEY}.Delete', f'r = {E1}.Delete: r = ch.Legend.LegendEntries.Count', f'{E1}.Select: r = TypeName(Selection)', f'{KEY}.Select: r = TypeName(Selection)',
    # Chart objects together.
    'r = ws.ChartObjects.Select', 'ws.ChartObjects.Select: r = TypeName(Selection)', 'ws.ChartObjects.Add 10, 10, 100, 100: ws.ChartObjects.Select: r = TypeName(Selection) & Selection.Count', 'r = TypeName(ws.ChartObjects.Duplicate)', 'ws.ChartObjects.Duplicate: r = ws.ChartObjects.Count', 'ws.ChartObjects.Copy: ws.Paste: r = ws.ChartObjects.Count', 'r = ws.ChartObjects.Copy', 'ws.ChartObjects.Cut: ws.Paste: r = ws.ChartObjects.Count', 'r = ws.ChartObjects(1).Duplicate.Name', 'r = ws.ChartObjects(1).Copy', 'r = TypeName(ws.ChartObjects(Array(1)))', 'r = TypeName(ws.ChartObjects.ShapeRange)', 'r = ws.ChartObjects.ShapeRange.Count',
]
READS = ['ch.SeriesCollection.Count', 'ws.ChartObjects.Count', 'ch.HasLegend', 'ch.HasTitle', 'ch.ChartType', 'ActiveSheet.Name']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Object, ch As Object, s As String, v As String, r As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))', '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    Set ch = ws.ChartObjects.Add(300, 100, 300, 200).Chart', '    ch.SetSourceData ws.Range("A1:C4"), xlColumns', '    On Error GoTo Failed', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    Application.CutCopyMode = False', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _part(path: Path) -> str:
    try:
        return chart_part(path)
    except AssertionError:
        return ''


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
              'cases': [[case, answer, part] for case, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
