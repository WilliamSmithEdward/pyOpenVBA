"""The Border of a chart's elements, and a series' ClearFormats: what each reads, and the chart part Excel saves.

Each case makes a workbook with categories in A2:A4 and two series in B and C, and a chart of them of the type it names. It
runs the case's statements, which set r to what they want recorded, reads READS, each as its value or "E" and an error
number, and saves the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_borders.py

writes tests/fixtures/chart_borders.json, which tests/test_excel_chart_borders.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_borders.json'
TYPES = {'column': 'xlColumnClustered', 'line': 'xlLine', 'lineMarkers': 'xlLineMarkers', 'scatter': 'xlXYScatter', 'pie': 'xlPie', 'area': 'xlArea'}
S1 = 'ch.SeriesCollection(1)'
ELEMENTS = {
    'value axis': 'ch.Axes(xlValue)', 'category axis': 'ch.Axes(xlCategory)', 'gridlines': 'ch.Axes(xlValue).MajorGridlines', 'series': S1, 'point': f'{S1}.Points(2)',
    'chart area': 'ch.ChartArea', 'plot area': 'ch.PlotArea', 'legend': 'ch.Legend', 'title': 'ch.ChartTitle', 'axis title': 'ch.Axes(xlValue).AxisTitle', 'trendline': f'{S1}.Trendlines(1)', 'error bars': f'{S1}.ErrorBars',
}
SETUP = {'title': 'ch.HasTitle = True: ', 'axis title': 'ch.Axes(xlValue).HasTitle = True: ', 'trendline': f'{S1}.Trendlines.Add: ', 'error bars': f'{S1}.ErrorBar xlY, xlBoth, xlFixedValue, 1: '}
SETS = ['.Color = 255', '.ColorIndex = 3', '.ColorIndex = xlNone', '.ColorIndex = xlAutomatic', '.Weight = xlHairline', '.Weight = xlThin', '.Weight = xlMedium', '.Weight = xlThick', '.LineStyle = xlNone', '.LineStyle = xlContinuous', '.LineStyle = xlDash', '.LineStyle = xlDot', '.LineStyle = xlDashDot', '.LineStyle = xlDashDotDot', '.LineStyle = xlAutomatic', '.LineStyle = xlDouble']
CASES: list[tuple[str, str]] = []
for name, element in ELEMENTS.items():
    before = SETUP.get(name, '')
    target = f'{element}.Border'
    CASES.append(('column', f'{before}r = TypeName({target})'))
    CASES.append(('column', f'{before}r = {target}.Color & " " & {target}.ColorIndex & " " & {target}.Weight & " " & {target}.LineStyle'))
    CASES += [('column', f'{before}{target}{change}: r = {target}.Color & " " & {target}.ColorIndex & " " & {target}.Weight & " " & {target}.LineStyle') for change in SETS]
    CASES.append(('column', f'{before}{target}.Color = 255: {target}.Weight = xlThick: {target}.LineStyle = xlDash: r = {element}.Format.Line.ForeColor.RGB & " " & {element}.Format.Line.Weight & " " & {element}.Format.Line.DashStyle & " " & {element}.Format.Line.Visible'))
    CASES.append(('column', f'{before}{element}.Format.Line.ForeColor.RGB = 255: {element}.Format.Line.Weight = 3: r = {target}.Color & " " & {target}.ColorIndex & " " & {target}.Weight & " " & {target}.LineStyle'))
for kind, statement in [(kind, f'{S1}.Border.Color = 255: r = {S1}.Border.Color & " " & {S1}.Border.Weight & " " & {S1}.Border.LineStyle') for kind in ('line', 'lineMarkers', 'scatter', 'pie', 'area')]:
    CASES.append((kind, statement))
CASES += [(kind, f'r = {S1}.Border.Color & " " & {S1}.Border.Weight & " " & {S1}.Border.LineStyle') for kind in ('line', 'lineMarkers', 'scatter', 'pie', 'area')]
CASES += [(kind, f'{S1}.Border.LineStyle = xlNone: r = {S1}.Border.LineStyle') for kind in ('line', 'scatter', 'pie')]
CASES += [(kind, f'{S1}.Border.Weight = xlThick: r = {S1}.Border.Weight') for kind in ('line', 'scatter', 'pie')]
# A series' ClearFormats, on what has been set on it.
LOOKS = ['', f'{S1}.Format.Fill.ForeColor.RGB = 255', f'{S1}.Format.Line.ForeColor.RGB = 255', f'{S1}.Format.Line.Weight = 3', f'{S1}.Format.Fill.Visible = msoFalse', f'{S1}.Points(2).Format.Fill.ForeColor.RGB = 255', f'{S1}.HasDataLabels = True', f'{S1}.Points(2).HasDataLabel = True', f'{S1}.Trendlines.Add', f'{S1}.Name = "z"', f'{S1}.InvertIfNegative = True']
CLEAR_READS = f'r = {S1}.Format.Fill.Visible & " " & {S1}.Format.Fill.ForeColor.RGB & " " & {S1}.Format.Line.Visible & " " & {S1}.Format.Line.ForeColor.RGB & " " & {S1}.Format.Line.Weight & " " & {S1}.HasDataLabels'
CASES += [('column', f'{look}{": " if look else ""}{S1}.ClearFormats: {CLEAR_READS}') for look in LOOKS]
LINES = [f'{S1}.MarkerStyle = xlMarkerStyleSquare', f'{S1}.MarkerSize = 12', f'{S1}.MarkerBackgroundColor = 255', f'{S1}.Smooth = True', f'{S1}.Format.Line.ForeColor.RGB = 255', f'{S1}.Format.Fill.ForeColor.RGB = 255', '']
CASES += [(kind, f'{look}{": " if look else ""}{S1}.ClearFormats: {CLEAR_READS}: r = r & " " & {S1}.MarkerStyle & " " & {S1}.MarkerSize') for kind in ('line', 'lineMarkers', 'scatter') for look in LINES]
CASES += [(kind, f'{look}{": " if look else ""}{S1}.ClearFormats: {CLEAR_READS}') for kind in ('pie', 'area') for look in ('', f'{S1}.Format.Fill.ForeColor.RGB = 255', f'{S1}.Points(2).Format.Fill.ForeColor.RGB = 255', f'{S1}.Format.Line.ForeColor.RGB = 255')]
CASES += [('column', f'r = {S1}.ClearFormats'), ('column', f'{S1}.ClearFormats: {S1}.ClearFormats: {CLEAR_READS}'), ('column', f'{S1}.ClearFormats: {S1}.Format.Fill.ForeColor.RGB = 255: {CLEAR_READS}'), ('column', f'{S1}.ClearFormats: ch.ChartType = xlLine: {CLEAR_READS}')]
READS = ['ch.SeriesCollection.Count', f'{S1}.Name']


def _probe() -> str:
    """The cases go in procedures of a hundred, since one holding them all is too large for VBA."""
    lines: list[str] = []
    for first in range(0, len(CASES), 100):
        lines += [f'Private Sub Act{first // 100}(ByVal which As Long, ws As Worksheet, ch As Object, r As Variant)', '    Select Case which']
        for index, (kind, statement) in enumerate(CASES[first:first + 100], first):
            lines += [f'    Case {index}', '        ch.SetSourceData ws.Range("A1:C4"), xlColumns', f'        ch.ChartType = {TYPES[kind]}', f'        {statement}']
        lines += ['    End Select', 'End Sub', '']
    lines += ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
              '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String, v As String, r As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
              '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))', '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
              '    Set ch = ws.ChartObjects.Add(300, 100, 300, 200).Chart', '    On Error GoTo Failed', '    Select Case which \\ 100']
    for block in range((len(CASES) + 99) // 100):
        lines += [f'    Case {block}: Act{block} which, ws, ch, r']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
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
