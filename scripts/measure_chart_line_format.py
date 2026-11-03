"""The Format of a chart's trendlines and error bars: what each member reads, and the chart part Excel saves.

Each case makes a workbook with categories in A2:A4 and two series in B and C, a column chart of them, and on the first
series a trendline or error bars. It runs the case's statements, which set r to what they want recorded, and saves the
workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_line_format.py

writes tests/fixtures/chart_line_format.json, which tests/test_excel_chart_line_format.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_line_format.json'
S1 = 'ch.SeriesCollection(1)'
OWNERS = {'trendline': (f'{S1}.Trendlines.Add', f'{S1}.Trendlines(1)'), 'error bars': (f'{S1}.ErrorBar xlY, xlBoth, xlFixedValue, 1', f'{S1}.ErrorBars')}
LINE_READS = ['Visible', 'ForeColor.RGB', 'ForeColor.Type', 'ForeColor.ObjectThemeColor', 'Weight', 'DashStyle', 'Style', 'Transparency', 'BeginArrowheadStyle', 'EndArrowheadStyle']
FILL_READS = ['Visible', 'ForeColor.RGB', 'Type', 'Transparency']
READ_ALL = ' & " " & '.join(f'L.{read}' for read in LINE_READS[:2] + LINE_READS[4:8])
STATEMENTS = [
    *[f'r = TypeName(F.{name})' for name in ('Parent', 'Line', 'Fill', 'Shadow', 'Glow', 'SoftEdge', 'ThreeD', 'TextFrame2', 'PictureEffects')],
    'r = TypeName(F)', *[f'r = L.{read}' for read in LINE_READS], *[f'r = F.Fill.{read}' for read in FILL_READS],
    # Setting the line, one thing at a time and together.
    *[f'L.ForeColor.RGB = {colour}: r = {READ_ALL}' for colour in (255, 0, 16777215)],
    *[f'L.Weight = {weight}: r = {READ_ALL}' for weight in (0, 0.25, 0.75, 1, 3, 6.5, 1584, 1585, -1)],
    *[f'L.DashStyle = {dash}: r = {READ_ALL}' for dash in range(0, 14)],
    *[f'L.Style = {style}: r = {READ_ALL}' for style in (1, 2, 3, 4, 5)],
    *[f'L.Transparency = {amount}: r = {READ_ALL}' for amount in (0, 0.5, 1, 2)],
    *[f'L.Visible = {value}: r = {READ_ALL}' for value in ('msoFalse', 'msoTrue', 'msoCTrue', 2)],
    f'L.ForeColor.RGB = 255: L.Visible = msoFalse: r = {READ_ALL}', f'L.Visible = msoFalse: L.ForeColor.RGB = 255: r = {READ_ALL}', f'L.Visible = msoFalse: L.Weight = 3: r = {READ_ALL}', f'L.Visible = msoFalse: L.Visible = msoTrue: r = {READ_ALL}', f'L.ForeColor.RGB = 255: L.Visible = msoFalse: L.Visible = msoTrue: r = {READ_ALL}',
    f'L.ForeColor.RGB = 255: L.Weight = 3: L.DashStyle = msoLineDash: r = {READ_ALL}', f'L.Weight = 3: L.ForeColor.RGB = 255: r = {READ_ALL}', f'L.DashStyle = msoLineDash: L.ForeColor.RGB = 255: L.Weight = 2: L.Transparency = 0.5: r = {READ_ALL}',
    f'L.ForeColor.ObjectThemeColor = msoThemeColorAccent2: r = {READ_ALL}', 'L.ForeColor.ObjectThemeColor = msoThemeColorAccent2: r = L.ForeColor.Type & " " & L.ForeColor.ObjectThemeColor', 'L.ForeColor.SchemeColor = 3: r = L.ForeColor.RGB', 'L.BeginArrowheadStyle = msoArrowheadTriangle: r = L.BeginArrowheadStyle', 'L.EndArrowheadStyle = msoArrowheadTriangle: r = L.EndArrowheadStyle',
    # What Border and Format see of each other.
    'L.ForeColor.RGB = 255: r = B.Color & " " & B.ColorIndex & " " & B.Weight & " " & B.LineStyle', 'L.Weight = 3: r = B.Color & " " & B.ColorIndex & " " & B.Weight & " " & B.LineStyle', 'L.DashStyle = msoLineDash: r = B.Color & " " & B.ColorIndex & " " & B.Weight & " " & B.LineStyle', 'L.Visible = msoFalse: r = B.Color & " " & B.ColorIndex & " " & B.Weight & " " & B.LineStyle',
    f'B.Color = 255: r = {READ_ALL}', f'B.Weight = xlThick: r = {READ_ALL}', f'B.LineStyle = xlDash: r = {READ_ALL}', f'B.LineStyle = xlDot: r = {READ_ALL}', f'B.LineStyle = xlDashDot: r = {READ_ALL}', f'B.LineStyle = xlDashDotDot: r = {READ_ALL}', f'B.ColorIndex = xlNone: r = {READ_ALL}', f'B.LineStyle = xlNone: r = {READ_ALL}', f'L.ForeColor.RGB = 255: B.LineStyle = xlAutomatic: r = {READ_ALL}', f'L.Weight = 3: B.Weight = xlThin: r = {READ_ALL}',
    # The fill and the rest, which a line has no use for.
    'F.Fill.ForeColor.RGB = 255: r = F.Fill.ForeColor.RGB', 'F.Fill.Visible = msoFalse: r = F.Fill.Visible', 'F.Fill.Visible = msoTrue: r = F.Fill.Visible', 'F.Shadow.Visible = msoTrue: r = F.Shadow.Visible', 'r = F.Shadow.Visible', 'r = F.Glow.Radius', 'r = F.SoftEdge.Type', 'r = F.Parent.Name',
    'L.ForeColor.RGB = 255: ch.ClearToMatchStyle: r = L.ForeColor.RGB', 'L.ForeColor.RGB = 255: ch.ChartType = xlLine: r = L.ForeColor.RGB', 'L.ForeColor.RGB = 255: ch.ChartStyle = 5: r = L.ForeColor.RGB',
]
CASES = [(owner, statement) for owner in OWNERS for statement in STATEMENTS]
CASES += [('error bars', 'ch.ChartType = xlXYScatter: L.ForeColor.RGB = 255: r = L.ForeColor.RGB'), ('error bars', 'ch.ChartType = xlXYScatter: B.Color = 255: r = B.Color'), ('trendline', f'{S1}.Trendlines.Add: L.ForeColor.RGB = 255: r = {S1}.Trendlines(2).Format.Line.ForeColor.RGB')]


def _probe() -> str:
    lines: list[str] = []
    for first in range(0, len(CASES), 100):
        lines += [f'Private Sub Act{first // 100}(ByVal which As Long, ws As Worksheet, ch As Object, r As Variant)', '    Dim F As Object, L As Object, B As Object', '    Select Case which']
        for index, (owner, statement) in enumerate(CASES[first:first + 100], first):
            make, target = OWNERS[owner]
            lines += [f'    Case {index}', f'        {make}', f'        Set F = {target}.Format: Set L = F.Line: Set B = {target}.Border', f'        {statement}']
        lines += ['    End Select', 'End Sub', '']
    lines += ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
              '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String, v As String, r As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
              '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))', '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
              '    Set ch = ws.ChartObjects.Add(300, 100, 300, 200).Chart', '    ch.SetSourceData ws.Range("A1:C4"), xlColumns', '    On Error GoTo Failed', '    Select Case which \\ 100']
    for block in range((len(CASES) + 99) // 100):
        lines += [f'    Case {block}: Act{block} which, ws, ch, r']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v',
              '    On Error GoTo 0', '    Tried = s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function',
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
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[f'{owner}: {statement}', answer, part] for (owner, statement), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
