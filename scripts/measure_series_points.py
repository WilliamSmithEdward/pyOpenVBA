"""Series.Points and Point: what a point reads on each kind of chart, what each change does to it and to its series, the
errors, and the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and C,
gives them to the chart with SetSourceData, sets the case's type and runs its statements with sr as the first series
and pt as its second point. It then reads READS, each as its value (TypeName for an object) or "E" and the error
number, and saves the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_series_points.py

writes tests/fixtures/series_points.json, which tests/test_excel_series_points.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/series_points.json'
TYPES = [51, 57, 52, 4, 65, 5, -4120, -4169, 1]
EVERY = ['', 'pt.Format.Fill.ForeColor.RGB = 255', 'pt.Format.Line.ForeColor.RGB = 255', 'pt.Format.Line.Weight = 3',
         'pt.Format.Fill.Visible = msoFalse', 'pt.Format.Line.Visible = msoFalse', 'pt.HasDataLabel = True',
         'pt.InvertIfNegative = True', 'pt.Explosion = 20', 'pt.MarkerStyle = xlMarkerStyleCircle',
         'pt.MarkerSize = 9', 'pt.MarkerBackgroundColor = 255', 'pt.MarkerForegroundColor = 255']
MORE = [
    (51, 'sr.Points(1).Format.Fill.ForeColor.RGB = 65280: pt.Format.Fill.ForeColor.RGB = 255'),
    (51, 'sr.Points(3).Format.Fill.ForeColor.RGB = 65280: sr.Points(1).Format.Fill.ForeColor.RGB = 255'),
    (51, 'pt.Format.Fill.ForeColor.RGB = 255: pt.Format.Fill.Transparency = 0.5'),
    (51, 'pt.Format.Fill.Transparency = 0.5'),
    (51, 'pt.Format.Fill.Solid'),
    (51, 'pt.Format.Fill.ForeColor.RGB = 255: pt.Format.Line.ForeColor.RGB = 65280: pt.Format.Line.Weight = 2'),
    (51, 'sr.Format.Fill.ForeColor.RGB = 65280'),
    (51, 'sr.Format.Fill.ForeColor.RGB = 65280: pt.Format.Fill.ForeColor.RGB = 255'),
    (51, 'pt.Format.Fill.ForeColor.RGB = 255: sr.Format.Fill.ForeColor.RGB = 65280'),
    (51, 'pt.Format.Line.ForeColor.RGB = 255: sr.Format.Line.ForeColor.RGB = 65280'),
    (51, 'pt.Format.Fill.ForeColor.RGB = 255: sr.InvertIfNegative = True'),
    (51, 'pt.Format.Fill.ForeColor.RGB = 255: pt.InvertIfNegative = True'),
    (51, 'pt.Format.Fill.ForeColor.RGB = 255: pt.ClearFormats'),
    (51, 'pt.ClearFormats'),
    (51, 'pt.Delete'),
    (51, 'pt.Format.Fill.ForeColor.RGB = 255: ws.Range("B3").Value = 7'),
    (51, 'pt.Format.Fill.ForeColor.RGB = 255: ch.ChartType = xlLine'),
    (51, 'pt.Format.Fill.ForeColor.RGB = -1'),
    (51, 'pt.Format.Line.Weight = -1'),
    (51, 'pt.Format.Fill.Transparency = 2'),
    (51, 'pt.HasDataLabel = True: pt.HasDataLabel = False'),
    (51, 'sr.HasDataLabels = True: pt.HasDataLabel = False'),
    (51, 'sr.HasDataLabels = True'),
    (5, 'pt.Format.Fill.ForeColor.RGB = 255: pt.Explosion = 20'),
    (5, 'pt.Explosion = 20: sr.Explosion = 10'),
    (5, 'sr.Explosion = 10: pt.Explosion = 20'),
    (5, 'pt.Explosion = 0'),
    (5, 'pt.Explosion = -1'),
    (4, 'pt.Format.Line.ForeColor.RGB = 255: pt.Format.Line.Weight = 2'),
    (65, 'pt.MarkerStyle = xlMarkerStyleSquare: pt.MarkerSize = 9: pt.MarkerBackgroundColor = 255'),
    (65, 'pt.MarkerSize = 1'),
]
CASES = [(kind, statements) for kind in TYPES for statements in EVERY] + MORE
READS = [
    'TypeName(sr.Points)', 'sr.Points.Count', 'TypeName(sr.Points.Parent)', 'TypeName(sr.Points.Application)',
    'TypeName(sr.Points(1))', 'TypeName(sr.Points.Item(2))', 'TypeName(sr.Points(0))', 'TypeName(sr.Points(4))',
    'TypeName(sr.Points("a"))',
    'pt.Name', 'TypeName(pt.Parent)', 'TypeName(pt.Application)', 'pt.Explosion', 'pt.HasDataLabel',
    'pt.InvertIfNegative', 'pt.MarkerStyle', 'pt.MarkerSize', 'pt.MarkerBackgroundColor', 'pt.MarkerForegroundColor',
    'pt.Shadow', 'pt.Has3DEffect', 'pt.SecondaryPlot', 'pt.IsTotal', 'TypeName(pt.DataLabel)',
    'TypeName(pt.Format)', 'TypeName(pt.Format.Parent)', 'pt.Format.Fill.Visible', 'pt.Format.Fill.ForeColor.RGB',
    'pt.Format.Fill.Transparency', 'pt.Format.Fill.Type', 'pt.Format.Fill.BackColor.RGB',
    'pt.Format.Fill.ForeColor.ObjectThemeColor', 'pt.Format.Fill.ForeColor.Type', 'pt.Format.Line.Visible',
    'pt.Format.Line.ForeColor.RGB', 'pt.Format.Line.Weight', 'pt.Format.Line.DashStyle', 'pt.Format.Line.Style',
    'pt.Format.Line.Transparency', 'pt.Format.Line.ForeColor.ObjectThemeColor', 'pt.Format.Line.ForeColor.Type',
    'sr.Format.Fill.ForeColor.RGB', 'sr.Format.Line.ForeColor.RGB', 'sr.HasDataLabels', 'sr.Explosion',
    'sr.Points(1).Format.Fill.ForeColor.RGB', 'sr.Points(1).HasDataLabel',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, sr As Object, pt As Object, s As String',
             '    Dim v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}', '        Set sr = ch.SeriesCollection(1)',
                  '        Set pt = sr.Points(2)']
        lines += [f'        {statements}'] if statements else []
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Tried = s',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered',
              'End Function', '', 'Public Function Probe(ByVal folder As String, ByVal first As Long, '
              'ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 40):
            last = min(first + 39, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
        parts = [chart_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{kind}: {statements}', answer, part]
                        for (kind, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    strip = re.compile(r'<c:(cat|val|xVal|yVal|tx)>.*?</c:\1>|<c:order[^>]*/>|<c:extLst>.*?</c:extLst>', re.DOTALL)
    for case, answer, part in record['cases']:
        series = re.findall(r'<c:ser>.*?</c:ser>', part, re.DOTALL)
        print(case, '\n  =', answer, '\n  1', strip.sub('', series[0]) if series else '')


if __name__ == '__main__':
    main()
