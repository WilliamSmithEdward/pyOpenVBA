"""Series.Format: what a series' Fill and Line read on each kind of chart, what each change does, the errors, and the
chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData, sets the case's type and runs its statements with sr as the first
series and s2 as the second. It then reads READS, each as its value (TypeName for an object) or "E" and the error
number, and saves the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_series_format.py

writes tests/fixtures/series_format.json, which tests/test_excel_series_format.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/series_format.json'
FILL = 'sr.Format.Fill.ForeColor.RGB = 255'
LINE = 'sr.Format.Line.ForeColor.RGB = 255'
CASES = [
    (51, ''), (57, ''), (52, ''), (4, ''), (65, ''), (5, ''), (-4120, ''), (-4169, ''), (1, ''),
    (51, FILL), (57, FILL), (4, FILL), (65, FILL), (5, FILL), (-4120, FILL), (-4169, FILL), (1, FILL),
    (51, LINE), (4, LINE), (65, LINE), (5, LINE), (-4169, LINE), (1, LINE),
    (51, 'sr.Format.Fill.Visible = msoFalse'), (4, 'sr.Format.Fill.Visible = msoFalse'),
    (51, 'sr.Format.Fill.Visible = msoFalse: sr.Format.Fill.Visible = msoTrue'),
    (51, 'sr.Format.Fill.Transparency = 0.5'),
    (51, 'sr.Format.Fill.ForeColor.RGB = 255: sr.Format.Fill.Transparency = 0.5'),
    (51, 'sr.Format.Fill.ForeColor.RGB = 255: sr.Format.Fill.Visible = msoFalse'),
    (51, 'sr.Format.Fill.Solid'),
    (51, 'sr.Format.Line.Visible = msoFalse'), (4, 'sr.Format.Line.Visible = msoFalse'),
    (51, 'sr.Format.Line.Visible = msoFalse: sr.Format.Line.Visible = msoTrue'),
    (4, 'sr.Format.Line.Visible = msoFalse: sr.Format.Line.Visible = msoTrue'),
    (51, 'sr.Format.Line.Weight = 3'), (4, 'sr.Format.Line.Weight = 3'),
    (51, 'sr.Format.Line.DashStyle = msoLineDash'), (4, 'sr.Format.Line.DashStyle = msoLineDash'),
    (4, 'sr.Format.Line.DashStyle = msoLineRoundDot'),
    (51, 'sr.Format.Line.Transparency = 0.5'), (4, 'sr.Format.Line.Transparency = 0.5'),
    (51, 'sr.Format.Line.ForeColor.RGB = 255: sr.Format.Line.Weight = 3'),
    (4, 'sr.Format.Line.ForeColor.RGB = 255: sr.Format.Line.Weight = 3: sr.Format.Line.DashStyle = msoLineDash'),
    (51, 'sr.Format.Fill.ForeColor.RGB = 255: sr.Format.Line.ForeColor.RGB = 65280'),
    (51, 'sr.Format.Fill.ForeColor.RGB = 255: s2.Format.Fill.ForeColor.RGB = 65280'),
    (51, 'sr.Format.Fill.ForeColor.RGB = -1'), (51, 'sr.Format.Line.Weight = -1'),
    (51, 'sr.Format.Fill.Transparency = 2'),
    (51, 'sr.Format.Fill.ForeColor.RGB = 255: sr.ChartType = xlLine'),
    (4, 'sr.Format.Line.ForeColor.RGB = 255: sr.MarkerStyle = xlMarkerStyleCircle'),
    (65, 'sr.MarkerForegroundColor = 255'), (65, 'sr.MarkerBackgroundColor = 255'),
    (51, 'sr.Format.Fill.ForeColor.RGB = 255: sr.HasDataLabels = True'),
    (51, 'sr.Format.Fill.ForeColor.RGB = 255: sr.InvertIfNegative = True'),
]
READS = [
    'TypeName(sr.Format)', 'TypeName(sr.Format.Parent)', 'TypeName(sr.Format.Application)',
    'sr.Format.Fill.Visible', 'sr.Format.Fill.ForeColor.RGB', 'sr.Format.Fill.Transparency', 'sr.Format.Fill.Type',
    'sr.Format.Fill.BackColor.RGB', 'sr.Format.Fill.ForeColor.ObjectThemeColor', 'sr.Format.Fill.ForeColor.Type',
    'sr.Format.Line.Visible', 'sr.Format.Line.ForeColor.RGB', 'sr.Format.Line.Weight', 'sr.Format.Line.DashStyle',
    'sr.Format.Line.Style', 'sr.Format.Line.Transparency', 'sr.Format.Line.ForeColor.ObjectThemeColor',
    'sr.Format.Line.ForeColor.Type', 'TypeName(sr.Format.Fill)', 'TypeName(sr.Format.Fill.Parent)',
    'TypeName(sr.Format.Line)', 'TypeName(sr.Format.Line.Parent)', 's2.Format.Fill.ForeColor.RGB',
    's2.Format.Fill.ForeColor.ObjectThemeColor', 's2.Format.Line.ForeColor.RGB',
    's2.Format.Line.ForeColor.ObjectThemeColor',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, sr As Object, s2 As Object, s As String, '
             'v As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}', '        Set sr = ch.SeriesCollection(1)',
                  '        Set s2 = ch.SeriesCollection(2)']
        lines += [f'        {statements}'] if statements else []
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Tried = s',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered',
              'End Function', '', 'Public Function Probe(ByVal folder As String) As String',
              '    Dim i As Long, s As String', '    Application.DisplayAlerts = False',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i, folder)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=290.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [chart_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{kind}: {statements}', answer, part]
                        for (kind, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        series = re.findall(r'<c:ser>.*?</c:ser>', part, re.DOTALL)
        looks = [re.sub(r'<c:(cat|val|xVal|yVal|tx)>.*?</c:\1>|<c:(idx|order)[^>]*/>|<c:extLst>.*?</c:extLst>', '', one)
                 for one in series[:2]]
        print(case, '\n  =', answer.replace(',FormatFill', ''), '\n  1', looks[0] if looks else '', '\n  2',
              looks[1] if len(looks) > 1 else '')


if __name__ == '__main__':
    main()
