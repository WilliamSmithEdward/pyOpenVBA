"""A sparkline group's Points and Axes, its DisplayBlanksAs, DisplayHidden and PlotBy, and the FormatColor of its
series and points: what each reads on a new group, what a set does, and the group a save then writes.

Excel stopped answering (RPC 0x800706BE) when a macro put Array(1, 3, -2, 4) in one row with Range.Value and then
added a sparkline over it, so the probe fills its cells through Transpose, as scripts/measure_sparklines.py does.

    python scripts/measure_sparkline_points.py

writes tests/fixtures/sparkline_points.json, which tests/test_excel_sparkline_points.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/sparkline_points.json'
#: Each case: a statement on a new line group over A1:D1 at E1, what is read, and whether the file is saved.
CASES = [
    ('', 'g.Points.Markers.Visible & " " & g.Points.Highpoint.Visible & " " & g.Points.Lowpoint.Visible & " " & '
         'g.Points.Firstpoint.Visible & " " & g.Points.Lastpoint.Visible & " " & g.Points.Negative.Visible', False),
    ('', 'g.Points.Markers.Color.Color & " " & g.Points.Highpoint.Color.Color & " " & g.Points.Negative.Color.Color',
     False),
    ('', 'TypeName(g.Points) & " " & TypeName(g.Points.Markers) & " " & TypeName(g.Points.Markers.Color) & " " & '
         'TypeName(g.Axes) & " " & TypeName(g.Axes.Horizontal) & " " & TypeName(g.Axes.Vertical) & " " & '
         'TypeName(g.Axes.Horizontal.Axis) & " " & TypeName(g.SeriesColor)', False),
    ('', 'g.Axes.Horizontal.IsDateAxis', False),
    ('', 'g.Axes.Horizontal.RightToLeft', False),
    ('', 'g.Axes.Horizontal.Axis.Visible & " " & g.Axes.Horizontal.Axis.Color.Color', False),
    ('', 'g.Axes.Vertical.MinScaleType & " " & g.Axes.Vertical.MaxScaleType', False),
    ('', 'g.Axes.Vertical.CustomMinScaleValue & " " & g.Axes.Vertical.CustomMaxScaleValue', False),
    ('', 'g.DisplayBlanksAs & " " & g.DisplayHidden & " " & g.PlotBy', False),
    ('', 'g.Points.Negative.Visible & " " & TypeName(g.Points.Parent)', False),
    ('g.Points.Markers.Visible = True', 'g.Points.Markers.Visible', True),
    ('g.Points.Highpoint.Visible = True', 'g.Points.Highpoint.Visible', True),
    ('g.Points.Lowpoint.Visible = True', 'g.Points.Lowpoint.Visible', True),
    ('g.Points.Firstpoint.Visible = True', 'g.Points.Firstpoint.Visible', True),
    ('g.Points.Lastpoint.Visible = True', 'g.Points.Lastpoint.Visible', True),
    ('g.Points.Negative.Visible = True', 'g.Points.Negative.Visible', True),
    ('g.Points.Markers.Color.Color = vbBlue', 'g.Points.Markers.Color.Color', True),
    ('g.Axes.Horizontal.Axis.Visible = True', 'g.Axes.Horizontal.Axis.Visible', True),
    ('g.Axes.Horizontal.Axis.Color.Color = vbRed', 'g.Axes.Horizontal.Axis.Color.Color', True),
    ('g.Axes.Vertical.MinScaleType = xlSparkScaleGroup', 'g.Axes.Vertical.MinScaleType', True),
    ('g.Axes.Vertical.MaxScaleType = xlSparkScaleCustom', 'g.Axes.Vertical.MaxScaleType', True),
    ('g.Axes.Vertical.MaxScaleType = xlSparkScaleCustom: g.Axes.Vertical.CustomMaxScaleValue = 7',
     'g.Axes.Vertical.CustomMaxScaleValue', True),
    ('g.Axes.Vertical.CustomMaxScaleValue = 7',
     'g.Axes.Vertical.CustomMaxScaleValue & " " & g.Axes.Vertical.MaxScaleType', True),
    ('g.Axes.Vertical.CustomMinScaleValue = -2.5',
     'g.Axes.Vertical.CustomMinScaleValue & " " & g.Axes.Vertical.MinScaleType', True),
    ('g.Axes.Vertical.CustomMaxScaleValue = 1E+300', 'g.Axes.Vertical.CustomMaxScaleValue', True),
    ('g.Axes.Vertical.MinScaleType = 9', 'g.Axes.Vertical.MinScaleType', False),
    ('g.DisplayBlanksAs = xlZero', 'g.DisplayBlanksAs', True),
    ('g.DisplayBlanksAs = xlInterpolated', 'g.DisplayBlanksAs', True),
    ('g.DisplayBlanksAs = xlNotPlotted', 'g.DisplayBlanksAs', True),
    ('g.DisplayBlanksAs = 7', 'g.DisplayBlanksAs', False),
    ('g.DisplayHidden = True', 'g.DisplayHidden', True),
    ('g.PlotBy = 1', 'g.PlotBy', False),
    ('g.Points.Markers.Visible = True: g.Points.Markers.Visible = False', 'g.Points.Markers.Visible', True),
    ('g.Points.Markers.Visible = "x"', 'g.Points.Markers.Visible', False),
    ('g.SeriesColor.ThemeColor = 5', 'g.SeriesColor.Color & " " & g.SeriesColor.ThemeColor', True),
    ('g.SeriesColor.ThemeColor = 5: g.SeriesColor.TintAndShade = 0.4',
     'g.SeriesColor.TintAndShade & " " & g.SeriesColor.Color', True),
    ('g.SeriesColor.TintAndShade = 0.5', 'g.SeriesColor.TintAndShade & " " & g.SeriesColor.Color', True),
    ('g.SeriesColor.ThemeColor = 5: g.SeriesColor.TintAndShade = -0.25',
     'g.SeriesColor.TintAndShade & " " & g.SeriesColor.Color', True),
    ('g.SeriesColor.ThemeColor = 5: g.SeriesColor.Color = vbRed',
     'g.SeriesColor.ThemeColor & " " & g.SeriesColor.Color', True),
    ('g.SeriesColor.ThemeColor = 0', 'g.SeriesColor.ThemeColor', False),
    ('g.SeriesColor.ThemeColor = 13', 'g.SeriesColor.ThemeColor', False),
    ('g.SeriesColor.TintAndShade = 2', 'g.SeriesColor.TintAndShade', False),
    ('g.Points.Negative.Visible = True: g.Type = xlSparkColumn', 'g.Type & " " & g.Points.Negative.Visible', True),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, g As Object', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    ws.Range("A1:D4").Value = Application.Transpose(Application.Transpose('
             'Array(Array(1, 3, -2, 4), Array(-1, 2, -3, 4), Array(5, 4, 3, 2), Array(2, 2, 2, 2))))',
             '    On Error GoTo Failed', '    Set g = ws.Range("E1").SparklineGroups.Add(xlSparkLine, "A1:D1")',
             '    Select Case which']
    for index, (statement, read, save) in enumerate(CASES):
        lines.append(f'    Case {index}')
        if statement:
            lines.append(f'        {statement}')
        lines.append(f'        Tried = {read}')
        if save:
            lines.append(f'        wb.SaveAs folder & "\\c{index}.xlsx", xlOpenXMLWorkbook')
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim s As String, i As Long',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        groups = {}
        for index, (_, _, save) in enumerate(CASES):
            if save:
                with zipfile.ZipFile(Path(folder) / f'c{index}.xlsx') as package:
                    sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
                found = re.search(r'<x14:sparklineGroup\b.*?</x14:sparklineGroup>', sheet, re.DOTALL)
                groups[str(index)] = found.group(0) if found else ''
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statement, read, save, answer]
                        for (statement, read, save), answer in zip(CASES, answers, strict=True)],
              'groups': groups}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for index, (statement, read, _, answer) in enumerate(record['cases']):
        print(index, (statement or read)[:60], '->', answer)


if __name__ == '__main__':
    main()
