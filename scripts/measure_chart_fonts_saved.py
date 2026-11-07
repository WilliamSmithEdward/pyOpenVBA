"""Fonts of charts Excel saved and opened again: what each element reads, wherever the file holds it.

    python scripts/measure_chart_fonts_saved.py

writes tests/fixtures/chart_fonts_saved.xlsx and tests/fixtures/chart_fonts_saved.json with what Excel reads from it, which
tests/test_excel_chart_font_gaps.py replays.
"""
from __future__ import annotations

import json
import shutil
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / 'tests/fixtures/chart_fonts_saved.json'
BOOK = ROOT / 'tests/fixtures/chart_fonts_saved.xlsx'
READS = ('Name', 'Size', 'Bold', 'Italic', 'Underline', 'Strikethrough', 'Color', 'ColorIndex', 'FontStyle', 'Subscript', 'Superscript', 'Background', 'ThemeColor', 'TintAndShade', 'ThemeFont')
ELEMENTS = ('ch.ChartTitle.Font', 'ch.Legend.Font', 'ch.Axes(1).TickLabels.Font', 'ch.Axes(2).TickLabels.Font', 'ch.SeriesCollection(1).DataLabels.Font', 'ch.ChartArea.Font', 'ch.Axes(2).AxisTitle.Font', 'ch.Axes(1).AxisTitle.Font', 'ch.Legend.LegendEntries(1).Font')
BASE = '''    Set ch = ws.ChartObjects.Add({0}, {1}, 360, 216).Chart
    ch.SetSourceData ws.Range("A1:D4"), xlColumns
    ch.ChartType = xlColumnClustered
    ch.HasTitle = True
    ch.HasLegend = True
    ch.Axes(2).HasTitle = True
    ch.Axes(1).HasTitle = True
    ch.SeriesCollection(1).HasDataLabels = True
'''
MAKE = 'Public Function Make(ByVal path As String) As String\n    Dim wb As Workbook, ws As Worksheet, ch As Object\n    Application.DisplayAlerts = False\n    Set wb = Workbooks.Add(xlWBATWorksheet)\n    Set ws = wb.Worksheets(1)\n    ws.Range("A1:D1").Value = Array("", "Rev", "Cost", "Rate")\n    ws.Range("A2:D2").Value = Array("x", 10, 4, 1)\n    ws.Range("A3:D3").Value = Array("y", 20, 6, 2)\n    ws.Range("A4:D4").Value = Array("z", 30, 8, 3)\n'
MAKE += "    ' 1: nothing set.\n" + BASE.format(48, 48)
MAKE += "    ' 2: each element its own.\n" + BASE.format(432, 48) + '''    ch.ChartTitle.Text = "Mine"
    With ch.ChartTitle.Font: .Size = 20: .Bold = False: .Italic = True: .Color = 255: .Name = "Arial": .Underline = xlUnderlineStyleDouble: End With
    With ch.Legend.Font: .Size = 8: .Bold = True: .Strikethrough = True: .ColorIndex = 5: End With
    With ch.Axes(1).TickLabels.Font: .Size = 7.5: .Underline = xlUnderlineStyleSingle: End With
    With ch.Axes(2).TickLabels.Font: .Name = "Courier New": .Subscript = True: End With
    With ch.SeriesCollection(1).DataLabels.Font: .Size = 12: .Color = 65280: End With
    With ch.Axes(2).AxisTitle.Font: .Size = 15: .Bold = False: .ColorIndex = xlColorIndexNone: End With
'''
MAKE += "    ' 3: the chart area, and one element over it.\n" + BASE.format(48, 288) + '''    With ch.ChartArea.Font: .Size = 14: .Bold = True: .Italic = True: .Color = 255: .Name = "Arial": End With
    ch.Legend.Font.Size = 9
    ch.Legend.Font.Bold = False
'''
MAKE += "    ' 4: as a recorded macro makes it.\n" + '''    Set ch = ws.Shapes.AddChart2(201, xlColumnClustered, 432, 288, 360, 216).Chart
    ch.SetSourceData ws.Range("A1:D4"), xlColumns
    ch.HasTitle = True
    ch.ChartTitle.Text = "Mine"
    ch.HasLegend = True
    ch.Axes(2).HasTitle = True
    ch.Axes(1).HasTitle = True
    ch.SeriesCollection(1).HasDataLabels = True
'''
MAKE += "    ' 5: a recorded macro's, with fonts set.\n" + '''    Set ch = ws.Shapes.AddChart2(201, xlColumnClustered, 816, 48, 360, 216).Chart
    ch.SetSourceData ws.Range("A1:D4"), xlColumns
    ch.HasTitle = True
    ch.ChartTitle.Text = "Mine"
    ch.HasLegend = True
    ch.Axes(2).HasTitle = True
    ch.Axes(1).HasTitle = True
    ch.SeriesCollection(1).HasDataLabels = True
    ch.ChartTitle.Font.Size = 22
    ch.Legend.Font.Bold = True
    ch.Axes(2).TickLabels.Font.Color = 255
    ch.ChartArea.Font.Italic = True
'''
MAKE += '    wb.SaveAs path, xlOpenXMLWorkbook\n    wb.Close False\n    Application.DisplayAlerts = True\n    Make = "ok"\nEnd Function\n'


def reader() -> str:
    lines = ['Private Function Say(o As Object) As String', '    Dim s As String, v As Variant', '    On Error Resume Next']
    for name in READS:
        lines += ['    v = "?": Err.Clear', f'    v = o.{name}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & v & ","']
    lines += ['    Err.Clear', '    Say = s', 'End Function', '', 'Public Function Reads(ByVal path As String) As String', '    Dim wb As Workbook, ch As Object, s As String, o As Object, i As Long', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Open(path)', '    For i = 1 To wb.Worksheets(1).ChartObjects.Count', '        Set ch = wb.Worksheets(1).ChartObjects(i).Chart', '        s = s & "#" & i & "~"', '        On Error Resume Next']
    for element in ELEMENTS:
        lines += ['        Set o = Nothing: Err.Clear', f'        Set o = {element}', '        If Err.Number Then s = s & "E" & Err.Number & "~" Else s = s & Say(o) & "~"']
    lines += ['        Err.Clear', '        On Error GoTo 0', '    Next', '    wb.Close False', '    Application.DisplayAlerts = True', '    Reads = s', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        made = Path(folder) / 'made.xlsx'
        excel.new_document()
        done = excel.run_vba(MAKE, 'Make', (str(made),), timeout=180.0, line_numbers=False)
        assert done.ok, done
        answers = excel.run_vba(reader(), 'Reads', (str(made),), timeout=180.0, line_numbers=False)
        assert answers.ok, answers
        shutil.copy(made, BOOK)
    strip_save_path(BOOK)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': reader(), 'answers': str(answers.value)}, indent=1) + '\n', encoding='utf-8')
    print(str(answers.value).replace('#', '\n#').replace('~', '\n  '))


if __name__ == '__main__':
    main()
