"""Combo charts in a workbook Excel saved and opened again: what they read, and that the model reads the same.

    python scripts/measure_combo_charts_saved.py

writes tests/fixtures/combo_chart_saved.xlsx, four charts on one sheet, and tests/fixtures/combo_charts_saved.json with what Excel
reads from them once opened, which tests/test_excel_combo_charts.py replays.
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

OUT = ROOT / 'tests/fixtures/combo_charts_saved.json'
BOOK = ROOT / 'tests/fixtures/combo_chart_saved.xlsx'
MAKE = r'''Public Function Make(ByVal path As String) As String
    Dim wb As Workbook, ws As Worksheet, ch As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1:D1").Value = Array("", "Rev", "Cost", "Rate")
    ws.Range("A2:D2").Value = Array("x", 10, 4, 1)
    ws.Range("A3:D3").Value = Array("y", 20, 6, 2)
    ws.Range("A4:D4").Value = Array("z", 30, 8, 3)
    ' A column and a line on secondary axes, titled.
    Set ch = ws.ChartObjects.Add(48, 48, 360, 216).Chart
    ch.SetSourceData ws.Range("A1:D4"), xlColumns
    ch.ChartType = xlColumnClustered
    ch.SeriesCollection(3).ChartType = xlLineMarkers
    ch.SeriesCollection(3).AxisGroup = xlSecondary
    ch.Axes(xlValue, xlSecondary).HasTitle = True
    ch.Axes(xlValue, xlSecondary).AxisTitle.Text = "Rate"
    ch.Axes(xlValue, xlSecondary).MaximumScale = 6
    ch.ChartGroups(1).GapWidth = 80
    ' Areas, columns and lines on one set of axes.
    Set ch = ws.ChartObjects.Add(48, 288, 360, 216).Chart
    ch.SetSourceData ws.Range("A1:D4"), xlColumns
    ch.ChartType = xlColumnStacked
    ch.SeriesCollection(1).ChartType = xlArea
    ch.SeriesCollection(3).ChartType = xlLine
    ' One kind on both axes, the secondary category axis showing.
    Set ch = ws.ChartObjects.Add(432, 48, 360, 216).Chart
    ch.SetSourceData ws.Range("A1:D4"), xlColumns
    ch.ChartType = xlColumnClustered
    ch.SeriesCollection(2).AxisGroup = xlSecondary
    ch.HasAxis(xlCategory, xlSecondary) = True
    ch.ChartGroups(2).GapWidth = 40
    ' As a recorded macro makes one.
    Set ch = ws.Shapes.AddChart2(201, xlColumnClustered, 432, 288, 360, 216).Chart
    ch.SetSourceData ws.Range("A1:D4"), xlColumns
    ch.SeriesCollection(2).ChartType = xlLine
    ch.SeriesCollection(2).AxisGroup = xlSecondary
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Make = "ok"
End Function
'''
EACH = ('ch.ChartType', 'ch.ChartGroups.Count', 'ch.Axes.Count', 'ch.HasAxis(1, 1)', 'ch.HasAxis(2, 1)', 'ch.HasAxis(1, 2)', 'ch.HasAxis(2, 2)', 'ch.SeriesCollection.Count', 'ch.ColumnGroups.Count', 'ch.LineGroups.Count', 'ch.AreaGroups.Count', 'ch.ChartGroups(1).AxisGroup', 'ch.ChartGroups(1).GapWidth', 'ch.ChartGroups(1).SeriesCollection.Count',
        'ch.ChartGroups(2).AxisGroup', 'ch.ChartGroups(2).GapWidth', 'ch.ChartGroups(2).SeriesCollection.Count', 'ch.ChartGroups(2).SeriesCollection(1).Name', 'ch.Axes(2, 2).AxisGroup', 'ch.Axes(2, 2).HasTitle', 'ch.Axes(2, 2).AxisTitle.Text', 'ch.Axes(2, 2).MaximumScaleIsAuto', 'ch.Axes(2, 2).Crosses', 'ch.Axes(2, 2).HasMajorGridlines', 'ch.Axes(2, 2).TickLabelPosition', 'ch.Axes(1, 2).Crosses', 'ch.Axes(1, 2).AxisGroup', 'ch.Axes(1, 1).AxisGroup', 'ch.Axes(2, 1).AxisGroup', 'ch.Axes(2, 1).Crosses')
SERIES = ('.Name', '.ChartType', '.AxisGroup', '.PlotOrder', '.Type', '.MarkerStyle', '.Formula')
EDITS = ('ch.SeriesCollection(1).ChartType = xlLine', 'ch.SeriesCollection(1).AxisGroup = xlSecondary', 'ch.Axes(2, 2).HasTitle = False', 'ch.ChartGroups(1).GapWidth = 33')


def reader() -> str:
    lines = ['Public Function Reads(ByVal path As String) As String', '    Dim wb As Workbook, ch As Object, s As String, v As Variant, i As Long, j As Long', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Open(path)', '    On Error Resume Next', '    For i = 1 To wb.Worksheets(1).ChartObjects.Count', '        Set ch = wb.Worksheets(1).ChartObjects(i).Chart', '        s = s & "#" & i & "|"']
    for expression in EACH:
        lines += ['        v = "E": Err.Clear', f'        v = {expression}', '        If Err.Number Then v = "E" & Err.Number', '        s = s & v & "|"']
    lines += ['        For j = 1 To ch.SeriesCollection.Count']
    for name in SERIES:
        lines += ['            v = "E": Err.Clear', f'            v = ch.SeriesCollection(j){name}', '            If Err.Number Then v = "E" & Err.Number', '            s = s & v & "|"']
    lines += ['        Next', '    Next', '    Err.Clear', '    s = s & "#scale|" & wb.Worksheets(1).ChartObjects(1).Chart.Axes(2, 2).MaximumScale', '    On Error GoTo 0', '    wb.Close False', '    Application.DisplayAlerts = True', '    Reads = s', 'End Function']
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
    print(str(answers.value).replace('#', '\n#'))


if __name__ == '__main__':
    main()
