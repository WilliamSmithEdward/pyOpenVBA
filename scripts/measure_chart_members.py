"""Members of a chart macros reach for: Delete, Refresh, ChartObjects, Shapes, Evaluate, Location, Resize and the rest, on an
embedded chart and on a chart sheet. What each answers or raises, what it leaves, and the sheets and charts afterwards.

Each case makes a workbook with categories in A2:A4 and two series in B and C on Data, a chart of them on Data, and a sheet
Other. It runs the case's statements, each of which sets the variable r to what it wants recorded, or raises. The case's
answer is "ok ", or "Err " and the error's number, then r, then READS, each as its value or "E" and an error number.

    python scripts/measure_chart_members.py

writes tests/fixtures/chart_members.json, which tests/test_excel_chart_members.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_members.json'
CASES = [
    # Taking it away.
    'ch.Delete', 'r = ch.Delete', 'ch.Delete: r = ch.Name', 'ch.Parent.Delete: r = ch.Name', 'co.Delete: r = d.ChartObjects.Count', 'ch.Delete: r = d.ChartObjects.Count', 'ch.Delete: r = co.Name', 'd.ChartObjects(1).Chart.Delete: r = d.Shapes.Count',
    # Drawing it again.
    'ch.Refresh', 'r = ch.Refresh', 'd.Range("B2").Value = 9: ch.Refresh: r = ch.SeriesCollection(1).Values()(1)', 'ch.Calculate',
    # What it holds.
    'r = TypeName(ch.ChartObjects)', 'r = ch.ChartObjects.Count', 'r = ch.ChartObjects().Count', 'r = TypeName(ch.ChartObjects(1))', 'r = TypeName(ch.Shapes)', 'r = ch.Shapes.Count', 'ch.Shapes.AddShape 1, 1, 1, 50, 50: r = ch.Shapes.Count', 'ch.Shapes.AddShape 1, 1, 1, 50, 50: r = ch.Shapes(1).Name', 'ch.Shapes.AddTextbox 1, 1, 1, 50, 50: r = ch.Shapes(1).Type', 'r = TypeName(ch.OLEObjects)', 'r = ch.OLEObjects.Count', 'r = TypeName(ch.Hyperlinks)', 'r = ch.Hyperlinks.Count',
    # What it works out.
    'r = TypeName(ch.Evaluate("Data!A1"))', 'r = ch.Evaluate("Data!A1").Value', 'r = ch.Evaluate("1+2")', 'r = ch.Evaluate("A1")', 'r = TypeName(ch.Evaluate("A1"))', 'r = ch.Evaluate("Series1")', 'r = TypeName(ch.Evaluate("Legend"))', 'r = ch.Evaluate("SUM(Data!B2:B4)")', 'r = [Data!B2]',
    # Where it lives.
    'r = TypeName(ch.Location(xlLocationAutomatic))', 'r = TypeName(ch.Location(xlLocationAsObject, "Other")): r = r & d.ChartObjects.Count & o.ChartObjects.Count', 'r = TypeName(ch.Location(xlLocationAsNewSheet, "Shown")): r = r & d.ChartObjects.Count & wb.Charts.Count & wb.Sheets.Count', 'r = TypeName(ch.Location(xlLocationAsNewSheet)): r = r & wb.Charts(1).Name', 'ch.Location xlLocationAsObject, "Nowhere"', 'ch.Location xlLocationAsObject, "Data": r = d.ChartObjects.Count', 'ch.Location 9', 'ch.Location xlLocationAsObject',
    # Moving and copying it.
    'ch.Move', 'ch.Move After:=o', 'ch.Copy', 'ch.Copy After:=o', 'r = wb.Charts.Add.Name', 'wb.Charts.Add: r = wb.Charts(1).Parent.Name', 'ch.Resize', 'ch.Resize 10, 10',
    # Its surfaces.
    'r = TypeName(ch.Walls)', 'r = TypeName(ch.Floor)', 'r = TypeName(ch.BackWall)', 'r = TypeName(ch.SideWall)', 'r = ch.Walls.Name', 'r = ch.Floor.Name', 'ch.ChartType = xl3DColumnClustered: r = TypeName(ch.Walls)', 'ch.ChartType = xl3DColumnClustered: r = ch.Floor.Name', 'ch.ChartType = xl3DColumnClustered: r = ch.BackWall.Name', 'ch.ChartType = xl3DColumnClustered: r = ch.SideWall.Name', 'ch.ChartType = xl3DColumnClustered: r = ch.Walls.Name',
    # Axis and series.
    'r = TypeName(ch.Axes(xlValue).Border)', 'r = ch.Axes(xlValue).Border.LineStyle', 'r = ch.Axes(xlValue).Border.Color', 'r = ch.Axes(xlValue).Border.Weight', 'ch.Axes(xlValue).Border.Color = 255: r = ch.Axes(xlValue).Border.Color', 'ch.Axes(xlValue).Border.Weight = xlThick: r = ch.Axes(xlValue).Border.Weight', 'ch.Axes(xlValue).Border.LineStyle = xlNone: r = ch.Axes(xlValue).Border.LineStyle', 'r = ch.Axes(xlCategory).MinorUnitScale',
    'r = TypeName(ch.SeriesCollection(1).LeaderLines)', 'ch.SeriesCollection(1).ClearFormats', 'ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = 255: ch.SeriesCollection(1).ClearFormats: r = ch.SeriesCollection(1).Format.Fill.ForeColor.RGB', 'r = ch.ClearToMatchColorStyle', 'ch.ClearToMatchColorStyle',
    # On a chart sheet.
    'r = TypeName(wb.Charts.Add)', 'r = TypeName(wb.Charts.Add.ChartObjects)', 'r = wb.Charts.Add.ChartObjects.Count', 'r = wb.Charts.Add.Shapes.Count', 'wb.Charts.Add.Delete: r = wb.Charts.Count', 'wb.Charts.Add.Refresh', 'r = TypeName(wb.Charts.Add.Location(xlLocationAsObject, "Other")): r = r & wb.Charts.Count & o.ChartObjects.Count',
]
READS = ['d.ChartObjects.Count', 'o.ChartObjects.Count', 'wb.Charts.Count', 'wb.Sheets.Count', 'ActiveSheet.Name', 'd.Shapes.Count', 'ch.Name', 'ch.SeriesCollection.Count', 'TypeName(ch.Parent)', 'ch.Parent.Name']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, d As Worksheet, o As Worksheet, ch As Object, co As Object, s As String, v As String, r As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set d = wb.Worksheets(1): d.Name = "Data"', '    Set o = wb.Worksheets.Add(After:=d): o.Name = "Other": d.Activate',
             '    d.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))', '    d.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    d.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    Set co = d.ChartObjects.Add(300, 100, 300, 200): Set ch = co.Chart', '    ch.SetSourceData d.Range("A1:C4"), xlColumns', '    On Error GoTo Failed', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    wb.Close False', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 30):
            last = min(first + 29, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (first, last), timeout=280.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[case, answer] for case, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
