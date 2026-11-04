"""More members the checklist had as UNASSESSED, in live Excel: Select and Border of chart elements, UniqueValues.CalcFor,
ColorScale.SetLastPriority, PivotAxis, Range.XPath on a workbook with no XML map, and the Creator of shape nodes and text
columns. Each row has a new workbook.

    python scripts/measure_unassessed_more.py

writes tests/fixtures/unassessed_more.json, which tests/test_excel_unassessed_more.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/unassessed_more.json"

S = "ch.SeriesCollection(1)"
LABELS = f"{S}.HasDataLabels = True\n"
SELECT = "Note TypeName(Selection)"
BARS = 'ch.ChartGroups(1).HasUpDownBars = True\n'
PIVOT = ('ws.Range("E1:F1").Value = Array("R", "N")\nws.Range("E2:F2").Value = Array("x", 1)\nws.Range("E3:F3").Value = Array("y", 2)\n'
         'Dim pt As Object\nSet pt = wb.PivotCaches.Create(xlDatabase, ws.Range("E1:F3")).CreatePivotTable(ws.Range("H3"), "P")\n'
         'pt.PivotFields("R").Orientation = xlRowField\npt.AddDataField pt.PivotFields("N"), "Total", xlSum\n')
UNIQUE = 'Dim u As Object\nSet u = ws.Range("A1:A3").FormatConditions.AddUniqueValues\n'
SCALE = 'Dim c As Object\nSet c = ws.Range("A1:A3").FormatConditions.AddColorScale(2)\n'
ROWS = {
    "labels_select": f'{LABELS}Note {S}.DataLabels.Select\n{SELECT}',
    "gridlines_select": 'ch.Axes(2).HasMajorGridlines = True\nNote ch.Axes(2).MajorGridlines.Select\n' + SELECT,
    "legend_entry_select": 'ch.HasLegend = True\nNote ch.Legend.LegendEntries(1).Select\n' + SELECT,
    "trendline_select": f'{S}.Trendlines.Add\nNote {S}.Trendlines(1).Select\n{SELECT}',
    "up_bars_border": f'{BARS}Note TypeName(ch.ChartGroups(1).UpBars.Border)',
    "down_bars_border": f'{BARS}Note TypeName(ch.ChartGroups(1).DownBars.Border)',
    "legend_border": 'ch.HasLegend = True\nNote TypeName(ch.Legend.Border)',
    "plot_border": 'Note TypeName(ch.PlotArea.Border)',
    "point_border": f'Note TypeName({S}.Points(1).Border)',
    "unique_calc_for": f'{UNIQUE}Note u.CalcFor',
    "unique_calc_for_set": f'{UNIQUE}u.CalcFor = 0\nNote u.CalcFor',
    "scale_last": f'{SCALE}Dim d As Object\nSet d = ws.Range("A1:A3").FormatConditions.AddColorScale(3)\nNote c.Priority & "," & d.Priority\n'
                  'Note TypeName(c.SetLastPriority)\nNote c.Priority & "," & d.Priority',
    "scale_last_alone": f'{SCALE}c.SetLastPriority\nNote c.Priority',
    "pivot_axis": f'{PIVOT}Note TypeName(pt.PivotRowAxis.Parent) & "," & TypeName(pt.PivotRowAxis.Application) & "," & pt.PivotRowAxis.Creator',
    "pivot_axis_column": f'{PIVOT}Note TypeName(pt.PivotColumnAxis.Parent) & "," & TypeName(pt.PivotColumnAxis.Application) & "," & pt.PivotColumnAxis.Creator',
    "xpath_type": 'Note TypeName(ws.Range("A1").XPath)',
    "xpath_value": 'Note "[" & ws.Range("A1").XPath.Value & "]"',
    "xpath_default": 'Note "[" & ws.Range("A1").XPath & "]"',
    "xpath_map": 'Note TypeName(ws.Range("A1").XPath.Map)',
    "xpath_repeating": 'Note ws.Range("A1").XPath.Repeating',
    "xpath_creator": 'Note ws.Range("A1").XPath.Creator',
    "xpath_clear": 'ws.Range("A1").XPath.Clear\nNote "ok"',
    "xpath_block": 'Note TypeName(ws.Range("A1:B2").XPath)',
    "nodes_creator": 'Dim sh As Object\nSet sh = ws.Shapes.BuildFreeform(0, 10, 10).ConvertToShape\nNote sh.Nodes.Creator',
    "node_creator": 'Dim b As Object, sh As Object\nSet b = ws.Shapes.BuildFreeform(0, 10, 10)\nb.AddNodes 0, 0, 40, 40\nSet sh = b.ConvertToShape\n'
                    'Note sh.Nodes(1).Creator',
    "text_column_creator": 'Dim sh As Object\nSet sh = ws.Shapes.AddTextbox(1, 10, 10, 50, 50)\nNote sh.TextFrame2.Column.Creator',
}

PROBE = '''Dim Log As String
Sub Note(ByVal s As Variant)
    Log = Log & CStr(s) & "|"
End Sub
Function Probe() As String
    Dim wb As Workbook, ws As Worksheet, ch As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = 1: ws.Range("A2").Value = 2: ws.Range("A3").Value = 3
    ws.Range("B1").Value = 4: ws.Range("B2").Value = 5: ws.Range("B3").Value = 6
    Set ch = ws.ChartObjects.Add(10, 10, 300, 200).Chart
    ch.SetSourceData ws.Range("A1:B3")
    ch.ChartType = xlLine
    On Error GoTo Failed
{action}
    GoTo Done
Failed:
    Log = Log & "!E" & Err.Number
Done:
    On Error Resume Next
    wb.Close False
    Probe = Log
End Function
'''


def main() -> None:
    rows: list[dict[str, str]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=60.0)) as excel:
        excel.new_document()
        for name, action in ROWS.items():
            result = excel.run_vba(PROBE.format(action=action), "Probe", timeout=60.0)
            value = str(result.value) if result.ok else f"failed {result.outcome}"
            rows.append({"name": name, "action": action, "result": value})
            print(name, value[:110], flush=True)
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "probe": PROBE, "rows": rows},
                              indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
