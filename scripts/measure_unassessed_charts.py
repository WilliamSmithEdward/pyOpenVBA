"""The chart members the checklist had as UNASSESSED, in live Excel, over one line chart of two series: DataTable,
FullSeriesCollection, LegendKey, CategoryCollection, LeaderLines, ChartCategory, and the Border, Name, Select and Copy of
titles, the chart area and its surfaces. DataTable's Border and Format are left to another branch. A row that selects
or copies holds live.lock.

    python scripts/measure_unassessed_charts.py

writes tests/fixtures/unassessed_charts.json, which tests/test_excel_unassessed_charts.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/unassessed_charts.json"

T = "ch.DataTable"
F = "ch.FullSeriesCollection"
K = "ch.Legend.LegendEntries(1).LegendKey"
C = "ch.ChartGroups(1).FullCategoryCollection"
L = "ch.SeriesCollection(1).LeaderLines"
TABLE = "ch.HasDataTable = True\n"
LEGEND = "ch.HasLegend = True\n"
LEADERS = "ch.SeriesCollection(1).HasDataLabels = True\nch.SeriesCollection(1).HasLeaderLines = True\n"
AXIS = "ch.Axes(1).HasTitle = True\n"
TITLE = 'ch.HasTitle = True\nch.ChartTitle.Text = "T"\n'
ROWS = {
    "table_types": f'{TABLE}Note TypeName({T}) & "," & TypeName({T}.Parent) & "," & TypeName({T}.Application) & "," & {T}.Creator',
    "table_borders": f'{TABLE}Note {T}.HasBorderHorizontal & "," & {T}.HasBorderVertical & "," & {T}.HasBorderOutline & "," & {T}.ShowLegendKey',
    "table_borders_set": f'{TABLE}{T}.HasBorderHorizontal = False\n{T}.HasBorderVertical = False\n{T}.HasBorderOutline = False\n{T}.ShowLegendKey = True\n'
                         f'Note {T}.HasBorderHorizontal & "," & {T}.HasBorderVertical & "," & {T}.HasBorderOutline & "," & {T}.ShowLegendKey',
    "table_font": f'{TABLE}Note TypeName({T}.Font)',
    "table_select": f'{TABLE}Note {T}.Select\nNote TypeName(Selection)',
    "table_select_then_cell": f'{TABLE}{T}.Select\nws.Range("A1").Select\nNote TypeName(Selection)',
    "table_delete": f'{TABLE}Note TypeName({T}.Delete)\nNote ch.HasDataTable',
    "table_none": f'Note TypeName({T})',
    "full_types": f'Note TypeName({F}) & "," & TypeName({F}.Parent) & "," & TypeName({F}.Application) & "," & {F}.Creator & "," & {F}.Count',
    "full_item": f'Note TypeName({F}(1)) & "," & {F}(1).Name & "," & {F}.Item(2).Name',
    "full_default": f'Note {F}.[_Default](2).Name',
    "full_item_out": f'Note TypeName({F}(9))',
    "full_item_name": f'Note {F}("Up").Name',
    "full_new_series": f'Note TypeName({F}.NewSeries)\nNote {F}.Count & "," & ch.SeriesCollection.Count',
    "full_add": f'Note TypeName({F}.Add(ws.Range("D1:D3")))\nNote {F}.Count',
    "full_extend": f'Note TypeName({F}.Extend(ws.Range("B4:C4")))\nNote {F}.Count',
    "full_filtered": f'{F}(1).IsFiltered = True\nNote {F}.Count & "," & ch.SeriesCollection.Count',
    "key_types": f'{LEGEND}Note TypeName({K}) & "," & TypeName({K}.Parent) & "," & {K}.Creator & "," & TypeName({K}.Format)',
    "key_marker": f'{LEGEND}Note {K}.MarkerStyle & "," & {K}.Smooth & "," & {K}.InvertIfNegative',
    "key_marker_set": f'{LEGEND}{K}.MarkerStyle = xlMarkerStyleCircle\nNote {K}.MarkerStyle & "," & ch.SeriesCollection(1).MarkerStyle',
    "key_smooth_set": f'{LEGEND}{K}.Smooth = True\nNote {K}.Smooth & "," & ch.SeriesCollection(1).Smooth',
    "key_invert_set": f'{LEGEND}{K}.InvertIfNegative = True\nNote {K}.InvertIfNegative',
    "key_smooth_off": f'{LEGEND}{K}.Smooth = True\n{K}.Smooth = False\nNote ch.SeriesCollection(1).Smooth',
    "key_smooth_text": f'{LEGEND}{K}.Smooth = "x"',
    "key_invert_series": f'{LEGEND}{K}.InvertIfNegative = True\nNote ch.SeriesCollection(1).InvertIfNegative',
    "key_invert_bar": f'ch.ChartType = xlColumnClustered\n{LEGEND}{K}.InvertIfNegative = True\nNote {K}.InvertIfNegative & "," & ch.SeriesCollection(1).InvertIfNegative',
    "key_select": f'{LEGEND}Note {K}.Select\nNote TypeName(Selection)',
    "key_delete": f'{LEGEND}{K}.Delete',
    "category_types": f'Note TypeName({C}) & "," & TypeName({C}.Parent) & "," & TypeName({C}.Application) & "," & {C}.Creator & "," & {C}.Count',
    "category_item": f'Note TypeName({C}(1)) & "," & {C}(1).Name & "," & {C}.Item(2).Name & "," & {C}.[_Default](1).Name',
    "category_item_out": f'Note TypeName({C}(9))',
    "category_members": f'Note TypeName({C}(1).Parent) & "," & TypeName({C}(1).Application) & "," & {C}(1).Creator & "," & {C}(1).IsFiltered',
    "category_filtered": f'{C}(1).IsFiltered = True\nNote {C}.Count & "," & ch.ChartGroups(1).CategoryCollection.Count & "," & {C}(1).IsFiltered',
    "category_group": 'Note TypeName(ch.ChartGroups(1).CategoryCollection) & "," & ch.ChartGroups(1).CategoryCollection.Count',
    "leaders_types": f'{LEADERS}Note TypeName({L}) & "," & TypeName({L}.Parent) & "," & TypeName({L}.Application) & "," & {L}.Creator',
    "leaders_border": f'{LEADERS}Note TypeName({L}.Border) & "," & TypeName({L}.Format)',
    "leaders_delete": f'{LEADERS}Note TypeName({L}.Delete)\nNote ch.SeriesCollection(1).HasLeaderLines',
    "leaders_none": f'Note TypeName({L})',
    "axis_title": f'{AXIS}Note TypeName(ch.Axes(1).AxisTitle.Border) & "," & ch.Axes(1).AxisTitle.Name',
    "axis_value": 'ch.Axes(2).HasTitle = True\nNote ch.Axes(2).AxisTitle.Name',
    "axis_category_text": f'{AXIS}ch.Axes(1).AxisTitle.Text = "Mine"\nNote ch.Axes(1).AxisTitle.Name',
    "axis_secondary": 'ch.SeriesCollection(2).AxisGroup = 2\nch.Axes(2, 2).HasTitle = True\nNote ch.Axes(2, 2).AxisTitle.Name',
    "axis_name_set": f'{AXIS}ch.Axes(1).AxisTitle.Name = "x"',
    "axis_title_select": f'{AXIS}Note ch.Axes(1).AxisTitle.Select\nNote TypeName(Selection)',
    "chart_area": 'Note TypeName(ch.ChartArea.Border)',
    "surface": 'Note ch.PlotArea.Name & "," & ch.PlotArea.Creator',
    "walls": 'ch.ChartType = xl3DColumn\nNote ch.Walls.Name & "," & ch.Walls.Creator & "," & ch.Floor.Name & "," & ch.Floor.Creator',
    "chart_title": f'{TITLE}Note TypeName(ch.ChartTitle.Border)',
    "chart_title_select": f'{TITLE}Note ch.ChartTitle.Select\nNote TypeName(Selection)',
    "held_by_chart": 'Note ch.Shapes.Count & "," & TypeName(ch.Shapes) & "," & TypeName(ch.Shapes.Application) & "," & TypeName(ch.Shapes.Parent) & "," & ch.Shapes.Creator',
}

PROBE = '''Dim Log As String
Sub Note(ByVal s As Variant)
    Log = Log & CStr(s) & "|"
End Sub
Function Probe() As String
    Dim wb As Workbook, ws As Worksheet, ch As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("", "Up", "Down")
    ws.Range("A2:C2").Value = Array("a", 1, 4)
    ws.Range("A3:C3").Value = Array("b", 2, 3)
    ws.Range("A4:C4").Value = Array("c", 5, 6)
    ws.Range("D1").Value = "More": ws.Range("D2").Value = 7: ws.Range("D3").Value = 8: ws.Range("D4").Value = 9
    Set ch = ws.ChartObjects.Add(10, 10, 300, 200).Chart
    ch.SetSourceData ws.Range("A1:C3")
    ch.ChartType = xlLineMarkers
    On Error GoTo Failed
{action}
    GoTo Done
Failed:
    Log = Log & "!E" & Err.Number
Done:
    On Error Resume Next
    Application.CutCopyMode = False
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
            print(name, value[:120], flush=True)
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "probe": PROBE, "rows": rows},
                              indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
