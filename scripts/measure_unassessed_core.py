"""Members the checklist had as UNASSESSED that macros reach most, measured in live Excel: Workbook.Calculate, Shape.Select,
Shapes.SelectAll and AddChart, ChartObject's Copy, Cut and Duplicate, Chart.ProtectGoalSeek and Series.Border. Rows that
copy or cut use the clipboard, so the run holds live.lock. Nothing here changes a user's option.

    python scripts/measure_unassessed_core.py

writes tests/fixtures/unassessed_core.json, which tests/test_excel_unassessed_core.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/unassessed_core.json"

SHAPE = 'Set sh = ws.Shapes.AddShape(1, 10, 10, 50, 50)\n'
SECOND = 'Set other = ws.Shapes.AddShape(1, 100, 10, 50, 50)\n'
CHART = 'Set co = ws.ChartObjects.Add(10, 10, 200, 100)\n'
SELECTION = 'Note TypeName(Selection)\nNote ActiveSheet.Name\n'
ROWS = {
    "book_calculate": 'ws.Range("A1").Formula = "=1+1"\nApplication.Calculation = xlCalculationManual\nws.Range("B1").Value = 5\n'
                      'ws.Range("A2").Formula = "=B1*2"\nws.Range("B1").Value = 7\nNote ws.Range("A2").Value\n'
                      'Note TypeName(wb.Calculate)\nNote ws.Range("A2").Value\nApplication.Calculation = xlCalculationAutomatic',
    "book_calculate_other_book": 'Dim b As Workbook\nSet b = Workbooks.Add\nb.Worksheets(1).Range("A1").Value = 3\n'
                                 'b.Worksheets(1).Range("A2").Formula = "=A1*2"\nApplication.Calculation = xlCalculationManual\n'
                                 'b.Worksheets(1).Range("A1").Value = 4\nwb.Calculate\nNote b.Worksheets(1).Range("A2").Value\n'
                                 'b.Calculate\nNote b.Worksheets(1).Range("A2").Value\nApplication.Calculation = xlCalculationAutomatic\nb.Close False',
    "shape_select": SHAPE + 'sh.Select\n' + SELECTION + 'Note Selection.Name\nNote Selection.ShapeRange.Count',
    "shape_select_replace_false": SHAPE + SECOND + 'sh.Select\nother.Select False\n' + SELECTION + 'Note Selection.ShapeRange.Count',
    "shape_select_replace_true": SHAPE + SECOND + 'sh.Select\nother.Select True\n' + SELECTION + 'Note Selection.Name',
    "shape_select_hidden_behind": SHAPE + 'sh.Visible = False\nwb.Worksheets.Add\nsh.Select\n' + SELECTION,
    "shape_select_other_sheet": SHAPE + 'wb.Worksheets.Add\nsh.Select\n' + SELECTION,
    "shape_select_hidden": SHAPE + 'sh.Visible = False\nsh.Select\n' + SELECTION,
    "shapes_select_all": SHAPE + SECOND + 'ws.Shapes.SelectAll\n' + SELECTION + 'Note Selection.ShapeRange.Count',
    "shapes_select_all_hidden": SHAPE + SECOND + 'sh.Visible = False\nws.Shapes.SelectAll\n' + SELECTION + 'Note Selection.Name',
    "shapes_select_all_every_hidden": SHAPE + 'sh.Visible = False\nws.Shapes.SelectAll\n' + SELECTION,
    "shapes_select_all_replaces": SHAPE + SECOND + 'sh.Select\nws.Shapes.SelectAll\n' + SELECTION + 'Note Selection.ShapeRange.Count',
    "shapes_select_all_with_chart": SHAPE + CHART + 'ws.Shapes.SelectAll\n' + SELECTION + 'Note Selection.ShapeRange.Count',
    "shapes_select_all_none": 'ws.Shapes.SelectAll\n' + SELECTION,
    "shapes_select_all_other_sheet": SHAPE + 'wb.Worksheets.Add\nws.Shapes.SelectAll\n' + SELECTION,
    "shapes_add_chart": 'Set sh = ws.Shapes.AddChart\nNote TypeName(sh)\nNote sh.Type\nNote sh.HasChart\nNote sh.Name\nNote sh.Chart.ChartType\n'
                        'Note sh.Width & "," & sh.Height',  # Left and Top centre it in the window, which is the screen's
    "shapes_add_chart_type": 'Set sh = ws.Shapes.AddChart(xlLine, 10, 20, 300, 200)\nNote sh.Chart.ChartType\nNote sh.Left & "," & sh.Top & "," & sh.Width & "," & sh.Height',
    "shapes_add_chart_bad_type": 'Set sh = ws.Shapes.AddChart(99999)',
    "co_duplicate": CHART + 'Dim d As Object\nSet d = co.Duplicate\nNote TypeName(d)\nNote d.Name\nNote d.Left & "," & d.Top\nNote ws.ChartObjects.Count',
    "co_copy_paste": CHART + 'co.Copy\nNote Application.CutCopyMode\nNote ws.ChartObjects.Count',
    "co_cut": CHART + 'co.Cut\nNote Application.CutCopyMode\nNote ws.ChartObjects.Count',
    "protect_goal_seek": CHART + 'Note co.Chart.ProtectGoalSeek\nco.Chart.ProtectGoalSeek = True\nNote co.Chart.ProtectGoalSeek',
    "series_border": 'ws.Range("A1").Value = 1\nws.Range("A2").Value = 2\n' + CHART + 'co.Chart.SeriesCollection.NewSeries\n'
                     'co.Chart.SeriesCollection(1).Values = ws.Range("A1:A2")\nNote TypeName(co.Chart.SeriesCollection(1).Border)',
}

PROBE = '''Dim Log As String
Sub Note(ByVal s As Variant)
    Log = Log & CStr(s) & "|"
End Sub
Function Probe() As String
    Dim wb As Workbook, ws As Worksheet, sh As Object, other As Object, co As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    On Error GoTo Failed
{action}
    GoTo Done
Failed:
    Log = Log & "!E" & Err.Number
Done:
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
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
            print(name, value[:170], flush=True)
            rows.append({"name": name, "action": action, "result": value})
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "probe": PROBE, "rows": rows},
                              indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
