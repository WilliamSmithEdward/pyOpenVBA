"""The effect members the checklist had as UNASSESSED, in live Excel: ShadowFormat, GlowFormat and SoftEdgeFormat
read on a new shape and after each setting, with what a bad value does, a chart element's ThreeDFormat, and the plain
members of Adjustments, CalloutFormat, FreeformBuilder and GroupShapes. Each row has a new workbook and shape.

    python scripts/measure_unassessed_effects.py

writes tests/fixtures/unassessed_effects.json, which tests/test_excel_unassessed_effects.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/unassessed_effects.json"

S, G, E = "sh.Shadow", "sh.Glow", "sh.SoftEdge"
CHART_AREA = 'Dim ca As Object\nSet ca = ws.ChartObjects.Add(10, 10, 200, 100).Chart.ChartArea\n'
SHADOW = f'Note {S}.Visible & "," & {S}.Type & "," & {S}.Blur & "," & {S}.OffsetX & "," & {S}.OffsetY'
GLOW = f'Note {G}.Radius & "," & {G}.Transparency'
EDGE = f'Note {E}.Type & "," & {E}.Radius'
ROWS = {
    "shadow_new": SHADOW,
    "shadow_types": f'Note TypeName({S}) & "," & TypeName({S}.Parent) & "," & {S}.Creator',
    "shadow_visible": f'{S}.Visible = True\n{SHADOW}',
    "shadow_visible_off": f'{S}.Visible = True\n{S}.Visible = False\n{SHADOW}',
    "shadow_blur": f'{S}.Blur = 4\n{SHADOW}',
    "shadow_blur_negative": f'{S}.Blur = -1\n{SHADOW}',
    "shadow_blur_far": f'{S}.Blur = 200\n{SHADOW}',
    "shadow_offset": f'{S}.OffsetX = 3\n{S}.OffsetY = -2\n{SHADOW}',
    "shadow_offset_far": f'{S}.OffsetX = 5000\n{SHADOW}',
    "shadow_type": f'{S}.Type = msoShadow21\n{SHADOW}',
    "shadow_type_2": f'{S}.Type = msoShadow1\n{SHADOW}',
    "shadow_type_bad": f'{S}.Type = 999\n{SHADOW}',
    "shadow_visible_text": f'{S}.Visible = "x"',
    "glow_new": GLOW,
    "glow_types": f'Note TypeName({G}) & "," & TypeName({G}.Color) & "," & TypeName({G}.Parent) & "," & {G}.Creator',
    "glow_radius": f'{G}.Radius = 5\n{GLOW}',
    "glow_radius_negative": f'{G}.Radius = -1\n{GLOW}',
    "glow_radius_far": f'{G}.Radius = 500\n{GLOW}',
    "glow_transparency": f'{G}.Radius = 5\n{G}.Transparency = 0.5\n{GLOW}',
    "glow_transparency_without_radius": f'{G}.Transparency = 0.5\n{GLOW}',
    "glow_transparency_far": f'{G}.Radius = 5\n{G}.Transparency = 2\n{GLOW}',
    "glow_radius_off": f'{G}.Radius = 5\n{G}.Radius = 0\n{GLOW}',
    "glow_color": f'{G}.Radius = 5\n{G}.Color.RGB = 255\nNote {G}.Color.RGB\n{GLOW}',
    "edge_new": EDGE,
    "edge_types": f'Note TypeName({E}) & "," & TypeName({E}.Parent) & "," & {E}.Creator',
    "edge_type_1": f'{E}.Type = 1\n{EDGE}',
    "edge_type_6": f'{E}.Type = 6\n{EDGE}',
    "edge_type_none": f'{E}.Type = 1\n{E}.Type = 0\n{EDGE}',
    "edge_type_bad": f'{E}.Type = 7\n{EDGE}',
    "edge_radius": f'{E}.Radius = 3\n{EDGE}',
    "edge_radius_matches": f'{E}.Radius = 2.5\n{EDGE}',
    "edge_radius_negative": f'{E}.Radius = -1\n{EDGE}',
    "edge_radius_far": f'{E}.Radius = 500\n{EDGE}',
    "three_chart": f'{CHART_AREA}Note ca.Format.ThreeD.Visible & "," & ca.Format.ThreeD.BevelTopType',
    "three_chart_types": f'{CHART_AREA}Note TypeName(ca.Format.ThreeD) & "," & TypeName(ca.Format.ThreeD.Parent) & "," & ca.Format.ThreeD.Creator',
    "three_shape": 'Note sh.ThreeD.Visible',
    "adjustments_creator": 'Note sh.Adjustments.Creator',
    "callout_creator": 'Dim c As Object\nSet c = ws.Shapes.AddCallout(2, 10, 10, 50, 50)\nNote c.Callout.Creator',
    "freeform_creator": 'Note ws.Shapes.BuildFreeform(0, 10, 10).Creator',
    "group_application": 'Dim o As Object\nSet o = ws.Shapes.AddShape(1, 60, 60, 20, 20)\n'
                         'Note TypeName(ws.Shapes.Range(Array(sh.Name, o.Name)).Group.GroupItems.Application)',
}

PROBE = '''Dim Log As String
Sub Note(ByVal s As Variant)
    Log = Log & CStr(s) & "|"
End Sub
Function Probe() As String
    Dim wb As Workbook, ws As Worksheet, sh As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    Set sh = ws.Shapes.AddShape(1, 10, 10, 50, 50)
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
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "probe": PROBE, "rows": rows},
                              indent=1) + "\n", encoding="utf-8")
    for row in rows:
        print(row["name"], row["result"][:110], flush=True)


if __name__ == "__main__":
    main()
