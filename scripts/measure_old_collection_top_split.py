"""A split for two rules tests/fixtures/old_collection_geometry.json left open, on a book a macro makes (two buttons at Top 70 and 130,
two arcs from Arcs.Add(10, 70, 60, 120) and Arcs.Add(10, 130, 60, 180)), at the display there is (StandardHeight recorded):

1. Buttons.Top set to 300, 0 and 37.5. Rival rules, written before measuring: the items move together so the topmost lands at 0, whatever
   the value (37.5 and 31.25 both gave 0); rival: so the topmost lands at the value less 37.5 (which also gives 0 for 37.5 and below);
   rival: so the topmost lands at the value, the 0 coming from something else.
2. One arc's Height set to 37.5 alone, and both through Arcs.Height = 75. Rival rules: an arc's Height takes half of what is set, alone
   too; rival: only through the collection.

    python scripts/measure_old_collection_top_split.py

writes tests/fixtures/old_collection_top_split.json, which tests/test_excel_old_collection_geometry.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/old_collection_top_split.json'
SOURCE = r'''Private Function Boxes(c As Object) As String
    Dim x As Object, s As String
    For Each x In c
        s = s & ";" & x.Left & "," & x.Top & "," & x.Width & "," & x.Height
    Next
    Boxes = Mid(s, 2)
End Function

Public Function Row(ByVal which As String) As String
    Dim wb As Workbook, ws As Object, s As String
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Buttons.Add 100, 70, 60, 30: ws.Buttons.Add 100, 130, 60, 30
    ws.Arcs.Add 10, 70, 60, 120: ws.Arcs.Add 10, 130, 60, 180
    s = ws.StandardHeight & "|"
    On Error Resume Next
    Select Case which
    Case "top 300": ws.Buttons.Top = 300
    Case "top 0": ws.Buttons.Top = 0
    Case "top 37.5": ws.Buttons.Top = 37.5
    Case "one arc height": ws.Arcs(1).Height = 37.5
    Case "arcs height": ws.Arcs.Height = 75
    End Select
    s = s & "E" & Err.Number & "|" & Boxes(ws.Buttons) & "|" & Boxes(ws.Arcs) & "|" & ws.StandardHeight
    wb.Close False
    Row = s
End Function
'''
CASES = ('none', 'top 300', 'top 0', 'top 37.5', 'one arc height', 'arcs height')


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    found: dict[str, str] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        for which in CASES:
            ran = excel.run_vba(SOURCE, 'Row', (which,), timeout=60.0, line_numbers=False)
            found[which] = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
            print(which, found[which], flush=True)
    heights = {one.split('|')[0] for one in found.values()} | {one.rsplit('|', 1)[-1] for one in found.values()}
    assert len(heights) == 1, heights
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': float(heights.pop()), 'source': SOURCE,
                               'cases': found}, indent=1) + '\n', encoding='utf-8')
    print('wrote', OUT.name, flush=True)


if __name__ == '__main__':
    main()
