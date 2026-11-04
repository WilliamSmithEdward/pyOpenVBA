"""A note box's Fill and ForeColor, held while the note goes. Prediction of the default rule for a held object: 424 on every line. The rival: Parent, Application and Creator still answer, as on a held Comment.

    python scripts/measure_gone_note_fill.py

writes tests/fixtures/gone_note_fill.json, which tests/test_excel_note_shapes.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/gone_note_fill.json'
SOURCE = r'''Public Function Tried() As String
    Dim ws As Worksheet, c As Comment, sh As Object, f As Object, k As Object, s As String, v As Variant
    Set ws = ActiveWorkbook.Worksheets(1)
    Set c = ws.Range("B2").AddComment("x")
    Set sh = c.Shape
    Set f = sh.Fill
    Set k = f.ForeColor
    s = "before " & k.RGB
    c.Delete
    On Error Resume Next
    Err.Clear: v = Empty: v = k.RGB
    s = s & "|col.RGB " & Err.Number & " " & CStr(v)
    Err.Clear: k.RGB = 255
    s = s & "|col.RGB= " & Err.Number
    Err.Clear: v = TypeName(k.Parent)
    s = s & "|col.Parent " & Err.Number & " " & CStr(v)
    Err.Clear: v = "": v = TypeName(f.ForeColor)
    s = s & "|fill.ForeColor " & Err.Number & " " & CStr(v)
    Err.Clear: v = "": v = TypeName(f.Parent)
    s = s & "|fill.Parent " & Err.Number & " " & CStr(v)
    Err.Clear: v = "": v = TypeName(k.Application)
    s = s & "|col.Application " & Err.Number & " " & CStr(v)
    Err.Clear: v = "": v = TypeName(f.Application)
    s = s & "|fill.Application " & Err.Number & " " & CStr(v)
    Err.Clear: v = "": v = k.Creator
    s = s & "|col.Creator " & Err.Number & " " & CStr(v)
    Err.Clear: v = "": v = f.Creator
    s = s & "|fill.Creator " & Err.Number & " " & CStr(v)
    On Error GoTo 0
    Tried = s
End Function
'''


def main() -> None:
    with ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        result = excel.run_vba(SOURCE, 'Tried', (), timeout=60.0)
        assert result.ok, result
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'answer': str(result.value)}, indent=1) + '\n', encoding='utf-8')
    print(result.value)


if __name__ == '__main__':
    main()
