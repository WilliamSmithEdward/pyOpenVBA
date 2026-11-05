"""ControlFormat.Value's type, typed and late, on a check box, a drop-down, a list box and a spinner. Rival rules: typed is Long throughout; late is a Double but for off.

    python scripts/measure_control_value_types.py

writes tests/fixtures/typed_answers/control_value_types.json, which tests/test_excel_typed_answers.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/typed_answers/control_value_types.json'
SOURCE = r'''Private Function Both(s As Shape) As String
    Dim o As Object, c As ControlFormat
    Set c = s.ControlFormat
    Set o = s
    Both = TypeName(c.Value) & ":" & c.Value & "/" & TypeName(o.ControlFormat.Value) & "/" & TypeName(s.ControlFormat.Value)
End Function

Public Function Tried() As String
    Dim wb As Workbook, ws As Worksheet, s As Shape, r As String, v As Variant
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = "a": ws.Range("A2").Value = "b"
    Set s = ws.Shapes.AddFormControl(xlCheckBox, 0, 0, 90, 20)
    For Each v In Array(xlOn, xlOff, xlMixed)
        s.ControlFormat.Value = v
        r = r & "check " & v & " " & Both(s) & "|"
    Next
    Set s = ws.Shapes.AddFormControl(xlDropDown, 0, 40, 90, 20)
    s.ControlFormat.ListFillRange = "$A$1:$A$2"
    s.ControlFormat.Value = 2
    r = r & "drop " & Both(s) & "|"
    s.ControlFormat.Value = 0
    r = r & "drop0 " & Both(s) & "|"
    Set s = ws.Shapes.AddFormControl(xlListBox, 0, 80, 90, 40)
    s.ControlFormat.ListFillRange = "$A$1:$A$2"
    s.ControlFormat.Value = 1
    r = r & "list " & Both(s) & "|"
    Set s = ws.Shapes.AddFormControl(xlSpinner, 100, 0, 20, 40)
    s.ControlFormat.Value = 3
    r = r & "spin " & Both(s) & "|"
    wb.Close False
    Tried = r
End Function
'''


def main() -> None:
    with ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        done = excel.run_vba(SOURCE, 'Tried', (), timeout=40.0, line_numbers=False)
        assert done.ok, done
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'answer': str(done.value)}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
