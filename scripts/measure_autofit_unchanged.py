"""Column AutoFit that finds the width a column already shows, on a 96-DPI display.

AutoFit changes a column only where the width it finds differs from the
one the column shows: a standard column fitted to the standard width stays
standard, and follows a new StandardWidth; a column of its own width keeps
it without best fit. This fits columns of abcdefgh in Aptos Narrow 11, 64
pixels, the standard width, and 63 with its first three characters bold,
from each starting width, records each column's width and UseStandardWidth,
saves the workbook to record the columns the file keeps, and then reads the
columns again under a wider StandardWidth.

    python scripts/measure_autofit_unchanged.py

writes tests/fixtures/autofit_unchanged.json, which
tests/test_excel_column_autofit.py replays. Excel must lay sheets out at 96
DPI, a 15pt standard row.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/autofit_unchanged.json"

PROBE = r'''Private Function State(ws As Worksheet, column As Long) As String
    State = CLng(ws.Columns(column).Width * 4 / 3) & "px std=" & ws.Columns(column).UseStandardWidth
End Function

Public Function Probe(path As String) As String
    Dim wb As Workbook, ws As Worksheet, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Styles("Normal").Font.Name = "Aptos Narrow"
    wb.Styles("Normal").Font.Size = 11
    Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = "abcdefgh"
    ws.Columns(1).AutoFit
    s = s & "fresh=" & State(ws, 1)
    ws.Range("B1").Value = "abcdefgh"
    ws.Columns(2).ColumnWidth = 20
    ws.Columns(2).AutoFit
    s = s & " from20=" & State(ws, 2)
    ws.Range("C1").Value = "abcdefgh"
    ws.Columns(3).ColumnWidth = 8.43
    s = s & " set843=" & State(ws, 3)
    ws.Columns(3).AutoFit
    s = s & " from843=" & State(ws, 3)
    ws.Range("D1").Value = "abcdefgh"
    ws.Range("D1").Characters(1, 3).Font.Bold = True
    ws.Columns(4).AutoFit
    s = s & " fit63=" & State(ws, 4)
    ws.Range("E1").Value = "abcdefgh"
    ws.Columns(5).Hidden = True
    ws.Columns(5).AutoFit
    s = s & " hidden=" & State(ws, 5) & " shown=" & Not ws.Columns(5).Hidden
    ws.Range("F1").Value = "abcdefgh"
    ws.Range("F1").Characters(1, 3).Font.Bold = True
    ws.Columns(6).AutoFit
    ws.Range("F1").Characters(1, 3).Font.Bold = False
    ws.Columns(6).AutoFit
    s = s & " back_to_64=" & State(ws, 6)
    ws.Columns(8).ColumnWidth = 20
    ws.Columns(8).ColumnWidth = 8.43
    s = s & " h_20_then_843=" & State(ws, 8)
    ws.Range("I1").Value = "abcdefgh"
    ws.Columns(9).ColumnWidth = 20
    ws.Columns(9).ColumnWidth = 8.43
    ws.Columns(9).AutoFit
    s = s & " i_custom64_fit=" & State(ws, 9)
    ws.Range("J1").Value = "abcdefgh"
    ws.Range("J1").Characters(1, 3).Font.Bold = True
    ws.Columns(10).ColumnWidth = 8.29
    s = s & " j_set=" & State(ws, 10)
    ws.Columns(10).AutoFit
    s = s & " j_fit63=" & State(ws, 10)
    wb.SaveAs path, xlOpenXMLWorkbook
    ws.StandardWidth = 10
    s = s & " after_standard_10: fresh=" & State(ws, 1) & " from20=" & State(ws, 2) & " from843=" & State(ws, 3) & _
        " hidden=" & State(ws, 5) & " back_to_64=" & State(ws, 6) & " G=" & State(ws, 7) & " H=" & State(ws, 8) & _
        " I=" & State(ws, 9)
    wb.Close False
    Application.DisplayAlerts = True
    Probe = s
End Function
'''


def main() -> None:
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        standard = excel.run_vba("Function H() As String\nH = CStr(ActiveSheet.StandardHeight)\nEnd Function", "H")
        assert standard.value == "15", f"Excel lays sheets out at another DPI: StandardHeight {standard.value}"
        version = excel.run_vba('Function V() As String\nV = Application.Version & "|" & Application.Build\n'
                                'End Function', "V")
        path = Path(folder) / "unchanged.xlsx"
        result = excel.run_vba(PROBE, "Probe", [str(path)], timeout=120.0)
        assert result.ok and not result.dialogs, f"{result.outcome}: {result.message} {result.error}"
        with zipfile.ZipFile(path) as package:
            sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    found = re.search(r"<cols>.*?</cols>", sheet)
    record = {"excel_version_build": version.value, "measured_at": datetime.now(timezone.utc).isoformat(),
              "probe": PROBE, "answer": result.value, "cols": found.group(0) if found else ""}
    OUT.write_text(json.dumps(record, indent=2) + "\n", encoding="utf-8")
    print(result.value)
    print(record["cols"])


if __name__ == "__main__":
    main()
