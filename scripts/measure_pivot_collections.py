"""The small collections of a pivot table over a list: Slicers, CalculatedFields and CalculatedItems, what each is, how
many it holds, what walking it finds and who its Application and Parent are.

    python scripts/measure_pivot_collections.py

writes tests/fixtures/pivot_collections.json, which tests/test_excel_pivot_collections.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/pivot_collections.json"
READS = ["TypeName(pt.Slicers)", "pt.Slicers.Count", "TypeName(pt.Slicers.Parent)", "TypeName(pt.Slicers.Application)",
         "pt.Slicers(1).Name",
         "TypeName(pt.CalculatedFields.Application)", "TypeName(pt.CalculatedFields.Parent)", "pt.CalculatedFields.Count",
         'TypeName(pt.PivotFields("Region").CalculatedItems.Application)',
         'TypeName(pt.PivotFields("Region").CalculatedItems.Parent)', 'pt.PivotFields("Region").CalculatedItems.Count',
         "Walked(pt.Slicers)", "Walked(pt.CalculatedFields)", 'Walked(pt.PivotFields("Region").CalculatedItems)']

PROBE = '''Function Walked(things As Object) As String
    Dim one As Object, n As Long
    For Each one In things
        n = n + 1
    Next
    Walked = "walked" & n
End Function

Function Probe() As String
    Dim wb As Workbook, data As Worksheet, ws As Worksheet, pt As Object, s As String, v As String
    Set wb = Workbooks.Add(xlWBATWorksheet): Set data = wb.Worksheets(1)
    data.Range("A1:B1").Value = Array("Region", "Amount")
    data.Range("A2:B2").Value = Array("East", 1)
    data.Range("A3:B3").Value = Array("West", 2)
    Set ws = wb.Worksheets.Add
    Set pt = wb.PivotCaches.Create(xlDatabase, data.Range("A1:B3")).CreatePivotTable(ws.Range("A3"), "P")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Amount"), "Total", xlSum
    On Error Resume Next
{reads}
    On Error GoTo 0
    Application.DisplayAlerts = False
    wb.Close False
    Probe = Mid$(s, 2)
End Function
'''


def source() -> str:
    reads = "\n".join(f'    Err.Clear\n    v = ""\n    v = {read}\n    If Err.Number Then v = "E" & Err.Number\n'
                      f'    s = s & "," & v' for read in READS)
    return PROBE.format(reads=reads)


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=60.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source(), "Probe", timeout=60.0)
        assert result.ok, result
    print(result.value)
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "reads": READS, "source": source(),
                               "result": str(result.value)}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
