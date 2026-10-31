"""Range.LocationInTable for every cell of a pivot table, with its PivotCellType beside it, in live Excel.

Excel opens a copy of each workbook scripts/measure_pivot_cells.py reads, and for every cell of each table's
TableRange2, for the table's whole TableRange1 and for the cell just right of the table, reads LocationInTable
and PivotCell.PivotCellType, an error's number where one is raised.

    python scripts/measure_location_in_table.py

writes tests/fixtures/location_in_table/, a JSON file for each workbook read, which
tests/test_excel_location_in_table.py replays.
"""

from __future__ import annotations

import json
import shutil
import tempfile
from pathlib import Path

from measure_pivot_cells import FILES
from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / "tests" / "fixtures"
OUT = FIXTURES / "location_in_table"

DUMP = '''Private Function Where(c As Object) As String
    Dim kind As String
    On Error Resume Next
    kind = c.PivotCell.PivotCellType
    If Err.Number <> 0 Then kind = "E" & Err.Number
    Err.Clear
    Where = c.LocationInTable
    If Err.Number <> 0 Then Where = "E" & Err.Number
    Where = Where & "|" & kind
End Function

Public Function Dump(path As String) As String
    Dim wb As Object, ws As Object, pt As Object, c As Object, whole As Object
    Dim tables() As String, cells() As String, n As Long, k As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path)
    ReDim tables(0 To 0)
    For Each ws In wb.Worksheets
        For Each pt In ws.PivotTables
            Set whole = pt.TableRange2
            ReDim cells(0 To whole.Cells.Count + 1)
            k = 0
            For Each c In whole.Cells
                cells(k) = c.Address(False, False) & "=" & Where(c)
                k = k + 1
            Next
            cells(k) = "whole=" & Where(pt.TableRange1)
            cells(k + 1) = "beside=" & Where(ws.Cells(whole.Row, whole.Column + whole.Columns.Count))
            ReDim Preserve tables(0 To n)
            tables(n) = ws.Name & "!" & pt.Name & "^" & Join(cells, "~")
            n = n + 1
        Next
    Next
    wb.Close False
    Dump = Join(tables, "`")
End Function
'''


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name in FILES:
            # Excel reads a copy, so that nothing it does touches the fixture.
            copy = Path(folder) / name.replace("/", "__")
            shutil.copy(FIXTURES / name, copy)
            result = excel.run_vba(DUMP, "Dump", (str(copy),), timeout=900.0)
            assert result.ok, f"{name}: {result.outcome}: {result.message} {result.error}"
            tables = dict(one.split("^", 1) for one in str(result.value).split("`") if "^" in one)
            record = {"file": name, "dump": DUMP,
                      "tables": {key: cells.split("~") for key, cells in tables.items()}}
            target = OUT / (Path(name).stem + ".json")
            target.write_text(json.dumps(record, indent=0, ensure_ascii=False) + "\n", encoding="utf-8")
            print(name, len(tables), "tables", sum(len(cells.split("~")) for cells in tables.values()), "cells")


if __name__ == "__main__":
    main()
