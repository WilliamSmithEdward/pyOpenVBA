"""What Range.PivotCell, PivotTable, PivotField and PivotItem answer for every cell of a pivot table, in live Excel.

Excel opens, read only, workbooks of pivot tables it saved for other
fixtures: every layout of tests/fixtures/pivot_layouts, the options,
calculated items, groupings, collapsed items, captions, filters and
sorts, and the small native files. For every cell of each table's
TableRange2 the macro reads its PivotCell's PivotCellType, PivotTable,
PivotField, PivotItem, DataField, RowItems, ColumnItems,
CustomSubtotalFunction and Range, and the range's own PivotField,
PivotItem and PivotTable, an error's number where one is raised; and the
same for the table's whole TableRange1 and for the cell just right of
the table. PivotRowLine and PivotColumnLine are left alone: reading the
LineType of a row label's PivotColumnLine stops Excel (RPC 0x800706BE,
pivots/native-multiple-data.xlsx).

    python scripts/measure_pivot_cells.py

writes tests/fixtures/pivot_cells/, a JSON file for each workbook read,
which tests/test_excel_pivot_cells.py replays.
"""

from __future__ import annotations

import json
import shutil
import tempfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / "tests" / "fixtures"
OUT = FIXTURES / "pivot_cells"

#: The workbooks read, by their place under tests/fixtures.
FILES = [
    "pivot_layouts/pivot_layouts.xlsx",
    "pivot_options/pivot_options.xlsx",
    "pivot_calculated_items/pivot_calculated_items.xlsx",
    "pivot_grouping/pivot_grouping.xlsx",
    "pivot_show_detail/pivot_show_detail.xlsx",
    "pivot_names/pivot_names.xlsx",
    "pivot_filters/pivot_filters.xlsx",
    "pivot_sorts/pivot_sorts.xlsx",
    "pivots/native-advanced-compact.xlsx",
    "pivots/native-advanced-date-groups.xlsx",
    "pivots/native-advanced-filtered-page.xlsx",
    "pivots/native-advanced-manual-group.xlsx",
    "pivots/native-advanced-no-totals.xlsx",
    "pivots/native-advanced-numeric-groups.xlsx",
    "pivots/native-advanced-outline.xlsx",
    "pivots/native-advanced-renamed.xlsx",
    "pivots/native-advanced-tabular.xlsx",
    "pivots/native-column.xlsx",
    "pivots/native-functions.xlsx",
    "pivots/native-multiple-data.xlsx",
    "pivots/native-page.xlsx",
    "pivots/native-shared-cache.xlsx",
    "pivots/native-simple.xlsx",
]

DUMP = '''Private Function Got(target As Object, name As String) As String
    On Error GoTo Failed
    Got = CStr(CallByName(target, name, VbGet))
    Exit Function
Failed:
    Got = "E" & Err.Number
End Function

Private Function NameOf(target As Object, name As String) As String
    Dim found As Object
    On Error GoTo Failed
    Set found = CallByName(target, name, VbGet)
    NameOf = found.Name
    Exit Function
Failed:
    NameOf = "E" & Err.Number
End Function

Private Function ListOf(target As Object, name As String) As String
    Dim found As Object, one As Object, out As String
    On Error GoTo Failed
    Set found = CallByName(target, name, VbGet)
    out = found.Count & ":"
    For Each one In found
        out = out & one.Name & ";"
    Next
    ListOf = out
    Exit Function
Failed:
    ListOf = out & "E" & Err.Number
End Function

Private Function PlaceOf(pc As Object) As String
    On Error GoTo Failed
    PlaceOf = pc.Range.Address(False, False)
    Exit Function
Failed:
    PlaceOf = "E" & Err.Number
End Function

Private Function CellOf(c As Object) As String
    Dim pc As Object, own As String
    own = NameOf(c, "PivotField") & "|" & NameOf(c, "PivotItem") & "|" & NameOf(c, "PivotTable")
    On Error Resume Next
    Set pc = c.PivotCell
    If Err.Number <> 0 Then
        CellOf = "E" & Err.Number & "|" & own
        Exit Function
    End If
    On Error GoTo 0
    CellOf = Got(pc, "PivotCellType") & "|" & NameOf(pc, "PivotTable") & "|" & NameOf(pc, "PivotField") & "|" & _
        NameOf(pc, "PivotItem") & "|" & NameOf(pc, "DataField") & "|" & ListOf(pc, "RowItems") & "|" & _
        ListOf(pc, "ColumnItems") & "|" & Got(pc, "CustomSubtotalFunction") & "|" & PlaceOf(pc) & "|" & own
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
                cells(k) = c.Address(False, False) & "=" & CellOf(c)
                k = k + 1
            Next
            cells(k) = "whole=" & CellOf(pt.TableRange1)
            cells(k + 1) = "beside=" & CellOf(ws.Cells(whole.Row, whole.Column + whole.Columns.Count))
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
