"""What renaming a sheet does in a workbook that holds a table: its formulas, its table's own, and the file.

Every case makes a new workbook: a table Sales on Sheet1 over A1:C4, headers H1 to H3 and numbers below, and Sheet2
beside it with numbers in A1 and B1; then a setup, then a rename. The macro reads, after the rename, the sheets'
names, the table's name and range, and the formula and value of each cell CASES names, and saves the workbook, whose
table part's calculated-column and totals-row formulas and workbook part's defined names the record keeps.

    python scripts/measure_rename_tables.py

writes tests/fixtures/rename_tables.json, which tests/test_excel_rename_tables.py replays.
"""

from __future__ import annotations

import html
import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "rename_tables.json"
#: Each case: its setup, the rename, and the cells read after it.
CASES: dict[str, tuple[list[str], str, list[str]]] = {
    "own_sheet": ([], 'ws.Name = "Main"', ["ws.Range(\"A2\")"]),
    "calculated_other": (['ws.Range("C2").Formula = "=Sheet2!$A$1*[@H1]"'], 'other.Name = "Data 2"',
                         ['ws.Range("C2")', 'ws.Range("C4")']),
    "calculated_own": (['ws.Range("C2").Formula = "=Sheet1!$A$2+[@H2]"'], 'ws.Name = "Main"',
                       ['ws.Range("C2")', 'ws.Range("C3")']),
    "totals_other": (["t.ShowTotals = True", 'ws.Range("B5").Formula = "=SUBTOTAL(109,[H2])+Sheet2!$B$1"'],
                     'other.Name = "Data 2"', ['ws.Range("B5")']),
    "formula_to_table_sheet": (['other.Range("A5").Formula = "=Sheet1!A2+SUM(Sales[H1])"'], 'ws.Name = "Main"',
                               ['other.Range("A5")']),
    "name_to_table": (['wb.Names.Add Name:="Total", RefersTo:="=SUM(Sales[H1])+Sheet1!$A$2"',
                       'other.Range("A6").Formula = "=Total"'], 'ws.Name = "Main"', ['other.Range("A6")']),
    "quoted_names": (['ws.Range("C2").Formula = "=Sheet2!$A$1+[@H1]"', "t.ShowTotals = True",
                      'ws.Range("A5").Formula = "=SUBTOTAL(109,[H1])+Sheet2!$A$1"'], "other.Name = \"It's\"",
                     ['ws.Range("C2")', 'ws.Range("A5")']),
}
HELPER = '''Private Function Shown(c As Object) As String
    If IsError(c.Value) Then
        Shown = c.Formula & "=" & "E" & CStr(CLng(c.Value))
    Else
        Shown = c.Formula & "=" & CStr(c.Value)
    End If
End Function
'''


def case_code(name: str) -> str:
    setup, rename, cells = CASES[name]
    lines = [f"Public Function Case_{name}(path As String) As String",
             "    Dim wb As Object, ws As Object, other As Object, t As Object, out As String",
             "    Set wb = Workbooks.Add(xlWBATWorksheet)",
             "    Set ws = wb.Worksheets(1)",
             "    Set other = wb.Worksheets.Add(After:=ws)",
             "    other.Range(\"A1\").Value = 10",
             "    other.Range(\"B1\").Value = 5",
             '    ws.Range("A1:C1").Value = Array("H1", "H2", "H3")',
             '    ws.Range("A2:B4").Value = 2',
             '    Set t = ws.ListObjects.Add(1, ws.Range("A1:C4"), , 1)',
             '    t.Name = "Sales"',
             *(f"    {line}" for line in setup),
             f"    {rename}",
             '    out = ws.Name & "|" & other.Name & "|" & t.Name & "|" & t.Range.Address',
             *(f'    out = out & "|" & Shown({cell})' for cell in cells),
             "    Application.DisplayAlerts = False",
             "    wb.SaveAs path, 51",
             "    wb.Close False",
             f"    Case_{name} = out",
             "End Function"]
    return "\n".join(lines) + "\n"


def saved_parts(path: Path) -> dict[str, list[str]]:
    """The saved workbook's table formulas, calculated and totals, and its defined names, as the file spells them."""
    with zipfile.ZipFile(path) as package:
        table = package.read("xl/tables/table1.xml").decode("utf-8")
        workbook = package.read("xl/workbook.xml").decode("utf-8")
    return {"calculated": [html.unescape(one) for one in
                           re.findall(r"<calculatedColumnFormula>(.*?)</calculatedColumnFormula>", table)],
            "totals": [html.unescape(one) for one in re.findall(r"<totalsRowFormula>(.*?)</totalsRowFormula>", table)],
            "names": [html.unescape(one) for one in re.findall(r"<definedName\b[^>]*>(.*?)</definedName>", workbook)]}


#: The workbook Excel saves with a table before a rename, which the loaded case opens again.
LOADED = OUT.parent / "rename_tables_loaded.xlsx"
LOADED_CODE = '''Public Function Make(path As String) As String
    Dim wb As Object, ws As Object, other As Object, t As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    Set other = wb.Worksheets.Add(After:=ws)
    other.Range("A1").Value = 10
    other.Range("B1").Value = 5
    ws.Range("A1:C1").Value = Array("H1", "H2", "H3")
    ws.Range("A2:B4").Value = 2
    Set t = ws.ListObjects.Add(1, ws.Range("A1:C4"), , 1)
    t.Name = "Sales"
    ws.Range("C2").Formula = "=Sheet2!$A$1*[@H1]"
    t.ShowTotals = True
    ws.Range("B5").Formula = "=SUBTOTAL(109,[H2])+Sheet2!$B$1"
    Application.DisplayAlerts = False
    wb.SaveAs path, 51
    wb.Close False
    Make = "made"
End Function

Public Function Rename(path As String, saved As String) As String
    Dim wb As Object, ws As Object, other As Object, out As String
    Set wb = Workbooks.Open(path)
    Set ws = wb.Worksheets(1)
    Set other = wb.Worksheets(2)
    other.Name = "Data 2"
    out = ws.Name & "|" & other.Name & "|" & ws.ListObjects(1).Name & "|" & ws.ListObjects(1).Range.Address
    out = out & "|" & ws.Range("C2").Formula & "|" & ws.Range("C4").Formula & "|" & ws.Range("B5").Formula
    out = out & "|" & ws.Range("B5").Value
    Application.DisplayAlerts = False
    wb.SaveAs saved, 51
    wb.Close False
    Rename = out
End Function
'''


def main() -> None:
    from fixture_workbook import copy_saved

    record: dict[str, object] = {"helper": HELPER, "cases": {}}
    cases: dict[str, object] = {}
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name in CASES:
            path = Path(folder) / f"{name}.xlsx"
            source = HELPER + case_code(name)
            result = excel.run_vba(source, f"Case_{name}", [str(path)], timeout=300.0)
            assert result.ok, f"{name}: {result.outcome}: {result.message} {result.error}"
            cases[name] = {"code": source, "cells": CASES[name][2], "answer": str(result.value),
                           "saved": saved_parts(path)}
            print(name, ascii(result.value), ascii(cases[name]["saved"]))  # type: ignore[index]
        made = Path(folder) / "made.xlsx"
        result = excel.run_vba(LOADED_CODE, "Make", [str(made)], timeout=300.0)
        assert result.ok, f"make: {result.outcome}: {result.message} {result.error}"
        copy_saved(made, LOADED)
        renamed = Path(folder) / "renamed.xlsx"
        result = excel.run_vba(LOADED_CODE, "Rename", [str(LOADED), str(renamed)], timeout=300.0)
        assert result.ok, f"rename: {result.outcome}: {result.message} {result.error}"
        record["loaded"] = {"code": LOADED_CODE, "answer": str(result.value), "saved": saved_parts(renamed)}
        print("loaded", ascii(result.value), ascii(record["loaded"]))
    record["cases"] = cases
    OUT.write_text(json.dumps(record, indent=1, ensure_ascii=True) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
