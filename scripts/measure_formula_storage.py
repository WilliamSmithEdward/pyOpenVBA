"""How Excel stores formulas in a saved sheet: which it shares, which are arrays, and what an edit does to them.

Excel saves a formula written to several cells at once as a shared
formula: the group's top-left cell carries the text and the group's
range, and every other cell only the group's index. An array formula is
kept in its block's first cell with the block beside it, the other cells
holding only their values. Each case below writes formulas from VBA, or
edits one of those workbooks, and saves the result; the formula cells
Excel wrote, and every formula Excel reads back after opening the file
again, go into tests/fixtures/formula_storage/.

    python scripts/measure_formula_storage.py

writes tests/fixtures/formula_storage/formula_storage.json and one .xlsx
per case, which tests/test_excel_formula_storage.py replays.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from fixture_workbook import strip_save_path

ROOT = Path(__file__).resolve().parent.parent
FOLDER = ROOT / "tests" / "fixtures" / "formula_storage"

#: The workbooks written from nothing: name, and what is written to a fresh sheet.
WRITTEN: list[tuple[str, str]] = [
    ("written", 'ws.Range("A1:A4").Value = Application.Transpose(Array(1, 2, 3, 4))\n'
                'ws.Range("B1:B3").Formula = "=A1*2"\n'
                'ws.Range("C1").Formula = "=A1*3"\n'
                'ws.Range("C1").AutoFill ws.Range("C1:C3")\n'
                'ws.Range("G1").Formula = "=A1"\n'
                'ws.Range("G2").Formula = "=A2"\n'
                'ws.Range("H1:I2").Formula = "=$A1+A$1"'),
    ("kinds", 'ws.Range("A1:A4").Value = Application.Transpose(Array(1, 2, 3, 4))\n'
              'ws.Range("K1:K2").Formula = "=5"\n'
              'ws.Range("L1:L2").FormulaR1C1 = "=RC1*2"\n'
              'ws.Range("M1:M2").Value = "=A1"\n'
              'ws.Range("O1").Formula = "=A1"\n'
              'ws.Range("O1:O3").FillDown\n'
              'ws.Range("P1").Formula = "=A1"\n'
              'ws.Range("P1").Copy ws.Range("P2:P3")\n'
              'ws.Range("Q1:Q3").Formula = "=SUM($A$1:$A$4)"\n'
              'ws.Range("R1:R2").Formula = "=A1&""x"""\n'
              'ws.Range("S1").Formula = "=""a""""b<&>"""'),
    ("arrays", 'ws.Range("A1:A4").Value = Application.Transpose(Array(1, 2, 3, 4))\n'
               'ws.Range("D1").FormulaArray = "=SUM(A1:A3*2)"\n'
               'ws.Range("E1:E3").FormulaArray = "=A1:A3*10"\n'
               'ws.Range("F1:F4").FormulaArray = "=A1:A3*10"\n'
               'ws.Range("G1:H2").FormulaArray = "={1,2}"\n'
               'ws.Range("I1").FormulaArray = "=A1:A3&""x"""'),
]

#: Edits made to one of those workbooks before it is saved again: name, the workbook, the edit.
EDITS: list[tuple[str, str, str]] = [
    ("value_edit", "written", 'ws.Range("A2").Value = 20'),
    ("split_formula", "written", 'ws.Range("B2").Formula = "=A2*5"'),
    ("split_clear", "written", 'ws.Range("B2").ClearContents'),
    ("master_formula", "written", 'ws.Range("B1").Formula = "=A1*7"'),
    ("master_clear", "written", 'ws.Range("B1").ClearContents'),
    ("insert_row", "written", "ws.Rows(2).Insert"),
    ("delete_row", "written", "ws.Rows(2).Delete"),
    ("rewrite", "written", 'ws.Range("B1:B3").Formula = "=A1*2"'),
    ("extend_fill", "written", 'ws.Range("B3").AutoFill ws.Range("B3:B4")'),
    ("array_value_edit", "arrays", 'ws.Range("A2").Value = 20'),
    ("array_cleared", "arrays", 'ws.Range("E1:E3").ClearContents'),
    ("array_insert_row", "arrays", "ws.Rows(1).Insert"),
]

#: Every formula cell of the block the cases use, read back after the file is opened again.
READ_BACK = """Private Function Formulas(ws As Object) As String
Dim cell As Object, out As String
For Each cell In ws.Range("A1:S5").Cells
    If cell.HasFormula Then out = out & cell.Address(False, False) & "=" & cell.Formula & "=" & cell.Text & ";"
Next
Formulas = out
End Function
"""


def module() -> str:
    lines = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False"]
    for name, writes in WRITTEN:
        path = FOLDER / f"{name}.xlsx"
        lines += ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", writes,
                  f'wb.SaveAs Filename:="{path}", FileFormat:=51', "wb.Close False"]
    for name, base, edit in EDITS:
        path = FOLDER / f"{name}.xlsx"
        lines += [f'Set wb = Workbooks.Open("{FOLDER / (base + ".xlsx")}")', "Set ws = wb.Worksheets(1)", edit,
                  f'wb.SaveAs Filename:="{path}", FileFormat:=51', "wb.Close False"]
    for name in [name for name, _ in WRITTEN] + [name for name, _, _ in EDITS]:
        lines += [f'Set wb = Workbooks.Open("{FOLDER / (name + ".xlsx")}")',
                  f'out = out & "{name}^" & Formulas(wb.Worksheets(1)) & "|"', "wb.Close False"]
    lines += ["Probe = out", "End Function"]
    return READ_BACK + "\n".join(lines) + "\n"


def formula_cells(path: Path) -> list[str]:
    """The cell elements holding a formula, as Excel wrote them, in order."""
    with zipfile.ZipFile(path) as package:
        xml = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    return re.findall(r"<c\b[^>]*>(?:(?!</c>).)*<f\b(?:(?!</c>).)*</c>", xml, re.DOTALL)


def main() -> None:
    FOLDER.mkdir(parents=True, exist_ok=True)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    for name in [name for name, _ in WRITTEN] + [name for name, _, _ in EDITS]:
        strip_save_path(FOLDER / f"{name}.xlsx")
    read = dict(part.split("^", 1) for part in str(result.value).split("|") if part)
    cases = []
    for name, base, action in [(name, "", action) for name, action in WRITTEN] + EDITS:
        formulas = [one.split("=", 1) for one in read[name].split(";") if one]
        cases.append({"name": name, "base": base, "action": action, "cells": formula_cells(FOLDER / f"{name}.xlsx"),
                      "read": {address: rest for address, rest in formulas}})
    record = {"reader": READ_BACK, "cases": cases}
    (FOLDER / "formula_storage.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for case in cases:
        print(f"--- {case['name']}: {case['action'].splitlines()[0]}")
        for cell in case["cells"]:
            print("   ", re.sub(r"<v>.*?</v>", "", cell))


if __name__ == "__main__":
    main()
