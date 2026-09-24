"""Which formulas Excel saves as shared when they name a sheet.

A formula written to several cells at once is saved once, as a shared
formula; here formulas naming another sheet, the sheet they are on, a
sheet in quotes, or none, are written to blocks of a sheet called Data,
and cells a shared formula reads are cut to another sheet, which moves
the formula's references there. The workbook is saved and each cell's
<c> recorded, and Formula read back from each.

    python scripts/measure_shared_sheets.py

writes tests/fixtures/formula_storage/sheets.json, which
tests/test_excel_formula_storage.py replays.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "formula_storage" / "sheets.json"

SETUP = ('ws.Name = "Data"\nws.Parent.Worksheets.Add(After:=ws).Name = "Other"\n'
         'ws.Parent.Worksheets.Add(After:=ws).Name = "My Sheet"\n'
         'ws.Range("A1:A4").Value = Application.Transpose(Array(1, 2, 3, 4))')
WRITES = """ws.Range("C2:C4").Formula = "=Other!A1*2"
ws.Range("D2:D4").Formula = "=Data!A1*2"
ws.Range("E2:E4").Formula = "=A1*2"
ws.Range("F2:F4").Formula = "=SUM(Other!A1:A2)"
ws.Range("G2:G4").Formula = "='My Sheet'!A1"
ws.Range("H2:H4").Formula = "=A1+Other!B1"
ws.Range("J1:J4").Formula = "=A1*3"
ws.Range("N1:N4").Value = Application.Transpose(Array(5, 6, 7, 8))
ws.Range("M1:M4").Formula = "=N1*5"
ws.Range("A1:A2").Cut ws.Parent.Worksheets("Other").Range("D1")
ws.Range("N3:N4").Cut ws.Parent.Worksheets("Other").Range("F1")
"""


def main() -> None:
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        path = Path(folder) / "sheets.xlsx"
        code = ("Public Function Report() As String\nDim wb As Object, ws As Object, one As Range, out As String\n"
                "Application.DisplayAlerts = False\nSet wb = Workbooks.Add(xlWBATWorksheet)\nSet ws = wb.Worksheets(1)\n"
                f"{SETUP}\n{WRITES}"
                'For Each one In ws.Range("C1:M4")\n'
                '    If one.HasFormula Then out = out & one.Address(False, False) & "=" & one.Formula & "|"\n'
                "Next\n"
                f'wb.SaveAs Filename:="{path}", FileFormat:=51\nwb.Close False\nReport = out\nEnd Function\n')
        result = excel.run_vba(code, "Report", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        with zipfile.ZipFile(path) as package:
            sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    formulas = dict(one.split("=", 1) for one in str(result.value).split("|") if one)
    cells = {match.group(1): match.group(0) for match in re.finditer(r'<c r="([C-M][1-4])".*?</c>', sheet)}
    OUT.write_text(json.dumps({"setup": SETUP, "writes": WRITES, "formulas": formulas, "cells": cells}, indent=1)
                   + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
