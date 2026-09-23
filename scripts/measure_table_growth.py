"""When a table grows on its own, and which references grow with it.

Excel's AutoCorrect widens a table when a value lands in the row just
under it or the column just right of it -- from VBA too -- and a
reference that ran to the table's last row can run on over the new one.
Each case below starts from Table1 over A1:C4 (Name, Qty, Price) with
formulas in column H naming parts of it by address, writes one thing in
live Excel, and reads back the table's range and those formulas.

    python scripts/measure_table_growth.py

writes tests/fixtures/tables/growth.json, which
tests/test_excel_table_growth.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "tables" / "growth.json"

SETUP = "\n".join([
    'ws.Range("A1:C1").Value = Array("Name", "Qty", "Price")',
    'ws.Range("A2:C2").Value = Array("apple", 3, 1.5)',
    'ws.Range("A3:C3").Value = Array("pear", 5, 2.25)',
    'ws.Range("A4:C4").Value = Array("plum", 7, 0.5)',
    'Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C4"), , xlYes)',
])

#: Formulas in column H, each naming part of the table or its surroundings by address.
REFERENCES: list[str] = [
    "=SUM(C2:C4)", "=SUM(C3:C4)", "=SUM(C1:C4)", "=COUNTA(A1:C4)", "=SUM($C$2:$C$4)", "=C4", "=SUM(C2:D4)",
    "=SUM(C2:C3)", "=C5", "=SUM(C2:C5)", "=SUM(B2:C4)", "=COUNTA(A2:A4)", "=SUM(C:C)", "=SUM(D2:D4)",
]

#: Each case: its name and what it writes on ws.
CASES: list[tuple[str, str]] = [
    ("none", ""),
    ("value_below", 'ws.Range("A5").Value = "kiwi"'),
    ("number_below_last_column", 'ws.Range("C5").Value = 9'),
    ("formula_below", 'ws.Range("B5").Formula = "=1+1"'),
    ("empty_below", 'ws.Range("A5").Value = ""'),
    ("row_below", 'ws.Range("A5:C5").Value = Array("kiwi", 9, 4)'),
    ("two_rows_below", 'ws.Range("A5:C6").Value = 1'),
    ("below_outside", 'ws.Range("D5").Value = 1'),
    ("two_below", 'ws.Range("A6").Value = "gap"'),
    ("value_right", 'ws.Range("D2").Value = 1'),
    ("header_right", 'ws.Range("D1").Value = "Extra"'),
    ("number_header_right", 'ws.Range("D1").Value = 5'),
    ("column_right", 'ws.Range("D1:D4").Value = 1'),
    ("right_outside", 'ws.Range("D5").Value = 1'),
    ("below_with_totals", 't.ShowTotals = True\nws.Range("A7").Value = "kiwi"'),
    ("under_totals", 't.ShowTotals = True\nws.Range("A6").Value = "kiwi"'),
    ("cells_below", 'ws.Cells(5, 1).Value = "kiwi"\nws.Cells(5, 2).Value = 9'),
    ("copy_below", 'ws.Range("A4:C4").Copy ws.Range("A5")'),
    ("fill_down", 'ws.Range("A4:C5").FillDown'),
    ("calculated_then_below", 't.ListColumns.Add\nt.ListColumns(4).DataBodyRange.Formula = "=[@Qty]*2"\n'
                              'ws.Range("A5").Value = "kiwi"'),
    ("second_table_below", 'ws.Range("A6:B6").Value = Array("Key", "Value")\nws.Range("A7:B7").Value = Array("a", 1)\n'
                           'ws.ListObjects.Add xlSrcRange, ws.Range("A6:B7"), , xlYes\nws.Range("A5").Value = "kiwi"'),
]

HELPERS = """Private Function Read(ws As Object) As String
    Dim k As Long, out As String
    out = ws.ListObjects(1).Range.Address & "~;~"
    For k = 1 To COUNT
        out = out & ws.Cells(k, 8).Formula & "~;~"
    Next
    out = out & ws.Range("D1").Formula & "~;~" & ws.Range("D5").Formula
    Read = out
End Function
""".replace("COUNT", str(len(REFERENCES)))


def module() -> str:
    lines = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, t As Object, out As String",
             "Application.DisplayAlerts = False"]
    for name, writing in CASES:
        lines += ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", SETUP]
        lines += [f'ws.Cells({row}, 8).Formula = "{formula}"' for row, formula in enumerate(REFERENCES, 1)]
        lines += ["Err.Clear", "On Error Resume Next", writing,
                  f'out = out & "{name}~:~" & Err.Number & "~;~" & Read(ws) & "~|~"', "On Error GoTo 0",
                  "wb.Close False"]
    lines += ["Probe = out", "End Function"]
    return HELPERS + "\n".join(lines) + "\n"


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = dict(part.split("~:~", 1) for part in str(result.value).split("~|~") if "~:~" in part)
    record = {"setup": SETUP, "references": REFERENCES, "reader": HELPERS,
              "cases": [{"name": name, "writing": writing, "answer": answers[name]} for name, writing in CASES]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for case in record["cases"]:
        print(f"{case['name']:24} {case['answer'].replace('~;~', ' | ')}")


if __name__ == "__main__":
    main()
