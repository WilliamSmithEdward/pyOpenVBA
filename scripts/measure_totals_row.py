"""A table's totals row: what ShowTotals and TotalsCalculation do to the sheet, the table and the file.

Each workbook below is made in live Excel from VBA -- a table over data,
its totals row shown, hidden or changed -- and asked questions about
what that did; then it is saved as xlsx and the parts Excel wrote are
read out of the package: the table part and the sheet's cells.

    python scripts/measure_totals_row.py

writes tests/fixtures/tables/totals_row/ -- one .xlsx per workbook and
totals_row.json -- which tests/test_excel_totals_row.py replays.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
FOLDER = ROOT / "tests" / "fixtures" / "tables" / "totals_row"

DATA = ('ws.Range("A1:C1").Value = Array("Name", "Qty", "Price")\n'
        'ws.Range("A2:C2").Value = Array("apple", 3, 1.5)\n'
        'ws.Range("A3:C3").Value = Array("pear", 5, 2.25)\n'
        'ws.Range("A4:C4").Value = Array("plum", 7, 0.5)\n')
TABLE = 'Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C4"), , xlYes)\n'

#: Each workbook: its name, the VBA that makes it on ws, and the expressions asked of it afterwards.
BOOKS: list[tuple[str, str, list[str]]] = [
    ("shown", DATA + TABLE + "t.ShowTotals = True",
     ['t.Range.Address', 't.TotalsRowRange.Address', 't.ListRows.Count', 'Cells3(ws, 5)',
      'Calcs(t)', 'ws.Range("A5").Formula & "/" & ws.Range("C5").FormulaR1C1']),
    ("shown_neighbours", DATA + 'ws.Range("A7").Value = "below"\nws.Range("E5").Value = "beside"\n'
                                'ws.Range("B6").Value = "under"\n' + TABLE + "t.ShowTotals = True",
     ['t.Range.Address', 'Cells3(ws, 5)', 'Cells3(ws, 6)', 'Cells3(ws, 7)', 'ws.Range("E5").Value & "/" & '
      'ws.Range("E6").Value']),
    ("shown_blocked", DATA + 'ws.Range("A5").Value = "blocked"\nws.Range("E5").Value = "beside"\n' + TABLE
     + "t.ShowTotals = True",
     ['t.Range.Address', 'Cells3(ws, 5)', 'Cells3(ws, 6)', 'ws.Range("E5").Value & "/" & ws.Range("E6").Value']),
    ("text_last", 'ws.Range("A1:C1").Value = Array("Name", "Qty", "Note")\n'
                  'ws.Range("A2:C2").Value = Array("apple", 3, "x")\nws.Range("A3:C3").Value = Array("pear", 5, "y")\n'
                  + TABLE + "t.ShowTotals = True",
     ['Cells3(ws, 4)', 'Calcs(t)']),
    ("number_first", 'ws.Range("A1:B1").Value = Array("Qty", "Price")\nws.Range("A2:B2").Value = Array(3, 1.5)\n'
                     'ws.Range("A3:B3").Value = Array(5, 2.25)\n'
                     'Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes)\nt.ShowTotals = True',
     ['Cells3(ws, 4)', 'Calcs(t)']),
    ("single", 'ws.Range("A1").Value = "Qty"\nws.Range("A2").Value = 3\nws.Range("A3").Value = 5\n'
               'Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:A3"), , xlYes)\nt.ShowTotals = True',
     ['Cells3(ws, 4)', 'Calcs(t)']),
    ("hidden_again", DATA + 'ws.Range("A7").Value = "below"\n' + TABLE + "t.ShowTotals = True\nt.ShowTotals = False",
     ['t.Range.Address', 'Cells3(ws, 5)', 'Cells3(ws, 6)', 'Cells3(ws, 7)', 'Calcs(t)']),
    ("shown_again", DATA + TABLE + "t.ShowTotals = True\nt.ListColumns(2).TotalsCalculation = xlTotalsCalculationMax\n"
                                   "t.ShowTotals = False\nt.ShowTotals = True",
     ['Cells3(ws, 5)', 'Calcs(t)']),
    ("functions", 'ws.Range("A1:K1").Value = Array("Label", "n0", "n1", "n2", "n3", "n4", "n5", "n6", "n7", "n8", '
                  '"n9")\n'
                  'ws.Range("A2:K2").Value = Array("a", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)\n'
                  'ws.Range("A3:K3").Value = Array("b", 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)\n'
                  'Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:K3"), , xlYes)\nt.ShowTotals = True\n'
                  'Dim k As Long\nFor k = 0 To 8\nt.ListColumns(k + 2).TotalsCalculation = k\nNext\n'
                  'On Error Resume Next\nt.ListColumns(11).TotalsCalculation = 9\n'
                  'ws.Range("M1").Value = Err.Number\nOn Error GoTo 0\n'
                  't.ListColumns(1).TotalsCalculation = xlTotalsCalculationCount',
     ['Calcs(t)', 'Formulas(ws, 4, 11)', 'ws.Range("M1").Value']),
    ("totals_references", DATA + TABLE + 't.ShowTotals = True\n'
                                         'ws.Range("E1").Formula = "=Table1[#Totals]"\n'
                                         'ws.Range("E2").Formula = "=Table1[[#Totals],[Price]]"\n'
                                         'ws.Range("E3").Formula = "=ROWS(Table1[[#Data],[#Totals]])"\n'
                                         'ws.Range("E4").Formula = "=ROWS(Table1[#All])"\n'
                                         'ws.Range("E6").Formula = "=SUM(Table1[Price])"',
     ['Formulas(ws, 1, 6)', 'Values(ws, 1, 6)']),
    ("totals_references_hidden", DATA + TABLE + 't.ShowTotals = True\n'
                                                'ws.Range("E1").Formula = "=Table1[#Totals]"\n'
                                                'ws.Range("E2").Formula = "=Table1[[#Totals],[Price]]"\n'
                                                'ws.Range("E3").Formula = "=ROWS(Table1[[#Data],[#Totals]])"\n'
                                                't.ShowTotals = False',
     ['Formulas(ws, 1, 3)', 'Values(ws, 1, 3)']),
    ("label_written", DATA + TABLE + 't.ShowTotals = True\nws.Range("A5").Value = "Sum"\n'
                                     'ws.Range("B5").Formula = "=SUBTOTAL(101,Table1[Qty])"',
     ['Cells3(ws, 5)', 'Calcs(t)']),
    ("custom_written", DATA + TABLE + 't.ShowTotals = True\nws.Range("B5").Formula = "=SUM(Table1[Qty])*2"\n'
                                      'ws.Range("C5").Value = 42',
     ['Cells3(ws, 5)', 'Calcs(t)']),
    ("cleared", DATA + TABLE + 't.ShowTotals = True\nws.Range("A5").ClearContents\nws.Range("C5").ClearContents',
     ['Cells3(ws, 5)', 'Calcs(t)']),
    ("set_while_hidden", DATA + TABLE + 't.ListColumns(2).TotalsCalculation = xlTotalsCalculationSum',
     ['Calcs(t)', 't.ShowTotals', 't.Range.Address', 'Cells3(ws, 5)']),
    ("set_while_hidden_then_shown", DATA + TABLE + 'On Error Resume Next\n'
                                                   't.ListColumns(2).TotalsCalculation = xlTotalsCalculationSum\n'
                                                   'On Error GoTo 0\nt.ShowTotals = True',
     ['Calcs(t)', 'Cells3(ws, 5)']),
    ("mixed_last", 'ws.Range("A1:B1").Value = Array("Name", "Price")\nws.Range("A2:B2").Value = Array("a", 1.5)\n'
                   'ws.Range("A3:B3").Value = Array("b", "x")\nws.Range("A4:B4").Value = Array("c", 0.5)\n'
                   'Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B4"), , xlYes)\nt.ShowTotals = True',
     ['Calcs(t)', 'Cells3(ws, 5)']),
    ("text_first_row", 'ws.Range("A1:B1").Value = Array("Name", "Price")\nws.Range("A2:B2").Value = Array("a", "x")\n'
                       'ws.Range("A3:B3").Value = Array("b", 1.5)\nws.Range("A4:B4").Value = Array("c", 0.5)\n'
                       'Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B4"), , xlYes)\nt.ShowTotals = True',
     ['Calcs(t)', 'Cells3(ws, 5)']),
    ("empty_last", 'ws.Range("A1:B1").Value = Array("Name", "Price")\nws.Range("A2").Value = "a"\n'
                   'ws.Range("A3").Value = "b"\n'
                   'Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes)\nt.ShowTotals = True',
     ['Calcs(t)', 'Cells3(ws, 4)']),
    ("renamed_column", DATA + TABLE + 't.ShowTotals = True\nws.Range("C1").Value = "Cost"',
     ['Cells3(ws, 5)', 'Calcs(t)']),
]

HELPERS = """Private Function Cells3(ws As Object, r As Long) As String
    Dim c As Long, out As String
    For c = 1 To 3
        out = out & ws.Cells(r, c).Formula & "=" & TypeName(ws.Cells(r, c).Value) & ":" & _
            CStr(ws.Cells(r, c).Text) & "/"
    Next
    Cells3 = out
End Function

Private Function Calcs(t As Object) As String
    Dim c As Object, out As String
    For Each c In t.ListColumns
        out = out & c.Name & ":" & c.TotalsCalculation & "/"
    Next
    Calcs = out
End Function

Private Function Formulas(ws As Object, r As Long, last As Long) As String
    Dim k As Long, out As String
    For k = 1 To last
        out = out & ws.Cells(r, k).Formula & "/"
    Next
    Formulas = out
End Function

Private Function Values(ws As Object, r As Long, last As Long) As String
    Dim k As Long, out As String
    For k = r To last
        out = out & CStr(ws.Range("E" & k).Text) & "/"
    Next
    Values = out
End Function
"""


def module() -> str:
    lines = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, t As Object, out As String",
             "Application.DisplayAlerts = False"]
    for index, (name, _, _) in enumerate(BOOKS):
        lines += [f'out = out & B{index}("{FOLDER / (name + ".xlsx")}")']
    lines += ["Probe = out", "End Function"]
    for index, (name, making, questions) in enumerate(BOOKS):
        lines += [f"Private Function B{index}(path As String) As String",
                  "Dim wb As Object, ws As Object, t As Object, out As String",
                  "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", making]
        for number, question in enumerate(questions):
            lines += [f'out = out & "{name}:{number}~:~" & Q{index}_{number}(ws, t) & "~|~"']
        lines += ["wb.SaveAs Filename:=path, FileFormat:=51", "wb.Close False", f"B{index} = out", "End Function"]
        for number, question in enumerate(questions):
            lines += [f"Private Function Q{index}_{number}(ws As Object, t As Object) As String", "On Error GoTo Bad",
                      f"Q{index}_{number} = {question}", "Exit Function", "Bad:",
                      f'Q{index}_{number} = "!" & Err.Number', "End Function"]
    return HELPERS + "\n".join(lines) + "\n"


def parts(path: Path) -> dict[str, str]:
    """The table part and the sheet's cells as Excel wrote them."""
    with zipfile.ZipFile(path) as package:
        out = {name: package.read(name).decode("utf-8") for name in package.namelist() if name.startswith("xl/tables/")}
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        found = re.search(r"<sheetData>.*?</sheetData>|<sheetData/>", sheet, re.DOTALL)
        out["sheetData"] = found.group(0) if found else ""
        strings = "xl/sharedStrings.xml"
        out["sharedStrings"] = package.read(strings).decode("utf-8") if strings in package.namelist() else ""
    return out


def main() -> None:
    FOLDER.mkdir(parents=True, exist_ok=True)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = dict(part.split("~:~", 1) for part in str(result.value).split("~|~") if "~:~" in part)
    record = {"helpers": HELPERS,
              "books": [{"name": name, "making": making, "questions": questions,
                         "answers": [answers[f"{name}:{number}"] for number in range(len(questions))],
                         "parts": parts(FOLDER / f"{name}.xlsx")}
                        for name, making, questions in BOOKS]}
    (FOLDER / "totals_row.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for book in record["books"]:
        print(f"--- {book['name']}")
        for question, answer in zip(book["questions"], book["answers"]):
            print(f"    {question[:50]:52} {answer}")


if __name__ == "__main__":
    main()
