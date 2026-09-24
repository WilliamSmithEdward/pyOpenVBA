"""Editing a table's rows and columns: ListRows, ListColumns, Resize, Unlist, Delete, calculated columns.

Each workbook below starts from the same sheet -- Table1 over A1:C4
(Name, Qty, Price), a few cells around it, and formulas in column G that
name the table -- is edited in live Excel from VBA, and is then read: the
table's range, a grid of the sheet's formulas and values, answers to a
few questions. It is saved as xlsx and the table part Excel wrote is
read out of the package.

    python scripts/measure_table_edits.py

writes tests/fixtures/tables/edits/ -- one .xlsx per workbook and
edits.json -- which tests/test_excel_table_edits.py replays.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from fixture_workbook import strip_save_path

ROOT = Path(__file__).resolve().parent.parent
FOLDER = ROOT / "tests" / "fixtures" / "tables" / "edits"

#: The sheet every workbook starts from.
SETUP = "\n".join([
    'ws.Range("A1:C1").Value = Array("Name", "Qty", "Price")',
    'ws.Range("A2:C2").Value = Array("apple", 3, 1.5)',
    'ws.Range("A3:C3").Value = Array("pear", 5, 2.25)',
    'ws.Range("A4:C4").Value = Array("plum", 7, 0.5)',
    'Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C4"), , xlYes)',
    'ws.Range("A7").Value = "below"',
    'ws.Range("E3").Value = "beside"',
    'ws.Range("G1").Formula = "=SUM(Table1[Qty])"',
    'ws.Range("G2").Formula = "=ROWS(Table1)"',
    'ws.Range("G3").Formula = "=Table1[@Price]"',
    'ws.Range("G4").Formula = "=SUM(Table1[[#All],[Price]])"',
    'ws.Range("G5").Formula = "=Table1[[#Headers],[Qty]]"',
    'ws.Range("G6").Formula = "=SUM(C2:C4)"',
])

#: Each workbook: its name and the VBA that edits it, with ws the sheet and t Table1.
BOOKS: list[tuple[str, str]] = [
    ("untouched", ""),
    ("row_added", "Set r = t.ListRows.Add\nws.Range(\"J1\").Value = r.Index"),
    ("row_added_at", "Set r = t.ListRows.Add(2)\nws.Range(\"J1\").Value = r.Index"),
    ("row_added_keep", "t.ListRows.Add AlwaysInsert:=False"),
    ("row_added_insert", "t.ListRows.Add AlwaysInsert:=True"),
    ("row_added_blocked", 'ws.Range("B5").Value = "blocked"\nt.ListRows.Add'),
    ("row_added_blocked_keep", 'ws.Range("B5").Value = "blocked"\nt.ListRows.Add AlwaysInsert:=False'),
    ("row_added_totals", "t.ShowTotals = True\nt.ListRows.Add"),
    ("row_added_values", "Set r = t.ListRows.Add\nr.Range.Value = Array(\"kiwi\", 9, 4)"),
    ("row_deleted", "t.ListRows(2).Delete"),
    ("row_deleted_last", "t.ListRows(3).Delete"),
    ("column_added", "Set c = t.ListColumns.Add\nws.Range(\"J1\").Value = c.Name & \"/\" & c.Index"),
    ("column_added_at", "Set c = t.ListColumns.Add(2)\nws.Range(\"J1\").Value = c.Name & \"/\" & c.Index"),
    ("column_added_twice", "t.ListColumns.Add\nt.ListColumns.Add"),
    ("column_deleted", "t.ListColumns(2).Delete"),
    ("resized_bigger", 'ws.Range("D1").Value = "Extra"\nws.Range("D2").Value = 1\nws.Range("A5:C5").Value = '
                       'Array("fig", 2, 1)\nt.Resize ws.Range("A1:D5")'),
    ("resized_smaller", 't.Resize ws.Range("A1:B3")'),
    ("resized_headerless", 'On Error Resume Next\nt.Resize ws.Range("A2:C4")\nws.Range("J1").Value = Err.Number'),
    ("resized_blank_header", 't.Resize ws.Range("A1:D4")'),
    ("unlisted", "t.Unlist"),
    ("deleted", "t.Delete"),
    ("calculated_one", "t.ListColumns.Add\nws.Range(\"D2\").Formula = \"=[@Qty]*2\""),
    ("calculated_body", "t.ListColumns.Add\nt.ListColumns(4).DataBodyRange.Formula = \"=[@Qty]*[@Price]\""),
    ("calculated_then_row", "t.ListColumns.Add\nt.ListColumns(4).DataBodyRange.Formula = \"=[@Qty]*2\"\n"
                            "t.ListRows.Add"),
    ("calculated_broken", "t.ListColumns.Add\nt.ListColumns(4).DataBodyRange.Formula = \"=[@Qty]*2\"\n"
                          "ws.Range(\"D3\").Value = 100"),
    ("calculated_over_values", 'ws.Range("C3").Formula = "=[@Qty]*10"'),
    ("typed_below", 'ws.Range("A5").Value = "kiwi"'),
    ("typed_right", 'ws.Range("D1").Value = "Extra"\nws.Range("D2").Value = 1'),
]

#: What is read from each workbook once it is edited.
QUESTIONS: list[str] = [
    "TableRange(ws)",
    'Grid(ws, "A1:G8")',
    'ws.Range("J1").Value',
    "Names(ws)",
]

HELPERS = """Private Function TableRange(ws As Object) As String
    If ws.ListObjects.Count = 0 Then
        TableRange = "none"
    Else
        TableRange = ws.ListObjects(1).Range.Address & " rows " & ws.ListObjects(1).ListRows.Count & _
            " totals " & ws.ListObjects(1).ShowTotals
    End If
End Function

Private Function Names(ws As Object) As String
    Dim c As Object, out As String
    If ws.ListObjects.Count = 0 Then Exit Function
    For Each c In ws.ListObjects(1).ListColumns
        out = out & c.Name & "/"
    Next
    Names = out
End Function

Private Function Grid(ws As Object, where As String) As String
    Dim r As Object, c As Object, out As String
    For Each r In ws.Range(where).Rows
        For Each c In r.Cells
            If c.HasFormula Then out = out & c.Formula & "=" Else out = out
            out = out & CStr(c.Text) & "~,~"
        Next
        out = out & "~/~"
    Next
    Grid = out
End Function
"""


def module() -> str:
    lines = ["Public Function Probe() As String", "Dim out As String", "Application.DisplayAlerts = False"]
    lines += [f'out = out & B{index}("{FOLDER / (name + ".xlsx")}")' for index, (name, _) in enumerate(BOOKS)]
    lines += ["Probe = out", "End Function"]
    for index, (name, editing) in enumerate(BOOKS):
        lines += [f"Private Function B{index}(path As String) As String",
                  "Dim wb As Object, ws As Object, t As Object, r As Object, c As Object, out As String",
                  "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", SETUP, "Err.Clear",
                  "On Error Resume Next", editing, f'If Err.Number <> 0 Then out = "{name}:error~:~" & Err.Number '
                                                  '& "~|~"', "On Error GoTo 0"]
        lines += [f'out = out & "{name}:{number}~:~" & Q{number}(ws) & "~|~"' for number in range(len(QUESTIONS))]
        lines += ["wb.SaveAs Filename:=path, FileFormat:=51", "wb.Close False", f"B{index} = out", "End Function"]
    for number, question in enumerate(QUESTIONS):
        lines += [f"Private Function Q{number}(ws As Object) As String", "On Error GoTo Bad",
                  f"Q{number} = {question}", "Exit Function", "Bad:", f'Q{number} = "!" & Err.Number', "End Function"]
    return HELPERS + "\n".join(lines) + "\n"


def parts(path: Path) -> dict[str, str]:
    """Each table part and the sheet's cells as Excel wrote them."""
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
    for name, _ in BOOKS:
        strip_save_path(FOLDER / f"{name}.xlsx")
    answers = dict(part.split("~:~", 1) for part in str(result.value).split("~|~") if "~:~" in part)
    record = {"helpers": HELPERS, "setup": SETUP, "questions": QUESTIONS,
              "books": [{"name": name, "editing": editing, "error": answers.get(f"{name}:error", ""),
                         "answers": [answers[f"{name}:{number}"] for number in range(len(QUESTIONS))],
                         "parts": parts(FOLDER / f"{name}.xlsx")} for name, editing in BOOKS]}
    (FOLDER / "edits.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for book in record["books"]:
        print(f"--- {book['name']} {book['error']}")
        for question, answer in zip(QUESTIONS, book["answers"]):
            print(f"    {question[:20]:22} {answer.replace('~,~', ' | ').replace('~/~', chr(10) + ' ' * 27)}")


if __name__ == "__main__":
    main()
