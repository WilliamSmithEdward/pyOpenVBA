"""Excel tables (ListObjects): what the object model answers about one, and what the file keeps of it.

Each workbook below is made in live Excel from VBA -- data written, a
table added over it with ListObjects.Add, some of its settings changed --
and saved as xlsx. Opened again, every table is read back through the
object model; and the parts Excel wrote for it -- the table part, the
sheet's tableParts element, the relationship and the content type -- are
read out of the package. Error cases run on a fresh sheet.

    python scripts/measure_tables.py

writes tests/fixtures/tables/ -- one .xlsx per workbook and tables.json --
which tests/test_excel_tables.py replays.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
FOLDER = ROOT / "tests" / "fixtures" / "tables"

#: Data every workbook starts from: headers in row 1 and three rows under them.
DATA = ('ws.Range("A1:C1").Value = Array("Name", "Qty", "Price")\n'
        'ws.Range("A2:C2").Value = Array("apple", 3, 1.5)\n'
        'ws.Range("A3:C3").Value = Array("pear", 5, 2.25)\n'
        'ws.Range("A4:C4").Value = Array("plum", 7, 0.5)\n')

#: Each workbook: its name, and the VBA that makes its tables on ws, its first sheet.
BOOKS: list[tuple[str, str]] = [
    ("basic", DATA + 'ws.ListObjects.Add xlSrcRange, ws.Range("A1:C4"), , xlYes'),
    ("totals", DATA + 'Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C4"), , xlYes)\n'
                      "t.ShowTotals = True"),
    ("no_headers", 'ws.Range("A2:B2").Value = Array(1, 2)\nws.Range("A3:B3").Value = Array(3, 4)\n'
                   'ws.ListObjects.Add xlSrcRange, ws.Range("A2:B3"), , xlNo'),
    ("named", DATA + 'Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C4"), , xlYes)\n'
                     't.Name = "Sales"\nt.TableStyle = "TableStyleLight9"\nt.ShowTableStyleRowStripes = False\n'
                     "t.ShowTableStyleFirstColumn = True"),
    ("two", DATA + 'ws.ListObjects.Add xlSrcRange, ws.Range("A1:C4"), , xlYes\n'
                   'ws.Range("E1:F1").Value = Array("Key", "Value")\nws.Range("E2:F2").Value = Array("a", 1)\n'
                   'ws.ListObjects.Add xlSrcRange, ws.Range("E1:F2"), , xlYes'),
    ("guess", DATA + 'ws.ListObjects.Add xlSrcRange, ws.Range("A1:C4")'),
    ("guess_numbers", 'ws.Range("A1:B1").Value = Array(1, 2)\nws.Range("A2:B2").Value = Array(3, 4)\n'
                      'ws.ListObjects.Add xlSrcRange, ws.Range("A1:B2")'),
    ("odd_headers", 'ws.Range("A1:D1").Value = Array("Name", "", "Name", 2020)\n'
                    'ws.Range("A2:D2").Value = Array("a", 1, 2, 3)\n'
                    'ws.ListObjects.Add xlSrcRange, ws.Range("A1:D2"), , xlYes'),
]

#: What is read back from every table of a reopened workbook.
READER = """Private Function Tables(ws As Object) As String
    Dim t As Object, c As Object, out As String
    out = ws.ListObjects.Count & ";"
    For Each t In ws.ListObjects
        out = out & t.Name & "," & t.DisplayName & "," & t.Range.Address & "," & Where(t.DataBodyRange) & "," & _
            Where(t.HeaderRowRange) & "," & Where(t.TotalsRowRange) & "," & t.ListColumns.Count & "," & _
            t.ListRows.Count & "," & t.ShowTotals & "," & t.ShowHeaders & "," & t.TableStyle & "," & _
            t.ShowAutoFilter & "," & t.ShowTableStyleRowStripes & "," & t.ShowTableStyleFirstColumn & "," & _
            t.ShowTableStyleLastColumn & "," & t.ShowTableStyleColumnStripes & "," & t.Parent.Name & ","
        For Each c In t.ListColumns
            out = out & c.Index & ":" & c.Name & ":" & Where(c.DataBodyRange) & ":" & c.Range.Address & ":" & _
                TypeName(c.Range.Cells(1).Value) & "/"
        Next
        out = out & "," & t.ListRows(1).Range.Address & "," & t.ListRows(1).Index & ";"
    Next
    Tables = out
End Function

Private Function Where(r As Object) As String
    If r Is Nothing Then Where = "Nothing" Else Where = r.Address
End Function
"""

#: Questions asked of the "two" workbook's first sheet: what each answers, or the error it raises.
QUESTIONS: list[tuple[str, str]] = [
    ("list_object_inside", 'ws.Range("B3").ListObject.Name'),
    ("list_object_header", 'ws.Range("A1").ListObject.Name'),
    ("list_object_other", 'ws.Range("F2").ListObject.Name'),
    ("list_object_outside", 'TypeName(ws.Range("H9").ListObject)'),
    ("by_name", 'ws.ListObjects("Table2").Range.Address'),
    ("by_index", "ws.ListObjects(2).Name"),
    ("missing_name", 'ws.ListObjects("Nope").Name'),
    ("missing_index", "ws.ListObjects(3).Name"),
    ("column_by_name", 'ws.ListObjects(1).ListColumns("Qty").Index'),
    ("column_missing", 'ws.ListObjects(1).ListColumns("Nope").Index'),
    ("row_missing", "ws.ListObjects(1).ListRows(4).Index"),
    ("typename", "TypeName(ws.ListObjects(1)) & TypeName(ws.ListObjects) & TypeName(ws.ListObjects(1).ListColumns) & "
                 "TypeName(ws.ListObjects(1).ListColumns(1)) & TypeName(ws.ListObjects(1).ListRows) & "
                 "TypeName(ws.ListObjects(1).ListRows(1))"),
    ("overlap", 'ws.ListObjects.Add(xlSrcRange, ws.Range("B2:C3"), , xlYes).Name'),
    ("default_member", "ws.ListObjects(1)"),
    ("sheet_names", 'ActiveWorkbook.Names.Count & "/" & ws.Range("Table1").Address & "/" & '
                    'ws.Range("Table1[Qty]").Address & "/" & ws.Range("Table1[#All]").Address & "/" & '
                    'ws.Range("Table1[#Headers]").Address'),
]


def module() -> str:
    lines = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, t As Object, out As String",
             "Application.DisplayAlerts = False"]
    for name, making in BOOKS:
        path = FOLDER / f"{name}.xlsx"
        lines += ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", making,
                  f'wb.SaveAs Filename:="{path}", FileFormat:=51', "wb.Close False"]
    for name, _ in BOOKS:
        lines += [f'Set wb = Workbooks.Open("{FOLDER / (name + ".xlsx")}")',
                  f'out = out & "{name}^" & Tables(wb.Worksheets(1)) & "|"', "wb.Close False"]
    lines += [f'Set wb = Workbooks.Open("{FOLDER / "two.xlsx"}")', "Set ws = wb.Worksheets(1)"]
    lines += [f'out = out & "{name}^" & Q{index}(ws) & "|"' for index, (name, _) in enumerate(QUESTIONS)]
    lines += ["wb.Close False", "Probe = out", "End Function"]
    for index, (_, question) in enumerate(QUESTIONS):
        lines += [f"Private Function Q{index}(ws As Object) As String", "On Error GoTo Bad",
                  f"Q{index} = {question}", "Exit Function", "Bad:", f'Q{index} = "!" & Err.Number', "End Function"]
    return READER + "\n".join(lines) + "\n"


def parts(path: Path) -> dict[str, str]:
    """The parts Excel wrote for the tables: each table part, the sheet's tableParts, its rels and the types."""
    with zipfile.ZipFile(path) as package:
        out = {name: package.read(name).decode("utf-8") for name in package.namelist() if name.startswith("xl/tables/")}
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        found = re.search(r"<tableParts\b.*?</tableParts>|<tableParts\b[^>]*/>", sheet, re.DOTALL)
        out["tableParts"] = found.group(0) if found else ""
        tail = re.search(r"</sheetData>(.*)</worksheet>", sheet, re.DOTALL)
        out["sheet_tail"] = tail.group(1) if tail else ""
        rels = "xl/worksheets/_rels/sheet1.xml.rels"
        out["sheet_rels"] = package.read(rels).decode("utf-8") if rels in package.namelist() else ""
        types = package.read("[Content_Types].xml").decode("utf-8")
        out["content_types"] = " ".join(re.findall(r"<Override\b[^>]*tables/[^>]*/>", types))
    return out


def main() -> None:
    FOLDER.mkdir(parents=True, exist_ok=True)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = dict(part.split("^", 1) for part in str(result.value).split("|") if "^" in part)
    record = {"reader": READER,
              "books": [{"name": name, "making": making, "read": answers[name], "parts": parts(FOLDER / f"{name}.xlsx")}
                        for name, making in BOOKS],
              "questions": [{"name": name, "question": question, "answer": answers[name]}
                            for name, question in QUESTIONS]}
    (FOLDER / "tables.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for book in record["books"]:
        print(f"--- {book['name']}: {book['read']}")
        for key, text in book["parts"].items():
            print(f"    {key}: {text[:400]}")
    for question in record["questions"]:
        print(f"{question['name']:22} {question['answer']}")


if __name__ == "__main__":
    main()
