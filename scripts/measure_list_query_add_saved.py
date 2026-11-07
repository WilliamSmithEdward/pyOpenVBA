"""Do books the model saved with a ListObject over a database open in Excel, and does Excel's own refresh of them work? The model makes
each book as a case of scripts/measure_list_query_add.py makes it, its Refresh function reading the rows Excel would, over an Access
file made for the run through Excel's own ACE OLEDB provider. The model's connection string is the one the macro set, where Excel saves
its provider's fuller one. No network source is ever named, and nothing here touches encryption, signing or a certificate.

Prediction: each opens with no repair, its table and its query table as the model wrote them, and Excel's refresh for four rows gives
the rows it gave in the books it made (tests/fixtures/list_query_add.json).

Second round, written before it ran: Excel's own books with a ListObject over a database (tests/fixtures/query_refresh_loaded/), the
table deleted or unlisted and saved, keep no connections part, no queryTable part and no ExternalData_1 name, as a table added in the
session drops all three; an unlisted table's cells stay, its header bold.

    python scripts/measure_list_query_add_saved.py

writes tests/fixtures/list_query_add_saved.json.
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from measure_list_query_add import CASE, CASES, FOLDER  # noqa: E402
from measure_query_refresh import PROVIDER  # noqa: E402
from measure_query_refresh_layout import MAKE  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tests/fixtures/list_query_add_saved.json"
LOADED_BOOKS = ROOT / "tests/fixtures/query_refresh_loaded"
#: Each book: the case that makes it, the copy it saves, and the command a check refreshes it with.
BOOKS = {
    "save": ("save", "", "SELECT * FROM T WHERE Id <= 4 ORDER BY Id"),
    "save_unrefreshed": ("save_unrefreshed", "", "SELECT * FROM T WHERE Id <= 4 ORDER BY Id"),
    "ids_three": ("ids", "6", "SELECT Id, Name, Id * 2 AS D FROM T WHERE Id <= 4 ORDER BY Id"),
}
CHECK = CASE.replace("{provider}", PROVIDER).replace("{steps}", "") + r'''
Function Check(ByVal path As String, ByVal command As String) As String
    Dim wb As Workbook, ws As Worksheet, q As QueryTable, s As String
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(path, UpdateLinks:=0)
    If Err.Number <> 0 Then Check = "open E" & Err.Number & " " & Err.Description: Exit Function
    Set ws = wb.Worksheets(1)
    s = "opened " & Info(wb, ws)
    Err.Clear
    Set q = ws.ListObjects(1).QueryTable
    q.CommandText = command
    q.Refresh False
    s = s & vbTab & "four E" & Err.Number & " " & Info(wb, ws)
    wb.Close False
    Check = s
End Function
'''


#: Excel's own books with a ListObject over a database (tests/fixtures/query_refresh_loaded/), each table deleted or unlisted and the
#: book saved: what the file keeps of the query table, its connection and its name.
LOADED = {f"{book}_{how}": (book, how) for book in ("access_table", "csv_table") for how in ("delete", "unlist")}
DROP = CHECK + r'''
Function Drop(ByVal path As String, ByVal out As String, ByVal how As String) As String
    Dim wb As Workbook, ws As Worksheet, s As String
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(path, UpdateLinks:=0)
    If Err.Number <> 0 Then Drop = "open E" & Err.Number & " " & Err.Description: Exit Function
    Set ws = wb.Worksheets(1)
    s = "opened " & Info(wb, ws)
    Err.Clear
    If how = "delete" Then ws.ListObjects(1).Delete Else ws.ListObjects(1).Unlist
    s = s & vbTab & how & " E" & Err.Number & " " & Info(wb, ws)
    Err.Clear
    wb.SaveAs out, 51
    s = s & vbTab & "saved E" & Err.Number
    wb.Close False
    Drop = s
End Function
'''


def kept(data: bytes) -> dict[str, object]:
    """What a save keeps that says what became of the table: its parts' names, the defined names and the content types."""
    import io
    import zipfile

    with zipfile.ZipFile(io.BytesIO(data)) as package:
        names = sorted(package.namelist())
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        types = package.read("[Content_Types].xml").decode("utf-8")
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    defined = re.search(r"<definedNames>.*?</definedNames>", workbook)
    return {"parts": names, "names": defined.group(0) if defined else "", "types": sorted(re.findall(r'PartName="([^"]*)"', types)),
            "tableParts": re.findall(r"<tableParts\b.*?(?:</tableParts>|/>)", sheet)}


def table_t(request: object) -> list[list[object]]:
    """The rows of T the command asks for, as ACE gives them: a table's name gives the whole table, and SQL text runs as SQL whatever
    CommandType says (tests/fixtures/list_query_add.json, table_then_sql)."""
    source = getattr(request, "source")
    if source.command_text == "T":
        fields, last = ["Id", "Name"], 6
    else:
        found = re.fullmatch(r"SELECT (.+) FROM T WHERE Id <= (\d+) ORDER BY Id", source.command_text)
        assert found is not None, source.command_text
        fields = ["Id", "Name"] if found.group(1) == "*" else ["D" if one.strip().startswith("Id *") else one.strip() for one in found.group(1).split(",")]
        last = int(found.group(2))
    return [list(fields), *[[{"Id": float(at), "Name": f"n{at}", "D": float(at * 2)}[one] for one in fields] for at in range(1, last + 1)]]


def main() -> None:
    from pyopenvba.apps.excel import ExcelApplication
    from pyvbaharness import ExcelSession, HarnessConfig

    measured: dict[str, list[str]] = {}
    with tempfile.TemporaryDirectory(prefix="pyopenvba_las_") as folder, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
        excel.new_document()
        file = str(Path(folder) / "rows.accdb")
        made = excel.run_vba(MAKE.replace("{provider}", PROVIDER), "MakeDb", (file,), timeout=120.0, line_numbers=False)
        assert made.ok, made.message
        for book, (case, copy, command) in BOOKS.items():
            app = ExcelApplication(sandbox=False)
            app.add_workbook()
            app.set_hook("Refresh", table_t)
            app.add_module(CASE.replace("{provider}", PROVIDER).replace("{steps}", "\n".join(f"    {step}" for step in CASES[case])), name="M")
            app.run("Go", file, str(Path(folder) / f"{book}~"))
            saved = Path(folder) / f"{book}~{copy}.xlsx"
            checked = excel.run_vba(CHECK, "Check", (str(saved), command), timeout=180.0, line_numbers=False)
            answer = str(checked.value) if checked.ok else f"[{checked.outcome}] {checked.message}"
            measured[book] = answer.replace(folder, FOLDER).split("\t")
            print(book, " / ".join(part[:150] for part in measured[book] if not part.startswith("|")), flush=True)
        dropped: dict[str, object] = {}
        for case, (book, how) in LOADED.items():
            given, out = Path(folder) / f"{case}.xlsx", Path(folder) / f"{case}_saved.xlsx"
            given.write_bytes((LOADED_BOOKS / f"{book}.xlsx").read_bytes())
            answer = excel.run_vba(DROP, "Drop", (str(given), str(out), how), timeout=180.0, line_numbers=False)
            text = str(answer.value) if answer.ok else f"[{answer.outcome}] {answer.message}"
            dropped[case] = {"rows": text.replace(folder, FOLDER).split("\t"), "saved": kept(out.read_bytes()) if out.exists() else None}
            print(case, text[:300].replace("\t", " / "), flush=True)
    text = json.dumps({"books": BOOKS, "cases": measured, "dropped": dropped, "measured_at": datetime.now(timezone.utc).isoformat()},
                      indent=1) + "\n"
    for leak in ("Users\\", "AppData", "William"):
        assert leak not in text, leak
    OUT.write_text(text, encoding="utf-8")


if __name__ == "__main__":
    main()
