"""Do books the model saved with a ListObject over a database open in Excel, and does Excel's own refresh of them work? The model makes
each book as a case of scripts/measure_list_query_add.py makes it, its Refresh function reading the rows Excel would, over an Access
file made for the run through Excel's own ACE OLEDB provider. The model's connection string is the one the macro set, where Excel saves
its provider's fuller one. No network source is ever named, and nothing here touches encryption, signing or a certificate.

Prediction: each opens with no repair, its table and its query table as the model wrote them, and Excel's refresh for four rows gives
the rows it gave in the books it made (tests/fixtures/list_query_add.json).

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


def table_t(request: object) -> list[list[object]]:
    """The rows of T the command asks for, as ACE gives them."""
    source = getattr(request, "source")
    if source.command_type == "xlCmdTable":
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
    text = json.dumps({"books": BOOKS, "cases": measured, "measured_at": datetime.now(timezone.utc).isoformat()}, indent=1) + "\n"
    for leak in ("Users\\", "AppData", "William"):
        assert leak not in text, leak
    OUT.write_text(text, encoding="utf-8")


if __name__ == "__main__":
    main()
