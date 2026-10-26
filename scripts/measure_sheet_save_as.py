"""Worksheet.SaveAs in live Excel: which sheet a text format writes, what it leaves of the workbook, and how it
saves a workbook format.

Each case makes a workbook of three sheets, Data, Other and Third, with the second active, calls SaveAs on a
sheet as the case says, and records the workbook after it -- its name, FileFormat, Saved, sheets, the active
sheet -- and its save events; the script records the bytes written, or the sheets of the package.

Predictions: A, a text format writes the sheet whose SaveAs runs, which becomes the active sheet and takes the
file's name. B, it writes the active sheet, as Workbook.SaveAs does. A workbook format saves every sheet. The save
events run as for Workbook.SaveAs.

    python scripts/measure_sheet_save_as.py

writes tests/fixtures/sheet_save_as.json, which tests/test_excel_sheet_save_as.py replays.
"""
from __future__ import annotations

import base64
import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/sheet_save_as.json"
#: Each case: its name, the file it saves as, and the SaveAs call, on ws, a sheet of wb.
CASES: list[tuple[str, str, str]] = [
    ("first_csv", "first.csv", 'wb.Worksheets("Data").SaveAs path, xlCSV'),
    ("active_csv", "active.csv", 'wb.Worksheets("Other").SaveAs path, xlCSV'),
    ("third_text", "third.txt", 'wb.Worksheets("Third").SaveAs Filename:=path, FileFormat:=xlUnicodeText'),
    ("first_xlsx", "first.xlsx", 'wb.Worksheets("Data").SaveAs path, xlOpenXMLWorkbook'),
    ("no_format", "noformat", 'wb.Worksheets("Data").SaveAs path'),
    ("named_args", "named.csv", 'wb.Worksheets("Data").SaveAs Filename:=path, FileFormat:=xlCSV, Local:=True, '
                                'AddToMru:=False, CreateBackup:=False'),
    ("utf8", "utf8.csv", 'wb.Worksheets("Third").SaveAs path, xlCSVUTF8'),
    ("msdos", "msdos.csv", 'wb.Worksheets("Data").SaveAs path, xlCSVMSDOS'),
    ("hidden_sheet", "hidden.csv", 'wb.Worksheets("Third").Visible = xlSheetHidden: '
                                   'wb.Worksheets("Third").SaveAs path, xlCSV'),
]

PROBE = r'''Private events As String

Private Function Book(wb As Workbook) As String
    Dim ws As Worksheet, s As String
    For Each ws In wb.Worksheets
        s = s & "," & ws.Name & IIf(ws.Visible = xlSheetVisible, "", "(hidden)")
    Next
    Book = wb.Name & "|" & wb.FileFormat & "|" & wb.Saved & "|" & wb.Worksheets.Count & "|" & Mid$(s, 2) & "|" & _
        wb.ActiveSheet.Name & "|" & ActiveWorkbook.Name
End Function

Public Function Probe(name As String, path As String) As String
    Dim wb As Workbook, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "Data"
    wb.Worksheets(1).Range("A1").Value = "data": wb.Worksheets(1).Range("B1").Value = 1
    wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "Other"
    wb.Worksheets("Other").Range("A1").Value = "other"
    wb.Worksheets.Add(After:=wb.Worksheets(2)).Name = "Third"
    wb.Worksheets("Third").Range("A1").Value = "third"
    wb.Worksheets("Other").Activate
    On Error Resume Next
    Select Case name
CASES
    End Select
    If Err.Number <> 0 Then s = "error=" & Err.Number & ":" & Err.Description & "|"
    On Error GoTo 0
    s = s & Book(wb)
    wb.Close False
    Application.DisplayAlerts = True
    Probe = s
End Function
'''


def probe() -> str:
    return PROBE.replace("CASES\n", "".join(f'        Case "{name}": {call}\n' for name, _, call in CASES))


def written(path: Path) -> dict[str, object]:
    """What a save wrote: a package's sheet names, or a text file's bytes."""
    data = path.read_bytes()
    if data[:2] == b"PK":
        with zipfile.ZipFile(path) as package:
            workbook = package.read("xl/workbook.xml").decode("utf-8")
        return {"file": path.name, "sheets": re.findall(r'<sheet\b[^>]*name="([^"]+)"', workbook)}
    return {"file": path.name, "data": base64.b64encode(data).decode("ascii")}


def main() -> None:
    source = probe()
    answers: dict[str, str] = {}
    files: dict[str, object] = {}
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        version = excel.run_vba('Function V() As String\nV = Application.Version & "|" & Application.Build\n'
                                'End Function', "V")
        for name, file, _ in CASES:
            path = Path(folder) / file
            result = excel.run_vba(source, "Probe", [name, str(path)], timeout=300.0)
            assert result.ok and not result.dialogs, f"{name}: {result.outcome}: {result.message} {result.error}"
            answers[name] = str(result.value)
            found = sorted(Path(folder).glob(file + "*"))
            files[name] = written(found[0]) if found else None
            print(name, answers[name], files[name], flush=True)
    OUT.write_text(json.dumps({"excel_version_build": version.value,
                               "measured_at": datetime.now(timezone.utc).isoformat(), "probe": source,
                               "cases": [list(case) for case in CASES], "answers": answers, "files": files},
                              indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
