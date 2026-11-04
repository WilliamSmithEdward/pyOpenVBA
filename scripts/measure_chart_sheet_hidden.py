"""A hidden chart sheet activated by a macro: what Activate and Select answer, what Excel's own save writes, and what Excel
reads back of it. The books are tests/fixtures/chart_sheets/hidden.xlsx and very_hidden.xlsx.

    python scripts/measure_chart_sheet_hidden.py

writes tests/fixtures/chart_sheets/hidden.json, which tests/test_excel_chart_sheets.py replays.
"""
from __future__ import annotations

import json
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tests/fixtures/chart_sheets"
ANSWERS = '''Function Answers(ByVal file As String) As String
    Dim wb As Workbook, s As String, c As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file)
    Set c = wb.Charts(1)
    On Error Resume Next
    c.Activate
    s = Err.Number & "|" & wb.ActiveSheet.Name & "|" & c.Visible: Err.Clear
    c.Select
    s = s & "|" & Err.Number: Err.Clear
    wb.Close False
    Answers = s
End Function
'''
SAVE = '''Sub Save1(ByVal file As String)
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file)
    wb.Charts(1).Activate
    wb.Save
    wb.Close False
End Sub
'''
READ = '''Function Read1(ByVal file As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file)
    Read1 = wb.ActiveSheet.Name & "|" & wb.Charts(1).Visible & "|" & wb.Windows(1).SelectedSheets.Count
    wb.Close False
End Function
'''


def views(file: Path) -> dict[str, object]:
    """What the save says of the tabs: the book's active and first tabs, each tab's entry and its view."""
    with zipfile.ZipFile(file) as package:
        book = package.read("xl/workbook.xml").decode("utf-8")
        view = re.search(r"<workbookView\b[^>]*>", book)
        assert view is not None
        return {"active": re.findall(r'\b(?:activeTab|firstSheet)="\d+"', view.group(0)), "sheets": re.findall(r"<sheet\b[^>]*>", book),
                "sheet": re.findall(r"<sheetView\b[^>]*>", package.read("xl/worksheets/sheet1.xml").decode("utf-8")),
                "chart": re.findall(r"<sheetView\b[^>]*>", package.read("xl/chartsheets/sheet1.xml").decode("utf-8"))}


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    folder = Path(tempfile.mkdtemp(prefix="pyopenvba_hidden_"))
    rows: dict[str, dict[str, object]] = {}
    try:
        for name in ("hidden", "very_hidden"):
            file = folder / f"{name}.xlsx"
            shutil.copy(OUT / file.name, file)
            with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                excel.new_document()
                answers = excel.run_vba(ANSWERS, "Answers", (str(file),), timeout=60.0, line_numbers=False)
                excel.run_vba(SAVE, "Save1", (str(file),), timeout=60.0)
                read = excel.run_vba(READ, "Read1", (str(file),), timeout=60.0)
            rows[name] = {"answers": str(answers.value), "read": str(read.value), "views": views(file)}
            print(name, rows[name], flush=True)
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    (OUT / "hidden.json").write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "answers": ANSWERS, "save": SAVE, "rows": rows}, indent=1) + "\n",
                                     encoding="utf-8")


if __name__ == "__main__":
    main()
