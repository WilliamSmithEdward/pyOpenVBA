"""What Excel answers of a workbook holding chart sheets: its tabs in order, what a macro walking Sheets, Worksheets and
Charts finds, names scoped to a sheet after a chart sheet, a formula across tabs, and what is left after a save. Each book
is saved by Excel here, kept in tests/fixtures/chart_sheets/ with its save path stripped, and read by READS and WALK. The
model then saves a copy, a cell edited and a chart sheet active, and Excel reads that as well: "after" and "activated".

    python scripts/measure_chart_sheets.py

writes tests/fixtures/chart_sheets/chart_sheets.json, which tests/test_excel_chart_sheets.py replays.
"""
from __future__ import annotations

import json
import shutil
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
OUT = ROOT / "tests/fixtures/chart_sheets"

DATA = '''ws.Range("A1:B1").Value = Array("k", "v")
    For i = 2 To 5
        ws.Cells(i, 1).Value = Chr$(95 + i): ws.Cells(i, 2).Value = i * 2
    Next
'''
#: Name: how the book is laid out, from Sheet1 holding the data; every book ends with Sheet1 active unless it says.
BOOKS: dict[str, str] = {
    "before": '''Set c = wb.Charts.Add(After:=wb.Sheets(wb.Sheets.Count)): c.SetSourceData ws.Range("A1:B5"): c.Name = "C"
    ws.Activate''',
    "last": '''Set c = wb.Charts.Add: c.SetSourceData ws.Range("A1:B5"): c.Name = "C"
    c.Move After:=ws
    ws.Activate''',
    "first": '''Set c = wb.Charts.Add(Before:=wb.Sheets(1)): c.SetSourceData ws.Range("A1:B5"): c.Name = "C"
    ws.Activate''',
    "middle": '''Set s = wb.Worksheets.Add(After:=ws): s.Name = "Two"
    Set c = wb.Charts.Add(After:=ws): c.SetSourceData ws.Range("A1:B5"): c.Name = "C"
    ws.Activate''',
    "active": '''Set c = wb.Charts.Add(After:=ws): c.SetSourceData ws.Range("A1:B5"): c.Name = "C"
    c.Activate''',
    "hidden": '''Set c = wb.Charts.Add(After:=ws): c.SetSourceData ws.Range("A1:B5"): c.Name = "C"
    ws.Activate
    c.Visible = 0''',
    "very_hidden": '''Set c = wb.Charts.Add(After:=ws): c.SetSourceData ws.Range("A1:B5"): c.Name = "C"
    ws.Activate
    c.Visible = 2''',
    "two_charts": '''Set c = wb.Charts.Add(After:=ws): c.SetSourceData ws.Range("A1:B5"): c.Name = "C"
    c.ChartType = 4: c.HasTitle = True: c.ChartTitle.Text = "First"
    Set c = wb.Charts.Add(Before:=ws): c.SetSourceData ws.Range("A1:B5"): c.Name = "D"
    ws.Activate''',
    "names": '''Set s = wb.Worksheets.Add(After:=ws): s.Name = "Two"
    Set c = wb.Charts.Add(Before:=s): c.SetSourceData ws.Range("A1:B5"): c.Name = "C"
    s.Range("A1").Value = 5
    s.Names.Add "Mine", "=Two!$A$1"
    ws.Names.Add "Yours", "=Sheet1!$B$2"
    wb.Names.Add "Book", "=Two!$A$1"
    s.PageSetup.PrintArea = "$A$1:$A$3"
    ws.Range("D1").Formula = "=SUM(Sheet1:Two!A1)"
    ws.Range("D2").Formula = "=Two!A1+1"
    ws.Range("D3").Formula = "=Two!Mine"
    ws.Activate''',
    "dressed": '''Set c = wb.Charts.Add(After:=ws): c.SetSourceData ws.Range("A1:B5"): c.Name = "C"
    c.Tab.Color = RGB(1, 2, 3)
    c.ChartType = 51: c.HasTitle = True: c.ChartTitle.Text = "T"
    c.PageSetup.Orientation = 1
    ws.Activate''',
    "protected": '''Set c = wb.Charts.Add(After:=ws): c.SetSourceData ws.Range("A1:B5"): c.Name = "C"
    c.Protect "pw", True, True
    ws.Activate''',
    "with_vba": '''Set c = wb.Charts.Add(After:=ws): c.SetSourceData ws.Range("A1:B5"): c.Name = "C"
    wb.VBProject.VBComponents(wb.VBProject.VBComponents.Count).CodeModule.AddFromString "Sub InChart()" & vbLf & "End Sub"
    ws.Activate''',
    "embedded_too": '''Set o = ws.ChartObjects.Add(10, 10, 200, 100): o.Chart.SetSourceData ws.Range("A1:B5")
    Set c = wb.Charts.Add(After:=ws): c.SetSourceData ws.Range("A1:B5"): c.Name = "C"
    ws.Activate''',
}
#: One read each, late bound as a macro with Sheets(...) has it; ``wb`` is the workbook, ``c`` the first chart sheet.
READS: list[str] = [
    "wb.Sheets.Count", "wb.Worksheets.Count", "wb.Charts.Count",
    'TypeName(wb.Sheets(1))', "wb.Sheets(1).Name", "wb.Sheets(1).Index", "wb.Sheets(1).Type", "wb.Sheets(1).Visible",
    'TypeName(wb.Sheets(2))', "wb.Sheets(2).Name", "wb.Sheets(2).Index", "wb.Sheets(2).Type", "wb.Sheets(2).Visible",
    "wb.Worksheets(1).Name", "wb.Worksheets(1).Index", "wb.Worksheets(wb.Worksheets.Count).Index",
    'wb.Worksheets("C").Name', 'wb.Sheets("C").Name', 'wb.Charts("C").Name', 'wb.Charts("Sheet1").Name',
    "wb.Charts(1).Name", "wb.Charts(1).Index", "wb.Charts(1).Visible", "wb.Charts(1).CodeName", "wb.Charts(1).ChartType",
    "wb.Charts(1).HasTitle", "wb.Charts(1).SeriesCollection.Count", "wb.Charts(1).SeriesCollection(1).Formula",
    "wb.Charts(1).ProtectContents", "wb.Charts(1).Tab.Color", "wb.Charts(1).Parent.Name = wb.Name",
    'TypeName(wb.ActiveSheet)', "wb.ActiveSheet.Name", 'TypeName(ActiveChart)', "ActiveChart Is Nothing",
    "wb.Sheets(1).Next.Name", "wb.Sheets(2).Previous.Name", "wb.Worksheets(1).Next.Name", "wb.Worksheets(1).Previous.Name",
    "wb.Names.Count", 'wb.Worksheets("Two").Names.Count', 'wb.Worksheets("Two").Names("Mine").RefersTo', 'wb.Worksheets("Two").PageSetup.PrintArea',
    'wb.Worksheets(1).Range("D1").Value', 'wb.Worksheets(1).Range("D1").Formula', 'wb.Worksheets(1).Range("D2").Value', 'wb.Worksheets(1).Range("D3").Value',
    'wb.Worksheets(1).ChartObjects.Count', 'wb.VBProject.VBComponents.Count', "wb.Sheets.Visible", 'TypeName(wb.Sheets)', 'TypeName(wb.Charts)',
    "wb.VBProject.VBComponents(2).Name", "wb.VBProject.VBComponents(3).Name", "wb.VBProject.VBComponents(3).Type", "wb.VBProject.VBComponents(3).CodeModule.CountOfLines",
]
#: What the model does before saving a copy for Excel to read: a cell the chart plots, and nothing else on the book.
EDIT = '''Sub Edit1(ByVal file As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(file)
    wb.Worksheets(1).Range("B2").Value = 99
    wb.Save
    wb.Close False
End Sub
Sub Activated(ByVal file As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(file)
    wb.Charts(1).Activate
    wb.Save
    wb.Close False
End Sub
'''
PLOTTED = '''Function Plotted(ByVal file As String) As String
    Dim wb As Workbook, v As Variant
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file)
    v = wb.Charts(1).SeriesCollection(1).Values
    Plotted = v(1) & "|" & wb.Worksheets(1).Range("B2").Value & "|" & wb.ActiveSheet.Name & "|" & wb.Windows(1).SelectedSheets.Count
    wb.Close False
End Function
'''
#: The macro: each of the sheets in turn, For Each over Sheets, Worksheets and Charts.
WALK = '''Function Walk(ByVal file As String) As String
    Dim wb As Workbook, s As Object, t As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file)
    For Each s In wb.Sheets
        t = t & "|" & TypeName(s) & ":" & s.Name & ":" & s.Index & ":" & s.Visible
    Next
    t = t & "||"
    For Each s In wb.Worksheets
        t = t & "|" & TypeName(s) & ":" & s.Name
    Next
    t = t & "||"
    For Each s In wb.Charts
        t = t & "|" & TypeName(s) & ":" & s.Name
    Next
    wb.Close False
    Walk = t
End Function
'''
READ = '''Private Function Tried(ByVal n As Long, wb As Workbook) As String
    Dim v As Variant
    On Error GoTo Failed
    Select Case n
{cases}
    End Select
    Tried = CStr(v)
    Exit Function
Failed:
    Tried = "E" & Err.Number
End Function
Function Read1(ByVal file As String) As String
    Dim wb As Workbook, i As Long, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file)
    For i = 0 To {last}
        s = s & vbTab & Tried(i, wb)
    Next
    wb.Close False
    Read1 = Mid$(s, 2)
End Function
'''
MAKE = '''Function Make(ByVal file As String, ByVal kind As Long) As String
    Dim wb As Workbook, ws As Worksheet, s As Worksheet, c As Object, o As Object, i As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(-4167): Set ws = wb.Worksheets(1)
    ''' + DATA + '''    {action}
    wb.SaveAs file, kind
    wb.Close False
    Make = "ok"
End Function
'''


def cases() -> str:
    return "\n".join(f"    Case {index}\n        v = {read}" for index, read in enumerate(READS))


def main() -> None:
    from fixture_workbook import strip_save_path
    from pyvbaharness import ExcelSession, HarnessConfig

    from pyopenvba.apps.excel import ExcelApplication

    OUT.mkdir(parents=True, exist_ok=True)
    folder = Path(tempfile.mkdtemp(prefix="pyopenvba_chartsheets_"))
    done = OUT / "chart_sheets.json"
    rows: list[dict[str, object]] = []
    try:
        for name, action in BOOKS.items():
            kind, suffix = (52, "xlsm") if name == "with_vba" else (51, "xlsx")
            file = folder / f"{name}.{suffix}"
            with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                excel.new_document()
                made = excel.run_vba(MAKE.format(action=action), "Make", (str(file), kind), timeout=90.0)
                if not made.ok:
                    rows.append({"name": name, "failed": f"{made.outcome} {made.message[:100]}"})
                    print(name, "failed", made.message[:100], flush=True)
                    continue
                read = excel.run_vba(READ.format(cases=cases(), last=len(READS) - 1), "Read1", (str(file),), timeout=90.0)
                walk = excel.run_vba(WALK, "Walk", (str(file),), timeout=90.0)
            fixture = OUT / file.name
            shutil.copy(file, fixture)
            strip_save_path(fixture)
            row: dict[str, object] = {"name": name, "file": file.name, "reads": str(read.value).split("\t") if read.ok else f"failed {read.message[:100]}",
                                      "walk": str(walk.value) if walk.ok else f"failed {walk.message[:100]}"}
            copies = {entry: folder / f"{name}_{entry}.{suffix}" for entry in ("Edit1", "Activated")}
            for entry, copy in copies.items():
                shutil.copy(file, copy)
                app = ExcelApplication(sandbox=False)
                app.add_workbook()
                app.add_module(EDIT, name="M")
                app.run(entry, copy.as_posix())
            with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                excel.new_document()
                again = excel.run_vba(READ.format(cases=cases(), last=len(READS) - 1), "Read1", (str(copies["Activated"]),), timeout=90.0)
                row["activated"] = str(again.value).split("\t") if again.ok else f"failed {again.message[:100]}"
                for key, copy in (("after", copies["Edit1"]), ("active", copies["Activated"]), ("before", file)):
                    plotted = excel.run_vba(PLOTTED, "Plotted", (str(copy),), timeout=90.0)
                    row[key] = str(plotted.value) if plotted.ok else f"failed {plotted.message[:100]}"
            rows.append(row)
            print(name, row["before"], "after", row["after"], "active", row["active"], flush=True)
            print(name, rows[-1]["walk"], flush=True)
            done.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "reads": READS, "rows": rows}, indent=1) + "\n", encoding="utf-8")
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    done.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "reads": READS, "rows": rows}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
