"""AutoFilter's dynamic filters on dates -- today, this week, last month, a quarter of any year -- as Excel has them.

xlFilterDynamic with Criteria1 from xlFilterToday (1) to
xlFilterAllDatesInPeriodDecember (32) keeps the dates in a period, most of
them counted from today. Every layout sits on a sheet of its own in one
new workbook: a header and rows of dates around today -- days, weeks,
months, quarters and years before and after it, the edges of this week,
month, quarter and year, a time on today, a date written as text, a
number in General and a blank -- then one dynamic filter. The dump reads
which rows are hidden, FilterMode, and the filter's On, Operator, Count
and Criteria1. The workbook is saved, its filters kept, and Excel opens it
and dumps every sheet again. The moment the workbook was made goes in the
record, for a replay to stand its clock there.

    python scripts/measure_dynamic_filters.py

writes tests/fixtures/dynamic_filters/: the workbook as Excel saved it,
and dynamic_filters.json, which tests/test_excel_dynamic_filters.py
replays.
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "dynamic_filters"

HELPER = '''Private Function Show(v As Variant) As String
    If IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsError(v) Then
        Show = "Error:" & CStr(CLng(v))
    ElseIf IsObject(v) Then
        Show = "Object"
    ElseIf IsArray(v) Then
        Show = "Array"
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function Filters(ws As Object) As String
    Dim out As String, f As Object
    On Error Resume Next
    out = "filtering=" & ws.FilterMode
    If ws.AutoFilterMode Then
        Set f = ws.AutoFilter.Filters(1)
        out = out & " f1:" & f.On
        If f.On Then
            Err.Clear
            out = out & " op=" & f.Operator & " n=" & f.Count
            If Err.Number <> 0 Then out = out & " E" & Err.Number
            Err.Clear
            out = out & " c1=" & Show(f.Criteria1)
            If Err.Number <> 0 Then out = out & "E" & Err.Number
        End If
    End If
    Filters = out
End Function

Private Function Dump(ws As Object) As String
    Dim out As String, r As Long
    For r = 1 To ws.UsedRange.Rows.Count + 1
        If ws.Rows(r).Hidden Then out = out & r & ","
    Next
    Dump = "hidden=" & out & " " & Filters(ws)
End Function
'''

#: Days before and after today the rows hold, and the edges of today's week, month, quarter and year, with a time
#: on today and three cells that are no dates.
TABLE = '''Private Sub Fill(ws As Object, today As Date)
Dim offsets As Variant, edges As Variant, one As Variant, week As Date, quarter As Date, r As Long
ws.Range("A1").Value = "When"
offsets = Array(-800, -400, -366, -365, -200, -100, -95, -92, -90, -60, -45, -31, -30, -15, -8, -7, -6, -3, -2, _
    -1, 0, 1, 2, 3, 6, 7, 8, 14, 15, 30, 31, 45, 60, 90, 92, 95, 100, 200, 365, 366, 400, 800)
r = 2
For Each one In offsets
    ws.Cells(r, 1).Value = today + one
    r = r + 1
Next
week = today - Weekday(today) + 1
quarter = DateSerial(Year(today), ((Month(today) - 1) \\ 3) * 3 + 1, 1)
edges = Array(week, week - 1, week + 6, week + 7, DateSerial(Year(today), Month(today), 1), _
    DateSerial(Year(today), Month(today) + 1, 0), DateSerial(Year(today), Month(today) + 1, 1), _
    DateSerial(Year(today), Month(today), 0), quarter, DateSerial(Year(quarter), Month(quarter) + 3, 0), _
    DateSerial(Year(quarter), Month(quarter) + 3, 1), quarter - 1, DateSerial(Year(today), 1, 1), _
    DateSerial(Year(today), 12, 31), DateSerial(Year(today) + 1, 1, 1), DateSerial(Year(today) - 1, 12, 31))
For Each one In edges
    ws.Cells(r, 1).Value = one
    r = r + 1
Next
ws.Cells(r, 1).Value = today + 0.5
ws.Cells(r + 1, 1).Value = "'" & Format(today, "m/d/yyyy")
ws.Cells(r + 2, 1).Value = CDbl(today)
ws.Cells(r + 2, 1).NumberFormat = "General"
ws.Cells(r + 4, 1).Value = "word"
End Sub

Private Function Layout(wb As Object, name As String, criteria As Long, today As Date) As String
Dim ws As Object, out As String
If wb.Worksheets(1).Name = "Sheet1" Then
    Set ws = wb.Worksheets(1)
Else
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
End If
ws.Name = name
Fill ws, today
On Error Resume Next
ws.Range("A1").CurrentRegion.AutoFilter Field:=1, Criteria1:=criteria, Operator:=xlFilterDynamic
If Err.Number <> 0 Then out = name & "^E" & Err.Number & " " & Err.Description & "|"
On Error GoTo 0
Layout = out & name & "^" & Dump(ws) & "|"
End Function
'''

#: The dynamic filters, by their criteria.
CRITERIA = {name: number for number, name in enumerate([
    "today", "yesterday", "tomorrow", "this_week", "last_week", "next_week", "this_month", "last_month",
    "next_month", "this_quarter", "last_quarter", "next_quarter", "this_year", "last_year", "next_year",
    "year_to_date", "quarter1", "quarter2", "quarter3", "quarter4", "january", "february", "march", "april", "may",
    "june", "july", "august", "september", "october", "november", "december"], 1)}

BUILD = HELPER + TABLE + '''
Public Function Build() As String
Dim wb As Object, ws As Object, out As String, today As Date
Application.DisplayAlerts = False
today = Date
out = "now=" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "|"
Set wb = Workbooks.Add(xlWBATWorksheet)
{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\dynamic_filters.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\dynamic_filters.xlsx")
out = out & "~"
For Each ws In wb.Worksheets
    out = out & ws.Name & "^" & Dump(ws) & "|"
Next
wb.Close False
Build = out
End Function
'''

LAYOUT = 'out = out & Layout(wb, "{name}", {criteria}, today)\n'


def file_filters(path: Path) -> dict[str, list[str]]:
    """Each sheet's autoFilter element as the file keeps it."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
        targets.update({key: value for value, key in re.findall(r'Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
        out: dict[str, list[str]] = {}
        for name, relationship in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
            text = package.read("xl/" + targets[relationship].lstrip("/").removeprefix("xl/")).decode("utf-8")
            out[name] = re.findall(r"<autoFilter\b[^>]*?/>|<autoFilter\b.*?</autoFilter>", text, re.DOTALL)
    return out


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    body = "".join(LAYOUT.replace("{name}", name).replace("{criteria}", str(number))
                   for name, number in CRITERIA.items())
    code = BUILD.replace("{layouts}", body).replace("{folder}", str(OUT))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    head, rest = str(result.value).split("|", 1)
    session, reopened = rest.split("~", 1)
    strip_save_path(OUT / "dynamic_filters.xlsx")
    entries = [one.split("^", 1) for one in session.split("|") if "^" in one]
    errors = {name: text for name, text in entries if text.startswith("E")}
    dumps = {name: text for name, text in entries if not text.startswith("E")}
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = file_filters(OUT / "dynamic_filters.xlsx")
    record = {"now": head.removeprefix("now="), "helper": HELPER, "table": TABLE, "layout": LAYOUT, "build": BUILD,
              "layouts": [{"name": name, "criteria": number, "error": errors.get(name, ""), "dump": dumps.get(name, ""),
                           "reopened": again.get(name, ""), "file": held.get(name, [])}
                          for name, number in CRITERIA.items()]}
    (OUT / "dynamic_filters.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                              encoding="utf-8")
    print("now", record["now"])
    for layout in record["layouts"]:
        print("==", layout["name"], layout["error"])
        print("   session ", layout["dump"])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"])
        print("   file    ", [re.sub(r' xr:uid="[^"]*"', "", one) for one in layout["file"]])


if __name__ == "__main__":
    main()
