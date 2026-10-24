"""Date fields in a pivot table, and filtering their items by date, as Excel lays the table out, reads and saves it.

A macro puts a date field on a pivot table's rows and adds a date filter
to its PivotFilters: a specific date, before, after or between dates, or
a period counted from today, yesterday to next year, a quarter or a month
of any year. Every layout sits on a sheet of its own in one new workbook,
a pivot table of its own over one source list of dates, regions and
amounts, the dates placed about today so that each period holds some:
rows by date and the sum of the amounts. The dump reads every cell of the
table, the first date cell's type and number format, each field's items
and whether each shows, and each field's filters: their type, values,
data field and whether active. A layout with times of day, or that
refreshes, has a source of its own. The workbook is saved, its pivot table
parts kept, and Excel opens it and dumps every table again.

    python scripts/measure_pivot_date_filters.py

writes tests/fixtures/pivot_date_filters/: the workbook as Excel saved it,
and pivot_date_filters.json, which tests/test_excel_pivot_date_filters.py
replays on the day it was measured.
"""

from __future__ import annotations

import datetime as dt
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

OUT = ROOT / "tests" / "fixtures" / "pivot_date_filters"

HELPER = '''Private note As String

Private Function Filters(f As Object) As String
    Dim out As String, pf As Object, n As Long
    On Error Resume Next
    Err.Clear
    n = f.PivotFilters.Count
    If Err.Number <> 0 Then Filters = " filters=E" & Err.Number: Exit Function
    out = " filters=" & n
    For Each pf In f.PivotFilters
        out = out & "[" & pf.FilterType
        Err.Clear
        out = out & " v1=" & CStr(pf.Value1)
        If Err.Number <> 0 Then out = out & " v1=E" & Err.Number
        Err.Clear
        out = out & " v2=" & CStr(pf.Value2)
        If Err.Number <> 0 Then out = out & " v2=E" & Err.Number
        Err.Clear
        out = out & " data=" & pf.DataField.Name
        If Err.Number <> 0 Then out = out & " data=E" & Err.Number
        Err.Clear
        out = out & " active=" & pf.Active
        If Err.Number <> 0 Then out = out & " active=E" & Err.Number
        Err.Clear
        out = out & " types=" & VarType(pf.Value1) & "/" & VarType(pf.Value2)
        If Err.Number <> 0 Then out = out & " types=E" & Err.Number
        Err.Clear
        out = out & " whole=" & pf.WholeDayFilter
        If Err.Number <> 0 Then out = out & " whole=E" & Err.Number
        out = out & "]"
    Next
    Filters = out
End Function

Private Function Dump(pt As Object) As String
    Dim out As String, c As Object, f As Object, it As Object, first As Object
    On Error Resume Next
    out = "range=" & pt.TableRange2.Address(False, False) & " cells="
    For Each c In pt.TableRange2.Cells
        out = out & CStr(c.Value) & ","
        If first Is Nothing And VarType(c.Value) = vbDate Then Set first = c
    Next
    If Not first Is Nothing Then out = out & " datecell=" & first.Address(False, False) & ":" & TypeName(first.Value) _
        & ":" & first.NumberFormat
    Err.Clear
    out = out & " valuecell=" & pt.DataBodyRange.Cells(1, 1).NumberFormat & " labelcell=" _
        & pt.RowRange.Cells(2, 1).NumberFormat
    For Each f In pt.PivotFields
        If f.Orientation <> xlHidden And f.Orientation <> xlDataField Then
            out = out & " " & f.Name & "@" & f.Orientation & ":"
            For Each it In f.PivotItems
                out = out & it.Name & IIf(it.Visible, "", "-h") & ","
            Next
            Err.Clear
            out = out & " fmt=" & f.NumberFormat
            If Err.Number <> 0 Then out = out & " fmt=E" & Err.Number
            out = out & Filters(f)
        End If
    Next
    For Each f In pt.DataFields
        out = out & " df=" & f.Name & ":" & f.NumberFormat
    Next
    Dump = out
End Function

Private Function Made(wb As Object, ws As Object, cache As Object, rowField As String, columnField As String, _
        pageField As String) As Object
    Dim pt As Object
    Set pt = cache.CreatePivotTable(ws.Range("A3"), ws.Name)
    pt.PivotFields(rowField).Orientation = xlRowField
    If Len(columnField) > 0 Then pt.PivotFields(columnField).Orientation = xlColumnField
    If Len(pageField) > 0 Then pt.PivotFields(pageField).Orientation = xlPageField
    pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum
    Set Made = pt
End Function

Private Function Layout(wb As Object, cache As Object, name As String, rowField As String, columnField As String, _
        pageField As String, source As String, step As Long) As String
    Dim ws As Object, pt As Object, out As String, it As Object, own As Object, held As Object, total As Object
    Dim r As Long
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    Set own = cache
    If Len(source) > 0 Then
        Set held = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        held.Name = source
        held.Range("A1:C13").Value = wb.Worksheets("data").Range("A1:C13").Value
        held.Range("A2:A13").NumberFormat = "m/d/yyyy"
        If Left(source, 5) = "timed" Then held.Range("A9").Value = Date + TimeSerial(12, 0, 0)
        If source = "fmt" Then held.Range("A2:A13").NumberFormat = "d-mmm-yyyy"
        If source = "mix" Then held.Range("A2").NumberFormat = "d-mmm-yyyy"
        If source = "dtf" Then held.Range("A2:A13").NumberFormat = "m/d/yyyy h:mm"
        If source = "mixl" Then held.Range("A13").NumberFormat = "d-mmm-yyyy"
        If source = "mixh" Then held.Range("A2:A7").NumberFormat = "d-mmm-yyyy"
        If Left(source, 4) = "cash" Then held.Range("C2:C13").NumberFormat = "$#,##0.00"
        If source = "numf" Then
            For r = 2 To 13
                held.Cells(r, 2).Value = (r Mod 4) * 10
            Next
            held.Range("B2:B13").NumberFormat = "0.0"
        End If
        If Left(source, 4) = "nums" Then
            For r = 2 To 13
                held.Cells(r, 2).Value = (r Mod 4) * 10
            Next
        End If
        Set own = wb.PivotCaches.Create(xlDatabase, source & "!R1C1:R13C3")
    End If
    Set pt = Made(wb, ws, own, rowField, columnField, pageField)
    Set total = pt.DataFields(1)
    note = ""
    On Error Resume Next
    Err.Clear
    Select Case step
{steps}
    End Select
    If Err.Number <> 0 Then out = "E" & Err.Number & " "
    If Len(note) > 0 Then out = out & "note=" & note & " "
    On Error GoTo 0
    Layout = name & "^" & out & Dump(pt) & "|"
End Function
'''

#: The source list: twelve rows of a date about today, a region and an amount.
SOURCE = '''Set data = wb.Worksheets(1)
data.Name = "data"
data.Range("A1:C1").Value = Array("Date", "Region", "Amount")
regions = Array("East", "West", "North", "South")
dates = Array(Date - 400, DateSerial(Year(Date) - 1, 12, 31), DateSerial(Year(Date), 1, 1), _
    DateSerial(Year(Date), 3, 15), DateSerial(Year(Date), Month(Date) - 1, 10), Date - 7, Date - 1, Date, Date + 1, _
    Date + 7, DateSerial(Year(Date), Month(Date) + 1, 5), DateSerial(Year(Date) + 1, 2, 1))
For r = 0 To 11
    data.Cells(r + 2, 1).Value = dates(r)
    data.Cells(r + 2, 2).Value = regions(r Mod 4)
    data.Cells(r + 2, 3).Value = (r + 1) * 10
Next
data.Range("A2:A13").NumberFormat = "m/d/yyyy"
Set cache = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R13C3")
'''

DATE = 'pt.PivotFields("Date")'


def dated(kind: str, first: str = "", second: str = "", *, extra: str = "") -> str:
    """A date filter of ``kind`` on the dates."""
    values = (f", Value1:={first}" if first else "") + (f", Value2:={second}" if second else "")
    return f"{DATE}.PivotFilters.Add Type:={kind}{values}{extra}"


def probe(text: str, statement: str) -> str:
    """A statement, and its error number into the note under ``text`` if it raises one."""
    return (f"Err.Clear\n{statement}\n"
            f'If Err.Number <> 0 Then note = note & "{text}=E" & Err.Number & ";"\nErr.Clear\n')


PERIODS = ("xlDateTomorrow xlDateToday xlDateYesterday xlDateNextWeek xlDateThisWeek xlDateLastWeek xlDateNextMonth "
           "xlDateThisMonth xlDateLastMonth xlDateNextQuarter xlDateThisQuarter xlDateLastQuarter xlDateNextYear "
           "xlDateThisYear xlDateLastYear xlYearToDate xlAllDatesInPeriodQuarter1 xlAllDatesInPeriodQuarter2 "
           "xlAllDatesInPeriodQuarter3 xlAllDatesInPeriodQuarter4 xlAllDatesInPeriodJanuary "
           "xlAllDatesInPeriodFebruary xlAllDatesInPeriodMarch xlAllDatesInPeriodOctober "
           "xlAllDatesInPeriodNovember xlAllDatesInPeriodDecember").split()

#: name -> (the row field, the column field, the page field, a source sheet of the table's own or "" for the shared
#: one, what the layout does to its table).
LAYOUTS: dict[str, tuple[str, str, str, str, str]] = {
    "dates": ("Date", "", "", "", ""),
    "specific": ("Date", "", "", "", dated("xlSpecificDate", "Date")),
    "not_specific": ("Date", "", "", "", dated("xlNotSpecificDate", "Date")),
    "before": ("Date", "", "", "", dated("xlBefore", "Date")),
    "before_equal": ("Date", "", "", "", dated("xlBeforeOrEqualTo", "Date")),
    "after": ("Date", "", "", "", dated("xlAfter", "Date")),
    "after_equal": ("Date", "", "", "", dated("xlAfterOrEqualTo", "Date")),
    "between": ("Date", "", "", "", dated("xlDateBetween", "Date - 7", "Date + 1")),
    "not_between": ("Date", "", "", "", dated("xlDateNotBetween", "Date - 7", "Date + 1")),
    **{period[2:].replace("AllDatesInPeriod", "in_").replace("Date", "").lower(): ("Date", "", "", "", dated(period))
       for period in PERIODS},
    "specific_text": ("Date", "", "", "", probe("add", dated("xlSpecificDate", 'Format(Date, "m/d/yyyy")'))),
    "specific_serial": ("Date", "", "", "", probe("add", dated("xlSpecificDate", "CDbl(Date)"))),
    "before_text": ("Date", "", "", "", probe("add", dated("xlBefore", '"1/1/2026"'))),
    "on_text_field": ("Region", "", "", "", probe("add", 'pt.PivotFields("Region").PivotFilters.Add '
                                                         'Type:=xlSpecificDate, Value1:=Date')),
    "period_on_text_field": ("Region", "", "", "", probe("add", 'pt.PivotFields("Region").PivotFilters.Add '
                                                                'Type:=xlDateToday')),
    "date_then_label": ("Date", "", "", "", dated("xlDateThisYear") + "\n"
                        + probe("label", f'{DATE}.PivotFilters.Add Type:=xlCaptionContains, Value1:="1"')),
    "label_then_date": ("Date", "", "", "", f'{DATE}.PivotFilters.Add Type:=xlCaptionContains, Value1:="1"\n'
                        + probe("date", dated("xlDateThisYear"))),
    "date_then_value": ("Date", "", "", "", dated("xlDateThisYear") + "\n"
                        + f"{DATE}.PivotFilters.Add Type:=xlValueIsGreaterThan, DataField:=total, Value1:=50"),
    "two_dates": ("Date", "", "", "", dated("xlDateThisYear") + "\n" + probe("second", dated("xlDateThisMonth"))),
    "clear_label": ("Date", "", "", "", dated("xlDateThisYear") + f"\n{DATE}.ClearLabelFilters"),
    "clear_value": ("Date", "", "", "", dated("xlDateThisYear") + f"\n{DATE}.ClearValueFilters"),
    "clear_all": ("Date", "", "", "", dated("xlDateThisYear") + f"\n{DATE}.ClearAllFilters"),
    "with_hidden": ("Date", "", "", "", dated("xlDateThisYear") + "\n"
                    f"{DATE}.PivotItems(9).Visible = False"),
    "column_dates": ("Region", "Date", "", "", dated("xlDateThisMonth")),
    "outer_dates": ("Date", "", "", "", 'pt.PivotFields("Region").Orientation = xlRowField\n'
                    + dated("xlAllDatesInPeriodQuarter4")),
    "inner_dates": ("Region", "", "", "", 'pt.PivotFields("Date").Orientation = xlRowField\n'
                    + dated("xlDateLastYear")),
    "page_dates": ("Region", "", "Date", "", probe("add", dated("xlDateToday"))),
    "missing_value": ("Date", "", "", "", probe("add", dated("xlSpecificDate"))),
    "between_one": ("Date", "", "", "", probe("add", dated("xlDateBetween", "Date"))),
    "period_with_value": ("Date", "", "", "", probe("add", dated("xlDateToday", "Date"))),
    "timed_specific": ("Date", "", "", "timed", dated("xlSpecificDate", "Date")),
    "timed_whole": ("Date", "", "", "timed2", dated("xlSpecificDate", "Date", extra=", WholeDayFilter:=True")),
    "timed_not_whole": ("Date", "", "", "timed3", dated("xlSpecificDate", "Date", extra=", WholeDayFilter:=False")),
    "timed_today": ("Date", "", "", "timed4", dated("xlDateToday")),
    "timed_after": ("Date", "", "", "timed5", dated("xlAfter", "Date")),
    "timed_before_equal": ("Date", "", "", "timed6", dated("xlBeforeOrEqualTo", "Date")),
    "refreshed": ("Date", "", "", "data2", dated("xlDateToday") + "\n"
                  'wb.Worksheets("data2").Range("A2").Value = Date\npt.PivotCache.Refresh'),
    "refreshed_new_date": ("Date", "", "", "data3", 'wb.Worksheets("data3").Range("A13").Value = Date - 1000\n'
                           "pt.PivotCache.Refresh"),
    "refreshed_numbers": ("Region", "", "", "nums", 'wb.Worksheets("nums").Range("B3,B7,B11").Value = 50\n'
                          "pt.PivotCache.Refresh"),
    "formatted": ("Date", "", "", "fmt", dated("xlDateThisMonth")),
    "mixed_format": ("Date", "", "", "mix", ""),
    "mixed_last": ("Date", "", "", "mixl", ""),
    "mixed_half": ("Date", "", "", "mixh", ""),
    "money": ("Date", "", "", "cash", ""),
    "money_count": ("Date", "", "", "cash2", 'pt.AddDataField pt.PivotFields("Amount"), "Count of Amount", xlCount\n'
                    'pt.AddDataField pt.PivotFields("Amount"), "Average of Amount", xlAverage'),
    "number_rows_format": ("Region", "", "", "numf", ""),
    "datetime_format": ("Date", "", "", "dtf", ""),
    "label_on_dates": ("Date", "", "", "", f'{DATE}.PivotFilters.Add Type:=xlCaptionContains, Value1:="/2025"'),
    "label_whole": ("Region", "", "", "", 'pt.PivotFields("Region").PivotFilters.Add Type:=xlCaptionContains, '
                                          'Value1:="t"'),
    "page_current_date": ("Region", "", "Date", "", 'pt.PivotFields("Date").CurrentPage = Format(Date, "m/d/yyyy")'),
    "sorted_dates": ("Date", "", "", "", f'{DATE}.AutoSort xlDescending, "Date"'),
    "timed_add2_whole": ("Date", "", "", "timed7", probe("add2", dated("xlSpecificDate", "Date",
                                                                       extra=", WholeDayFilter:=True")
                                                         .replace(".Add ", ".Add2 "))),
    "timed_add2_not_whole": ("Date", "", "", "timed8", probe("add2", dated("xlSpecificDate", "Date",
                                                                           extra=", WholeDayFilter:=False")
                                                             .replace(".Add ", ".Add2 "))),
    "add2_today": ("Date", "", "", "", probe("add2", dated("xlDateToday").replace(".Add ", ".Add2 "))),
}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, data As Object, cache As Object, out As String, r As Long, regions As Variant, dates As Variant
Dim ws As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{source}{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\pivot_date_filters.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\pivot_date_filters.xlsx")
out = out & "~"
For Each ws In wb.Worksheets
    If ws.PivotTables.Count > 0 Then out = out & ws.Name & "^" & Dump(ws.PivotTables(1)) & "|"
Next
wb.Close False
Build = Format(Date, "yyyy-mm-dd") & "@" & out
End Function
'''

LAYOUT = 'out = out & Layout(wb, cache, "{name}", "{row}", "{column}", "{page}", "{source}", {step})\n'


def tables(path: Path) -> dict[str, tuple[str, str]]:
    """Each pivot table's pivotFields, rowItems, colItems and filters, and its cache's cacheFields, as the file
    keeps them, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, tuple[str, str]] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            kept = [found.group() for tag in ("pivotFields", "rowItems", "colItems")
                    if (found := re.search(rf"<{tag}\b.*?</{tag}>", text, re.DOTALL))]
            filters = re.search(r'<filters count="\d+">.*?</filter></filters>', text, re.DOTALL)
            rels = package.read(part.replace("pivotTables/", "pivotTables/_rels/") + ".rels").decode("utf-8")
            target = re.search(r'Target="\.\./pivotCache/([^"]+)"', rels)
            cache = package.read(f"xl/pivotCache/{target.group(1)}").decode("utf-8") if target else ""
            cached = re.search(r"<cacheFields\b.*?</cacheFields>", cache, re.DOTALL)
            out[name.group(1) if name else part] = ("".join(kept) + (filters.group() if filters else ""),
                                                    cached.group() if cached else "")
    return out


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    steps = "".join(f"    Case {index}\n" + "".join(f"        {line}\n" for line in step.splitlines())
                    for index, (_name, (_row, _column, _page, _source, step)) in enumerate(LAYOUTS.items(), 1))
    body = "".join(LAYOUT.replace("{name}", name).replace("{row}", row).replace("{column}", column)
                   .replace("{page}", page).replace("{source}", source).replace("{step}", str(index))
                   for index, (name, (row, column, page, source, _step)) in enumerate(LAYOUTS.items(), 1))
    code = (BUILD.replace("{steps}", steps).replace("{source}", SOURCE).replace("{layouts}", body)
            .replace("{folder}", str(OUT)))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    today, rest = str(result.value).split("@", 1)
    assert today == dt.date.today().isoformat(), (today, dt.date.today())
    session, reopened = rest.split("~", 1)
    strip_save_path(OUT / "pivot_date_filters.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = tables(OUT / "pivot_date_filters.xlsx")
    record = {"today": today, "build": BUILD, "source": SOURCE, "layout": LAYOUT, "steps": steps,
              "layouts": [{"name": name, "row": row, "column": column, "page": page, "own_source": source,
                           "step": step, "dump": dumps.get(name, ""), "reopened": again.get(name, ""),
                           "file": held.get(name, ("", ""))[0], "cache": held.get(name, ("", ""))[1]}
                          for name, (row, column, page, source, step) in LAYOUTS.items()]}
    (OUT / "pivot_date_filters.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                                 encoding="utf-8")
    print("today", today)
    for layout in record["layouts"]:
        print("==", layout["name"])
        print("   session ", layout["dump"])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"])
        print("   file    ", re.sub(r"\s+", " ", str(layout["file"]))[:1500])


if __name__ == "__main__":
    main()
