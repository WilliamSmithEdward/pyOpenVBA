"""A pivot table's fields grouped, by number ranges, date periods and chosen items, as Excel groups, reads and saves
them.

A macro groups a field of numbers by a step, a field of dates by days,
months, quarters or years, or by seconds, minutes and hours where its
values carry times, or chosen items of a field, through Range.Group on a
cell of the field, and ungroups it again. Every layout sits on a sheet of
its own in one new workbook, a pivot table of its own over a cache of its
own, since a grouping belongs to its cache, over one of three source
lists: regions, products, dates, amounts and units; fractions, negatives,
large and formatted numbers; and stamps with times. The dump reads every
cell of the table, a quote marking text that reads as a number, each
field's name, orientation, source name and items. The workbook is saved,
its pivot table and cache parts kept, and Excel opens it and dumps every
table again.

    python scripts/measure_pivot_grouping.py

writes tests/fixtures/pivot_grouping/: the workbook as Excel saved it, and
pivot_grouping.json, which tests/test_excel_pivot_grouping.py replays.
"""

from __future__ import annotations

import json
import posixpath
import re
import shutil
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "pivot_grouping"

HELPER = '''Private note As String

Private Function Props(o As Object, names As String) As String
    Dim n As Variant, out As String, v As Variant
    On Error Resume Next
    For Each n In Split(names, " ")
        Err.Clear
        v = Empty
        v = CallByName(o, CStr(n), VbGet)
        If Err.Number <> 0 Then
            out = out & " " & n & "=E" & Err.Number
        Else
            out = out & " " & n & "=" & CStr(v)
        End If
    Next
    Props = out
End Function

Private Function Dump(pt As Object) As String
    Dim out As String, c As Object, f As Object, it As Object
    On Error Resume Next
    out = "range=" & pt.TableRange2.Address(False, False) & " cells="
    For Each c In pt.TableRange2.Cells
        If VarType(c.Value2) = vbString Then
            If IsNumeric(c.Value2) Then out = out & "'"
        End If
        out = out & CStr(c.Value2)
        If c.NumberFormat <> "General" Then out = out & "~" & c.NumberFormat
        out = out & ","
    Next
    out = out & " fields=" & pt.PivotFields.Count
    For Each f In pt.PivotFields
        out = out & " [" & f.Name & ":" & Props(f, "Orientation SourceName")
        If f.Orientation = xlRowField Or f.Orientation = xlColumnField Then
            For Each it In f.PivotItems
                out = out & " " & it.Name
            Next
        End If
        out = out & "]"
    Next
    Dump = out
End Function

Private Function Made(ws As Object, cache As Object, rowFields As String, columnFields As String) As Object
    Dim pt As Object, n As Variant
    Set pt = cache.CreatePivotTable(ws.Range("A3"), ws.Name)
    For Each n In Split(rowFields, ",")
        pt.PivotFields(CStr(n)).Orientation = xlRowField
    Next
    For Each n In Split(columnFields, ",")
        pt.PivotFields(CStr(n)).Orientation = xlColumnField
    Next
    pt.AddDataField pt.PivotFields("Units"), "Sum of Units", xlSum
    Set Made = pt
End Function

Private Function Layout(wb As Object, name As String, source As String, rowFields As String, _
        columnFields As String, step As Long) As String
    Dim ws As Object, pt As Object, out As String, cache As Object, pt2 As Object
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    Set cache = wb.PivotCaches.Create(xlDatabase, source)
    Set pt = Made(ws, cache, rowFields, columnFields)
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

SOURCE = '''Set data = wb.Worksheets(1)
data.Name = "data"
data.Range("A1:E1").Value = Array("Region", "Product", "Date", "Amount", "Units")
rows = Array("East|Apple|2024|1|15|10|1", "East|Banana|2024|2|20|20|2", "East|Cherry|2024|3|5|30|3", _
    "West|Apple|2024|5|11|40|4", "West|Banana|2024|7|30|50|5", "North|Apple|2024|8|2|60|6", _
    "North|Cherry|2024|10|19|70|7", "South|Banana|2024|12|24|80|8", "South|Cherry|2025|1|3|90|9", _
    "East|Apple|2025|2|14|15|1", "West|Apple|2025|3|31|25|2", "North|Apple|2025|4|1|35|3", _
    "South|Apple|2025|6|15|45|4", "West|Banana|2025|8|8|55|5", "South|Banana|2025|9|9|65|6", _
    "North|Cherry|2025|11|30|75|7")
For r = 0 To 15
    parts = Split(rows(r), "|")
    data.Cells(r + 2, 1).Value = parts(0)
    data.Cells(r + 2, 2).Value = parts(1)
    data.Cells(r + 2, 3).Value = DateSerial(CLng(parts(2)), CLng(parts(3)), CLng(parts(4)))
    data.Cells(r + 2, 4).Value = CLng(parts(5))
    data.Cells(r + 2, 5).Value = CLng(parts(6))
Next
Set data = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
data.Name = "numbers"
data.Range("A1:D1").Value = Array("Value", "Big", "Money", "Units")
rows = Array("-5.5|1000|1.5", "-1|2500|2", "0|12000|3.25", "0.25|45000|10", "3.75|46000|11.75", "7|70000|20", _
    "12|89999|25.5", "19.5|90000|30")
For r = 0 To 7
    parts = Split(rows(r), "|")
    data.Cells(r + 2, 1).Value = Val(parts(0))
    data.Cells(r + 2, 2).Value = Val(parts(1))
    data.Cells(r + 2, 3).Value = Val(parts(2))
    data.Cells(r + 2, 4).Value = r + 1
Next
data.Range("C2:C9").NumberFormat = "$#,##0.00"
Set data = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
data.Name = "times"
data.Range("A1:B1").Value = Array("Stamp", "Units")
rows = Array("2024|1|15|0|15|30", "2024|1|15|9|45|10", "2024|1|16|13|5|59", "2024|2|1|23|59|0", _
    "2024|3|3|12|0|0", "2024|3|3|12|30|45")
For r = 0 To 5
    parts = Split(rows(r), "|")
    data.Cells(r + 2, 1).Value = DateSerial(CLng(parts(0)), CLng(parts(1)), CLng(parts(2))) + _
        TimeSerial(CLng(parts(3)), CLng(parts(4)), CLng(parts(5)))
    data.Cells(r + 2, 2).Value = r + 1
Next
Set data = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
data.Name = "mixed"
data.Range("A1:G1").Value = Array("Whole", "Timed", "Units", "Iso", "Same", "Gappy", "Texty")
data.Range("A2:G2").Value = Array(DateSerial(2024, 1, 15) + TimeSerial(9, 30, 0), DateSerial(2024, 1, 15), 1, _
    DateSerial(2024, 1, 15), 5, 1, 1)
data.Range("A3:G3").Value = Array(DateSerial(2024, 2, 20), DateSerial(2024, 2, 20), 2, DateSerial(2024, 2, 20), 5, _
    Empty, "x")
data.Range("A4:G4").Value = Array(DateSerial(2024, 3, 5), DateSerial(2024, 3, 5) + TimeSerial(6, 0, 0), 3, _
    DateSerial(2024, 3, 5), 5, 3, 3)
data.Range("D2:D4").NumberFormat = "yyyy-mm-dd"
Set data = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
data.Name = "spans"
data.Range("A1:{last}1").Value = Array({names}, "Units")
data.Range("A2:{last}2").Value = Array({lows}, 1)
data.Range("A3:{last}3").Value = Array({highs}, 2)
'''

#: Two values a field of its own holds, for the step Excel chooses when none is given.
SPANS = [(0, 3), (0, 3.1), (0, 3.2), (0, 3.5), (0, 4), (0, 5), (0, 6), (0, 7.9), (0, 25), (0, 30), (0, 31), (0, 32),
         (0, 35), (0, 40), (0, 50), (0, 80), (0, 0.3), (0, 0.5), (0, 0.05), (0, 1), (0, 1000), (0, 3000), (0, 3500),
         (100, 103), (100, 200)]


def span_name(low: float, high: float) -> str:
    return f"S{low:g}_{high:g}".replace(".", "p")


def column_letter(number: int) -> str:
    text = ""
    while number:
        number, rest = divmod(number - 1, 26)
        text = chr(65 + rest) + text
    return text


SOURCE = (SOURCE.replace("{last}", column_letter(len(SPANS) + 1))
          .replace("{names}", ", ".join(f'"{span_name(low, high)}"' for low, high in SPANS))
          .replace("{lows}", ", ".join(f"{low:g}" for low, _high in SPANS))
          .replace("{highs}", ", ".join(f"{high:g}" for _low, high in SPANS)))
SOURCES = {"data": "data!R1C1:R17C5", "numbers": "numbers!R1C1:R9C4", "times": "times!R1C1:R7C2",
           "mixed": "mixed!R1C1:R4C7", "spans": f"spans!R1C1:R3C{len(SPANS) + 1}"}


def probe(text: str, statement: str) -> str:
    return (f"Err.Clear\n{statement}\n"
            f'If Err.Number <> 0 Then note = note & "{text}=E" & Err.Number & ";"\nErr.Clear\n')


def lines(*statements: str) -> str:
    return "\n".join(statements)


def note(text: str, expression: str) -> str:
    return (f'Err.Clear\nnote = note & "{text}=" & {expression} & ";"\n'
            f'If Err.Number <> 0 Then note = note & "{text}=E" & Err.Number & ";"\nErr.Clear\n')


def cell(field: str) -> str:
    """The first of a field's item cells, where Range.Group groups the field."""
    return f'pt.PivotFields("{field}").DataRange.Cells(1)'


def label(field: str, item: str) -> str:
    return f'pt.PivotFields("{field}").PivotItems("{item}").LabelRange'


def union(field: str, *items: str) -> str:
    return "Union(" + ", ".join(label(field, item) for item in items) + ")"


def item_ranges(field: str, item: str) -> tuple[str, str]:
    """Notes of an item's label cells and data cells."""
    found = f'pt.PivotFields("{field}").PivotItems("{item}")'
    return (note(f"{field}.{item}.label", f"{found}.LabelRange.Address(False, False)"),
            note(f"{field}.{item}.data", f"{found}.DataRange.Address(False, False)"))


def refreshed(step: str) -> str:
    """The step, then the source's last record changed to a new region, date and amount, the cache refreshed and
    the record put back, which leaves the table as the refresh made it."""
    return lines(step, 'Set pt2 = wb.Worksheets("data")', 'old = pt2.Range("A17:D17").Value',
                 'pt2.Range("A17:D17").Value = Array("Central", "Apple", DateSerial(2026, 3, 1), 120)',
                 probe("refresh", "pt.PivotCache.Refresh"), 'pt2.Range("A17:D17").Value = old')


def grouped(field: str, arguments: str, text: str = "group") -> str:
    return probe(text, f"{cell(field)}.Group {arguments}")


def periods(*chosen: str) -> str:
    names = ("seconds", "minutes", "hours", "days", "months", "quarters", "years")
    return "Array(" + ", ".join("True" if name in chosen else "False" for name in names) + ")"


DATES = "Date"
AMOUNTS = "Amount"
MONTHS_YEARS = f"Start:=True, End:=True, Periods:={periods('months', 'years')}"
EAST_WEST = union("Region", "East", "West") + ".Group"

#: name -> (the source, the row fields, the column fields, what the layout does to its table).
LAYOUTS: dict[str, tuple[str, str, str, str]] = {
    "normal_numbers": ("data", AMOUNTS, "", ""),
    "normal_dates": ("data", DATES, "", ""),
    # Numbers by a step.
    "numbers_by_10": ("data", AMOUNTS, "", grouped(AMOUNTS, "Start:=0, End:=100, By:=10")),
    "numbers_auto": ("data", AMOUNTS, "", grouped(AMOUNTS, "Start:=True, End:=True, By:=25")),
    "numbers_inside": ("data", AMOUNTS, "", grouped(AMOUNTS, "Start:=20, End:=60, By:=20")),
    "numbers_fraction": ("data", AMOUNTS, "", grouped(AMOUNTS, "Start:=10, End:=90, By:=12.5")),
    "numbers_columns": ("data", "Region", AMOUNTS, grouped(AMOUNTS, "Start:=0, End:=100, By:=50")),
    "numbers_under_region": ("data", f"Region,{AMOUNTS}", "",
                             probe("group", f'{label(AMOUNTS, "10")}.Cells(1).Group Start:=0, End:=100, By:=50')),
    "numbers_ungroup": ("data", AMOUNTS, "", lines(grouped(AMOUNTS, "Start:=0, End:=100, By:=10"),
                                                   probe("ungroup", f"{cell(AMOUNTS)}.Ungroup"))),
    "numbers_default": ("data", AMOUNTS, "", probe("group", f"{cell(AMOUNTS)}.Group")),
    "numbers_regroup": ("data", AMOUNTS, "", lines(grouped(AMOUNTS, "Start:=0, End:=100, By:=10", "first"),
                                                   grouped(AMOUNTS, "Start:=True, End:=True, By:=25", "second"))),
    "numbers_page": ("data", "Region", "", lines(probe("page", f'pt.PivotFields("{AMOUNTS}").Orientation = 3'),
                                                 grouped(AMOUNTS, "Start:=0, End:=100, By:=50"))),
    "numbers_hide": ("data", AMOUNTS, "", lines(grouped(AMOUNTS, "Start:=0, End:=100, By:=10"),
                                                probe("hide", f'pt.PivotFields("{AMOUNTS}").PivotItems("20-29")'
                                                              '.Visible = False'))),
    "value_by_5": ("numbers", "Value", "", grouped("Value", "Start:=-10, End:=20, By:=5")),
    "value_auto_5": ("numbers", "Value", "", grouped("Value", "Start:=True, End:=True, By:=5")),
    "value_by_2_5": ("numbers", "Value", "", grouped("Value", "Start:=True, End:=True, By:=2.5")),
    "value_default": ("numbers", "Value", "", probe("group", f'{cell("Value")}.Group')),
    "big_default": ("numbers", "Big", "", probe("group", f'{cell("Big")}.Group')),
    "big_by_10000": ("numbers", "Big", "", grouped("Big", "Start:=0, End:=100000, By:=10000")),
    "money_by_10": ("numbers", "Money", "", grouped("Money", "Start:=0, End:=30, By:=10")),
    "money_default": ("numbers", "Money", "", probe("group", f'{cell("Money")}.Group')),
    # Dates by periods.
    "dates_months": ("data", DATES, "", grouped(DATES, f"Start:=True, End:=True, Periods:={periods('months')}")),
    "dates_quarters": ("data", DATES, "", grouped(DATES, f"Start:=True, End:=True, Periods:={periods('quarters')}")),
    "dates_years": ("data", DATES, "", grouped(DATES, f"Start:=True, End:=True, Periods:={periods('years')}")),
    "dates_days": ("data", DATES, "", grouped(DATES, f"Start:=True, End:=True, Periods:={periods('days')}")),
    "dates_weeks": ("data", DATES, "", grouped(DATES, f"Start:=True, End:=True, By:=7, Periods:={periods('days')}")),
    "dates_by_10": ("data", DATES, "", grouped(DATES, f"Start:=True, End:=True, By:=10, Periods:={periods('days')}")),
    "dates_months_years": ("data", DATES, "", grouped(DATES, MONTHS_YEARS)),
    "dates_quarters_years": ("data", DATES, "", grouped(DATES, f"Start:=True, End:=True, Periods:="
                                                               f"{periods('quarters', 'years')}")),
    "dates_all": ("data", DATES, "", grouped(DATES, f"Start:=True, End:=True, Periods:="
                                                    f"{periods('days', 'months', 'quarters', 'years')}")),
    "dates_from": ("data", DATES, "", grouped(DATES, "Start:=DateSerial(2024, 6, 1), End:=DateSerial(2025, 6, 30), "
                                                     f"Periods:={periods('months')}")),
    "dates_end_on_data": ("data", DATES, "", grouped(DATES, "Start:=DateSerial(2024, 1, 15), "
                                                            "End:=DateSerial(2025, 6, 15), "
                                                            f"Periods:={periods('months')}")),
    "dates_end_midnight": ("data", DATES, "", grouped(DATES, "Start:=True, End:=DateSerial(2026, 1, 1), "
                                                             f"Periods:={periods('years')}")),
    "dates_start_time": ("data", DATES, "", grouped(DATES, "Start:=DateSerial(2024, 3, 1) + TimeSerial(12, 0, 0), "
                                                           f"End:=True, Periods:={periods('months')}")),
    "dates_default": ("data", DATES, "", probe("group", f"{cell(DATES)}.Group")),
    "dates_no_periods": ("data", DATES, "", grouped(DATES, "Start:=True, End:=True")),
    "dates_columns": ("data", "Region", DATES, grouped(DATES, f"Start:=True, End:=True, Periods:={periods('years')}")),
    "dates_inner": ("data", f"Region,{DATES}", "",
                    probe("group", f'pt.PivotFields("{DATES}").PivotItems(1).LabelRange.Cells(1).Group {MONTHS_YEARS}')),
    "dates_regroup": ("data", DATES, "", lines(grouped(DATES, f"Start:=True, End:=True, Periods:={periods('months')}",
                                                       "first"),
                                               probe("second", f'{cell("Months (Date)")}.Group Start:=True, End:=True, '
                                                               f'Periods:={periods("years")}'))),
    "dates_ungroup": ("data", DATES, "", lines(grouped(DATES, MONTHS_YEARS),
                                               probe("ungroup", f"{cell(DATES)}.Ungroup"))),
    "dates_ungroup_years": ("data", DATES, "", lines(grouped(DATES, MONTHS_YEARS),
                                                     probe("ungroup", f'{cell("Years (Date)")}.Ungroup'))),
    "dates_ungroup_months": ("data", DATES, "", lines(grouped(DATES, MONTHS_YEARS),
                                                      probe("ungroup", f'{cell("Months (Date)")}.Ungroup'))),
    "dates_hide": ("data", DATES, "", lines(grouped(DATES, MONTHS_YEARS),
                                            probe("hide", 'pt.PivotFields("Months (Date)").PivotItems("Jan")'
                                                          '.Visible = False'))),
    "times_hours": ("times", "Stamp", "", grouped("Stamp", f"Start:=True, End:=True, Periods:={periods('hours')}")),
    "times_minutes": ("times", "Stamp", "", grouped("Stamp", f"Start:=True, End:=True, Periods:={periods('minutes')}")),
    "times_seconds": ("times", "Stamp", "", grouped("Stamp", f"Start:=True, End:=True, Periods:={periods('seconds')}")),
    "times_days_hours": ("times", "Stamp", "", grouped("Stamp", f"Start:=True, End:=True, Periods:="
                                                                f"{periods('days', 'hours')}")),
    # Chosen items.
    "manual_group": ("data", "Region", "", probe("group", EAST_WEST)),
    "manual_order": ("data", "Region", "", probe("group", union("Region", "East", "North") + ".Group")),
    "manual_two": ("data", "Region", "", lines(probe("first", EAST_WEST),
                                               probe("second", union("Region", "North", "South") + ".Group"))),
    "manual_ungroup": ("data", "Region", "", lines(probe("group", EAST_WEST),
                                                   probe("ungroup", f'{label("Region2", "Group1")}.Ungroup'))),
    "manual_renumber": ("data", "Region", "", lines(probe("first", EAST_WEST),
                                                    probe("second", union("Region", "North", "South") + ".Group"),
                                                    probe("ungroup", f'{label("Region2", "Group1")}.Ungroup'),
                                                    probe("third", EAST_WEST))),
    "manual_second_level": ("data", "Region", "", lines(probe("first", EAST_WEST),
                                                        probe("second", union("Region2", "Group1", "North")
                                                              + ".Group"))),
    "manual_overlap": ("data", "Region", "", lines(probe("first", EAST_WEST),
                                                   probe("second", union("Region", "West", "North") + ".Group"))),
    "manual_single": ("data", "Region", "", probe("group", f'{label("Region", "East")}.Group')),
    "manual_columns": ("data", "", "Region", probe("group", EAST_WEST)),
    "manual_inner": ("data", "Product,Region", "", probe("group", EAST_WEST)),
    "manual_numbers": ("data", AMOUNTS, "", probe("group", union(AMOUNTS, "10", "20") + ".Group")),
    "manual_dates": ("data", DATES, "", probe("group", f'Union(pt.PivotFields("{DATES}").PivotItems(1).LabelRange, '
                                                       f'pt.PivotFields("{DATES}").PivotItems(2).LabelRange).Group')),
    "manual_rename": ("data", "Region", "", lines(probe("group", EAST_WEST),
                                                  probe("item", 'pt.PivotFields("Region2").PivotItems("Group1").Name'
                                                                ' = "Coasts"'),
                                                  probe("field", 'pt.PivotFields("Region2").Caption = "Area"'))),
    "manual_hide": ("data", "Region", "", lines(probe("group", EAST_WEST),
                                                probe("hide", 'pt.PivotFields("Region2").PivotItems("Group1")'
                                                              '.Visible = False'))),
    # What the groups say of each other.
    "relations_manual": ("data", "Region", "", lines(
        probe("group", EAST_WEST),
        note("child", 'pt.PivotFields("Region2").ChildField.Name'),
        note("parent", 'pt.PivotFields("Region").ParentField.Name'),
        note("childitems", 'pt.PivotFields("Region2").PivotItems("Group1").ChildItems.Count'),
        note("childitem", 'pt.PivotFields("Region2").PivotItems("Group1").ChildItems(1).Name'),
        note("parentitem", 'pt.PivotFields("Region").PivotItems("East").ParentItem.Name'),
        note("level", 'pt.PivotFields("Region2").GroupLevel'),
        note("base_level", 'pt.PivotFields("Region").GroupLevel'),
        note("total", 'pt.PivotFields("Region").TotalLevels'),
        note("parentitems", 'pt.PivotFields("Region2").ParentItems.Count'),
        note("item_source", 'pt.PivotFields("Region2").PivotItems("Group1").SourceName'))),
    "relations_dates": ("data", DATES, "", lines(
        grouped(DATES, MONTHS_YEARS),
        note("child", 'pt.PivotFields("Years (Date)").ChildField.Name'),
        note("parent", 'pt.PivotFields("Months (Date)").ParentField.Name'),
        note("date_parent", f'pt.PivotFields("{DATES}").ParentField.Name'),
        note("level", 'pt.PivotFields("Months (Date)").GroupLevel'),
        note("items", 'pt.PivotFields("Years (Date)").PivotItems.Count'),
        note("first", 'pt.PivotFields("Years (Date)").PivotItems(1).Name'),
        note("date_items", f'pt.PivotFields("{DATES}").PivotItems.Count'))),
    "relations_numbers": ("data", AMOUNTS, "", lines(
        grouped(AMOUNTS, "Start:=0, End:=100, By:=10"),
        note("value", f'pt.PivotFields("{AMOUNTS}").PivotItems("10-19").Value'),
        note("source", f'pt.PivotFields("{AMOUNTS}").PivotItems("10-19").SourceName'),
        note("count", f'pt.PivotFields("{AMOUNTS}").PivotItems.Count'),
        note("first", f'pt.PivotFields("{AMOUNTS}").PivotItems(1).Name'))),
    "shared_cache": ("data", AMOUNTS, "", lines(
        'Set pt2 = cache.CreatePivotTable(ws.Range("H3"), name & "_2")',
        'pt2.PivotFields("Region").Orientation = xlRowField',
        'pt2.AddDataField pt2.PivotFields("Units"), "Sum of Units", xlSum',
        grouped(AMOUNTS, "Start:=0, End:=100, By:=50"),
        'note = note & "second=" & Dump(pt2) & ";"')),
    # Refusals.
    "errors": ("data", "Region", "", lines(probe("text", f'{cell("Region")}.Group Start:=0, End:=10, By:=1'),
                                           probe("zero", f'{cell("Region")}.Group Start:=True, End:=True, By:=0'),
                                           probe("data", "pt.DataBodyRange.Cells(1).Group Start:=0, End:=100, By:=10"),
                                           probe("header", "pt.TableRange1.Cells(1).Group Start:=0, End:=100, By:=10"),
                                           probe("grand", "pt.TableRange1.Cells(pt.TableRange1.Rows.Count, 1).Group"),
                                           probe("ungroup", f'{label("Region", "East")}.Ungroup'))),
    "number_errors": ("data", AMOUNTS, "", lines(probe("zero", f"{cell(AMOUNTS)}.Group Start:=0, End:=100, By:=0"),
                                                 probe("reversed", f"{cell(AMOUNTS)}.Group Start:=100, End:=0, By:=10"),
                                                 probe("periods", f"{cell(AMOUNTS)}.Group Start:=True, End:=True, "
                                                                  f"Periods:={periods('months')}"))),
    "date_errors": ("data", DATES, "", lines(probe("none", f"{cell(DATES)}.Group Start:=True, End:=True, "
                                                           f"Periods:={periods()}"),
                                             probe("reversed", f"{cell(DATES)}.Group Start:=DateSerial(2025, 1, 1), "
                                                               "End:=DateSerial(2024, 1, 1), "
                                                               f"Periods:={periods('months')}"),
                                             probe("short", f"{cell(DATES)}.Group Start:=True, End:=True, "
                                                            "Periods:=Array(True, False)"))),
    # Round three: where a group sorts and lands, a day step ending on a value, and times beside whole days.
    "manual_ungroup_one": ("data", "Region", "", lines(probe("first", EAST_WEST),
                                                       probe("second", union("Region", "North", "South") + ".Group"),
                                                       probe("ungroup", f'{label("Region2", "Group1")}.Ungroup'))),
    "manual_late_numbers": ("data", AMOUNTS, "", probe("group", union(AMOUNTS, "80", "90") + ".Group")),
    "manual_late_text": ("data", "Region", "", probe("group", union("Region", "South", "West") + ".Group")),
    "manual_three": ("data", "Region", "", probe("group", union("Region", "East", "West", "North") + ".Group")),
    "manual_swallow": ("data", "Region", "", lines(probe("first", EAST_WEST),
                                                   probe("second", union("Region", "East", "West", "North")
                                                         + ".Group"))),
    "dates_weeks_end_on_data": ("data", DATES, "", grouped(DATES, "Start:=DateSerial(2024, 1, 15), "
                                                                  "End:=DateSerial(2025, 3, 31), By:=7, "
                                                                  f"Periods:={periods('days')}")),
    "mixed_whole": ("mixed", "Whole", "", grouped("Whole", f"Start:=True, End:=True, Periods:={periods('hours')}")),
    "mixed_timed": ("mixed", "Timed", "", grouped("Timed", f"Start:=True, End:=True, Periods:={periods('hours')}")),
    # Round four: a value between End and the last group's top, one value alone, blanks and text, and a date format.
    "numbers_end_inside": ("data", AMOUNTS, "", grouped(AMOUNTS, "Start:=0, End:=82, By:=10")),
    "dates_by_7_months": ("data", DATES, "", grouped(DATES, f"Start:=True, End:=True, By:=7, Periods:="
                                                            f"{periods('days', 'months')}")),
    "mixed_same": ("mixed", "Same", "", lines(probe("default", f'{cell("Same")}.Group'),
                                              note("items", 'pt.PivotFields("Same").PivotItems.Count'))),
    "mixed_same_by": ("mixed", "Same", "", grouped("Same", "Start:=True, End:=True, By:=1")),
    "mixed_gappy": ("mixed", "Gappy", "", grouped("Gappy", "Start:=True, End:=True, By:=1")),
    "mixed_texty": ("mixed", "Texty", "", grouped("Texty", "Start:=True, End:=True, By:=1")),
    "mixed_iso_weeks": ("mixed", "Iso", "", grouped("Iso", f"Start:=True, End:=True, By:=7, Periods:={periods('days')}")),
    "mixed_iso_months": ("mixed", "Iso", "", grouped("Iso", f"Start:=True, End:=True, Periods:={periods('months')}")),
    **{f"span_{span_name(low, high)}": ("spans", span_name(low, high), "",
                                        probe("group", f"{cell(span_name(low, high))}.Group"))
       for low, high in SPANS},
    # Where an item's label and data cells are, which Range.Group of a union of items reads.
    "ranges_single": ("data", "Region", "", lines(*item_ranges("Region", "East"))),
    "ranges_compact": ("data", "Product,Region", "", lines(*item_ranges("Product", "Apple"),
                                                            *item_ranges("Region", "East"))),
    "ranges_outline": ("data", "Product,Region", "", lines("pt.RowAxisLayout 2", *item_ranges("Product", "Apple"),
                                                            *item_ranges("Region", "East"))),
    "ranges_tabular": ("data", "Product,Region", "", lines("pt.RowAxisLayout 1", *item_ranges("Product", "Apple"),
                                                            *item_ranges("Region", "East"))),
    "ranges_columns": ("data", "Product", "Region", lines(*item_ranges("Product", "Apple"),
                                                          *item_ranges("Region", "East"))),
    "ranges_two_columns": ("data", "", "Product,Region", lines(*item_ranges("Product", "Apple"),
                                                               *item_ranges("Region", "East"))),
    "ranges_values_rows": ("data", "Region", "", lines(f'pt.AddDataField pt.PivotFields("{AMOUNTS}"), '
                                                       '"Sum of Amount", xlSum',
                                                       'pt.DataPivotField.Orientation = xlRowField',
                                                       *item_ranges("Region", "East"))),
    "ranges_hidden": ("data", "Region", "", lines('pt.PivotFields("Region").PivotItems("East").Visible = False',
                                                  *item_ranges("Region", "East"))),
    "ranges_page": ("data", "Product", "", lines('pt.PivotFields("Region").Orientation = xlPageField',
                                                 *item_ranges("Region", "East"))),
    "ranges_collapsed": ("data", "Product,Region", "", lines(
        'pt.PivotFields("Product").PivotItems("Apple").ShowDetail = False', *item_ranges("Product", "Apple"),
        *item_ranges("Region", "East"))),
    "ranges_grouped": ("data", "Region", "", lines(probe("group", EAST_WEST), *item_ranges("Region2", "Group1"),
                                                   *item_ranges("Region", "East"), *item_ranges("Region", "North"))),
    # Round six: the rest of what groups say of each other, a refresh, groups other tables share, and a
    # calculated field beside them.
    "relations_more": ("data", "Region", "", lines(
        note("level_alone", 'pt.PivotFields("Region").GroupLevel'),
        note("total_alone", 'pt.PivotFields("Region").TotalLevels'),
        note("parent_alone", 'pt.PivotFields("Region").ParentField.Name'),
        note("child_alone", 'pt.PivotFields("Region").ChildField.Name'),
        note("parentitems_alone", 'pt.PivotFields("Region").ParentItems.Count'),
        note("parentitem_alone", 'pt.PivotFields("Region").PivotItems("East").ParentItem.Name'),
        note("childitems_alone", 'pt.PivotFields("Region").PivotItems("East").ChildItems.Count'),
        probe("group", EAST_WEST),
        note("group_total", 'pt.PivotFields("Region2").TotalLevels'),
        note("single_children", 'pt.PivotFields("Region2").PivotItems("North").ChildItems.Count'),
        note("single_child", 'pt.PivotFields("Region2").PivotItems("North").ChildItems(1).Name'),
        note("single_parent", 'pt.PivotFields("Region").PivotItems("North").ParentItem.Name'),
        note("base_parentitems", 'pt.PivotFields("Region").ParentItems.Count'),
        note("group_parentitem", 'pt.PivotFields("Region2").PivotItems("Group1").ParentItem.Name'),
        note("base_childitems", 'pt.PivotFields("Region").PivotItems("East").ChildItems.Count'))),
    "relations_levels": ("data", "Region", "", lines(
        probe("first", EAST_WEST), probe("second", union("Region2", "Group1", "North") + ".Group"),
        note("top", 'pt.PivotFields("Region3").GroupLevel'), note("middle", 'pt.PivotFields("Region2").GroupLevel'),
        note("bottom", 'pt.PivotFields("Region").GroupLevel'), note("total", 'pt.PivotFields("Region3").TotalLevels'),
        note("middle_parent", 'pt.PivotFields("Region2").ParentField.Name'),
        note("top_child", 'pt.PivotFields("Region3").ChildField.Name'),
        note("east_parent", 'pt.PivotFields("Region2").PivotItems("Group1").ParentItem.Name'))),
    "relations_date_levels": ("data", DATES, "", lines(
        grouped(DATES, MONTHS_YEARS),
        note("date_level", f'pt.PivotFields("{DATES}").GroupLevel'),
        note("date_total", f'pt.PivotFields("{DATES}").TotalLevels'),
        note("years_level", 'pt.PivotFields("Years (Date)").GroupLevel'),
        note("years_total", 'pt.PivotFields("Years (Date)").TotalLevels'),
        note("jan_parent", 'pt.PivotFields("Months (Date)").PivotItems("Jan").ParentItem.Name'),
        note("year_children", 'pt.PivotFields("Years (Date)").PivotItems("2024").ChildItems.Count'))),
    "relations_range": ("data", AMOUNTS, "", lines(
        grouped(AMOUNTS, "Start:=0, End:=100, By:=10"),
        note("level", f'pt.PivotFields("{AMOUNTS}").GroupLevel'),
        note("total", f'pt.PivotFields("{AMOUNTS}").TotalLevels'),
        note("parent", f'pt.PivotFields("{AMOUNTS}").ParentField.Name'),
        note("child", f'pt.PivotFields("{AMOUNTS}").ChildField.Name'),
        note("children", f'pt.PivotFields("{AMOUNTS}").PivotItems("10-19").ChildItems.Count'))),
    "refresh_numbers_auto": ("data", AMOUNTS, "", refreshed(grouped(AMOUNTS, "Start:=True, End:=True, By:=25"))),
    "refresh_numbers_fixed": ("data", AMOUNTS, "", refreshed(grouped(AMOUNTS, "Start:=0, End:=100, By:=10"))),
    "refresh_dates": ("data", DATES, "", refreshed(grouped(DATES, MONTHS_YEARS))),
    "refresh_dates_fixed": ("data", DATES, "", refreshed(grouped(DATES, "Start:=DateSerial(2024, 1, 1), "
                                                                       "End:=DateSerial(2025, 12, 31), "
                                                                       f"Periods:={periods('years')}"))),
    "refresh_manual": ("data", "Region", "", refreshed(probe("group", EAST_WEST))),
    "shared_range": ("data", AMOUNTS, "", lines(
        'Set pt2 = cache.CreatePivotTable(ws.Range("H3"), name & "_2")',
        f'pt2.PivotFields("{AMOUNTS}").Orientation = xlRowField',
        'pt2.AddDataField pt2.PivotFields("Units"), "Sum of Units", xlSum',
        grouped(AMOUNTS, "Start:=0, End:=100, By:=50"),
        'note = note & "second=" & Dump(pt2) & ";"')),
    "shared_dates": ("data", DATES, "", lines(
        'Set pt2 = cache.CreatePivotTable(ws.Range("H3"), name & "_2")',
        f'pt2.PivotFields("{DATES}").Orientation = xlRowField',
        'pt2.AddDataField pt2.PivotFields("Units"), "Sum of Units", xlSum',
        grouped(DATES, MONTHS_YEARS),
        'note = note & "second=" & Dump(pt2) & ";"')),
    "shared_manual": ("data", "Region", "", lines(
        'Set pt2 = cache.CreatePivotTable(ws.Range("H3"), name & "_2")',
        'pt2.PivotFields("Region").Orientation = xlRowField',
        'pt2.AddDataField pt2.PivotFields("Units"), "Sum of Units", xlSum',
        probe("group", EAST_WEST),
        'note = note & "second=" & Dump(pt2) & ";"')),
    "calculated_dates": ("data", DATES, "", lines(
        'pt.CalculatedFields.Add "Double", "=Units*2"',
        grouped(DATES, MONTHS_YEARS))),
    "calculated_manual": ("data", "Region", "", lines(
        'pt.CalculatedFields.Add "Double", "=Units*2"',
        probe("group", EAST_WEST))),
}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, data As Object, out As String, r As Long, rows As Variant, parts As Variant
Dim ws As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{source}{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\pivot_grouping.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\pivot_grouping.xlsx")
out = out & "@@"
For Each ws In wb.Worksheets
    If ws.PivotTables.Count > 0 Then out = out & ws.Name & "^" & Dump(ws.PivotTables(ws.Name)) & "|"
Next
wb.Close False
Build = out
End Function
'''

LAYOUT = 'out = out & Layout(wb, "{name}", "{source}", "{row}", "{column}", {step})\n'
REVISIONS = re.compile(r' xmlns:mc="[^"]*"| mc:Ignorable="[^"]*"| xmlns:xr="[^"]*"| xr:uid="[^"]*"| cacheId="\d+"'
                       r'| refreshedDate="[^"]*"| refreshedBy="[^"]*"| r:id="[^"]*"')


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's definition and its cache's fields, as the file keeps them, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8")
            rels = package.read(posixpath.join("xl/pivotTables/_rels", posixpath.basename(part) + ".rels"))
            target = re.search(rb'Target="([^"]*pivotCacheDefinition\d+\.xml)"', rels)
            cache = ""
            if target:
                cache = package.read(posixpath.normpath(posixpath.join("xl/pivotTables",
                                                                       target.group(1).decode()))).decode("utf-8")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            body = re.sub(r"^<\?xml[^>]*\?>\s*", "", text)
            fields = re.search(r"<cacheFields\b.*?</cacheFields>", cache, re.DOTALL)
            out[name.group(1) if name else part] = REVISIONS.sub("", body + (fields.group() if fields else ""))
    return out


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    steps = "".join(f"    Case {index}\n" + "".join(f"        {line}\n" for line in step.splitlines())
                    for index, (_name, (_source, _row, _column, step)) in enumerate(LAYOUTS.items(), 1))
    body = "".join(LAYOUT.replace("{name}", name).replace("{source}", SOURCES[source]).replace("{row}", row)
                   .replace("{column}", column).replace("{step}", str(index))
                   for index, (name, (source, row, column, _step)) in enumerate(LAYOUTS.items(), 1))
    code = (BUILD.replace("{steps}", steps).replace("{source}", SOURCE).replace("{layouts}", body)
            .replace("{folder}", str(OUT)))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    session, reopened = str(result.value).split("@@", 1)
    strip_save_path(OUT / "pivot_grouping.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = tables(OUT / "pivot_grouping.xlsx")
    record = {"build": BUILD, "source": SOURCE, "layout": LAYOUT, "steps": steps, "sources": SOURCES,
              "layouts": [{"name": name, "source": source, "row": row, "column": column, "step": step,
                           "dump": dumps.get(name, ""), "reopened": again.get(name, ""), "file": held.get(name, "")}
                          for name, (source, row, column, step) in LAYOUTS.items()]}
    (OUT / "pivot_grouping.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                             encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"])
        print("   session ", layout["dump"][:900])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"][:900])


if __name__ == "__main__":
    main()
