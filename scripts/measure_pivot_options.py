"""A pivot table's options, as Excel reads, sets, shows and saves them.

A macro sets a pivot table's style and its style options, the text it
shows for empty and error values, its page fields' order and wrap, its
flags, tag and alternative text, and lays it out again by AddFields,
ClearTable, Location and ManualUpdate. Every layout sits on a sheet of its
own in one new workbook, a pivot table of its own over one source list
of regions, products, channels, amounts and units: regions on the rows,
products on the columns, the sum of the amounts. The dump reads every
cell of the table, its ranges, every option and the fields it shows. A
layout with error values has a source of its own. The workbook is saved,
its pivot table and cache parts kept, and Excel opens it and dumps every
table again.

    python scripts/measure_pivot_options.py

writes tests/fixtures/pivot_options/: the workbook as Excel saved it, and
pivot_options.json, which tests/test_excel_pivot_options.py replays.
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

OUT = ROOT / "tests" / "fixtures" / "pivot_options"

OPTIONS = ("TableStyle2 ShowTableStyleRowHeaders ShowTableStyleColumnHeaders ShowTableStyleRowStripes "
           "ShowTableStyleColumnStripes ShowTableStyleLastColumn NullString DisplayNullString ErrorString "
           "DisplayErrorString ManualUpdate Location Tag AlternativeText Summary DisplayContextTooltips "
           "DisplayEmptyColumn DisplayEmptyRow DisplayImmediateItems EnableDataValueEditing EnableDrilldown "
           "EnableFieldDialog EnableFieldList EnableWizard FieldListSortAscending HasAutoFormat InGridDropZones "
           "PageFieldOrder PageFieldWrapCount PreserveFormatting PrintTitles RepeatItemsOnEachPrintedPage SaveData "
           "ShowDrillIndicators ShowValuesRow SortUsingCustomLists SubtotalHiddenPageItems")

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
            ' The harness hands a NUL back as a space; spell it out.
            out = out & " " & n & "=" & Replace(CStr(v), Chr(0), "<NUL>")
        End If
    Next
    Props = out
End Function

Private Function Spots(o As Object, names As String) As String
    Dim n As Variant, out As String, r As Object
    On Error Resume Next
    For Each n In Split(names, " ")
        Err.Clear
        Set r = Nothing
        Set r = CallByName(o, CStr(n), VbGet)
        If Err.Number <> 0 Then
            out = out & " " & n & "=E" & Err.Number
        ElseIf r Is Nothing Then
            out = out & " " & n & "=-"
        Else
            out = out & " " & n & "=" & r.Address(False, False)
        End If
    Next
    Spots = out
End Function

Private Function Dump(pt As Object) As String
    Dim out As String, c As Object, f As Object
    On Error Resume Next
    out = "range=" & pt.TableRange2.Address(False, False) & " cells="
    For Each c In pt.TableRange2.Cells
        out = out & CStr(c.Value)
        If VarType(c.Value) = vbError Then out = out & "~e"
        out = out & ","
    Next
    out = out & Spots(pt, "TableRange1 RowRange ColumnRange DataBodyRange PageRange")
    out = out & Props(pt, "{options}")
    Err.Clear
    out = out & " visible="
    For Each f In pt.VisibleFields
        out = out & f.Name & ":" & f.Orientation & ","
    Next
    If Err.Number <> 0 Then out = out & "E" & Err.Number
    Err.Clear
    out = out & " fields="
    For Each f In pt.PivotFields
        If f.Orientation <> 0 Then out = out & f.Name & ":" & f.Orientation & ":" & f.Position & ","
    Next
    For Each f In pt.DataFields
        out = out & f.Name & ":data,"
    Next
    Dump = out
End Function

Private Function Made(ws As Object, cache As Object, rowFields As String, columnFields As String, _
        pageFields As String) As Object
    Dim pt As Object, n As Variant
    Set pt = cache.CreatePivotTable(ws.Range("A5"), ws.Name)
    For Each n In Split(rowFields, ",")
        pt.PivotFields(CStr(n)).Orientation = xlRowField
    Next
    For Each n In Split(columnFields, ",")
        pt.PivotFields(CStr(n)).Orientation = xlColumnField
    Next
    For Each n In Split(pageFields, ",")
        pt.PivotFields(CStr(n)).Orientation = xlPageField
    Next
    pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum
    Set Made = pt
End Function

Private Function Layout(wb As Object, cache As Object, name As String, rowFields As String, columnFields As String, _
        pageFields As String, source As String, step As Long) As String
    Dim ws As Object, pt As Object, out As String, own As Object, held As Object, total As Object, x As Variant
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    Set own = cache
    If Len(source) > 0 Then
        Set held = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        held.Name = source
        held.Range("A1:E17").Value = wb.Worksheets("data").Range("A1:E17").Value
        If Left(source, 10) = "src_errors" Then held.Range("D3").Value = CVErr(xlErrNA)
        Set own = wb.PivotCaches.Create(xlDatabase, source & "!R1C1:R17C5")
    End If
    Set pt = Made(ws, own, rowFields, columnFields, pageFields)
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
'''.replace("{options}", OPTIONS)

#: The source list of tests/fixtures/pivot_layouts: sixteen rows of a region, a product, a channel, an amount and
#: units, West without cherries and North without bananas.
SOURCE = '''Set data = wb.Worksheets(1)
data.Name = "data"
data.Range("A1:E1").Value = Array("Region", "Product", "Channel", "Amount", "Units")
rows = Array("East|Apple|Online|10|1", "East|Banana|Store|20|2", "East|Cherry|Online|30|3", _
    "West|Apple|Store|40|4", "West|Banana|Online|50|5", "North|Apple|Online|60|6", "North|Cherry|Store|70|7", _
    "South|Banana|Store|80|8", "South|Cherry|Online|90|9", "East|Apple|Store|15|1", "West|Apple|Online|25|2", _
    "North|Apple|Store|35|3", "South|Apple|Store|45|4", "West|Banana|Store|55|5", "South|Banana|Online|65|6", _
    "North|Cherry|Online|75|7")
For r = 0 To 15
    parts = Split(rows(r), "|")
    data.Cells(r + 2, 1).Value = parts(0)
    data.Cells(r + 2, 2).Value = parts(1)
    data.Cells(r + 2, 3).Value = parts(2)
    data.Cells(r + 2, 4).Value = CLng(parts(3))
    data.Cells(r + 2, 5).Value = CLng(parts(4))
Next
Set cache = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R17C5")
'''


def probe(text: str, statement: str) -> str:
    """A statement, and its error number into the note under ``text`` if it raises one."""
    return (f"Err.Clear\n{statement}\n"
            f'If Err.Number <> 0 Then note = note & "{text}=E" & Err.Number & ";"\nErr.Clear\n')


def lines(*statements: str) -> str:
    return "\n".join(statements)


def flags(*pairs: tuple[str, str]) -> str:
    """Each option set, its error number noted under its name."""
    return lines(*(probe(name, f"pt.{name} = {value}") for name, value in pairs))


UNITS = 'pt.AddDataField pt.PivotFields("Units"), "Sum of Units", xlSum'

#: name -> (the row fields, the column fields, the page fields, a source sheet of the table's own or "" for the
#: shared one, what the layout does to its table).
LAYOUTS: dict[str, tuple[str, str, str, str, str]] = {
    "defaults": ("Region", "Product", "", "", ""),
    "style": ("Region", "Product", "", "", lines('pt.TableStyle2 = "PivotStyleMedium9"',
                                               "pt.ShowTableStyleRowStripes = True",
                                               "pt.ShowTableStyleColumnStripes = True",
                                               "pt.ShowTableStyleRowHeaders = False",
                                               "pt.ShowTableStyleColumnHeaders = False",
                                               "pt.ShowTableStyleLastColumn = False")),
    "style_lower": ("Region", "Product", "", "", 'pt.TableStyle2 = "pivotstylemedium4"'),
    "style_table": ("Region", "Product", "", "", 'pt.TableStyle2 = "TableStyleMedium2"'),
    "null_string": ("Region", "Product", "", "", 'pt.NullString = "-"'),
    "null_zero": ("Region", "Product", "", "", 'pt.NullString = "0"'),
    "null_off": ("Region", "Product", "", "", lines('pt.NullString = "-"', "pt.DisplayNullString = False")),
    "error_default": ("Region", "Product", "", "src_errors1", ""),
    "error_string": ("Region", "Product", "", "src_errors2", lines('pt.ErrorString = "err"',
                                                                   "pt.DisplayErrorString = True")),
    "error_string_only": ("Region", "Product", "", "src_errors3", 'pt.ErrorString = "err"'),
    "error_off": ("Region", "Product", "", "src_errors4", lines('pt.ErrorString = "err"',
                                                                "pt.DisplayErrorString = True",
                                                                "pt.DisplayErrorString = False")),
    "manual_pending": ("Region", "Product", "", "", lines("pt.ManualUpdate = True",
                                                          'pt.PivotFields("Channel").Orientation = xlRowField')),
    "manual_done": ("Region", "Product", "", "", lines("pt.ManualUpdate = True",
                                                       'pt.PivotFields("Channel").Orientation = xlRowField',
                                                       "pt.ManualUpdate = False")),
    "add_fields": ("Region", "Product", "", "", 'pt.AddFields RowFields:="Product", ColumnFields:="Channel"'),
    "add_fields_array": ("Region", "Product", "", "", 'pt.AddFields RowFields:=Array("Region", "Product"), '
                                                      'PageFields:="Channel"'),
    "add_fields_to_table": ("Region", "Product", "", "", 'pt.AddFields RowFields:="Channel", AddToTable:=True'),
    "add_fields_rows_only": ("Region", "Product", "", "", 'pt.AddFields RowFields:="Channel"'),
    "add_fields_none": ("Region", "Product", "", "", probe("none", "pt.AddFields")),
    "add_fields_bad": ("Region", "Product", "", "", probe("bad", 'pt.AddFields RowFields:="Nope"')),
    "add_fields_values": ("Region", "Product", "", "", lines(UNITS, 'pt.AddFields RowFields:=Array("Region", '
                                                                    '"Data"), ColumnFields:="Product"')),
    "add_fields_result": ("Region", "Product", "", "", lines('x = pt.AddFields(RowFields:="Channel")',
                                                             'note = note & "result=" & TypeName(x) & ":" & '
                                                             'CStr(x) & ";"')),
    "clear_table": ("Region", "Product", "", "", "pt.ClearTable"),
    "clear_then_add": ("Region", "Product", "", "", lines("pt.ClearTable", 'pt.PivotFields("Channel").Orientation '
                                                          "= xlRowField", 'pt.AddDataField pt.PivotFields("Units"), '
                                                          '"Sum of Units", xlSum')),
    "clear_with_page": ("Region", "Product", "Channel", "", "pt.ClearTable"),
    "location": ("Region", "Product", "", "", 'pt.Location = ws.Range("D8").Address'),
    "location_text": ("Region", "Product", "", "", 'pt.Location = "$F$10"'),
    "location_relative": ("Region", "Product", "", "", 'pt.Location = "G2"'),
    "location_sheet": ("Region", "Product", "", "", 'pt.Location = "\'" & ws.Name & "\'!$B$12"'),
    "location_bad": ("Region", "Product", "", "", probe("bad", 'pt.Location = "zz"')),
    "location_page": ("Region", "Product", "Channel", "", 'pt.Location = "$C$9"'),
    "page_order": ("Region", "", "Channel,Product", "", lines("pt.PageFieldOrder = xlOverThenDown",
                                                             "pt.PageFieldWrapCount = 1")),
    "page_wrap": ("Region", "", "Channel,Product", "", "pt.PageFieldWrapCount = 1"),
    "page_errors": ("Region", "", "Channel,Product", "", lines(probe("order", "pt.PageFieldOrder = 3"),
                                                              probe("wrap", "pt.PageFieldWrapCount = -1"),
                                                              probe("big", "pt.PageFieldWrapCount = 256"))),
    "hidden_page_items": ("Region", "", "Channel", "", lines('pt.PivotFields("Channel").EnableMultiplePageItems '
                                                             "= True", 'pt.PivotFields("Channel").PivotItems('
                                                             '"Online").Visible = False',
                                                             "pt.SubtotalHiddenPageItems = True")),
    "values_row": ("Region", "", "", "", lines(UNITS, "pt.ShowValuesRow = True")),
    "grid_drop_zones": ("Region", "Product", "", "", "pt.InGridDropZones = True"),
    "flags": ("Region", "Product", "", "", flags(
        ("DisplayContextTooltips", "False"), ("DisplayEmptyColumn", "True"), ("DisplayEmptyRow", "True"),
        ("DisplayImmediateItems", "False"), ("EnableDataValueEditing", "True"), ("EnableDrilldown", "False"),
        ("EnableFieldDialog", "False"), ("EnableFieldList", "False"), ("EnableWizard", "False"),
        ("FieldListSortAscending", "True"), ("HasAutoFormat", "False"), ("PreserveFormatting", "False"),
        ("PrintTitles", "True"), ("RepeatItemsOnEachPrintedPage", "False"), ("SaveData", "False"),
        ("ShowDrillIndicators", "False"), ("SortUsingCustomLists", "False"), ("Tag", '"t"'),
        ("AlternativeText", '"alt"'), ("Summary", '"sum"'))),
    "update": ("Region", "Product", "", "", "pt.Update"),
    "null_letter": ("Region", "Product", "", "", 'pt.NullString = "x"'),
    "null_refresh": ("Region", "Product", "", "src_null1", lines('pt.NullString = "-"', "pt.RefreshTable")),
    "null_display_off": ("Region", "Product", "", "", "pt.DisplayNullString = False"),
    "null_back": ("Region", "Product", "", "", lines('pt.NullString = "-"', 'pt.NullString = ""')),
    "null_then_layout": ("Region", "Product", "", "", lines('pt.NullString = "-"',
                                                            'pt.PivotFields("Channel").Orientation = xlRowField')),
    "null_before": ("Region", "Product", "", "", lines("pt.ClearTable", 'pt.NullString = "-"',
                                                       'pt.PivotFields("Region").Orientation = xlRowField',
                                                       'pt.PivotFields("Product").Orientation = xlColumnField',
                                                       'pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", '
                                                       "xlSum")),
    "manual_rows_only": ("Region", "", "", "", lines("pt.ManualUpdate = True",
                                                     'pt.PivotFields("Channel").Orientation = xlRowField')),
    "manual_page": ("Region", "Product", "Channel", "", lines("pt.ManualUpdate = True", UNITS)),
    "manual_tabular": ("Region", "Product", "", "", lines("pt.RowAxisLayout xlTabularRow", "pt.ManualUpdate = True",
                                                          'pt.PivotFields("Channel").Orientation = xlRowField')),
    "manual_two_data": ("Region", "Product", "", "", lines(UNITS, "pt.ManualUpdate = True",
                                                           'pt.PivotFields("Channel").Orientation = xlRowField')),
    "manual_unchanged": ("Region", "Product", "", "", "pt.ManualUpdate = True"),
    "manual_style": ("Region", "Product", "", "", lines("pt.ManualUpdate = True",
                                                        'pt.TableStyle2 = "PivotStyleMedium2"')),
    "manual_hidden_item": ("Region", "Product", "", "", lines("pt.ManualUpdate = True",
                                                              'pt.PivotFields("Region").PivotItems("East").Visible = '
                                                              "False")),
    "subtotal_hidden": ("Region", "", "Channel", "", probe("set", "pt.SubtotalHiddenPageItems = True")),
    "immediate_items": ("Region", "Product", "", "", probe("set", "pt.DisplayImmediateItems = False")),
    "save_data": ("Region", "Product", "", "", probe("set", "pt.SaveData = False")),
    "values_row_columns": ("Region", "Product", "", "", lines(UNITS, "pt.ShowValuesRow = True")),
    "grid_rows": ("Region,Channel", "Product", "", "", "pt.InGridDropZones = True"),
    "visible_values": ("Region", "Product", "Channel", "", UNITS),
    "add_fields_pages": ("Region", "Product", "", "", 'pt.AddFields PageFields:=Array("Channel", "Product")'),
    "add_fields_data_source": ("Region", "Product", "", "", 'pt.AddFields RowFields:="Amount"'),
    "add_fields_twice": ("Region", "Product", "", "", probe("twice", 'pt.AddFields RowFields:=Array("Region", '
                                                                      '"Region")')),
    "clear_filters": ("Region", "Product", "", "", lines('pt.PivotFields("Region").PivotFilters.Add '
                                                         'Type:=xlCaptionContains, Value1:="t"', "pt.ClearTable")),
}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, data As Object, cache As Object, out As String, r As Long, rows As Variant, parts As Variant
Dim ws As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{source}{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\pivot_options.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\pivot_options.xlsx")
out = out & "@@"
For Each ws In wb.Worksheets
    If ws.PivotTables.Count > 0 Then out = out & ws.Name & "^" & Dump(ws.PivotTables(1)) & "|"
Next
wb.Close False
Build = out
End Function
'''

LAYOUT = 'out = out & Layout(wb, cache, "{name}", "{row}", "{column}", "{page}", "{source}", {step})\n'


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's attributes, style and extensions, and its cache's attributes, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            kept = [found.group() for pattern in (r"<pivotTableDefinition\b[^>]*>", r"<location\b[^>]*/>",
                                                  r"<pivotTableStyleInfo\b[^>]*/>",
                                                  r"<extLst>(?:(?!</extLst>).)*</extLst>(?=</pivotTableDefinition>)")
                    if (found := re.search(pattern, text, re.DOTALL))]
            rels = package.read(part.replace("pivotTables/", "pivotTables/_rels/") + ".rels").decode("utf-8")
            target = re.search(r'Target="\.\./pivotCache/([^"]+)"', rels)
            cache = package.read(f"xl/pivotCache/{target.group(1)}").decode("utf-8") if target else ""
            cached = re.search(r"<pivotCacheDefinition\b[^>]*>", cache)
            out[name.group(1) if name else part] = "".join(kept) + (cached.group() if cached else "")
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
    session, reopened = str(result.value).split("@@", 1)
    strip_save_path(OUT / "pivot_options.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = tables(OUT / "pivot_options.xlsx")
    record = {"build": BUILD, "source": SOURCE, "layout": LAYOUT, "steps": steps,
              "layouts": [{"name": name, "row": row, "column": column, "page": page, "own_source": source,
                           "step": step, "dump": dumps.get(name, ""), "reopened": again.get(name, ""),
                           "file": held.get(name, "")}
                          for name, (row, column, page, source, step) in LAYOUTS.items()]}
    (OUT / "pivot_options.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                            encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"])
        print("   session ", layout["dump"])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"])
        print("   file    ", re.sub(r"\s+", " ", str(layout["file"]))[:2500])


if __name__ == "__main__":
    main()
