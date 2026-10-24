"""AutoFilter by dates grouped by year, month, day, hour, minute or second, as Excel filters, reads back and saves.

A macro the recorder writes for a date column filtered by its drop-down
names the dates as Criteria2:=Array(level, "date", ...) with
Operator:=xlFilterValues, the level 0 for a year to 5 for a second.
Every layout sits on a sheet of its own in one new workbook: a header
and twelve rows whose first column holds dates and times across years,
months and days, a date written as text, a number in General, a blank
and a word, and whose second holds words; then one filter. The dump reads
which rows are hidden, the sheet's FilterMode, and each filter's On,
Operator, Count, Criteria1 and Criteria2. The workbook is saved, the
file's filters kept, and Excel opens it and dumps every sheet again.

    python scripts/measure_date_filters.py

writes tests/fixtures/date_filters/: the workbook as Excel saved it, and
date_filters.json, which tests/test_excel_date_filters.py replays.
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

OUT = ROOT / "tests" / "fixtures" / "date_filters"

HELPER = '''Private Function Show(v As Variant) As String
    Dim i As Long, out As String
    If IsArray(v) Then
        out = "Array("
        For i = LBound(v) To UBound(v)
            out = out & Show(v(i)) & ","
        Next
        Show = out & ")"
    ElseIf IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsError(v) Then
        Show = "Error:" & CStr(CLng(v))
    ElseIf IsObject(v) Then
        Show = "Object"
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function Filters(ws As Object) As String
    Dim out As String, i As Long, f As Object, part As String
    On Error Resume Next
    out = "filtering=" & ws.FilterMode
    If ws.AutoFilterMode Then
        For i = 1 To ws.AutoFilter.Filters.Count
            Set f = ws.AutoFilter.Filters(i)
            part = " f" & i & ":" & f.On
            If f.On Then
                Err.Clear
                part = part & " op=" & f.Operator & " n=" & f.Count
                If Err.Number <> 0 Then part = part & " E" & Err.Number
                Err.Clear
                part = part & " c1=" & Show(f.Criteria1)
                If Err.Number <> 0 Then part = part & "E" & Err.Number
                Err.Clear
                part = part & " c2=" & Show(f.Criteria2)
                If Err.Number <> 0 Then part = part & "E" & Err.Number
            End If
            out = out & part
        Next
    End If
    Filters = out
End Function

Private Function Dump(ws As Object) As String
    Dim out As String, r As Long
    For r = 1 To 14
        If ws.Rows(r).Hidden Then out = out & r & ","
    Next
    Dump = "hidden=" & out & " " & Filters(ws)
End Function
'''

#: The rows every layout filters.
TABLE = '''ws.Range("A1:B1").Value = Array("When", "Word")
ws.Range("A2").Value = #3/15/2025#
ws.Range("A3").Value = #3/15/2025 10:30:15 AM#
ws.Range("A4").Value = #3/15/2025 10:45:00 AM#
ws.Range("A5").Value = #3/16/2025#
ws.Range("A6").Value = #4/1/2025#
ws.Range("A7").Value = #3/15/2024#
ws.Range("A8").Value = #12/31/2025 11:59:59 PM#
ws.Range("A9").Value = "'3/15/2025"
ws.Range("A10").Value = 45731
ws.Range("A10").NumberFormat = "General"
ws.Range("A12").Value = "word"
ws.Range("A13").Value = #1/1/2026#
For r = 2 To 13
    ws.Cells(r, 2).Value = "w" & r
Next
'''


def filt(arguments: str) -> str:
    return f'ws.Range("A1:B13").AutoFilter {arguments}\n'


#: name -> what the layout does once its sheet holds the rows.
LAYOUTS: dict[str, str] = {
    "year": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(0, "2025")'),
    "year_as_date": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(0, "6/30/2025")'),
    "month": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(1, "3/1/2025")'),
    "month_mid": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(1, "3/20/2025")'),
    "day": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(2, "3/15/2025")'),
    "hour": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(3, "3/15/2025 10:00:00")'),
    "minute": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(4, "3/15/2025 10:30:00")'),
    "second": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(5, "3/15/2025 10:30:15")'),
    "two_days": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(2, "3/15/2025", 2, "4/1/2025")'),
    "day_and_year": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(2, "3/16/2025", 0, "1/1/2024")'),
    "day_with_time": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(2, "3/15/2025 10:30:00")'),
    "day_and_hour": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(2, "3/16/2025", '
                         '3, "3/15/2025 10:00:00")'),
    "dates_and_words": filt('Field:=1, Criteria1:="word", Operator:=xlFilterValues, '
                            'Criteria2:=Array(2, "3/15/2025")'),
    "dates_and_one_word": filt('Field:=1, Criteria1:=Array("word"), Operator:=xlFilterValues, '
                               'Criteria2:=Array(2, "3/15/2025")'),
    "dates_and_two_words": filt('Field:=1, Criteria1:=Array("word", "x"), Operator:=xlFilterValues, '
                                'Criteria2:=Array(1, "1/1/2026")'),
    "dates_and_three_words": filt('Field:=1, Criteria1:=Array("word", "x", "y"), Operator:=xlFilterValues, '
                                  'Criteria2:=Array(1, "1/1/2026")'),
    "words_only": filt('Field:=1, Criteria1:=Array("word", "x", "y"), Operator:=xlFilterValues'),
    "dates_and_blank": filt('Field:=1, Criteria1:="=", Operator:=xlFilterValues, '
                            'Criteria2:=Array(0, "1/1/2026")'),
    "recorder_style": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(0, "1/1/2024", 1, "4/1/2025", '
                           '2, "12/31/2025")'),
    "refiltered": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(0, "1/1/2025")')
    + filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(0, "1/1/2024")'),
    "order_years": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(0, "1/1/2024", 0, "1/1/2026", '
                        '0, "1/1/1999")'),
    "order_days": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(2, "3/16/2025", 2, "3/15/2025", '
                       '2, "3/15/2024")'),
    "order_levels": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(5, "3/15/2025 10:30:15", '
                         '0, "1/1/2023", 4, "3/15/2025 10:31:00", 1, "2/1/2025", 2, "1/1/2026", '
                         '3, "3/15/2025 09:00:00")'),
    "order_overlap": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(1, "3/1/2025", 0, "1/1/2025")'),
    "order_duplicate": filt('Field:=1, Operator:=xlFilterValues, Criteria2:=Array(2, "3/15/2025", '
                            '2, "3/15/2025")'),
    "order_words": filt('Field:=1, Criteria1:=Array("x", "word", "W2"), Operator:=xlFilterValues, '
                        'Criteria2:=Array(0, "1/1/2024")'),
}
#: A sheet of the saved file given each column filter in the classic form, a filters element of dateGroupItems,
#: for Excel to open: the name of the copy, the sheet, and the filterColumn's content.
CLASSIC: dict[str, tuple[str, str]] = {
    "classic_day": ("day", '<filters><dateGroupItem year="2025" month="3" day="15" dateTimeGrouping="day"/>'
                           '</filters>'),
    "classic_month_and_word": ("dates_and_two_words", '<filters><filter val="word"/><dateGroupItem year="2026" '
                                                      'month="1" dateTimeGrouping="month"/></filters>'),
    "classic_year": ("year_as_date", '<filters><dateGroupItem year="2025" dateTimeGrouping="year"/></filters>'),
}

READ = HELPER + '''
Public Function Read(path As String, saved As String) As String
Dim wb As Object, ws As Object, out As String
Application.DisplayAlerts = False
Set wb = Workbooks.Open(path)
For Each ws In wb.Worksheets
    out = out & ws.Name & "^" & Dump(ws) & "|"
Next
wb.SaveAs saved, 51
wb.Close False
Read = out
End Function
'''

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, ws As Object, out As String, r As Long, first As Boolean
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
first = True
{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\date_filters.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\date_filters.xlsx")
out = out & "~"
For Each ws In wb.Worksheets
    out = out & ws.Name & "^" & Dump(ws) & "|"
Next
wb.Close False
Build = out
End Function
'''

LAYOUT = '''If first Then
    Set ws = wb.Worksheets(1)
    first = False
Else
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
End If
ws.Name = "{name}"
{table}On Error Resume Next
{steps}If Err.Number <> 0 Then out = out & "{name}^E" & Err.Number & " " & Err.Description & "|"
On Error GoTo 0
out = out & "{name}^" & Dump(ws) & "|"
'''


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
    body = "".join(LAYOUT.replace("{name}", name).replace("{table}", TABLE).replace("{steps}", steps)
                   for name, steps in LAYOUTS.items())
    code = BUILD.replace("{layouts}", body).replace("{folder}", str(OUT))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    session, reopened = str(result.value).split("~", 1)
    strip_save_path(OUT / "date_filters.xlsx")
    entries = [one.split("^", 1) for one in session.split("|") if "^" in one]
    errors = {name: text for name, text in entries if text.startswith("E")}
    dumps = {name: text for name, text in entries if not text.startswith("E")}
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = file_filters(OUT / "date_filters.xlsx")
    classic: list[dict[str, object]] = []
    for name, (sheet, content) in CLASSIC.items():
        classic_copy(OUT / "date_filters.xlsx", OUT / f"{name}.xlsx", sheet, content)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name, (sheet, content) in CLASSIC.items():
            result = excel.run_vba(READ, "Read", (str(OUT / f"{name}.xlsx"), str(OUT / f"{name}_resaved.xlsx")),
                                   timeout=600.0)
            assert result.ok, f"{result.outcome}: {result.message} {result.error}"
            strip_save_path(OUT / f"{name}_resaved.xlsx")
            shown = dict(one.split("^", 1) for one in str(result.value).split("|") if "^" in one)
            classic.append({"name": name, "sheet": sheet, "content": content, "dump": shown[sheet],
                            "resaved": file_filters(OUT / f"{name}_resaved.xlsx")[sheet]})
    record = {"helper": HELPER, "table": TABLE, "layout": LAYOUT, "build": BUILD, "read": READ,
              "layouts": [{"name": name, "steps": steps, "error": errors.get(name, ""), "dump": dumps.get(name, ""),
                           "reopened": again.get(name, ""), "file": held.get(name, [])}
                          for name, steps in LAYOUTS.items()],
              "classic": classic}
    (OUT / "date_filters.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"], layout["error"])
        print("   session ", layout["dump"])
        print("   reopened", layout["reopened"])
        print("   file    ", [re.sub(r' xr:uid="[^"]*"', "", one) for one in layout["file"]])
    for one in classic:
        print("== classic", one["name"])
        print("   dump    ", one["dump"])
        print("   resaved ", [re.sub(r' xr:uid="[^"]*"', "", element) for element in one["resaved"]])


def classic_copy(source: Path, target: Path, sheet: str, content: str) -> None:
    """``source`` with the first column filter of ``sheet`` holding ``content``."""
    with zipfile.ZipFile(source) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
        targets.update({key: value for value, key in re.findall(r'Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
        relationship = dict(re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook))[sheet]
        part = "xl/" + targets[relationship].lstrip("/").removeprefix("xl/")
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
            for info in package.infolist():
                data = package.read(info.filename)
                if info.filename == part:
                    text = data.decode("utf-8")
                    text = re.sub(r'(<filterColumn colId="0">).*?(</filterColumn>)',
                                  lambda match: match.group(1) + content + match.group(2), text, count=1,
                                  flags=re.DOTALL)
                    data = text.encode("utf-8")
                out.writestr(info, data)


if __name__ == "__main__":
    main()
