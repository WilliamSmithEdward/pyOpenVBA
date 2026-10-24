"""Which icon a conditional format's icon set shows each cell, as AutoFilter by icon finds it, and how it saves.

Excel shows an icon set's icons on the screen alone; no property reads
the icon a cell shows. AutoFilter with Operator:=xlFilterIcon and an Icon
as Criteria1 keeps the rows showing it, and xlFilterNoIcon those showing
none, so filtering by each icon in turn reads every cell's. Every layout
sits on a sheet of its own in one new workbook: numbers 10 to 100 under a
header, a word, a blank and an error, an icon set over them -- the
default traffic lights, arrows, five arrows, criteria on numbers, one
above rather than at its value, percentiles, a formula, the order turned
round, the values hidden, an icon taken out -- then a filter by each icon,
of the set or of another set, and by none. The dump after each reads the
hidden rows and the filter's On, Operator, Count and what Criteria1 is.
The last filter of each sheet, on an icon or on none, is saved, and Excel
opens the file and dumps every sheet again.

    python scripts/measure_icon_filters.py

writes tests/fixtures/icon_filters/: the workbook as Excel saved it, and
icon_filters.json, which tests/test_excel_icon_filters.py replays.
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

OUT = ROOT / "tests" / "fixtures" / "icon_filters"

HELPER = '''Private Function Filters(ws As Object) As String
    Dim out As String, f As Object, c As Object
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
            Set c = Nothing
            Set c = f.Criteria1
            If Err.Number <> 0 Then
                out = out & " c1=E" & Err.Number
            ElseIf c Is Nothing Then
                out = out & " c1=Nothing"
            Else
                out = out & " c1=" & TypeName(c) & ":" & c.Index
                If Err.Number <> 0 Then out = out & "E" & Err.Number
            End If
        End If
    End If
    Filters = out
End Function

Private Function Dump(ws As Object) As String
    Dim out As String, r As Long
    For r = 1 To 15
        If ws.Rows(r).Hidden Then out = out & r & ","
    Next
    Dump = "hidden=" & out & " " & Filters(ws)
End Function

Private Function Fill(ws As Object) As Object
    Dim r As Long
    ws.Range("A1").Value = "N"
    For r = 2 To 11
        ws.Cells(r, 1).Value = (r - 1) * 10
    Next
    ws.Range("A12").Value = "word"
    ws.Range("A14").Formula = "=1/0"
    ws.Range("A15").Value = 55
    Set Fill = ws.Range("A2:A15").FormatConditions.AddIconSetCondition
End Function

Private Function Layout(wb As Object, name As String, setId As Long, count As Long, filterSet As Long, _
        last As Long) As String
    Dim ws As Object, f As Object, out As String, k As Long, shown As Long
    If wb.Worksheets(1).Name = "Sheet1" Then
        Set ws = wb.Worksheets(1)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    End If
    ws.Name = name
    Set f = Fill(ws)
    f.IconSet = wb.IconSets(setId)
    Select Case name
{cases}
    End Select
    shown = wb.IconSets(filterSet).Count
    For k = 1 To shown
        On Error Resume Next
        Err.Clear
        ws.Range("A1:A15").AutoFilter Field:=1, Criteria1:=wb.IconSets(filterSet).Item(k), Operator:=xlFilterIcon
        If Err.Number <> 0 Then out = out & "icon" & k & "=E" & Err.Number & ";"
        On Error GoTo 0
        out = out & "icon" & k & "=" & Dump(ws) & ";"
    Next
    On Error Resume Next
    Err.Clear
    ws.Range("A1:A15").AutoFilter Field:=1, Operator:=xlFilterNoIcon
    If Err.Number <> 0 Then out = out & "none=E" & Err.Number & ";"
    On Error GoTo 0
    out = out & "none=" & Dump(ws) & ";"
    If last > 0 Then
        ws.Range("A1:A15").AutoFilter Field:=1, Criteria1:=wb.IconSets(filterSet).Item(last), Operator:=xlFilterIcon
    End If
    Layout = name & "^" & out & "|"
End Function
'''

#: name -> (set, icons, what the layout does to the set's criteria).
LAYOUTS: dict[str, tuple[int, int, str]] = {
    "lights": (4, 3, ""),
    "arrows": (1, 3, ""),
    "five_arrows": (14, 5, ""),
    "numbers": (1, 3, 'f.IconCriteria(2).Type = xlConditionValueNumber\nf.IconCriteria(2).Value = 30\n'
                      'f.IconCriteria(3).Type = xlConditionValueNumber\nf.IconCriteria(3).Value = 70\n'),
    "above": (1, 3, 'f.IconCriteria(2).Type = xlConditionValueNumber\nf.IconCriteria(2).Value = 30\n'
                    'f.IconCriteria(2).Operator = xlGreater\nf.IconCriteria(3).Type = xlConditionValueNumber\n'
                    'f.IconCriteria(3).Value = 70\nf.IconCriteria(3).Operator = xlGreater\n'),
    "percentiles": (1, 3, 'f.IconCriteria(2).Type = xlConditionValuePercentile\nf.IconCriteria(2).Value = 50\n'
                          'f.IconCriteria(3).Type = xlConditionValuePercentile\nf.IconCriteria(3).Value = 90\n'),
    "formula": (1, 3, 'f.IconCriteria(2).Type = xlConditionValueFormula\n'
                      'f.IconCriteria(2).Value = "=AVERAGE($A$2:$A$11)"\n'),
    "reversed": (1, 3, "f.ReverseOrder = True\n"),
    "icons_only": (1, 3, "f.ShowIconOnly = True\n"),
    "no_icon": (1, 3, "f.IconCriteria(2).Icon = xlIconNoCellIcon\n"),
    "saved_no_icon": (1, 3, ""),
    "other_set": (1, 3, ""),
    "lights_on_arrows": (1, 3, ""),
}
#: The layouts filtered by the icons of another set than their own, and those whose last filter, the one saved,
#: is on no icon; the rest are saved filtered by their set's last icon.
FILTER_SETS = {"other_set": 14, "lights_on_arrows": 4}
NO_ICON_LAST = {"saved_no_icon"}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, ws As Object, out As String
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\icon_filters.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\icon_filters.xlsx")
out = out & "~"
For Each ws In wb.Worksheets
    out = out & ws.Name & "^" & Dump(ws) & "|"
Next
wb.Close False
Build = out
End Function
'''

LAYOUT = 'out = out & Layout(wb, "{name}", {set}, {count}, {filter_set}, {last})\n'


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
    cases = "".join(f'    Case "{name}"\n' + "".join(f"        {line}\n" for line in steps.splitlines())
                    for name, (_set, _count, steps) in LAYOUTS.items() if steps)
    body = "".join(LAYOUT.replace("{name}", name).replace("{set}", str(set_id)).replace("{count}", str(count))
                   .replace("{filter_set}", str(FILTER_SETS.get(name, set_id)))
                   .replace("{last}", "0" if name in NO_ICON_LAST else str(count))
                   for name, (set_id, count, _steps) in LAYOUTS.items())
    code = BUILD.replace("{cases}", cases).replace("{layouts}", body).replace("{folder}", str(OUT))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    session, reopened = str(result.value).split("~", 1)
    strip_save_path(OUT / "icon_filters.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = file_filters(OUT / "icon_filters.xlsx")
    record = {"build": BUILD, "layout": LAYOUT, "cases": cases,
              "layouts": [{"name": name, "set": set_id, "count": count, "steps": steps,
                           "filter_set": FILTER_SETS.get(name, set_id), "last": 0 if name in NO_ICON_LAST else count,
                           "dumps": dumps.get(name, ""), "reopened": again.get(name, ""), "file": held.get(name, [])}
                          for name, (set_id, count, steps) in LAYOUTS.items()]}
    (OUT / "icon_filters.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"])
        for part in str(layout["dumps"]).split(";"):
            if part:
                print("  ", part)
        print("   reopened", layout["reopened"])
        print("   file    ", [re.sub(r' xr:uid="[^"]*"', "", one) for one in layout["file"]])


if __name__ == "__main__":
    main()
