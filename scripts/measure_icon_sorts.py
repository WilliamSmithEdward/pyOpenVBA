"""Sorting on the icon a conditional format's icon set shows, as the Sort object a recorded macro uses does it.

Every layout sits on a sheet of its own in one new workbook: a list of
eight rows under a header, named r2 to r9 over numbers an icon set of
three arrows covers, sorted by an icon one way or another. The dump reads
each row's name in order, then the sort fields the sheet keeps: what each
sorts on, its order, what its SortOnValue is and the index of the icon it
holds. The layouts ask where rows of the icon go, on top or at the bottom,
how two icons and an icon beside values combine, what an icon no row
shows, an icon of another set and a field never given an icon do. Then
the workbook is saved and opened again: each sheet's sortState goes beside
the fields as Excel reads them back.

    python scripts/measure_icon_sorts.py

writes tests/fixtures/icon_sorts/: icon_sorts.xlsx as Excel saved it, and
icon_sorts.json, which tests/test_excel_icon_sorts.py replays.
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

OUT = ROOT / "tests" / "fixtures" / "icon_sorts"

HELPER = '''Private Function Dump(ws As Object) As String
    Dim out As String, r As Long, f As Object, i As Long, v As Object
    For r = 2 To 9
        out = out & ws.Cells(r, 1).Value & ","
    Next
    On Error Resume Next
    For i = 1 To ws.Sort.SortFields.Count
        Set f = ws.Sort.SortFields(i)
        out = out & " f" & i & ":" & f.Key.Address & ":" & f.SortOn & ":" & f.Order
        Err.Clear
        Set v = Nothing
        Set v = f.SortOnValue
        If Err.Number <> 0 Then
            out = out & ":E" & Err.Number
        ElseIf v Is Nothing Then
            out = out & ":Nothing"
        Else
            out = out & ":" & TypeName(v)
            out = out & ":" & v.Index
            If Err.Number <> 0 Then out = out & ":E" & Err.Number
        End If
    Next
    Dump = out
End Function

Private Sub Fill(ws As Object)
    Dim i As Long, f As Object
    ws.Range("A1:B1").Value = Array("Name", "N")
    For i = 2 To 9
        ws.Cells(i, 1).Value = "r" & i
        ws.Cells(i, 2).Value = Choose(i - 1, 50, 30, 80, 10, 70, 20, 60, 40)
    Next
    Set f = ws.Range("B2:B9").FormatConditions.AddIconSetCondition
    f.IconSet = ws.Parent.IconSets(xl3Arrows)
End Sub

Private Sub Sorted(ws As Object)
    ws.Sort.SetRange ws.Range("A1:B9")
    ws.Sort.Header = xlYes
    ws.Sort.Apply
End Sub
'''

#: name -> what the layout does once its sheet holds the list.
LAYOUTS: dict[str, str] = {
    "up_on_top": 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add(Key:=ws.Range("B2:B9"), SortOn:=xlSortOnIcon, '
                 'Order:=xlAscending).SetIcon wb.IconSets(xl3Arrows).Item(3)\nSorted ws\n',
    "up_at_bottom": 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add(Key:=ws.Range("B2:B9"), SortOn:=xlSortOnIcon, '
                    'Order:=xlDescending).SetIcon wb.IconSets(xl3Arrows).Item(3)\nSorted ws\n',
    "down_on_top": 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add(Key:=ws.Range("B2:B9"), SortOn:=xlSortOnIcon, '
                   'Order:=xlAscending).SetIcon wb.IconSets(xl3Arrows).Item(1)\nSorted ws\n',
    "two_icons": 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add(Key:=ws.Range("B2:B9"), SortOn:=xlSortOnIcon, '
                 'Order:=xlAscending).SetIcon wb.IconSets(xl3Arrows).Item(1)\n'
                 'ws.Sort.SortFields.Add(Key:=ws.Range("B2:B9"), SortOn:=xlSortOnIcon, '
                 'Order:=xlAscending).SetIcon wb.IconSets(xl3Arrows).Item(2)\nSorted ws\n',
    "icon_then_values": 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add(Key:=ws.Range("B2:B9"), '
                        'SortOn:=xlSortOnIcon, Order:=xlAscending).SetIcon wb.IconSets(xl3Arrows).Item(2)\n'
                        'ws.Sort.SortFields.Add Key:=ws.Range("B2:B9"), SortOn:=xlSortOnValues, '
                        'Order:=xlAscending\nSorted ws\n',
    "other_set": 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add(Key:=ws.Range("B2:B9"), SortOn:=xlSortOnIcon, '
                 'Order:=xlAscending).SetIcon wb.IconSets(xl5Arrows).Item(5)\nSorted ws\n',
    "never_set": 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add Key:=ws.Range("B2:B9"), SortOn:=xlSortOnIcon, '
                 'Order:=xlAscending\nSorted ws\n',
    "middle_on_top": 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add(Key:=ws.Range("B2:B9"), '
                     'SortOn:=xlSortOnIcon, Order:=xlAscending).SetIcon wb.IconSets(xl3Arrows).Item(2)\nSorted ws\n',
}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, ws As Object, out As String
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\icon_sorts.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\icon_sorts.xlsx")
out = out & "~"
For Each ws In wb.Worksheets
    out = out & ws.Name & "^" & Dump(ws) & "|"
Next
wb.Close False
Build = out
End Function
'''

LAYOUT = '''If wb.Worksheets(1).Name = "Sheet1" Then
    Set ws = wb.Worksheets(1)
Else
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
End If
ws.Name = "{name}"
Fill ws
On Error Resume Next
Err.Clear
{steps}If Err.Number <> 0 Then out = out & "{name}^E" & Err.Number & " " & Err.Description & "|"
On Error GoTo 0
out = out & "{name}^" & Dump(ws) & "|"
'''


def sort_states(path: Path) -> dict[str, list[str]]:
    """Each sheet's sortState element as the file keeps it."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
        targets.update({key: value for value, key in re.findall(r'Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
        out: dict[str, list[str]] = {}
        for name, relationship in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
            text = package.read("xl/" + targets[relationship].lstrip("/").removeprefix("xl/")).decode("utf-8")
            out[name] = re.findall(r"<sortState\b.*?</sortState>|<sortState\b[^>]*/>", text, re.DOTALL)
    return out


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    body = "".join(LAYOUT.replace("{name}", name).replace("{steps}", steps) for name, steps in LAYOUTS.items())
    code = BUILD.replace("{layouts}", body).replace("{folder}", str(OUT))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    session, reopened = str(result.value).split("~", 1)
    strip_save_path(OUT / "icon_sorts.xlsx")
    entries = [one.split("^", 1) for one in session.split("|") if "^" in one]
    errors = {name: text for name, text in entries if text.startswith("E")}
    dumps = {name: text for name, text in entries if not text.startswith("E")}
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = sort_states(OUT / "icon_sorts.xlsx")
    record = {"build": BUILD, "layout": LAYOUT,
              "layouts": [{"name": name, "steps": steps, "error": errors.get(name, ""), "dump": dumps.get(name, ""),
                           "reopened": again.get(name, ""), "file": held.get(name, [])}
                          for name, steps in LAYOUTS.items()]}
    (OUT / "icon_sorts.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"], layout["error"])
        print("   session ", layout["dump"])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"])
        print("   file    ", layout["file"])


if __name__ == "__main__":
    main()
