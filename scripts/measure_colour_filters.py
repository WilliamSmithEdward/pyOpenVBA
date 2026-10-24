"""AutoFilter by a cell's fill or font colour, as Excel filters, reads back, saves and opens it again.

Every layout sits on a sheet of its own in one new workbook: a header and
ten rows whose first column has fills -- yellow, red, white, none, a theme
colour, a pattern over yellow -- whose second has font colours -- red,
blue, black, automatic -- and whose third a conditional format paints
green over 5; then one or two filters by colour, xlFilterCellColor,
xlFilterFontColor, xlFilterNoFill or xlFilterAutomaticFontColor, on the
sheet's range or on a table's. The dump reads which rows are hidden, the
sheet's FilterMode, and each filter's On, Operator, Count, what Criteria1
is and its Color. The workbook is saved, the file's filters and their
dxfs kept, and Excel opens it and dumps every sheet again, and saves it.

    python scripts/measure_colour_filters.py

writes tests/fixtures/colour_filters/: the workbook as Excel saved it and
saved again, and colour_filters.json, which
tests/test_excel_colour_filters.py replays.
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

OUT = ROOT / "tests" / "fixtures" / "colour_filters"

HELPER = '''Private Function Filters(ws As Object) As String
    Dim out As String, i As Long, f As Object, part As String, c As Object
    On Error Resume Next
    out = "filtering=" & ws.FilterMode
    If ws.AutoFilterMode Or ws.ListObjects.Count > 0 Then
        Dim af As Object
        Set af = Nothing
        If ws.ListObjects.Count > 0 Then Set af = ws.ListObjects(1).AutoFilter Else Set af = ws.AutoFilter
        For i = 1 To af.Filters.Count
            Set f = af.Filters(i)
            part = " f" & i & ":" & f.On
            If f.On Then
                Err.Clear
                part = part & " op=" & f.Operator & " n=" & f.Count
                If Err.Number <> 0 Then part = part & " E" & Err.Number
                Err.Clear
                Set c = Nothing
                Set c = f.Criteria1
                If Err.Number = 424 Then
                    Err.Clear
                    part = part & " c1=" & TypeName(f.Criteria1) & ":" & CStr(f.Criteria1)
                    If Err.Number <> 0 Then part = part & "E" & Err.Number
                ElseIf Err.Number <> 0 Then
                    part = part & " c1=E" & Err.Number
                ElseIf c Is Nothing Then
                    part = part & " c1=" & TypeName(f.Criteria1)
                Else
                    part = part & " c1=" & TypeName(c) & ":" & c.Color & ":" & c.Pattern
                    If Err.Number <> 0 Then part = part & "E" & Err.Number
                End If
                Err.Clear
                part = part & " c2=" & TypeName(f.Criteria2)
                If Err.Number <> 0 Then part = part & "E" & Err.Number
            End If
            out = out & part
        Next
    End If
    Filters = out
End Function

Private Function Dump(ws As Object) As String
    Dim out As String, r As Long
    For r = 1 To 12
        If ws.Rows(r).Hidden Then out = out & r & ","
    Next
    Dump = "hidden=" & out & " " & Filters(ws)
End Function
'''

#: The rows every layout filters.
TABLE = '''ws.Range("A1:D1").Value = Array("Fill", "Font", "Rule", "Value")
For r = 2 To 11
    ws.Cells(r, 1).Value = "f" & r
    ws.Cells(r, 2).Value = "t" & r
    ws.Cells(r, 3).Value = r - 1
    ws.Cells(r, 4).Value = r * 10
Next
ws.Range("A2,A5").Interior.Color = RGB(255, 255, 0)
ws.Range("A3").Interior.Color = RGB(255, 0, 0)
ws.Range("A6").Interior.Color = RGB(255, 255, 255)
ws.Range("A7").Interior.ThemeColor = xlThemeColorAccent1
ws.Range("A8").Interior.Pattern = xlGray50
ws.Range("A8").Interior.PatternColor = RGB(255, 0, 0)
ws.Range("A8").Interior.Color = RGB(255, 255, 0)
ws.Range("B2,B5").Font.Color = RGB(255, 0, 0)
ws.Range("B3").Font.Color = RGB(0, 0, 255)
ws.Range("B6").Font.Color = RGB(0, 0, 0)
Set f = ws.Range("C2:C11").FormatConditions.Add(xlCellValue, xlGreater, "5")
f.Interior.Color = RGB(0, 255, 0)
'''


def filt(arguments: str) -> str:
    return f'ws.Range("A1:D11").AutoFilter {arguments}\n'


#: name -> what the layout does once its sheet holds the rows.
LAYOUTS: dict[str, str] = {
    "cell_yellow": filt("Field:=1, Criteria1:=RGB(255, 255, 0), Operator:=xlFilterCellColor"),
    "cell_red": filt("Field:=1, Criteria1:=RGB(255, 0, 0), Operator:=xlFilterCellColor"),
    "cell_white": filt("Field:=1, Criteria1:=RGB(255, 255, 255), Operator:=xlFilterCellColor"),
    "cell_absent": filt("Field:=1, Criteria1:=RGB(1, 2, 3), Operator:=xlFilterCellColor"),
    "no_fill": filt("Field:=1, Operator:=xlFilterNoFill"),
    "theme_fill": filt('Field:=1, Criteria1:=ws.Range("A7").Interior.Color, Operator:=xlFilterCellColor'),
    "font_red": filt("Field:=2, Criteria1:=RGB(255, 0, 0), Operator:=xlFilterFontColor"),
    "font_black": filt("Field:=2, Criteria1:=RGB(0, 0, 0), Operator:=xlFilterFontColor"),
    "font_automatic": filt("Field:=2, Operator:=xlFilterAutomaticFontColor"),
    "rule_green": filt("Field:=3, Criteria1:=RGB(0, 255, 0), Operator:=xlFilterCellColor"),
    "rule_no_fill": filt("Field:=3, Operator:=xlFilterNoFill"),
    "two_fields": filt("Field:=1, Criteria1:=RGB(255, 255, 0), Operator:=xlFilterCellColor")
    + filt("Field:=2, Criteria1:=RGB(255, 0, 0), Operator:=xlFilterFontColor"),
    "refiltered": filt("Field:=1, Criteria1:=RGB(255, 255, 0), Operator:=xlFilterCellColor")
    + filt("Field:=1, Criteria1:=RGB(255, 0, 0), Operator:=xlFilterCellColor"),
    "shown_again": filt("Field:=1, Criteria1:=RGB(255, 255, 0), Operator:=xlFilterCellColor")
    + "ws.ShowAllData\n",
    "table": 'Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:D11"), , xlYes)\n'
    'lo.TableStyle = ""\n'
    "lo.Range.AutoFilter Field:=1, Criteria1:=RGB(255, 255, 0), Operator:=xlFilterCellColor\n",
}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, ws As Object, f As Object, lo As Object, out As String, r As Long, first As Boolean
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
first = True
{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\colour_filters.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\colour_filters.xlsx")
out = out & "~"
For Each ws In wb.Worksheets
    out = out & ws.Name & "^" & Dump(ws) & "|"
Next
wb.SaveAs "{folder}\\colour_filters_resaved.xlsx", 51
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


def file_filters(path: Path) -> dict[str, object]:
    """Each sheet's filter as the file keeps it, its colour filters naming the dxf they hold, and the dxfs."""
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        found = re.search(r"<dxfs\b[^>]*?(?:/>|>(.*?)</dxfs>)", styles, re.DOTALL)
        dxfs = re.findall(r"<dxf>.*?</dxf>|<dxf/>", found.group(1) or "", re.DOTALL) if found else []
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
        targets.update({key: value for value, key in re.findall(r'Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
        sheets: dict[str, object] = {}
        for name, relationship in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
            part = "xl/" + targets[relationship].lstrip("/").removeprefix("xl/")
            text = package.read(part).decode("utf-8")
            tables = [package.read("xl/tables/" + target.rsplit("/", 1)[-1]).decode("utf-8")
                      for target in re.findall(r'Target="([^"]*tables/[^"]+)"',
                                               package.read(part.replace("worksheets/", "worksheets/_rels/")
                                                            + ".rels").decode("utf-8"))] \
                if f"{part.replace('worksheets/', 'worksheets/_rels/')}.rels" in package.namelist() else []
            filters = re.findall(r"<autoFilter\b[^>]*?/>|<autoFilter\b.*?</autoFilter>", "".join([text, *tables]),
                                 re.DOTALL)
            sheets[name] = [re.sub(r'dxfId="(\d+)"', lambda match: f'dxf="{dxfs[int(match.group(1))]}"', one)
                            for one in filters]
    return {"dxfs": dxfs, "sheets": sheets}


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
    for name in ("colour_filters", "colour_filters_resaved"):
        strip_save_path(OUT / f"{name}.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    record = {"helper": HELPER, "table": TABLE, "layout": LAYOUT, "build": BUILD,
              "layouts": [{"name": name, "steps": steps, "dump": dumps.get(name, ""), "reopened": again.get(name, "")}
                          for name, steps in LAYOUTS.items()],
              "file": file_filters(OUT / "colour_filters.xlsx"),
              "resaved": file_filters(OUT / "colour_filters_resaved.xlsx")}
    (OUT / "colour_filters.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"])
        print("   session ", layout["dump"])
        print("   reopened", layout["reopened"])
        print("   file    ", record["file"]["sheets"].get(layout["name"]))


if __name__ == "__main__":
    main()
