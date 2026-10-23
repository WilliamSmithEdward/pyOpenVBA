"""Sorting on a cell's colour or its font's, as the Sort object a recorded macro uses does it.

Every layout sits on a sheet of its own in one new workbook: a list of
eight rows under a header, some of whose names are filled red, blue or
yellow or written in red or blue, sorted by colour one way or another.
The dump reads each row's name, fill and font colour in order, then the
sort fields the sheet keeps: what each sorts on, its order, what its
SortOnValue is and the colour it holds, and what the last call returned
or the error it raised. The layouts ask where rows of the colour go, on
top or at the bottom, how two colours and a colour beside values combine,
what a colour no row has, white, a conditional format's fill and a theme
colour do, and what SortOnValue holds before it is set.

Then the workbook is saved and opened again: each sheet's sortState and
the stylesheet's dxfs go beside the fields as Excel reads them back, and
sorting again from the fields read back.

    python scripts/measure_sort_colour.py

writes tests/fixtures/sort_colour/: sort_colour.xlsx as Excel saved it,
and sort_colour.json, which tests/test_excel_sort_colour.py replays.
"""

from __future__ import annotations

import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "sort_colour"

HELPER = '''Private Function Show(v As Variant) As String
    If IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsError(v) Then
        Show = "Error:" & CStr(CLng(v))
    ElseIf IsObject(v) Then
        If v Is Nothing Then Show = "Nothing" Else Show = "Object"
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function Dump(ws As Object) As String
    Dim out As String, r As Long, f As Object, i As Long
    For r = 2 To 9
        out = out & ws.Cells(r, 1).Value & ":" & ws.Cells(r, 1).Interior.Color & ":" & ws.Cells(r, 1).Font.Color & ","
    Next
    On Error Resume Next
    For i = 1 To ws.Sort.SortFields.Count
        Set f = ws.Sort.SortFields(i)
        out = out & " f" & i & ":" & f.Key.Address & ":" & f.SortOn & ":" & f.Order & ":" & TypeName(f.SortOnValue)
        Err.Clear
        out = out & ":" & f.SortOnValue.Color
        If Err.Number <> 0 Then out = out & ":E" & Err.Number
    Next
    Dump = out
End Function
'''

#: The list every layout sorts: names r2 to r9 over numbers, some names filled or coloured.
LIST = '''ws.Range("A1:C1").Value = Array("Name", "Note", "N")
For i = 2 To 9
    ws.Cells(i, 1).Value = "r" & i
    ws.Cells(i, 3).Value = Choose(i - 1, 5, 3, 8, 1, 7, 2, 6, 4)
Next
ws.Range("A3,A6").Interior.Color = RGB(255, 0, 0)
ws.Range("A5").Interior.Color = RGB(0, 0, 255)
ws.Range("A8").Interior.Color = RGB(255, 255, 0)
ws.Range("A4,A7").Font.Color = RGB(255, 0, 0)
ws.Range("A9").Font.Color = RGB(0, 0, 255)
'''
RED, BLUE, YELLOW, WHITE = "RGB(255, 0, 0)", "RGB(0, 0, 255)", "RGB(255, 255, 0)", "RGB(255, 255, 255)"
CELL, FONT, VALUES = "xlSortOnCellColor", "xlSortOnFontColor", "xlSortOnValues"


def field(key: str, on: str, order: str = "xlAscending", colour: str | None = None) -> str:
    added = f'ws.Sort.SortFields.Add(Key:=ws.Range("{key}"), SortOn:={on}, Order:={order})'
    return f"{added}.SortOnValue.Color = {colour}\n" if colour else f"Set v = {added}\n"


def sort(*fields: str, header: str = "xlYes", block: str = "A1:C9") -> str:
    return ("ws.Sort.SortFields.Clear\n" + "".join(fields) +
            f'With ws.Sort\n.SetRange ws.Range("{block}")\n.Header = {header}\n.Apply\nEnd With\n')


#: name -> setup; each works on ws, the layout's own sheet, keeping the last call's result in v.
LAYOUTS: dict[str, str] = {
    "cell_on_top": sort(field("A2:A9", CELL, colour=RED)),
    "cell_on_bottom": sort(field("A2:A9", CELL, "xlDescending", RED)),
    "two_colours": sort(field("A2:A9", CELL, colour=RED), field("A2:A9", CELL, colour=BLUE)),
    "two_colours_bottom": sort(field("A2:A9", CELL, colour=RED), field("A2:A9", CELL, "xlDescending", YELLOW)),
    "colour_then_values": sort(field("A2:A9", CELL, colour=RED), field("C2:C9", VALUES)),
    "values_then_colour": sort(field("C2:C9", VALUES), field("A2:A9", CELL, colour=RED)),
    "font_on_top": sort(field("A2:A9", FONT, colour=RED)),
    "font_on_bottom": sort(field("A2:A9", FONT, "xlDescending", BLUE)),
    "no_match": sort(field("A2:A9", CELL, colour="RGB(0, 255, 0)")),
    "white": sort(field("A2:A9", CELL, colour=WHITE)),
    "white_font": sort(field("A2:A9", FONT, colour="RGB(0, 0, 0)")),
    "no_header": sort(field("A1:A9", CELL, colour=RED), header="xlNo", block="A1:C9"),
    "conditional": 'ws.Range("C2:C9").FormatConditions.Add(xlCellValue, xlGreater, "5").Interior.Color = ' + RED +
                   "\n" + sort(field("C2:C9", CELL, colour=RED)),
    "theme": 'ws.Range("A4").Interior.ThemeColor = xlThemeColorAccent1\n'
             'ws.Range("A4").Interior.TintAndShade = 0\n' + sort(field("A2:A9", CELL, colour='ws.Range("A4").Interior.Color')),
    "pattern": 'ws.Range("A7").Interior.Pattern = xlGray50\nws.Range("A7").Interior.Color = ' + RED + "\n" +
               sort(field("A2:A9", CELL, colour=RED)),
    "before_set": field("A2:A9", CELL) + field("A2:A9", FONT) + field("A2:A9", VALUES),
    # The colour SortOnValue holds before it is set: no fill, and the automatic font colour.
    "no_fill_on_top": sort(field("A2:A9", CELL)),
    "no_fill_on_bottom": sort(field("A2:A9", CELL, "xlDescending")),
    "automatic_font_on_bottom": sort(field("A2:A9", FONT, "xlDescending")),
    "table": 'With ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C9"), , xlYes)\n.Sort.SortFields.Clear\n'
             '.Sort.SortFields.Add(Key:=.ListColumns(1).DataBodyRange, SortOn:=xlSortOnCellColor, '
             'Order:=xlAscending).SortOnValue.Color = ' + RED + '\n.Sort.Header = xlYes\n.Sort.Apply\n'
             'v = TypeName(.Sort.SortFields(1).SortOnValue) & ":" & .Sort.SortFields(1).SortOnValue.Color\n'
             'End With\n',
    "values_value": field("C2:C9", VALUES) + "v = TypeName(ws.Sort.SortFields(1).SortOnValue)\n",
    "colour_after_values": sort(field("A2:A9", CELL, colour=RED)) + "ws.Sort.SortFields(1).SortOnValue.Color = " +
                           BLUE + "\nws.Sort.Apply\n",
    "range_sort_after": sort(field("A2:A9", CELL, colour=RED)) +
                        'ws.Range("A1:C9").Sort Key1:=ws.Range("C1"), Order1:=xlDescending, Header:=xlYes\n',
    # Apply with something missing, which a sheet opened from a file with no sort settings meets.
    "apply_empty": "ws.Sort.SortFields.Clear\nws.Sort.Apply\n",
    "apply_no_fields": 'ws.Sort.SortFields.Clear\nws.Sort.SetRange ws.Range("A1:C9")\nws.Sort.Apply\n',
    "apply_no_range": "ws.Sort.SortFields.Clear\n" + field("A2:A9", CELL, colour=RED) + "ws.Sort.Apply\n",
}


def case_code(index: int, setup: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String",
             "Dim failed As String, v As Variant, i As Long", "On Error Resume Next", "Err.Clear",
             *setup.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
             f'Case{index} = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    return "\n".join(lines) + "\n"


def _element(text: str, tag: str) -> str:
    found = re.search(rf"<{tag}\b[^>]*/>|<{tag}\b.*?</{tag}>", text, re.DOTALL)
    return found.group() if found else ""


def _targets(package: zipfile.ZipFile, rels: str) -> dict[str, str]:
    targets: dict[str, str] = {}
    for element in re.findall(r"<Relationship\b[^>]*>", package.read(rels).decode("utf-8")):
        attributes = dict(re.findall(r'(\w+)="([^"]*)"', element))
        targets[attributes["Id"]] = attributes["Target"]
    return targets


def file_parts(package: zipfile.ZipFile) -> tuple[dict[str, str], str]:
    """Each sheet's sortState, or its table's, by sheet name, and the stylesheet's dxfs."""
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    targets = _targets(package, "xl/_rels/workbook.xml.rels")
    found: dict[str, str] = {}
    for name, rid in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
        part = posixpath.normpath(posixpath.join("xl", targets[rid].lstrip("/").removeprefix("xl/")))
        found[name] = _element(package.read(part).decode("utf-8"), "sortState")
        rels = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        if rels in package.namelist():
            for target in _targets(package, rels).values():
                if "/tables/" in "/" + target or target.startswith("../tables/"):
                    table = package.read(posixpath.normpath(posixpath.join(posixpath.dirname(part), target)))
                    found[name] += _element(table.decode("utf-8"), "sortState")
    return found, _element(package.read("xl/styles.xml").decode("utf-8"), "dxfs")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "sort_colour.xlsx"
    if path.exists():
        path.unlink()
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, setup) in enumerate(LAYOUTS.items()):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{name}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, LIST + setup))
    # Then the saved file, opened again, read the same way, and each sheet sorted again by what it kept.
    build += [f'wb.SaveAs "{path}", 51', "wb.Close False", 'out = out & "#"', f'Set wb = Workbooks.Open("{path}")',
              "For Each ws In wb.Worksheets", '    out = out & ws.Name & "^" & Dump(ws) & "|"', "Next",
              'out = out & "#"', "On Error Resume Next", "For Each ws In wb.Worksheets", "    Err.Clear",
              "    ws.Sort.Apply",
              '    out = out & ws.Name & "^" & IIf(Err.Number <> 0, "S" & Err.Number & ";", "") & Dump(ws) & "|"',
              "Next", "On Error GoTo 0", "wb.Close False", "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    strip_save_path(path)
    before, after, again = str(result.value).split("#")
    answers = before.split("|")[: len(LAYOUTS)]
    reopened = dict(one.split("^", 1) for one in after.split("|") if one)
    applied = dict(one.split("^", 1) for one in again.split("|") if one)
    with zipfile.ZipFile(path) as package:
        states, dxfs = file_parts(package)
    record = {"helper": HELPER, "list": LIST, "dxfs": dxfs,
              "layouts": [{"name": name, "setup": setup, "answers": answer, "reopened": reopened.get(name, ""),
                           "applied_again": applied.get(name, ""), "sort_state": states.get(name, "")}
                          for (name, setup), answer in zip(LAYOUTS.items(), answers, strict=True)]}
    (OUT / "sort_colour.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for one in record["layouts"]:
        print(f"{one['name']}: {one['answers']}")
        if one["answers"].split(";", 1)[1] != one["reopened"]:
            print(f"    reopened: {one['reopened']}")
        if one["applied_again"] != one["reopened"]:
            print(f"    applied again: {one['applied_again']}")
        print(f"    {one['sort_state']}")
    print("dxfs:", dxfs)


if __name__ == "__main__":
    main()
