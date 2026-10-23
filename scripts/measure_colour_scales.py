"""Colour scales, as FormatConditions.AddColorScale makes them and as Excel colours the cells.

Every layout sits on a sheet of its own in one new workbook: the numbers 0
to 100 in tens down A1:A11, A12 blank, a colour scale over A1:A12 made one
way or another. The dump reads each cell's DisplayFormat.Interior.Color,
then every rule on the sheet: its TypeName, Type, Priority and range, and
for a colour scale each criterion's Index, Type, Value and FormatColor --
Color, ThemeColor and TintAndShade -- and what the last call returned or
the error it raised. The layouts ask for the defaults of two and three
colours, numbers, percents, percentiles and formulas as criteria, colours
and theme colours set, blanks, text, errors and logicals among the
numbers, all one value, negatives, and how a scale meets other rules. Two
more colour 0 to 100 in steps of one, to pin how Excel blends.

Then the workbook is saved and opened again: each sheet's
conditionalFormatting elements go beside the dumps Excel reads back.

    python scripts/measure_colour_scales.py

writes tests/fixtures/colour_scales/: colour_scales.xlsx as Excel saved
it, and colour_scales.json, which tests/test_excel_colour_scales.py
replays.
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

OUT = ROOT / "tests" / "fixtures" / "colour_scales"

HELPER = '''Private Function Show(v As Variant) As String
    If IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsError(v) Then
        Show = "Error:" & CStr(CLng(v))
    ElseIf IsObject(v) Then
        If v Is Nothing Then Show = "Nothing" Else Show = "Object:" & TypeName(v)
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function Dump(ws As Object) As String
    Dim out As String, r As Long, f As Object, i As Long, c As Object
    For r = 1 To 12
        out = out & ws.Cells(r, 1).DisplayFormat.Interior.Color & ","
    Next
    On Error Resume Next
    out = out & " n=" & ws.Cells.FormatConditions.Count
    For i = 1 To ws.Cells.FormatConditions.Count
        Set f = ws.Cells.FormatConditions(i)
        out = out & " [" & TypeName(f) & ":" & f.Type & ":" & f.Priority & ":" & f.AppliesTo.Address
        If TypeName(f) = "ColorScale" Then
            out = out & ":" & f.ColorScaleCriteria.Count
            For Each c In f.ColorScaleCriteria
                out = out & " c" & c.Index & "=" & c.Type
                Err.Clear: out = out & "/" & c.Value: If Err.Number <> 0 Then out = out & "/E" & Err.Number
                Err.Clear: out = out & "/" & c.FormatColor.Color: If Err.Number <> 0 Then out = out & "/E" & Err.Number
                Err.Clear: out = out & "/" & c.FormatColor.ThemeColor: If Err.Number <> 0 Then out = out & "/E" & Err.Number
                Err.Clear: out = out & "/" & c.FormatColor.TintAndShade: If Err.Number <> 0 Then out = out & "/E" & Err.Number
            Next
        End If
        out = out & "]"
    Next
    Dump = out
End Function

Private Function Colours(ws As Object, column As Long) As String
    Dim out As String, r As Long
    For r = 1 To 101
        out = out & ws.Cells(r, column).DisplayFormat.Interior.Color & ","
    Next
    Colours = out
End Function
'''

#: The numbers every layout colours: 0 to 100 in tens down A1:A11, A12 blank.
LIST = "For i = 1 To 11: ws.Cells(i, 1).Value = (i - 1) * 10: Next\n"
SCALE = 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(ColorScaleType:={})\n'


def criterion(index: int, kind: str, value: object | None = None) -> str:
    code = f"f.ColorScaleCriteria({index}).Type = {kind}\n"
    return code + (f"f.ColorScaleCriteria({index}).Value = {value}\n" if value is not None else "")


def colour(index: int, value: str) -> str:
    return f"f.ColorScaleCriteria({index}).FormatColor.Color = {value}\n"


FILL_RULE = ('Set g = ws.Range("A1:A12").FormatConditions.Add(xlCellValue, xlGreater, "50")\n'
             "g.Interior.Color = RGB(0, 0, 255)\n")
#: name -> setup; each works on ws, the layout's own sheet, its scale in f, keeping the last call's result in v.
LAYOUTS: dict[str, str] = {
    "two_default": SCALE.format(2),
    "three_default": SCALE.format(3),
    "two_numbers": SCALE.format(2) + criterion(1, "xlConditionValueNumber", 20) +
                   criterion(2, "xlConditionValueNumber", 80),
    "three_percent": SCALE.format(3) + criterion(1, "xlConditionValuePercent", 10) +
                     criterion(2, "xlConditionValuePercent", 50) + criterion(3, "xlConditionValuePercent", 90),
    "three_percentile": SCALE.format(3) + criterion(1, "xlConditionValuePercentile", 25) +
                        criterion(2, "xlConditionValuePercentile", 50) + criterion(3, "xlConditionValuePercentile", 75),
    "three_mid_number": SCALE.format(3) + criterion(2, "xlConditionValueNumber", 30),
    "colours": SCALE.format(3) + colour(1, "RGB(0, 0, 255)") + colour(2, "RGB(255, 255, 255)") +
               colour(3, "RGB(255, 0, 0)"),
    "theme_colour": SCALE.format(2) + "f.ColorScaleCriteria(1).FormatColor.ThemeColor = xlThemeColorAccent1\n"
                                      "f.ColorScaleCriteria(1).FormatColor.TintAndShade = 0.4\n",
    "formula": 'ws.Range("C1").Value = 30\n' + SCALE.format(2) + criterion(1, "xlConditionValueFormula", '"=$C$1"'),
    "odd_cells": 'ws.Range("A3").Value = "text"\nws.Range("A5").ClearContents\nws.Range("A7").Formula = "=1/0"\n'
                 'ws.Range("A9").Value = True\n' + SCALE.format(3),
    "all_same": 'ws.Range("A1:A11").Value = 5\n' + SCALE.format(3),
    "negatives": "For i = 1 To 11: ws.Cells(i, 1).Value = (i - 6) * 10: Next\n" + SCALE.format(3),
    "one_number": 'ws.Range("A1:A11").ClearContents\nws.Range("A4").Value = 7\n' + SCALE.format(2),
    "two_scales": SCALE.format(2) + "Set g = " + SCALE.format(3).removeprefix("Set f = "),
    "scale_then_fill": SCALE.format(3) + FILL_RULE,
    "fill_then_scale": FILL_RULE + SCALE.format(3),
    "fill_first_stops": SCALE.format(3) + FILL_RULE + "g.SetFirstPriority\ng.StopIfTrue = True\n",
    "type_names": SCALE.format(3) + 'v = TypeName(f) & ":" & TypeName(f.ColorScaleCriteria) & ":" & '
                                    'TypeName(f.ColorScaleCriteria(1)) & ":" & TypeName(f.ColorScaleCriteria(1).FormatColor)\n',
    "stop_if_true": SCALE.format(3) + "v = f.StopIfTrue\nf.StopIfTrue = True\n",
    "criterion_zero": SCALE.format(3) + "v = f.ColorScaleCriteria(0).Type\n",
    "criterion_past": SCALE.format(2) + "v = f.ColorScaleCriteria(3).Type\n",
    "scale_type_four": SCALE.format(4),
    "scale_omitted": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale\n',
    "value_of_lowest": SCALE.format(2) + "v = f.ColorScaleCriteria(1).Value\n",
    "lowest_given_value": SCALE.format(2) + "f.ColorScaleCriteria(1).Value = 5\n",
    "applies_moved": SCALE.format(3) + 'f.ModifyAppliesToRange ws.Range("A1:A6")\n',
    "deleted": SCALE.format(3) + "f.Delete\n",
    "gradient_two": "For i = 1 To 101: ws.Cells(i, 2).Value = i - 1: Next\n"
                    'ws.Range("B1:B101").FormatConditions.AddColorScale 2\nv = Colours(ws, 2)\n',
    "gradient_three": "For i = 1 To 101: ws.Cells(i, 2).Value = i - 1: Next\n"
                      'ws.Range("B1:B101").FormatConditions.AddColorScale 3\nv = Colours(ws, 2)\n',
}


def case_code(index: int, setup: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String",
             "Dim failed As String, v As Variant, i As Long, f As Object, g As Object", "On Error Resume Next",
             "Err.Clear", *setup.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"',
             "On Error GoTo 0", f'Case{index} = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    return "\n".join(lines) + "\n"


def _targets(package: zipfile.ZipFile, rels: str) -> dict[str, str]:
    targets: dict[str, str] = {}
    for element in re.findall(r"<Relationship\b[^>]*>", package.read(rels).decode("utf-8")):
        attributes = dict(re.findall(r'(\w+)="([^"]*)"', element))
        targets[attributes["Id"]] = attributes["Target"]
    return targets


def file_parts(package: zipfile.ZipFile) -> dict[str, str]:
    """Each sheet's conditionalFormatting elements, and any extLst, by sheet name."""
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    targets = _targets(package, "xl/_rels/workbook.xml.rels")
    found: dict[str, str] = {}
    for name, rid in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
        part = posixpath.normpath(posixpath.join("xl", targets[rid].lstrip("/").removeprefix("xl/")))
        text = package.read(part).decode("utf-8")
        found[name] = "".join(re.findall(r"<conditionalFormatting\b.*?</conditionalFormatting>", text, re.DOTALL)) + \
            "".join(re.findall(r"<extLst>.*?</extLst>", text, re.DOTALL))
    return found


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "colour_scales.xlsx"
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
    build += [f'wb.SaveAs "{path}", 51', "wb.Close False", 'out = out & "#"', f'Set wb = Workbooks.Open("{path}")',
              "For Each ws In wb.Worksheets", '    out = out & ws.Name & "^" & Dump(ws) & "|"', "Next",
              "wb.Close False", "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    strip_save_path(path)
    before, after = str(result.value).split("#")
    answers = before.split("|")[: len(LAYOUTS)]
    reopened = dict(one.split("^", 1) for one in after.split("|") if one)
    with zipfile.ZipFile(path) as package:
        parts = file_parts(package)
    record = {"helper": HELPER, "list": LIST,
              "layouts": [{"name": name, "setup": setup, "answers": answer, "reopened": reopened.get(name, ""),
                           "file": parts.get(name, "")}
                          for (name, setup), answer in zip(LAYOUTS.items(), answers, strict=True)]}
    (OUT / "colour_scales.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for one in record["layouts"]:
        print(f"{one['name']}: {one['answers'][:700]}")
        if one["answers"].split(";", 1)[1] != one["reopened"]:
            print(f"    reopened: {one['reopened'][:500]}")
        print(f"    {one['file'][:600]}")


if __name__ == "__main__":
    main()
