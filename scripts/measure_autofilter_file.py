"""How Excel writes an AutoFilter to its file.

Every case filters the table scripts/measure_autofilter.py uses, on a
sheet of its own, and reads the filter back the way that probe does; then
the workbook is saved. What the file holds -- each sheet's autoFilter
element and the workbook's defined names -- goes beside what Excel read.

    python scripts/measure_autofilter_file.py

writes tests/fixtures/autofilter_file/: autofilter.xlsx as Excel saved it,
and autofilter_file.json with each sheet's reading, its autoFilter XML and
the workbook's definedNames; tests/test_excel_autofilter_file.py replays
it.
"""

from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from measure_autofilter import HELPER, NUMBERS, TABLE, filt  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "autofilter_file"

#: name -> setup, run on the table.
CASES: dict[str, str] = {
    "arrows": filt(),
    "equals": filt('Field:=1, Criteria1:="apple"'),
    "wildcard": filt('Field:=1, Criteria1:="a*"'),
    "not_equal": filt('Field:=1, Criteria1:="<>apple"'),
    "greater": filt('Field:=2, Criteria1:=">2"'),
    "between": filt('Field:=2, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=10"'),
    "either": filt('Field:=1, Criteria1:="apple", Operator:=xlOr, Criteria2:="cherry"'),
    "values_three": filt('Field:=1, Criteria1:=Array("apple", "cherry", "banana"), Operator:=xlFilterValues'),
    "values_two": filt('Field:=1, Criteria1:=Array("apple", "cherry"), Operator:=xlFilterValues'),
    "values_blank": filt('Field:=4, Criteria1:=Array("x", "y", ""), Operator:=xlFilterValues'),
    "top_items": NUMBERS + filt("Field:=1, Criteria1:=3, Operator:=xlTop10Items", "F1:F9"),
    "top_percent": NUMBERS + filt("Field:=1, Criteria1:=25, Operator:=xlTop10Percent", "F1:F9"),
    "bottom_items": NUMBERS + filt("Field:=1, Criteria1:=2, Operator:=xlBottom10Items", "F1:F9"),
    "above_average": NUMBERS + filt("Field:=1, Criteria1:=xlFilterAboveAverage, Operator:=xlFilterDynamic", "F1:F9"),
    "below_average": NUMBERS + filt("Field:=1, Criteria1:=xlFilterBelowAverage, Operator:=xlFilterDynamic", "F1:F9"),
    "blanks": filt('Field:=4, Criteria1:="="'),
    "non_blanks": filt('Field:=4, Criteria1:="<>"'),
    "two_fields": filt('Field:=1, Criteria1:="cherry"') + filt('Field:=4, Criteria1:="y"'),
    "date_after": filt('Field:=3, Criteria1:=">1/15/2020"'),
    "date_equals": filt('Field:=3, Criteria1:="1/15/2020"'),
    "text_greater": filt('Field:=1, Criteria1:=">b"'),
    "number_equals": filt('Field:=2, Criteria1:="1"'),
    "number_shown": filt('Field:=2, Criteria1:="1.00"'),
    "not_number": filt('Field:=2, Criteria1:="<>2"'),
    "any_text": filt('Field:=1, Criteria1:="*"'),
    "shown_all": filt('Field:=1, Criteria1:="apple"') + "ws.ShowAllData\n",
    "turned_off": filt('Field:=1, Criteria1:="apple"') + "ws.AutoFilterMode = False\n",
    "escaped": filt('Field:=1, Criteria1:="b~*"'),
    "time_after": filt('Field:=3, Criteria1:=">1/15/2020 5:00"'),
    "hidden_button": filt('Field:=1, Criteria1:="apple", VisibleDropDown:=False'),
    "or_custom": filt('Field:=1, Criteria1:="a*", Operator:=xlOr, Criteria2:="cherry"'),
    "and_plain": filt('Field:=1, Criteria1:="apple", Operator:=xlAnd, Criteria2:="Apple"'),
    "or_blank": filt('Field:=4, Criteria1:="=", Operator:=xlOr, Criteria2:="x"'),
    "decimal_greater": filt('Field:=2, Criteria1:=">2.50"'),
    "grown_below": filt('Field:=1, Criteria1:="apple"') + 'ws.Range("A14").Value = "apple"\n',
    "grown_reapplied": filt('Field:=1, Criteria1:="apple"') + 'ws.Range("A14").Value = "pear"\nws.AutoFilter.ApplyFilter\n',
    "sub_range": filt('Field:=1, Criteria1:="apple"', "A1:D5"),
    "sub_range_arrows": filt(reference="A1:D5"),
    "sub_range_then_field": filt(reference="A1:D5") + filt('Field:=1, Criteria1:="apple"', "A1"),
    "sub_range_refiltered": filt('Field:=1, Criteria1:="apple"', "A1:D5") + filt('Field:=2, Criteria1:=">1"', "A1:D5"),
    "sub_range_mid": filt('Field:=1, Criteria1:="apple"', "A3:D5"),
}


def case_code(index: int, setup: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String", "Dim failed As String, v As Variant",
             "On Error Resume Next", "Err.Clear", *setup.splitlines(),
             'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
             f'Case{index} = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    return "\n".join(lines) + "\n"


def autofilter_xml(package: zipfile.ZipFile) -> dict[str, str]:
    """Each sheet's autoFilter element, by sheet name, empty where it has none."""
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    relationships = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    targets = dict(re.findall(r'<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', relationships))
    targets |= {rid: target for target, rid in re.findall(r'<Relationship\b[^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"',
                                                           relationships)}
    found: dict[str, str] = {}
    for name, rid in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
        text = package.read("xl/" + targets[rid].lstrip("/").removeprefix("xl/")).decode("utf-8")
        element = re.search(r"<autoFilter\b[^>]*/>|<autoFilter\b.*?</autoFilter>", text, re.DOTALL)
        found[name] = element.group() if element else ""
    return found


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "autofilter.xlsx"
    if path.exists():
        path.unlink()
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, setup) in enumerate(CASES.items()):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{name}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, TABLE + setup))
    # Then the saved file, opened again, read the same way.
    build += [f'wb.SaveAs "{path}", 51', "wb.Close False", 'out = out & "#"', f'Set wb = Workbooks.Open("{path}")',
              "For Each ws In wb.Worksheets", '    out = out & Dump(ws) & "|"', "Next", "wb.Close False"]
    # And a workbook whose names mix the filter's with names of both scopes, to see the order they are written in.
    names_path = OUT / "names_order.xlsx"
    if names_path.exists():
        names_path.unlink()
    build += ["Set wb = Workbooks.Add(xlWBATWorksheet)", 'wb.Worksheets(1).Name = "Zed"',
              'wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "Abc"', 'wb.Worksheets.Add(After:=wb.Worksheets(2)).Name = "Mid"',
              "For Each ws In wb.Worksheets", '    ws.Range("A1:B3").Value = 1', '    ws.Range("A1:B3").AutoFilter', "Next",
              'wb.Names.Add Name:="beta", RefersTo:="=Zed!$A$1"', 'wb.Names.Add Name:="Alpha", RefersTo:="=Zed!$A$2"',
              'wb.Names.Add Name:="zulu", RefersTo:="=Zed!$A$3"', 'wb.Names.Add Name:="_under", RefersTo:="=Zed!$A$4"',
              'wb.Worksheets("Mid").Names.Add Name:="beta", RefersTo:="=Mid!$B$1"',
              'wb.Worksheets("Abc").Names.Add Name:="Local", RefersTo:="=Abc!$B$2"',
              'wb.Worksheets("Zed").PageSetup.PrintArea = "$A$1:$B$3"',
              f'wb.SaveAs "{names_path}", 51', "wb.Close False", "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    before, after = str(result.value).split("#")
    readings = before.split("|")[: len(CASES)]
    reopened = after.split("|")[: len(CASES)]
    with zipfile.ZipFile(path) as package:
        elements = autofilter_xml(package)
        names = re.findall(r"<definedName\b.*?</definedName>", package.read("xl/workbook.xml").decode("utf-8"))
    with zipfile.ZipFile(OUT / "names_order.xlsx") as package:
        order = re.findall(r"<definedName\b.*?</definedName>", package.read("xl/workbook.xml").decode("utf-8"))
    for one in order:
        print(one)
    record = {"helper": HELPER, "table": TABLE, "names": names, "names_order": order,
              "cases": [{"name": name, "setup": setup, "reading": reading, "reopened": again, "xml": elements[name]}
                        for (name, setup), reading, again in zip(CASES.items(), readings, reopened, strict=True)]}
    (OUT / "autofilter_file.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for (name, _), reading, again in zip(CASES.items(), readings, reopened, strict=True):
        if reading.split(";", 1)[1] != again:
            print(f"{name}:\n   before {reading.split(';', 1)[1]}\n   after  {again}")


if __name__ == "__main__":
    main()
