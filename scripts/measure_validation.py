"""Data validation: what Range.Validation answers, and what a file keeps of it.

Each workbook below is given validation rules from VBA in live Excel --
Validation.Add with each type, operator and argument, Modify, Delete,
settings changed, cells inserted or copied around them -- and every cell
in READ is described through the Validation object: each property, or the
error reading it raises. The workbook is saved as xlsx and the
dataValidations element Excel wrote is read out of the package.

    python scripts/measure_validation.py

writes tests/fixtures/validation/ -- one .xlsx per workbook and
validation.json -- which tests/test_excel_validation.py replays.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from fixture_workbook import strip_save_path

ROOT = Path(__file__).resolve().parent.parent
FOLDER = ROOT / "tests" / "fixtures" / "validation"

#: The properties each cell's Validation is described by, in order.
PROPERTIES = ["Type", "AlertStyle", "Operator", "Formula1", "Formula2", "IgnoreBlank", "InCellDropdown", "ShowInput",
              "ShowError", "InputTitle", "InputMessage", "ErrorTitle", "ErrorMessage", "Value"]

#: The cells described after each workbook's VBA has run.
READ = ["A1", "A3", "A5", "A6", "B1", "C1", "C3"]

#: Each workbook: its name and the VBA that sets it up on ws, its first sheet.
BOOKS: list[tuple[str, str]] = [
    ("none", ""),
    ("list_literal", 'ws.Range("A1:A5").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, '
                     'Operator:=xlBetween, Formula1:="red,green,blue"'),
    ("list_only_type", 'ws.Range("A1:A5").Validation.Add xlValidateList, , , "red,green,blue"'),
    ("list_range", 'ws.Range("D1:D3").Value = Application.Transpose(Array("x", "y", "z"))\n'
                   'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="=$D$1:$D$3"'),
    ("list_relative", 'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="=D1:D3"'),
    ("whole_between", 'ws.Range("A1:A5").Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, "1", '
                      '"10"'),
    ("whole_numbers", 'ws.Range("A1:A5").Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, 1, 10'),
    ("decimal_greater", 'ws.Range("A1:A5").Validation.Add xlValidateDecimal, xlValidAlertWarning, xlGreater, "0.5"'),
    ("date_from", 'ws.Range("A1:A5").Validation.Add xlValidateDate, xlValidAlertInformation, xlGreaterEqual, '
                  '"1/2/2020"'),
    ("time_between", 'ws.Range("A1:A5").Validation.Add xlValidateTime, Formula1:="9:00", Formula2:="17:30"'),
    ("text_length", 'ws.Range("A1:A5").Validation.Add xlValidateTextLength, , xlLessEqual, "5"'),
    ("custom_relative", 'ws.Range("A1:A5").Validation.Add xlValidateCustom, Formula1:="=A1>B1"'),
    ("any_value", 'ws.Range("A1:A5").Validation.Add xlValidateInputOnly'),
    ("messages", 'With ws.Range("A1:A5").Validation\n.Add xlValidateList, Formula1:="a,b"\n'
                 '.InputTitle = "Pick"\n.InputMessage = "One of a or b"\n.ErrorTitle = "Wrong"\n'
                 '.ErrorMessage = "Not a or b"\n.ShowInput = False\n.IgnoreBlank = False\n'
                 '.InCellDropdown = False\nEnd With'),
    ("added_twice", 'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="a,b"\nOn Error Resume Next\n'
                    'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="c,d"\n'
                    'ws.Range("E1").Value = Err.Number\nOn Error GoTo 0'),
    ("added_over_part", 'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="a,b"\nOn Error Resume Next\n'
                        'ws.Range("A3:C3").Validation.Add xlValidateList, Formula1:="c,d"\n'
                        'ws.Range("E1").Value = Err.Number\nOn Error GoTo 0'),
    ("deleted", 'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="a,b"\n'
                'ws.Range("A1:A5").Validation.Delete'),
    ("deleted_part", 'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="a,b"\n'
                     'ws.Range("A3").Validation.Delete'),
    ("modified", 'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="a,b"\n'
                 'ws.Range("A1:A5").Validation.Modify xlValidateWholeNumber, xlValidAlertStop, xlGreater, "3"'),
    ("value_checked", 'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="red,green"\n'
                      'ws.Range("A1").Value = "red"\nws.Range("A3").Value = "blue"'),
    ("value_numbers", 'ws.Range("A1:A5").Validation.Add xlValidateWholeNumber, , xlBetween, "1", "10"\n'
                      'ws.Range("A1").Value = 5\nws.Range("A3").Value = 50\nws.Range("A5").Value = 2.5'),
    ("rows_inserted", 'ws.Range("A1:A5").Validation.Add xlValidateCustom, Formula1:="=A1>B1"\nws.Rows(3).Insert'),
    ("rows_deleted", 'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="a,b"\nws.Rows(2).Delete'),
    ("copied", 'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="a,b"\nws.Range("A1").Copy ws.Range("C1")'),
    ("cleared", 'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="a,b"\nws.Range("A1:A2").Clear'),
    ("clear_contents", 'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="a,b"\n'
                       'ws.Range("A1:A2").ClearContents'),
    ("two_rules", 'ws.Range("C1:C3").Validation.Add xlValidateList, Formula1:="a,b"\n'
                  'ws.Range("A1:A5").Validation.Add xlValidateWholeNumber, , xlGreater, "0"'),
    ("block", 'ws.Range("A1:C3").Validation.Add xlValidateList, Formula1:="a,b"'),
    ("union", 'ws.Range("A1,A3,C3").Validation.Add xlValidateList, Formula1:="a,b"'),
    ("other_sheet_list", 'ws.Parent.Worksheets.Add After:=ws\nws.Parent.Worksheets(2).Range("A1:A2").Value = '
                         'Application.Transpose(Array("p", "q"))\n'
                         'ws.Range("A1:A5").Validation.Add xlValidateList, Formula1:="=Sheet2!$A$1:$A$2"'),
    ("value_case", 'ws.Range("A1:A6").Validation.Add xlValidateList, Formula1:="red,green"\n'
                   'ws.Range("A1").Value = "RED"\nws.Range("A3").Value = " red"\nws.Range("A5").Value = "green "\n'
                   'ws.Range("A6").Value = 1'),
    ("value_list_range", 'ws.Range("D1:D3").Value = Application.Transpose(Array("x", 2, "z"))\n'
                         'ws.Range("A1:A6").Validation.Add xlValidateList, Formula1:="=$D$1:$D$3"\n'
                         'ws.Range("A1").Value = "X"\nws.Range("A3").Value = "w"\nws.Range("A5").Value = 2\n'
                         'ws.Range("A6").Value = "2"'),
    ("value_custom", 'ws.Range("A1:A6").Validation.Add xlValidateCustom, Formula1:="=A1>B1"\n'
                     'ws.Range("A1:B1").Value = Array(5, 3)\nws.Range("A3:B3").Value = Array(1, 2)\n'
                     'ws.Range("B5").Value = 1\nws.Range("A6:B6").Value = Array("a", 1)'),
    ("value_text_length", 'ws.Range("A1:A6").Validation.Add xlValidateTextLength, , xlLessEqual, "5"\n'
                          'ws.Range("A1").Value = "abc"\nws.Range("A3").Value = "abcdefg"\n'
                          'ws.Range("A5").Value = 123456\nws.Range("A6").Value = 12345'),
    ("value_date", 'ws.Range("A1:A6").Validation.Add xlValidateDate, , xlGreaterEqual, "1/2/2020"\n'
                   'ws.Range("A1").Value = DateSerial(2020, 1, 3)\nws.Range("A3").Value = DateSerial(2019, 1, 1)\n'
                   'ws.Range("A5").Value = "text"\nws.Range("A6").Value = 43833'),
    ("value_decimal", 'ws.Range("A1:A6").Validation.Add xlValidateDecimal, , xlGreater, "0.5"\n'
                      'ws.Range("A1").Value = 0.6\nws.Range("A3").Value = "abc"\nws.Range("A5").Value = True\n'
                      'ws.Range("A6").Value = "0.7"'),
    ("value_whole_text", 'ws.Range("A1:A6").Validation.Add xlValidateWholeNumber, , xlBetween, "1", "10"\n'
                         'ws.Range("A1").Value = "5"\nws.Range("A3").Value = 0\nws.Range("A5").Value = 10\n'
                         'ws.Range("A6").Value = -1'),
    ("mixed_range", 'ws.Range("A1").Validation.Add xlValidateList, Formula1:="a,b"\n'
                    'ws.Range("A3").Validation.Add xlValidateList, Formula1:="a,b"\nOn Error Resume Next\n'
                    'ws.Range("E1").Value = "type " & ws.Range("A1:A3").Validation.Type\n'
                    'If Err.Number <> 0 Then ws.Range("E1").Value = "!" & Err.Number\nErr.Clear\n'
                    'ws.Range("E1").Value = ws.Range("E1").Value & "/" & ws.Range("A1,A3").Validation.Type\n'
                    'If Err.Number <> 0 Then ws.Range("E1").Value = ws.Range("E1").Value & "/!" & Err.Number\n'
                    'On Error GoTo 0'),
]

HELPERS = "\n".join(
    f"""Private Function R_{name}(v As Object) As String
    On Error GoTo Bad
    R_{name} = CStr(v.{name})
    Exit Function
Bad:
    R_{name} = "!" & Err.Number
End Function
""" for name in PROPERTIES) + """
Private Function Describe(c As Object) As String
    Dim v As Object
    Set v = c.Validation
    Describe = """ + ' & "~;~" & '.join(f"R_{name}(v)" for name in PROPERTIES) + """
End Function
"""


def module() -> str:
    lines = ["Public Function Probe() As String", "Dim out As String", "Application.DisplayAlerts = False"]
    lines += [f'out = out & B{index}("{FOLDER / (name + ".xlsx")}")' for index, (name, _) in enumerate(BOOKS)]
    lines += ["Probe = out", "End Function"]
    for index, (name, making) in enumerate(BOOKS):
        lines += [f"Private Function B{index}(path As String) As String",
                  "Dim wb As Object, ws As Object, out As String",
                  "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", making,
                  f'out = "{name}:E1~:~" & ws.Range("E1").Value & "~|~"']
        lines += [f'out = out & "{name}:{cell}~:~" & Describe(ws.Range("{cell}")) & "~|~"' for cell in READ]
        lines += ["wb.SaveAs Filename:=path, FileFormat:=51", "wb.Close False", f"B{index} = out", "End Function"]
    return HELPERS + "\n".join(lines) + "\n"


def saved(path: Path) -> str:
    """The dataValidations element Excel wrote in the first sheet, or ""."""
    with zipfile.ZipFile(path) as package:
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    found = re.search(r"<dataValidations\b.*?</dataValidations>", sheet, re.DOTALL)
    return found.group(0) if found else ""


def main() -> None:
    FOLDER.mkdir(parents=True, exist_ok=True)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    for name, _ in BOOKS:
        strip_save_path(FOLDER / f"{name}.xlsx")
    answers = dict(part.split("~:~", 1) for part in str(result.value).split("~|~") if "~:~" in part)
    record = {"helpers": HELPERS, "properties": PROPERTIES, "read": READ,
              "books": [{"name": name, "making": making, "e1": answers[f"{name}:E1"],
                         "cells": {cell: answers[f"{name}:{cell}"] for cell in READ},
                         "saved": saved(FOLDER / f"{name}.xlsx")} for name, making in BOOKS]}
    (FOLDER / "validation.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for book in record["books"]:
        print(f"--- {book['name']} e1={book['e1']}")
        for cell, answer in book["cells"].items():
            print(f"    {cell:4} {answer.replace('~;~', ' | ')}")
        print(f"    saved {book['saved']}")


if __name__ == "__main__":
    main()
