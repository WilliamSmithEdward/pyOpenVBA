"""Worksheet settings macros set beside Protect or on dashboards, in live Excel: each one's value on a new sheet, what
each setting a macro gives it reads back as, what a few of them do, and which a save keeps.

Each property is read, set to each value in its list, an error's number and description noted where a setting
raises one, and read back after each. ScrollArea's values each follow one it takes, so the reading shows what a
value did. A few cases read what ScrollArea does to a selection, what EnableCalculation does to a formula and what
DisplayRightToLeft does to a cell's place. Last, every setting is made on one sheet, which is copied, and the
workbook is saved, its sheetPr and sheetView kept, and opened again.

    python scripts/measure_worksheet_settings.py

writes tests/fixtures/worksheet_settings.json, which tests/test_excel_worksheet_settings.py replays.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "worksheet_settings.json"

FLAG_VALUES = ["Not start", "0", "1", '"True"', '"False"', "True", "Null", "True", '"abc"', "False", '"1"', "True",
               '"0"', "True", '"yes"', "False", "0.4", "False", "2", "True", "Empty", "False", '""']


def after(baseline: str, values: list[str]) -> list[str]:
    """Each of ``values`` set after ``baseline``."""
    return [one for value in values for one in (baseline, value)]


#: property -> the values set on ``ws``, as VBA expressions; ``start`` is the value it began with.
SETTINGS: dict[str, list[str]] = {
    "ScrollArea": after('"G7"', ['"A1:D10"', '"a1:d10"', '"A1"', '"$A$1:$D$10"', '"A1:B2,D4:E5"', '"1:5"', '"A:C"',
                                 '"Sheet1!C3:D4"', '"Other!A1:B2"', '"Spot"', '"foo"', '""', "Null", "5", "Empty",
                                 'ws.Range("A1:B2")', '"A1:XFD1048576"', '"R1C1:R2C2"', '" A1:B2 "', '"A1 B2"',
                                 '"B2:A1"', '"A1:B2 B1:C2"', '"A1,A1"', "True"]),
    "EnableAutoFilter": FLAG_VALUES,
    "EnableOutlining": FLAG_VALUES,
    "EnablePivotTable": FLAG_VALUES,
    "EnableCalculation": FLAG_VALUES,
    "TransitionExpEval": FLAG_VALUES,
    "TransitionFormEntry": FLAG_VALUES,
    "DisplayRightToLeft": FLAG_VALUES,
}

HELPER = '''Private Function Shown(value As Variant) As String
    If IsNull(value) Then
        Shown = "Null"
    ElseIf IsObject(value) Then
        Shown = "Object:" & TypeName(value)
    Else
        Shown = TypeName(value) & ":" & CStr(value)
    End If
End Function

Private Function Failure() As String
    If Err.Number <> 0 Then Failure = "E" & Err.Number & ":" & Err.Description
    Err.Clear
End Function

Private Function Book() As Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "Other"
    wb.Names.Add "Spot", "=Sheet1!$E$5:$F$6"
    wb.Worksheets(1).Activate
    Set Book = wb
End Function
'''

#: name -> (steps on a new workbook's first sheet ``ws``, the readings after them).
EFFECTS: dict[str, tuple[list[str], list[str]]] = {
    "scroll_select": (['ws.ScrollArea = "A1:D10"', 'ws.Range("Z100").Select'],
                      ["ActiveCell.Address", "Selection.Address", "ws.ScrollArea"]),
    "scroll_active": (['ws.Range("Z100").Select', 'ws.ScrollArea = "A1:D10"'],
                      ["ActiveCell.Address", "Selection.Address"]),
    "scroll_activate": (['ws.ScrollArea = "A1:D10"', 'ws.Range("Z100").Activate'], ["ActiveCell.Address"]),
    "scroll_goto": (['ws.ScrollArea = "A1:D10"', 'Application.Goto ws.Range("Z100")'], ["ActiveCell.Address"]),
    "scroll_write": (['ws.ScrollArea = "A1:D10"', 'ws.Range("Z100").Value = 5'], ['ws.Range("Z100").Value']),
    "calc_off": (['ws.Range("A1").Value = 1', 'ws.Range("B1").Formula = "=A1*2"', "ws.EnableCalculation = False",
                  'ws.Range("A1").Value = 5'],
                 ['ws.Range("B1").Value', "ws.EnableCalculation"]),
    "calc_off_calculate": (['ws.Range("A1").Value = 1', 'ws.Range("B1").Formula = "=A1*2"',
                            "ws.EnableCalculation = False", 'ws.Range("A1").Value = 5', "ws.Calculate"],
                           ['ws.Range("B1").Value']),
    "calc_off_application": (['ws.Range("A1").Value = 1', 'ws.Range("B1").Formula = "=A1*2"',
                              "ws.EnableCalculation = False", 'ws.Range("A1").Value = 5', "Application.Calculate"],
                             ['ws.Range("B1").Value']),
    "calc_off_full": (['ws.Range("A1").Value = 1', 'ws.Range("B1").Formula = "=A1*2"',
                       "ws.EnableCalculation = False", 'ws.Range("A1").Value = 5', "Application.CalculateFull"],
                      ['ws.Range("B1").Value']),
    "calc_back_on": (['ws.Range("A1").Value = 1', 'ws.Range("B1").Formula = "=A1*2"', "ws.EnableCalculation = False",
                      'ws.Range("A1").Value = 5', "ws.EnableCalculation = True"],
                     ['ws.Range("B1").Value']),
    "calc_new_formula": (["ws.EnableCalculation = False", 'ws.Range("A1").Value = 3',
                          'ws.Range("B1").Formula = "=A1*2"'],
                         ['ws.Range("B1").Value']),
    "calc_range_calculate": (['ws.Range("A1").Value = 1', 'ws.Range("B1").Formula = "=A1*2"',
                              "ws.EnableCalculation = False", 'ws.Range("A1").Value = 5', 'ws.Range("B1").Calculate'],
                             ['ws.Range("B1").Value']),
    "calc_other_sheet": (['ws.Range("A1").Value = 1', 'ws.Range("B1").Formula = "=A1*2"',
                          'wb.Worksheets("Other").Range("C1").Formula = "=Sheet1!B1+Sheet1!A1"',
                          "ws.EnableCalculation = False", 'ws.Range("A1").Value = 5'],
                         ['ws.Range("B1").Value', 'wb.Worksheets("Other").Range("C1").Value']),
    "calc_reads_other": (['wb.Worksheets("Other").Range("A1").Value = 1', 'ws.Range("B1").Formula = "=Other!A1*2"',
                          "ws.EnableCalculation = False", 'wb.Worksheets("Other").Range("A1").Value = 5'],
                         ['ws.Range("B1").Value']),
    "calc_volatile": (['ws.Range("B1").Formula = "=ROWS(A1:A3)+COUNTA(A1:A3)"', "ws.EnableCalculation = False",
                       'ws.Range("A1").Value = 5', "Application.CalculateFull"],
                      ['ws.Range("B1").Value']),
    "calc_over_constant": (['ws.Range("A1").Value = 3', 'ws.Range("B1").Value = 7', "ws.EnableCalculation = False",
                            'ws.Range("B1").Formula = "=A1*2"', 'ws.Range("D1").Formula = "=""a"""'],
                           ['ws.Range("B1").Value', 'ws.Range("D1").Value', 'ws.Range("B1").Text']),
    "calc_over_constant_back_on": (['ws.Range("A1").Value = 3', 'ws.Range("B1").Value = 7',
                                    "ws.EnableCalculation = False", 'ws.Range("B1").Formula = "=A1*2"',
                                    'ws.Range("D1").Formula = "=""a"""', "ws.EnableCalculation = True"],
                                   ['ws.Range("B1").Value', 'ws.Range("D1").Value']),
    "calc_manual_back_on": (["Application.Calculation = xlCalculationManual", 'ws.Range("A1").Value = 3',
                             'ws.Range("B1").Formula = "=A1*2"', "ws.Calculate", "ws.EnableCalculation = False",
                             'ws.Range("A1").Value = 5', "ws.EnableCalculation = True"],
                            ['ws.Range("B1").Value']),
    "calc_manual_back_on_calculate": (["Application.Calculation = xlCalculationManual", 'ws.Range("A1").Value = 3',
                                       'ws.Range("B1").Formula = "=A1*2"', "ws.Calculate",
                                       "ws.EnableCalculation = False", 'ws.Range("A1").Value = 5',
                                       "ws.EnableCalculation = True", "ws.Calculate"],
                                      ['ws.Range("B1").Value']),
    "rtl_left": (["ws.DisplayRightToLeft = True"],
                 ['ws.Range("B1").Left', 'ws.Range("B1").Width', 'ws.Range("A2").Top', "ActiveWindow.DisplayRightToLeft"]),
    "rtl_window": (["ActiveWindow.DisplayRightToLeft = True"], ["ws.DisplayRightToLeft"]),
}

#: Every setting made on one sheet, which is then copied, saved and opened again.
EVERYTHING = ['ws.ScrollArea = "A1:D10"', "ws.EnableAutoFilter = True", "ws.EnableOutlining = True",
              "ws.EnablePivotTable = True", "ws.EnableCalculation = False",
              "ws.TransitionExpEval = True", "ws.TransitionFormEntry = True", "ws.DisplayRightToLeft = True"]
READINGS = ["ScrollArea", "EnableAutoFilter", "EnableOutlining", "EnablePivotTable", "EnableCalculation",
            "TransitionExpEval", "TransitionFormEntry", "DisplayRightToLeft"]


def setting_code(index: int, name: str, values: list[str]) -> str:
    lines = [f"Private Function Setting{index}() As String",
             "    Dim wb As Workbook, ws As Worksheet, start As Variant, out As String",
             "    Set wb = Book()", "    Set ws = wb.Worksheets(1)", f"    start = ws.{name}",
             f'    out = "{name}^" & Shown(start)', "    On Error Resume Next"]
    for value in values:
        shown = value.replace('"', '""')
        lines += ["    Err.Clear", f"    ws.{name} = {value}",
                  f'    out = out & "^{shown}=>" & Failure() & "|" & Shown(ws.{name})']
    lines += ["    On Error GoTo 0", "    wb.Close False", f'    Setting{index} = out & "~"', "End Function"]
    return "\n".join(lines) + "\n"


def effect_code(index: int, name: str, steps: list[str], readings: list[str]) -> str:
    lines = [f"Private Function Effect{index}() As String",
             "    Dim wb As Workbook, ws As Worksheet, out As String",
             "    Set wb = Book()", "    Set ws = wb.Worksheets(1)", "    On Error Resume Next", f'    out = "{name}^"']
    for step in steps:
        lines += ["    Err.Clear", f"    {step}", "    out = out & Failure()"]
    for reading in readings:
        lines += [f'    out = out & "^" & Shown({reading}) & Failure()']
    lines += ["    ActiveWindow.DisplayRightToLeft = False", "    Application.Calculation = xlCalculationAutomatic",
              "    On Error GoTo 0", "    wb.Close False", f'    Effect{index} = out & "~"', "End Function"]
    return "\n".join(lines) + "\n"


def readings_of(target: str) -> str:
    return " & \"|\" & ".join(f"Shown({target}.{name})" for name in READINGS)


SAVED = '''Public Function Saved(folder As String) As String
    Dim wb As Workbook, ws As Worksheet, out As String
    Application.DisplayAlerts = False
    Set wb = Book()
    Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = 1
    ws.Range("B1").Formula = "=A1*2"
{steps}
    ws.Range("A1").Value = 5
    ws.Range("C1").Formula = "=A1+1"
    ws.Copy After:=ws
    out = "set^^" & {set_readings} & "~copy^^" & {copy_readings}
    wb.SaveAs folder & "\\saved.xlsx", 51
    out = out & "~after_save^^" & {set_readings}
    out = out & "~values_after_save^^" & Shown(ws.Range("B1").Value) & "|" & Shown(ws.Range("C1").Value)
    wb.Close False
    Set wb = Workbooks.Open(folder & "\\saved.xlsx")
    out = out & "~reopened^^" & {reopen_readings} & "~reopened_copy^^" & {reopen_copy}
    out = out & "~values_reopened^^" & Shown(wb.Worksheets(1).Range("B1").Value) & "|" & _
        Shown(wb.Worksheets(1).Range("C1").Value)
    wb.Worksheets(1).Range("A1").Value = 9
    out = out & "~values_edited^^" & Shown(wb.Worksheets(1).Range("B1").Value) & "|" & _
        Shown(wb.Worksheets(1).Range("C1").Value)
    wb.Close False
    Saved = out & "~"
End Function
'''


def main() -> None:
    code = HELPER + "".join(setting_code(index, name, values) for index, (name, values) in enumerate(SETTINGS.items()))
    code += "".join(effect_code(index, name, steps, readings)
                    for index, (name, (steps, readings)) in enumerate(EFFECTS.items()))
    calls = "".join(f"    out = out & Setting{index}()\n" for index in range(len(SETTINGS)))
    calls += "".join(f"    out = out & Effect{index}()\n" for index in range(len(EFFECTS)))
    code += ("Public Function Probe() As String\n    Dim out As String\n    Application.DisplayAlerts = False\n"
             + calls + "    Probe = out\nEnd Function\n")
    saved = HELPER + SAVED.replace("{steps}", "\n".join("    " + step for step in EVERYTHING)) \
        .replace("{set_readings}", readings_of("ws")).replace("{copy_readings}", readings_of("wb.Worksheets(2)")) \
        .replace("{reopen_readings}", readings_of("wb.Worksheets(1)")) \
        .replace("{reopen_copy}", readings_of("wb.Worksheets(2)"))
    with tempfile.TemporaryDirectory(prefix="pyopenvba-worksheet-") as name:
        folder = Path(name).resolve()
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            probe = excel.run_vba(code, "Probe", timeout=900.0)
            assert probe.ok, f"{probe.outcome}: {probe.message} {probe.error}"
            stored = excel.run_vba(saved, "Saved", [str(folder)], timeout=300.0)
            assert stored.ok, f"{stored.outcome}: {stored.message} {stored.error}"
        parts: dict[str, list[str]] = {}
        with zipfile.ZipFile(folder / "saved.xlsx") as package:
            for index in (1, 2):
                sheet = package.read(f"xl/worksheets/sheet{index}.xml").decode("utf-8")
                parts[f"sheet{index}"] = re.findall(r"<sheetPr\b[^>]*>|<sheetView\b[^>]*>|<sheetProtection\b[^>]*>",
                                                    sheet)
                # What the sheet's formulas hold in the file, the frozen sheet's as it kept them.
                parts[f"values{index}"] = re.findall(r'<c r="[BC]1"[^>]*>.*?<v>(.*?)</v>', sheet)
    cases = [entry.split("^") for entry in (str(probe.value) + str(stored.value)).split("~") if entry]
    record = {"helper": HELPER, "probe": code, "saved": saved, "settings": SETTINGS,
              "effects": {key: [steps, readings] for key, (steps, readings) in EFFECTS.items()},
              "everything": EVERYTHING, "readings": READINGS, "cases": cases, "parts": parts}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for case in cases:
        print(ascii(case)[:900])
    print(parts)


if __name__ == "__main__":
    main()
