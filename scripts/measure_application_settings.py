"""Application settings everyday macros switch, in live Excel: each one's value as Excel starts, and what each
setting a macro gives it reads back as.

Each property is read, set to each value in its list, an error's number and description noted where a setting
raises one, read back after each, and set back to the value it started with, which is checked, since Excel keeps
many of these in the user's settings. A few cases read how two settings bear on each other, and
Worksheet.DisplayPageBreaks is set, saved and read again from the file. DisplayRecentFiles, which can clear the
recent-files list when turned off, and AutomationSecurity, a security setting, are left alone.

    python scripts/measure_application_settings.py

writes tests/fixtures/application_settings.json, which tests/test_excel_application_settings.py replays.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "application_settings.json"

#: Each value a setting may not take comes after one it may, so its reading shows whether it was ignored.
FLAG_VALUES = ["Not start", "0", "1", '"True"', '"False"', "True", "Null", "True", '"abc"', "False", '"1"', "True",
               '"0"', "True", '"yes"', "False", "0.4", "False", "2", "True", "Empty", "False", '""']


def after(baseline: str, values: list[str]) -> list[str]:
    """Each of ``values`` set after ``baseline``."""
    return [one for value in values for one in (baseline, value)]


#: property -> the values set, as VBA expressions; ``start`` is the value it began with.
SETTINGS: dict[str, list[str]] = {
    "DisplayStatusBar": FLAG_VALUES,
    "DisplayFormulaBar": FLAG_VALUES,
    "DisplayScrollBars": FLAG_VALUES,
    "DisplayNoteIndicator": FLAG_VALUES,
    "DisplayFullScreen": FLAG_VALUES,
    "EnableAnimations": FLAG_VALUES,
    "PrintCommunication": FLAG_VALUES,
    "Interactive": FLAG_VALUES,
    "AskToUpdateLinks": FLAG_VALUES,
    "AlertBeforeOverwriting": FLAG_VALUES,
    "CellDragAndDrop": FLAG_VALUES,
    "CopyObjectsWithCells": FLAG_VALUES,
    "EditDirectlyInCell": FLAG_VALUES,
    "EnableAutoComplete": FLAG_VALUES,
    "FixedDecimal": FLAG_VALUES,
    "MoveAfterReturn": FLAG_VALUES,
    "UserControl": FLAG_VALUES,
    **{name: FLAG_VALUES for name in (
        "AutoFormatAsYouTypeReplaceHyperlinks", "AutoPercentEntry", "ChartDataPointTrack", "ConstrainNumeric",
        "ControlCharacters", "DeferAsyncQueries", "DisplayFormulaAutoComplete", "DisplayFunctionToolTips",
        "DisplayInsertOptions", "DisplayPasteOptions", "EnableCheckFileExtensions", "EnableLargeOperationAlert",
        "EnableLivePreview", "EnableMacroAnimations", "EnableSound", "ExtendList", "FlashFill", "FlashFillMode",
        "FormatStaleValues", "GenerateGetPivotData", "MapPaperSize", "PivotTableSelection", "PromptForSummaryInfo",
        "RollZoom", "ShowChartTipNames", "ShowChartTipValues", "ShowConvertToDataType",
        "ShowQuickAnalysis", "ShowToolTips", "TransitionNavigKeys",
        "WarnOnFunctionNameConflict")},
    "Cursor": ["2", "1", "3", "-4143", *after("2", ["5", "0", "-1", "4", "-4142", '"3"', "Null", "1.5", "2.5",
                                                      "3.5", '"abc"', "Empty"])],
    "EnableCancelKey": ["0", "2", "1", *after("2", ["3", "4", "5", "6", "-1", "-2", "-3", "255", "256", "65536",
                                                      '"0"', "Null", "0.5", "1.5", '"abc"', "Empty"]),
                        *after("0", ["-1", "True", "-2", "3"]), *after("1", ["-1", "True", "False", "-2", "3"])],
    "MoveAfterReturnDirection": ["-4162", "-4161", "-4159", "-4121",
                                 *after("-4162", ["0", "1", "2", "3", "4", "5", "-1", "-4160", "-4163", "-4120",
                                                  "-4122", "65536", '"-4161"', "Null", '"abc"', "Empty"])],
    "DisplayCommentIndicator": ["0", "1", "-1", *after("1", ["2", "-2", "3", '"0"', "Null", "0.5", '"abc"',
                                                              "Empty"])],
    "FixedDecimalPlaces": ["4", "0", "-1", "15", "16", *after("5", ["100", "-100", "32767", "1.6", "2.5", '"3"',
                                                                     "Null", '"abc"', "Empty", "127", "128", "255",
                                                                     "256", "1000", "10000", "32766", "-128", "-129",
                                                                     "-1000", "-32768", "-32769", "32768", "300",
                                                                     "301", "-300", "-301"])],
    "Caption": ['"My caption"', '""', "5", "True", *after('"x"', ["Null", "Empty"]), "String(300, \"c\")",
                "String(200, \"d\")", "String(201, \"e\")"],
}

PROBE = '''Private Function Shown(value As Variant) As String
    If IsNull(value) Then
        Shown = "Null"
    Else
        Shown = TypeName(value) & ":" & CStr(value)
    End If
End Function

Private Function Failure() As String
    If Err.Number <> 0 Then Failure = "E" & Err.Number & ":" & Err.Description
    Err.Clear
End Function
'''

#: Two settings read together, and the steps between: name -> (steps, readings).
TOGETHER: dict[str, tuple[list[str], list[str]]] = {
    "indicator_off": (["Application.DisplayCommentIndicator = 0"],
                      ["Application.DisplayNoteIndicator", "Application.DisplayCommentIndicator"]),
    "indicator_note_on": (["Application.DisplayCommentIndicator = 0", "Application.DisplayNoteIndicator = True"],
                          ["Application.DisplayNoteIndicator", "Application.DisplayCommentIndicator"]),
    "indicator_note_off": (["Application.DisplayCommentIndicator = 1", "Application.DisplayNoteIndicator = False"],
                           ["Application.DisplayNoteIndicator", "Application.DisplayCommentIndicator"]),
    "indicator_shown": (["Application.DisplayCommentIndicator = 1"],
                        ["Application.DisplayNoteIndicator", "Application.DisplayCommentIndicator"]),
    "indicator_shown_note_on": (["Application.DisplayCommentIndicator = 1", "Application.DisplayNoteIndicator = True"],
                                ["Application.DisplayNoteIndicator", "Application.DisplayCommentIndicator"]),
    "status_hidden": (["Application.DisplayStatusBar = False", 'Application.StatusBar = "busy"'],
                      ["Application.StatusBar", "Application.DisplayStatusBar"]),
    "print_off": (["Application.PrintCommunication = False", "ws.PageSetup.Orientation = 2",
                   "ws.PageSetup.LeftMargin = 36"],
                  ["ws.PageSetup.Orientation", "ws.PageSetup.LeftMargin", "Application.PrintCommunication"]),
    "print_back_on": (["Application.PrintCommunication = False", "ws.PageSetup.Orientation = 2",
                       "Application.PrintCommunication = True"],
                      ["ws.PageSetup.Orientation"]),
    "print_last_wins": (["Application.PrintCommunication = False", "ws.PageSetup.Orientation = 2",
                         "ws.PageSetup.Orientation = 1", "ws.PageSetup.LeftMargin = 36", "ws.PageSetup.LeftMargin = 72",
                         "Application.PrintCommunication = True"],
                        ["ws.PageSetup.Orientation", "ws.PageSetup.LeftMargin"]),
    "print_header": (["Application.PrintCommunication = False", 'ws.PageSetup.CenterHeader = "x"',
                      "ws.PageSetup.DifferentFirstPageHeaderFooter = True"],
                     ["ws.PageSetup.CenterHeader", "ws.PageSetup.DifferentFirstPageHeaderFooter"]),
    "print_header_back_on": (["Application.PrintCommunication = False", 'ws.PageSetup.CenterHeader = "x"',
                              "Application.PrintCommunication = True"],
                             ["ws.PageSetup.CenterHeader"]),
    "print_header_too_long": (["Application.PrintCommunication = False", 'ws.PageSetup.CenterHeader = String(300, "a")',
                               "Application.PrintCommunication = True"],
                              ["ws.PageSetup.CenterHeader"]),
    "print_zoom_invalid": (["Application.PrintCommunication = False", "ws.PageSetup.Zoom = 5",
                            "Application.PrintCommunication = True"],
                           ["ws.PageSetup.Zoom"]),
    "print_area": (["Application.PrintCommunication = False", 'ws.PageSetup.PrintArea = "$A$1:$B$2"'],
                   ["ws.PageSetup.PrintArea"]),
    "print_twice_off": (["Application.PrintCommunication = False", "ws.PageSetup.Orientation = 2",
                         "Application.PrintCommunication = False"],
                        ["ws.PageSetup.Orientation"]),
    "print_other_object": (["Application.PrintCommunication = False", "ws.PageSetup.Orientation = 2"],
                           ["wb.Worksheets(1).PageSetup.Orientation"]),
    "print_page_text": (["Application.PrintCommunication = False", 'ws.PageSetup.FirstPage.CenterHeader.Text = "F"'],
                        ["ws.PageSetup.FirstPage.CenterHeader.Text"]),
    "print_page_text_back_on": (["Application.PrintCommunication = False",
                                 'ws.PageSetup.FirstPage.CenterHeader.Text = "F"', "Application.PrintCommunication = True"],
                                ["ws.PageSetup.FirstPage.CenterHeader.Text"]),
    "print_titles": (["Application.PrintCommunication = False", 'ws.PageSetup.PrintTitleRows = "$1:$2"'],
                     ["ws.PageSetup.PrintTitleRows"]),
    "print_coercion": (["Application.PrintCommunication = False", 'ws.PageSetup.Orientation = "abc"',
                        "ws.PageSetup.LeftMargin = Null", "Application.PrintCommunication = True"],
                       ["ws.PageSetup.Orientation", "ws.PageSetup.LeftMargin"]),
    "print_two_sheets": (["Application.PrintCommunication = False", "wb.Worksheets.Add", "ws.PageSetup.Orientation = 2",
                          "wb.Worksheets(1).PageSetup.Orientation = 2", "Application.PrintCommunication = True"],
                         ["ws.PageSetup.Orientation", "wb.Worksheets(1).PageSetup.Orientation"]),
    "print_book_closed": (["Application.PrintCommunication = False", "ws.PageSetup.Orientation = 2", "wb.Close False",
                           "Application.PrintCommunication = True"],
                          ["Application.PrintCommunication"]),
    "fixed_decimal_entry": (["Application.FixedDecimal = True", 'ws.Range("A1").Value = 1234', 'ws.Range("A2").Formula = "=5"'],
                            ['ws.Range("A1").Value', 'ws.Range("A2").Value']),
    "decimal_places_alone": (["Application.FixedDecimalPlaces = 3"], ["Application.FixedDecimal"]),
    "caption_window": (['wb.Windows(1).Caption = "Named window"', 'Application.Caption = "Mine"'],
                       ["Application.Caption"]),
}


def setting_code(index: int, name: str, values: list[str]) -> str:
    lines = [f"Private Function Setting{index}() As String",
             "    Dim start As Variant, out As String",
             f"    start = Application.{name}",
             f'    out = "{name}^" & Shown(start)',
             "    On Error Resume Next"]
    for value in values:
        shown = value.replace('"', '""')
        lines += ["    Err.Clear", f"    Application.{name} = {value}",
                  f'    out = out & "^{shown}=>" & Failure() & "|" & Shown(Application.{name})']
    # The caption reads with the active workbook's name before it, and an empty one is Excel's own.
    back = '""' if name == "Caption" else "start"
    lines += ["    Err.Clear", f"    Application.{name} = {back}",
              f'    out = out & "^restored=" & (Application.{name} = start) & Failure()',
              "    On Error GoTo 0", f'    Setting{index} = out & "~"', "End Function"]
    return "\n".join(lines) + "\n"


def together_code(index: int, name: str, steps: list[str], readings: list[str]) -> str:
    """Steps on a new workbook's sheet ``ws``, the readings after them, and every setting the steps touch set back."""
    touched = sorted({found for step in steps for found in re.findall(r"Application\.(\w+)", step)})
    lines = [f"Private Function Together{index}() As String",
             "    Dim wb As Object, ws As Object, out As String",
             *[f"    Dim start{n} As Variant" for n in range(len(touched))],
             *[f"    start{n} = Application.{prop}" for n, prop in enumerate(touched)],
             "    Set wb = Workbooks.Add(xlWBATWorksheet)", "    Set ws = wb.Worksheets(1)",
             "    On Error Resume Next", f'    out = "{name}^"']
    for step in steps:
        lines += ["    Err.Clear", f"    {step}", '    out = out & Failure()']
    for reading in readings:
        lines += [f'    out = out & "^" & Shown({reading}) & Failure()']
    # The status bar goes back to Excel's own with False, its reading FALSE no setting of it, and the caption with an
    # empty one, its reading holding the window's caption.
    lines += [*[f"    Application.{prop} = " + {"StatusBar": "False", "Caption": "\"\""}.get(prop, f"start{n}")
                for n, prop in enumerate(touched)],
              *[f'    out = out & "^restored {prop}=" & (Application.{prop} = start{n})'
                for n, prop in enumerate(touched)],
              "    wb.Close False", "    On Error GoTo 0", f'    Together{index} = out & "~"', "End Function"]
    return "\n".join(lines) + "\n"


BREAKS = '''Public Function Breaks(folder As String) As String
    Dim wb As Workbook, out As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    out = "page_breaks^" & Shown(wb.Worksheets(1).DisplayPageBreaks)
    wb.Worksheets(1).DisplayPageBreaks = True
    out = out & "^" & Shown(wb.Worksheets(1).DisplayPageBreaks)
    On Error Resume Next
    wb.Worksheets(1).DisplayPageBreaks = Null
    out = out & "^null=>" & Failure() & "|" & Shown(wb.Worksheets(1).DisplayPageBreaks)
    wb.Worksheets(1).DisplayPageBreaks = 0
    out = out & "^zero=>" & Failure() & "|" & Shown(wb.Worksheets(1).DisplayPageBreaks)
    wb.Worksheets(1).DisplayPageBreaks = True
    On Error GoTo 0
    wb.SaveAs folder & "\\breaks.xlsx", 51
    wb.Close False
    Set wb = Workbooks.Open(folder & "\\breaks.xlsx")
    out = out & "^reopened=" & Shown(wb.Worksheets(1).DisplayPageBreaks)
    wb.Close False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Application.PrintCommunication = False
    wb.Worksheets(1).PageSetup.Orientation = 2
    wb.SaveAs folder & "\\queued.xlsx", 51
    out = out & "~queued^after_save=" & Shown(wb.Worksheets(1).PageSetup.Orientation)
    Application.PrintCommunication = True
    out = out & "^back_on=" & Shown(wb.Worksheets(1).PageSetup.Orientation)
    wb.Close False
    Set wb = Workbooks.Open(folder & "\\queued.xlsx")
    out = out & "^reopened=" & Shown(wb.Worksheets(1).PageSetup.Orientation)
    wb.Close False
    Breaks = out & "~"
End Function
'''


def main() -> None:
    code = PROBE + "".join(setting_code(index, name, values) for index, (name, values) in enumerate(SETTINGS.items()))
    code += "".join(together_code(index, name, steps, readings)
                    for index, (name, (steps, readings)) in enumerate(TOGETHER.items()))
    calls = "".join(f"    out = out & Setting{index}()\n" for index in range(len(SETTINGS)))
    calls += "".join(f"    out = out & Together{index}()\n" for index in range(len(TOGETHER)))
    # The status bar as Excel starts, before any setting here touches it.
    code += ("Public Function Probe() As String\n    Dim out As String\n"
             '    out = "status_start^" & Shown(Application.StatusBar) & "~"\n' + calls + "    Probe = out\nEnd Function\n")
    with tempfile.TemporaryDirectory(prefix="pyopenvba-settings-") as name:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            probe = excel.run_vba(code, "Probe", timeout=600.0)
            assert probe.ok, f"{probe.outcome}: {probe.message} {probe.error}"
            breaks = excel.run_vba(PROBE + BREAKS, "Breaks", [str(Path(name).resolve())], timeout=120.0)
            assert breaks.ok, f"{breaks.outcome}: {breaks.message} {breaks.error}"
            with zipfile.ZipFile(Path(name) / "breaks.xlsx") as package:
                sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
            views = re.findall(r"<sheetView\b[^>]*>", sheet)
            with zipfile.ZipFile(Path(name) / "queued.xlsx") as package:
                queued = re.findall(r"<pageSetup\b[^>]*>", package.read("xl/worksheets/sheet1.xml").decode("utf-8"))
    cases = [entry.split("^") for entry in (str(probe.value) + str(breaks.value)).split("~") if entry]
    unrestored = [case[0] for case in cases[1:len(SETTINGS) + 1] if "restored=True" not in case[-1]]
    assert not unrestored, f"not set back: {unrestored}"
    record = {"probe": code, "breaks": BREAKS, "settings": SETTINGS,
              "together": {key: [steps, readings] for key, (steps, readings) in TOGETHER.items()},
              "cases": cases, "sheet_views": views, "queued_page_setup": queued}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for case in cases:
        print(ascii(case))
    print(views, queued)


if __name__ == "__main__":
    main()
