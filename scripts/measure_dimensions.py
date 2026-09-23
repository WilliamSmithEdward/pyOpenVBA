"""Row heights, column widths and hidden rows and columns as Excel's object model reports them.

Every probe runs in a workbook of its own, sets something, and reads a
list of properties back. A read records the value's type as well as its
text, Null for a mixed range, and E<number> for an error, so the replay
in tests/test_excel_dimensions.py holds the model to all three. Excel
rounds a height or width to whole pixels of the display it runs on, so
the probes record StandardHeight and the replay requires the 96-DPI
display the model emulates (15pt rows).

    python scripts/measure_dimensions.py

writes tests/fixtures/dimensions.json.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent

HELPER = '''Private Function Show(v As Variant) As String
    If IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function
'''

HEIGHTS = (0.05, 0.1, 0.2, 0.25, 0.3, 0.5, 0.63, 0.7, 0.825, 0.9, 1, 1.025, 1.1, 1.3, 1.5, 2, 3.3, 7.77, 10, 12.3,
           14.9, 15, 15.1, 15.3, 15.5, 15.7, 15.9, 19.9, 20, 20.1, 20.15625, 20.2, 20.37, 33.33, 100.01, 123.456,
           409.4, 409.5, 409.52, 0.0925, 0.28, 0.47, 0.66, 0.84, 0.0937, 0.2812)
WIDTHS = (0, 0.01, 0.04166, 0.0417, 0.05, 0.08, 0.2, 0.25, 0.3, 0.5, 0.7, 0.9, 0.9583, 0.95834, 0.99, 1, 1.05,
          1.07, 1.08, 1.5, 2, 2.5, 3.3, 5.07, 5.08, 8, 8.43, 8.4999, 8.5, 8.57, 8.64, 8.65, 9, 10, 10.07, 10.08, 10.1,
          10.21, 10.22, 12.345, 20.5, 100.07, 254.9, 255)


def reads(target: str, names: tuple[str, ...]) -> list[str]:
    return [f"{target}.{name}" for name in names]


ROW = ("RowHeight", "Height", "Hidden", "UseStandardHeight")
COL = ("ColumnWidth", "Width", "Hidden", "UseStandardWidth")

PROBES: dict[str, tuple[str, list[str]]] = {
    # --- a sheet nobody has touched ----------------------------------------------------
    "defaults": ("", ["ws.StandardHeight", "ws.StandardWidth", *reads("ws.Rows(1)", ROW), *reads("ws.Columns(1)", COL),
                      'ws.Range("A1").Left', 'ws.Range("A1").Top', 'ws.Range("A1").Width', 'ws.Range("A1").Height',
                      'ws.Range("B1").Left', 'ws.Range("A2").Top', 'ws.Range("C3").RowHeight',
                      'ws.Range("C3").ColumnWidth', 'ws.Range("C3").UseStandardHeight',
                      'ws.Range("C3").UseStandardWidth', "ws.UsedRange.Address"]),
    # --- setting a height or a width ------------------------------------------------------
    "row_height_values": ("For Each h In Array(" + ", ".join(str(one) for one in HEIGHTS) + ")\n"
                          "i = i + 1\nws.Rows(i).RowHeight = h\n"
                          'Report = Report & ws.Rows(i).RowHeight & ":" & ws.Rows(i).Height & ":" & ws.Rows(i).Hidden '
                          '& ":" & ws.Rows(i).UseStandardHeight & ","\nNext h', ["ws.UsedRange.Address"]),
    "row_height_grid": ("For i = 0 To 2184\nws.Rows(i + 1).RowHeight = i * 0.1875\n"
                        'Report = Report & ws.Rows(i + 1).RowHeight & ":" & ws.Rows(i + 1).Height & ","\nNext i', []),
    "column_width_values": ("For Each h In Array(" + ", ".join(str(one) for one in WIDTHS) + ")\n"
                            "i = i + 1\nws.Columns(i).ColumnWidth = h\n"
                            'Report = Report & ws.Columns(i).ColumnWidth & ":" & ws.Columns(i).Width & ":" & '
                            'ws.Columns(i).Hidden & ":" & ws.Columns(i).UseStandardWidth & ","\nNext h',
                            ["ws.UsedRange.Address"]),
    "column_width_grid": ("For i = 0 To 1790\nIf i < 12 Then\nws.Columns(i + 1).ColumnWidth = i / 12\nElse\n"
                          "ws.Columns(i + 1).ColumnWidth = (i - 5) / 7\nEnd If\n"
                          'Report = Report & ws.Columns(i + 1).ColumnWidth & ":" & ws.Columns(i + 1).Width & ","\n'
                          "Next i", []),
    "row_height_bad": ("On Error Resume Next\nFor Each h In Array(-1, -0.01, -0.02, -0.025, -0.03, 409.55, 410, 1E+300, "
                       '"abc", "12", "12.5", True, False)\ni = i + 1\nErr.Clear\nws.Rows(i).RowHeight = h\n'
                       'Report = Report & Err.Number & ":" & ws.Rows(i).RowHeight & ":" & ws.Rows(i).Hidden & ","\n'
                       "Next h", []),
    "column_width_bad": ("On Error Resume Next\nFor Each h In Array(-1, -0.001, 255.0001, 255.001, 256, 1E+300, "
                         '"abc", "12", "12.5", True, False)\ni = i + 1\nErr.Clear\nws.Columns(i).ColumnWidth = h\n'
                         'Report = Report & Err.Number & ":" & ws.Columns(i).ColumnWidth & ":" & ws.Columns(i).Hidden '
                         '& ","\nNext h', []),
    "row_height_on_cells": ('ws.Range("B2:C3").RowHeight = 20\nws.Range("A5,A7").RowHeight = 30',
                            [*reads("ws.Rows(2)", ROW), *reads("ws.Rows(3)", ROW), *reads("ws.Rows(5)", ROW),
                             *reads("ws.Rows(6)", ROW), *reads("ws.Rows(7)", ROW), "ws.UsedRange.Address"]),
    "column_width_on_cells": ('ws.Range("B2:C3").ColumnWidth = 20\nws.Range("E1,G1").ColumnWidth = 30',
                              [*reads("ws.Columns(2)", COL), *reads("ws.Columns(3)", COL), *reads("ws.Columns(5)", COL),
                               *reads("ws.Columns(6)", COL), *reads("ws.Columns(7)", COL), "ws.UsedRange.Address"]),
    "sizes_on_many": ('ws.Rows("1:1048576").RowHeight = 20\nws.Columns("A:XFD").ColumnWidth = 12',
                      [*reads("ws.Rows(5)", ROW), *reads("ws.Rows(1048576)", ROW), *reads("ws.Columns(5)", COL),
                       *reads("ws.Columns(16384)", COL), "ws.StandardHeight", "ws.StandardWidth",
                       "ws.Cells.RowHeight", "ws.Cells.ColumnWidth", "ws.UsedRange.Address"]),
    # --- hiding --------------------------------------------------------------------------
    "hide_custom_row": ("ws.Rows(2).RowHeight = 20\nws.Rows(2).Hidden = True",
                        [*reads("ws.Rows(2)", ROW), 'ws.Range("A3").Top']),
    "hide_row_by_height": ("ws.Rows(2).RowHeight = 20\nws.Rows(2).RowHeight = 0", reads("ws.Rows(2)", ROW)),
    "unhide_row_after_zero": ("ws.Rows(2).RowHeight = 20\nws.Rows(2).RowHeight = 0\nws.Rows(2).Hidden = False",
                              reads("ws.Rows(2)", ROW)),
    "unhide_custom_row": ("ws.Rows(2).RowHeight = 20\nws.Rows(2).Hidden = True\nws.Rows(2).Hidden = False",
                          reads("ws.Rows(2)", ROW)),
    "height_unhides_row": ("ws.Rows(2).Hidden = True\nws.Rows(2).RowHeight = 30", reads("ws.Rows(2)", ROW)),
    "hide_standard_row": ("ws.Rows(2).Hidden = True", [*reads("ws.Rows(2)", ROW), "ws.UsedRange.Address"]),
    "unhide_standard_row": ("ws.Rows(2).Hidden = True\nws.Rows(2).Hidden = False", reads("ws.Rows(2)", ROW)),
    "zero_pixel_row": ("ws.Rows(2).RowHeight = 0.1", reads("ws.Rows(2)", ROW)),
    "hide_custom_column": ("ws.Columns(2).ColumnWidth = 20\nws.Columns(2).Hidden = True",
                           [*reads("ws.Columns(2)", COL), 'ws.Range("C1").Left']),
    "hide_column_by_width": ("ws.Columns(2).ColumnWidth = 20\nws.Columns(2).ColumnWidth = 0", reads("ws.Columns(2)", COL)),
    "unhide_column_after_zero": ("ws.Columns(2).ColumnWidth = 20\nws.Columns(2).ColumnWidth = 0\n"
                                 "ws.Columns(2).Hidden = False", reads("ws.Columns(2)", COL)),
    "unhide_custom_column": ("ws.Columns(2).ColumnWidth = 20\nws.Columns(2).Hidden = True\nws.Columns(2).Hidden = False",
                             reads("ws.Columns(2)", COL)),
    "width_unhides_column": ("ws.Columns(2).Hidden = True\nws.Columns(2).ColumnWidth = 30", reads("ws.Columns(2)", COL)),
    "hide_standard_column": ("ws.Columns(2).Hidden = True", [*reads("ws.Columns(2)", COL), "ws.UsedRange.Address"]),
    "unhide_standard_column": ("ws.Columns(2).Hidden = True\nws.Columns(2).Hidden = False",
                               [*reads("ws.Columns(2)", COL), "ws.StandardWidth"]),
    "zero_pixel_column": ("ws.Columns(2).ColumnWidth = 0.04", reads("ws.Columns(2)", COL)),
    "hidden_several": ('ws.Rows("12:13").Hidden = True\nws.Columns("H:I").Hidden = True\n'
                       'ws.Range("A20:B21").EntireRow.Hidden = True\nws.Range("K1:L2").EntireColumn.Hidden = True',
                       ['ws.Rows("12:13").Hidden', "ws.Rows(12).Hidden", "ws.Rows(13).Hidden", 'ws.Rows("11:12").Hidden',
                        'ws.Rows("13:14").Hidden', 'ws.Columns("H:I").Hidden', 'ws.Columns("G:H").Hidden',
                        'ws.Columns("I:J").Hidden', "ws.Rows(20).Hidden", "ws.Rows(21).Hidden", "ws.Columns(11).Hidden",
                        "ws.Columns(12).Hidden", 'ws.Columns("H:I").ColumnWidth', 'ws.Columns("H:I").Width',
                        'ws.Rows("12:13").RowHeight', 'ws.Rows("12:13").Height', "ws.UsedRange.Address"]),
    "hidden_needs_whole": ("", ['ws.Range("A1").Hidden', 'ws.Range("A1:B2").Hidden', "ws.Cells.Hidden",
                                'ws.Range("A1:A3").EntireRow.Hidden', 'ws.Range("A1:C1").EntireColumn.Hidden',
                                'ws.Range("1:2,4:5").Hidden', 'ws.Range("A:A,C:C").Hidden']),
    "hidden_set_on_cells": ('ws.Range("A6").Hidden = True', ["ws.Rows(6).Hidden", "ws.Columns(1).Hidden"]),
    "hidden_values": ('On Error Resume Next\nFor Each h In Array(5, -1, 0, "True", "yes", "0", 0.4, Empty)\n'
                      "i = i + 1\nErr.Clear\nws.Rows(i).Hidden = h\n"
                      'Report = Report & Err.Number & ":" & ws.Rows(i).Hidden & ","\nNext h', []),
    "hidden_geometry": ('ws.Rows(3).Hidden = True\nws.Columns(3).Hidden = True',
                        ['ws.Range("A4").Top', 'ws.Range("D1").Left', 'ws.Range("C3").Width', 'ws.Range("C3").Height',
                         'ws.Range("B2:D4").Width', 'ws.Range("B2:D4").Height', 'ws.Rows(3).Top', 'ws.Columns(3).Left']),
    # --- standard sizes ------------------------------------------------------------------------
    "use_standard_height": ("ws.Rows(2).RowHeight = 20\nws.Rows(3).RowHeight = 15\nws.Rows(4).UseStandardHeight = False\n"
                            "ws.Rows(5).RowHeight = 20\nws.Rows(5).UseStandardHeight = True",
                            [*reads("ws.Rows(2)", ROW), *reads("ws.Rows(3)", ROW), *reads("ws.Rows(4)", ROW),
                             *reads("ws.Rows(5)", ROW), 'ws.Range("A1:A2").UseStandardHeight',
                             'ws.Range("A2:A3").UseStandardHeight', 'ws.Rows("1:2").UseStandardHeight',
                             'ws.Rows("2:3").UseStandardHeight', "ws.UsedRange.Address"]),
    "use_standard_width": ("ws.Columns(2).ColumnWidth = 20\nws.Columns(3).ColumnWidth = 8.43\n"
                           "ws.Columns(4).UseStandardWidth = False\nws.Columns(5).ColumnWidth = 20\n"
                           "ws.Columns(5).UseStandardWidth = True\nws.Columns(6).ColumnWidth = 8.4999",
                           [*reads("ws.Columns(2)", COL), *reads("ws.Columns(3)", COL), *reads("ws.Columns(4)", COL),
                            *reads("ws.Columns(5)", COL), *reads("ws.Columns(6)", COL),
                            'ws.Range("A1:B1").UseStandardWidth', 'ws.Columns("A:B").UseStandardWidth',
                            'ws.Columns("B:C").UseStandardWidth']),
    "use_standard_bad": ('On Error Resume Next\nws.Rows(2).UseStandardHeight = "maybe"\n'
                         'Report = Report & Err.Number & ";"\nErr.Clear\nws.Columns(2).UseStandardWidth = "maybe"\n'
                         'Report = Report & Err.Number & ";"', ["ws.Rows(2).UseStandardHeight",
                                                                "ws.Columns(2).UseStandardWidth"]),
    "standard_width": ("ws.Columns(2).ColumnWidth = 20\nws.Columns(3).ColumnWidth = 8.43\nws.StandardWidth = 12",
                       ["ws.StandardWidth", *reads("ws.Columns(1)", COL), *reads("ws.Columns(2)", COL),
                        *reads("ws.Columns(3)", COL), 'ws.Range("B1").Left']),
    "standard_width_rounds": ("ws.StandardWidth = 12.1", ["ws.StandardWidth", *reads("ws.Columns(1)", COL)]),
    "standard_width_zero": ("ws.StandardWidth = 0", ["ws.StandardWidth", *reads("ws.Columns(1)", COL)]),
    "standard_width_back": ("ws.StandardWidth = 12\nws.StandardWidth = 8.43", ["ws.StandardWidth",
                                                                              *reads("ws.Columns(1)", COL)]),
    "standard_width_bad": ('On Error Resume Next\nFor Each h In Array(256, -1, 255, "abc", 1E+300)\nErr.Clear\n'
                           'ws.StandardWidth = h\nReport = Report & Err.Number & ":" & ws.StandardWidth & ","\n'
                           "Next h", []),
    "standard_height_read_only": ("ws.StandardHeight = 20", ["ws.StandardHeight"]),
    "use_standard_shows_row": ("ws.Rows(3).RowHeight = 20\nws.Rows(3).Hidden = True\nws.Rows(3).UseStandardHeight = True",
                               reads("ws.Rows(3)", ROW)),
    "custom_column_to_standard": ("ws.Columns(3).ColumnWidth = 20\nws.Columns(3).ColumnWidth = 8.43\n"
                                  "ws.Columns(4).ColumnWidth = 20\nws.Columns(4).Hidden = True\n"
                                  "ws.Columns(4).ColumnWidth = 8.43", [*reads("ws.Columns(3)", COL),
                                                                       *reads("ws.Columns(4)", COL)]),
    "hidden_custom_column_counts": ("ws.Columns(3).ColumnWidth = 20\nws.Columns(3).Hidden = True",
                                    ["ws.UsedRange.Address"]),
    # --- every row or column at once ------------------------------------------------------
    "hide_every_row": ('ws.Rows("1:1048576").Hidden = True', ["ws.StandardHeight", *reads("ws.Rows(1)", ROW),
                                                             *reads("ws.Rows(1048576)", ROW), 'ws.Range("A5").Top',
                                                             "ws.UsedRange.Address"]),
    "show_every_row_again": ('ws.Rows("1:1048576").Hidden = True\nws.Rows("1:1048576").Hidden = False',
                             ["ws.StandardHeight", *reads("ws.Rows(1)", ROW), "ws.UsedRange.Address"]),
    "show_one_of_every_row": ('ws.Rows("1:1048576").Hidden = True\nws.Rows(5).Hidden = False',
                              ["ws.StandardHeight", *reads("ws.Rows(4)", ROW), *reads("ws.Rows(5)", ROW),
                               'ws.Range("A7").Top', "ws.UsedRange.Address"]),
    "hide_rows_to_the_end": ('ws.Rows(2).RowHeight = 20\nws.Rows("3:1048576").Hidden = True',
                             ["ws.StandardHeight", *reads("ws.Rows(1)", ROW), *reads("ws.Rows(2)", ROW),
                              *reads("ws.Rows(3)", ROW), 'ws.Range("A9").Top', "ws.UsedRange.Address"]),
    "hide_every_column": ('ws.Columns(3).ColumnWidth = 20\nws.Columns("A:XFD").Hidden = True',
                          ["ws.StandardWidth", *reads("ws.Columns(1)", COL), *reads("ws.Columns(3)", COL),
                           'ws.Range("E1").Left', "ws.UsedRange.Address"]),
    "show_every_column_again": ('ws.Columns("A:XFD").Hidden = True\nws.Columns("A:XFD").Hidden = False',
                                ["ws.StandardWidth", *reads("ws.Columns(1)", COL)]),
    # --- where things are ----------------------------------------------------------------------
    "geometry": ("ws.Rows(2).RowHeight = 20\nws.Columns(2).ColumnWidth = 20",
                 ['ws.Range("B2").Left', 'ws.Range("B2").Top', 'ws.Range("C3").Left', 'ws.Range("C3").Top',
                  'ws.Range("B2:C3").Width', 'ws.Range("B2:C3").Height', 'ws.Range("A1:C3").Width',
                  'ws.Range("A1,C3").Width', 'ws.Range("A1,C3").Height', 'ws.Range("C3,A1").Left',
                  'ws.Range("C3,A1").Top', 'ws.Range("XFD1").Left', 'ws.Range("A1048576").Top',
                  'ws.Columns("A:C").Width', 'ws.Rows("1:3").Height', "ws.Cells.Width", "ws.Cells.Height",
                  "ws.Columns(2).Left", "ws.Rows(2).Top", "ws.Columns(2).Top", "ws.Rows(2).Left"]),
    # --- reading several rows or columns at once -------------------------------------------------
    "mixed_no_cells": ("ws.Rows(20).RowHeight = 20\nws.Rows(21).RowHeight = 30\n"
                       "ws.Columns(11).ColumnWidth = 20\nws.Columns(12).ColumnWidth = 30",
                       ['ws.Rows("20:21").RowHeight', 'ws.Range("A20:B21").RowHeight', 'ws.Rows("19:20").RowHeight',
                        'ws.Columns("K:L").ColumnWidth', 'ws.Range("K1:L1").ColumnWidth', "ws.Cells.RowHeight",
                        "ws.Cells.ColumnWidth", 'ws.Rows("20:21").Height', 'ws.Columns("K:L").Width',
                        'ws.Range("A20,A21").RowHeight', 'ws.Range("K1,L1").ColumnWidth', "ws.UsedRange.Address"]),
}

# Rows 20 and 21 at 20pt and 30pt, columns K and L at 20 and 30 characters,
# and a cell or two placed around them: whether a multi-row or multi-column
# read comes back Null depends on where the sheet's cells are.
ROW_TARGETS = ['ws.Rows("19:20")', 'ws.Rows("20:21")', 'ws.Rows("19:21")', 'ws.Rows("21:22")', 'ws.Range("A19:A20")',
               'ws.Range("A19:B20")', 'ws.Range("A20:B21")', 'ws.Range("B19:B21")', 'ws.Range("A1:A20")',
               'ws.Rows("1:25")']
COL_TARGETS = ['ws.Columns("J:K")', 'ws.Columns("K:L")', 'ws.Columns("I:K")', 'ws.Columns("J:L")', 'ws.Columns("L:M")',
               'ws.Range("J1:K1")', 'ws.Range("J5:K5")', 'ws.Range("J1:K10")', 'ws.Range("J5:K9")',
               'ws.Range("K1:L1")', 'ws.Range("K5:L9")', 'ws.Range("J9:K9")']
for placement in ("A20", "B20", "A19", "A21", "A22", "A1", "A20 B21", "A19 B21", "A21 B19", "A18", "Z20", "Z19",
                  "A20 A21", "C5 D30"):
    setup = "ws.Rows(20).RowHeight = 20\nws.Rows(21).RowHeight = 30\n" + "".join(
        f'ws.Range("{cell}").Value = 1\n' for cell in placement.split())
    PROBES[f"mixed_rows_{placement.replace(' ', '_')}"] = (
        setup.rstrip("\n"), [*[f"{target}.RowHeight" for target in ROW_TARGETS], "ws.UsedRange.Address"])
for placement in ("K1", "K5", "L5", "J1", "I5", "K5 L9", "K5 L5", "K1 L9", "K9 L1", "J5 K9", "J1 K1", "K5 K9",
                  "L5 L9", "J9 K5", "K5 J20", "A5 K9", "Z5 K9", "J5 L9"):
    setup = "ws.Columns(11).ColumnWidth = 20\nws.Columns(12).ColumnWidth = 30\n" + "".join(
        f'ws.Range("{cell}").Value = 1\n' for cell in placement.split())
    PROBES[f"mixed_columns_{placement.replace(' ', '_')}"] = (
        setup.rstrip("\n"), [*[f"{target}.ColumnWidth" for target in COL_TARGETS], "ws.UsedRange.Address"])
# Hidden and custom rows side by side, with a cell in column C so the reads compare.
PAIR_ROWS = {"def": "", "c15": "ws.Rows({r}).RowHeight = 15", "c20": "ws.Rows({r}).RowHeight = 20",
             "hid": "ws.Rows({r}).Hidden = True", "h20": "ws.Rows({r}).RowHeight = 20\nws.Rows({r}).Hidden = True",
             "q1": "ws.Rows({r}).RowHeight = 0.1"}
PAIR_COLS = {"def": "", "c20": "ws.Columns({c}).ColumnWidth = 20", "hid": "ws.Columns({c}).Hidden = True",
             "h20": "ws.Columns({c}).ColumnWidth = 20\nws.Columns({c}).Hidden = True",
             "p1": "ws.Columns({c}).ColumnWidth = 0.05"}
for first, first_code in PAIR_ROWS.items():
    for second, second_code in PAIR_ROWS.items():
        setup = "\n".join(line for line in (first_code.format(r=2), second_code.format(r=3),
                                            'ws.Range("C1").Value = 1', 'ws.Range("C9").Value = 1') if line)
        PROBES[f"pair_rows_{first}_{second}"] = (setup, [f"ws.Rows(\"2:3\").{name}" for name in ROW])
for first, first_code in PAIR_COLS.items():
    for second, second_code in PAIR_COLS.items():
        setup = "\n".join(line for line in (first_code.format(c=2), second_code.format(c=3),
                                            'ws.Range("A3").Value = 1', 'ws.Range("I3").Value = 1') if line)
        PROBES[f"pair_columns_{first}_{second}"] = (setup, [f"ws.Columns(\"B:C\").{name}" for name in COL])

PROBES.update({
    # --- AutoFit, where the text does not decide the answer -----------------------------------
    "autofit_row_custom": ("ws.Rows(2).RowHeight = 30\nws.Rows(2).AutoFit", reads("ws.Rows(2)", ROW)),
    "autofit_rows_many": ('ws.Rows(2).RowHeight = 30\nws.Rows(4).RowHeight = 5\nws.Rows("1:5").AutoFit',
                          [*reads("ws.Rows(2)", ROW), *reads("ws.Rows(4)", ROW)]),
    "autofit_hidden_row": ("ws.Rows(2).RowHeight = 30\nws.Rows(2).Hidden = True\nws.Rows(2).AutoFit",
                           reads("ws.Rows(2)", ROW)),
    "autofit_empty_column": ("ws.Columns(3).ColumnWidth = 30\nws.Columns(3).AutoFit", reads("ws.Columns(3)", COL)),
    "autofit_hidden_column": ("ws.Columns(3).ColumnWidth = 30\nws.Columns(3).Hidden = True\nws.Columns(3).AutoFit",
                              reads("ws.Columns(3)", COL)),
    "autofit_needs_whole": ('ws.Range("F1").Value = "Hello"\nws.Range("F1").AutoFit', ["ws.Columns(6).ColumnWidth"]),
    "autofit_cells_rows": ('ws.Rows(2).RowHeight = 30\nws.Range("A2:B2").EntireRow.AutoFit', reads("ws.Rows(2)", ROW)),
    # --- rows and columns moving -----------------------------------------------------------------
    "insert_row_below_custom": ("ws.Rows(3).RowHeight = 30\nws.Rows(5).Hidden = True\nws.Rows(4).Insert",
                                [*reads("ws.Rows(3)", ROW), *reads("ws.Rows(4)", ROW), *reads("ws.Rows(5)", ROW),
                                 *reads("ws.Rows(6)", ROW)]),
    "insert_row_at_custom": ("ws.Rows(3).RowHeight = 30\nws.Rows(3).Insert",
                             [*reads("ws.Rows(3)", ROW), *reads("ws.Rows(4)", ROW)]),
    "insert_row_below_hidden": ("ws.Rows(3).RowHeight = 30\nws.Rows(3).Hidden = True\nws.Rows(4).Insert",
                                [*reads("ws.Rows(3)", ROW), *reads("ws.Rows(4)", ROW)]),
    "delete_rows": ("ws.Rows(3).RowHeight = 30\nws.Rows(6).Hidden = True\nws.Rows(4).RowHeight = 40\n"
                    'ws.Rows("4:5").Delete',
                    [*reads("ws.Rows(3)", ROW), *reads("ws.Rows(4)", ROW), *reads("ws.Rows(5)", ROW)]),
    "insert_column_right_of_custom": ("ws.Columns(3).ColumnWidth = 30\nws.Columns(5).Hidden = True\n"
                                      "ws.Columns(4).Insert",
                                      [*reads("ws.Columns(3)", COL), *reads("ws.Columns(4)", COL),
                                       *reads("ws.Columns(5)", COL), *reads("ws.Columns(6)", COL)]),
    "insert_column_at_custom": ("ws.Columns(3).ColumnWidth = 30\nws.Columns(3).Insert",
                                [*reads("ws.Columns(3)", COL), *reads("ws.Columns(4)", COL)]),
    "delete_columns": ("ws.Columns(3).ColumnWidth = 30\nws.Columns(6).Hidden = True\nws.Columns(4).ColumnWidth = 40\n"
                       'ws.Columns("D:E").Delete',
                       [*reads("ws.Columns(3)", COL), *reads("ws.Columns(4)", COL), *reads("ws.Columns(5)", COL)]),
    "copy_rows": ('ws.Rows(2).RowHeight = 30\nws.Range("A2").Value = 1\nws.Rows(2).Copy ws.Rows(6)\n'
                  'ws.Range("A2").Copy ws.Range("A8")', [*reads("ws.Rows(6)", ROW), *reads("ws.Rows(8)", ROW)]),
    "copy_columns": ('ws.Columns(2).ColumnWidth = 30\nws.Range("B1").Value = 1\nws.Columns(2).Copy ws.Columns(6)\n'
                     'ws.Range("B1").Copy ws.Range("H1")', [*reads("ws.Columns(6)", COL), *reads("ws.Columns(8)", COL)]),
    "copy_sheet": ("ws.Rows(2).RowHeight = 30\nws.Columns(2).ColumnWidth = 30\nws.Rows(3).Hidden = True\n"
                   "ws.Columns(3).Hidden = True\nws.StandardWidth = 12\nws.Copy After:=ws",
                   ["wb.Worksheets(2).Rows(2).RowHeight", "wb.Worksheets(2).Columns(2).ColumnWidth",
                    "wb.Worksheets(2).Rows(3).Hidden", "wb.Worksheets(2).Columns(3).Hidden",
                    "wb.Worksheets(2).StandardWidth", "wb.Worksheets(2).Columns(4).ColumnWidth"]),
    # --- rows as tall as their fonts ---------------------------------------------------------------
    "font_grows_row": ('ws.Range("A1").Font.Size = 20',
                       [*reads("ws.Rows(1)", ROW), 'ws.Range("A2").Top', 'ws.Rows("1:2").Height',
                        "ws.StandardHeight", "ws.UsedRange.Address"]),
    "font_grows_row_with_value": ('ws.Range("A1").Value = "x"\nws.Range("A1").Font.Size = 20', reads("ws.Rows(1)", ROW)),
    "font_back_to_normal": ('ws.Range("A1").Font.Size = 20\nws.Range("A1").Font.Size = 11', reads("ws.Rows(1)", ROW)),
    "font_smaller_than_normal": ('ws.Range("A1").Font.Size = 8', reads("ws.Rows(1)", ROW)),
    "font_other_face": ('ws.Range("A1").Font.Name = "Segoe UI"\nws.Range("A2").Font.Name = "Tahoma"\n'
                        'ws.Range("A2").Font.Size = 16\nws.Range("A3").Font.Name = "Arial"\n'
                        'ws.Range("A3").Font.Size = 20\nws.Range("A3").Font.Bold = True\n'
                        'ws.Range("A4").Font.Name = "calibri light"\nws.Range("A4").Font.Size = 30',
                        [*reads("ws.Rows(1)", ROW), *reads("ws.Rows(2)", ROW), *reads("ws.Rows(3)", ROW),
                         *reads("ws.Rows(4)", ROW)]),
    "font_same_face_sizes": ('ws.Range("A1").Font.Size = 20\nws.Range("B1").Font.Size = 26\n'
                             'ws.Range("C1").Font.Size = 14\nws.Range("D1").Font.Name = "Aptos"\n'
                             'ws.Range("D1").Font.Size = 24', reads("ws.Rows(1)", ROW)),
    "font_odd_sizes": ('ws.Range("A1").Font.Size = 12.375\nws.Range("A2").Font.Size = 12.4\n'
                       'ws.Range("A3").Font.Size = 10.3\nws.Range("A4").Font.Size = 409',
                       [*reads("ws.Rows(1)", ROW), *reads("ws.Rows(2)", ROW), *reads("ws.Rows(3)", ROW),
                        *reads("ws.Rows(4)", ROW)]),
    "font_custom_row": ('ws.Rows(1).RowHeight = 12\nws.Range("A1").Font.Size = 20', reads("ws.Rows(1)", ROW)),
    "font_autofit": ('ws.Rows(1).RowHeight = 12\nws.Range("A1").Font.Size = 20\nws.Rows(1).AutoFit',
                     reads("ws.Rows(1)", ROW)),
    "font_use_standard": ('ws.Range("A1").Font.Size = 20\nws.Rows(1).UseStandardHeight = True\n'
                          'ws.Range("A2").Font.Size = 20\nws.Rows(2).UseStandardHeight = False\n'
                          'ws.Range("A2").Font.Size = 30', [*reads("ws.Rows(1)", ROW), *reads("ws.Rows(2)", ROW)]),
    "font_clearing": ('ws.Range("A1").Font.Size = 20\nws.Range("A1").ClearFormats\n'
                      'ws.Range("A2").Value = 1\nws.Range("A2").Font.Size = 20\nws.Range("A2").ClearContents\n'
                      'ws.Range("A3").Value = 1\nws.Range("A3").Font.Size = 20\nws.Range("A3").Clear',
                      [*reads("ws.Rows(1)", ROW), *reads("ws.Rows(2)", ROW), *reads("ws.Rows(3)", ROW)]),
    "font_hidden_row": ('ws.Range("A1").Font.Size = 20\nws.Rows(1).Hidden = True\nws.Range("A2").Font.Size = 20\n'
                        "ws.Rows(2).Hidden = True\nws.Rows(2).Hidden = False",
                        [*reads("ws.Rows(1)", ROW), *reads("ws.Rows(2)", ROW)]),
    "font_merged": ('ws.Range("A1:A2").Merge\nws.Range("A1").Font.Size = 20\nws.Range("C4:D4").Merge\n'
                    'ws.Range("C4").Font.Size = 20',
                    [*reads("ws.Rows(1)", ROW), *reads("ws.Rows(2)", ROW), *reads("ws.Rows(4)", ROW)]),
    "font_wrap_and_turn_without_text": ('ws.Range("A1").Font.Size = 20\nws.Range("A1").WrapText = True\n'
                                        'ws.Range("A2").Font.Size = 20\nws.Range("A2").Orientation = 90',
                                        [*reads("ws.Rows(1)", ROW), *reads("ws.Rows(2)", ROW)]),
    "font_on_a_block": ('ws.Range("A1:C3").Font.Size = 16', [*reads("ws.Rows(1)", ROW), *reads("ws.Rows(3)", ROW),
                                                             *reads("ws.Rows(4)", ROW), 'ws.Range("A5").Top']),
    "font_rows_below_move": ('ws.Range("A2").Font.Size = 30\nws.Range("A4").Font.Name = "Georgia"\n'
                             'ws.Range("A4").Font.Size = 40', ['ws.Range("A3").Top', 'ws.Range("A10").Top',
                                                               'ws.Rows("1:5").Height']),
    "font_multi_row_read": ('ws.Range("A1").Font.Size = 20\nws.Range("A3").Value = 1',
                            ['ws.Rows("1:2").RowHeight', 'ws.Rows("2:3").RowHeight', 'ws.Rows("1:2").UseStandardHeight',
                             'ws.Range("A1:A2").RowHeight']),
    "font_copied": ('ws.Range("A1").Font.Size = 20\nws.Range("A1").Copy ws.Range("A5")\n'
                    "ws.Rows(1).Copy ws.Rows(7)", [*reads("ws.Rows(5)", ROW), *reads("ws.Rows(7)", ROW)]),
    "font_insert_below": ('ws.Range("A2").Font.Size = 20\nws.Rows(2).Insert',
                          [*reads("ws.Rows(2)", ROW), *reads("ws.Rows(3)", ROW)]),
    # --- what the model reports unsupported: Excel's answers, recorded ----------------------------
    "unmeasured_mixed_faces": ('ws.Range("A1").Font.Size = 20\nws.Range("B1").Font.Name = "Arial"\n'
                               'ws.Range("B1").Font.Size = 26\nws.Range("A2").Font.Name = "Segoe UI"\n'
                               'ws.Range("B2").Font.Name = "Tahoma"\nws.Range("B2").Font.Size = 16',
                               [*reads("ws.Rows(1)", ROW), *reads("ws.Rows(2)", ROW)]),
    "unmeasured_face": ('ws.Range("A1").Font.Name = "Segoe Print"\nws.Range("A1").Font.Size = 20',
                        reads("ws.Rows(1)", ROW)),
    "unmeasured_superscript": ('ws.Range("A1").Font.Size = 20\nws.Range("A1").Font.Superscript = True\n'
                               'ws.Range("A2").Font.Size = 20\nws.Range("A2").Font.Subscript = True',
                               [*reads("ws.Rows(1)", ROW), *reads("ws.Rows(2)", ROW)]),
    "unmeasured_wrapped_text": ('ws.Range("A1").Value = "x" & vbLf & "y"\nws.Range("A1").WrapText = True',
                                reads("ws.Rows(1)", ROW)),
    "unmeasured_turned_text": ('ws.Range("A1").Value = "turned"\nws.Range("A1").Font.Size = 20\n'
                               'ws.Range("A1").Orientation = 90', reads("ws.Rows(1)", ROW)),
})


def body(setup: str, reads_: list[str]) -> str:
    # The setup runs under Resume Next too, so a property Excel refuses
    # records its error number as S<n> instead of ending the probe.
    lines = ["Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
             "On Error Resume Next", setup, 'If Err.Number <> 0 Then Report = Report & "S" & Err.Number & ";"']
    for expression in reads_:
        lines += ["Err.Clear", "v = Empty", f"v = {expression}",
                  'If Err.Number <> 0 Then Report = Report & "E" & Err.Number & ";" '
                  'Else Report = Report & Show(v) & ";"']
    lines.append("wb.Close False")
    return "\n".join(line for line in lines if line) + "\n"


def module(text: str) -> str:
    return (HELPER + "Public Function Report() As String\n"
            "Dim v As Variant, i As Long, h As Variant, wb As Object, ws As Object\n" + text + "End Function\n")


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        probe = excel.run_vba(module("Report = Workbooks.Add(xlWBATWorksheet).Worksheets(1).StandardHeight\n"
                                     "ActiveWorkbook.Close False\n"), "Report", timeout=60.0)
        assert probe.ok and str(probe.value) == "15", f"needs a 96-DPI display (StandardHeight {probe.value})"
        for name, (setup, reads_) in PROBES.items():
            code = module(body(setup, reads_))
            result = excel.run_vba(code, "Report", timeout=300.0)
            reported = str(result.value) if result.ok else f"ERROR {result.outcome}: {result.message}"
            records.append({"name": name, "code": code, "reported": reported})
            print(name, reported[:160], flush=True)
    (ROOT / "tests/fixtures/dimensions.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
