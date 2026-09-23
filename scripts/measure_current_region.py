"""Range.CurrentRegion, and End over cells that hold only a format, as Excel's object model answers them.

Every layout sits on a sheet of its own in one new workbook: a setup that
fills some cells, and reads of CurrentRegion (or End) from several
starting ranges. Most layouts are built to ask one question -- whether a
diagonal neighbour joins a region, whether a formula returning "" or a
cell with only a format does, what a merged cell or a hidden row does, and
what a whole row, a whole column or several areas answer. Eight more are
seeded random grids read from every other cell, which pins the walk
itself. A read records the value's type and text, and E<number> for an
error, as the other probes do.

    python scripts/measure_current_region.py

writes tests/fixtures/current_region.json, which
tests/test_excel_current_region.py replays.
"""

from __future__ import annotations

import json
import random
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "current_region.json"

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


def region(*starts: str) -> list[str]:
    return [f'ws.Range("{start}").CurrentRegion.Address' for start in starts]


def fill(*references: str, value: str = "1") -> str:
    return "\n".join(f'ws.Range("{reference}").Value = {value}' for reference in references)


#: name -> (setup, reads). Every setup works on ws, the layout's own sheet.
LAYOUTS: dict[str, tuple[str, list[str]]] = {
    "block": (fill("B2:D4"), region("B2", "C3", "D4", "A1", "E5", "F6", "E2", "B2:C3", "A1:E5", "C2:C6")),
    "diagonal_chain": (fill("B2", "C3", "D4"), region("B2", "C3", "D4", "E5", "A1", "B4", "D2")),
    "gap_row": (fill("B2:D2", "B4:D4"), region("B2", "C3", "B4", "A3", "E3")),
    "gap_column": (fill("B2:B4", "D2:D4"), region("B2", "C3", "D2", "C1", "C5")),
    "l_shape": (fill("B2:B5", "C5:E5"), region("B2", "E5", "C3", "D4", "F4")),
    "single": (fill("C3"), region("C3", "B2", "D4", "C2", "E3", "E5")),
    "empty_sheet": ("", region("C3", "A1", "XFD1048576", "B2:C3")),
    "formula_empty_string": ('ws.Range("B2").Value = 1\nws.Range("C2").Formula = "="""""\nws.Range("D2").Value = 2',
                             region("B2", "D2", "C3")),
    "space_value": ('ws.Range("B2").Value = 1\nws.Range("C2").Value = " "\nws.Range("D2").Value = 2', region("B2", "D2")),
    "empty_string_value": ('ws.Range("B2").Value = 1\nws.Range("C2").Value = ""\nws.Range("D2").Value = 2',
                           region("B2", "D2")),
    "apostrophe": ('ws.Range("B2").Value = 1\nws.Range("C2").Value = "\'"\nws.Range("D2").Value = 2', region("B2", "D2")),
    "error_value": ('ws.Range("B2").Value = 1\nws.Range("C2").Formula = "=1/0"\nws.Range("D2").Value = 2',
                    region("B2", "D2")),
    "cleared": ('ws.Range("B2:D2").Value = 1\nws.Range("C2").ClearContents', region("B2", "D2")),
    "formatted_gap": ('ws.Range("B2").Value = 1\nws.Range("C2").Font.Bold = True\nws.Range("D2").Value = 2',
                      region("B2", "C2", "D2")),
    "crossing_formats": ('ws.Range("B5").Value = 1\nws.Range("D5").Value = 2\nws.Rows(5).Font.Bold = True\n'
                         'ws.Columns("C").Font.Italic = True', region("B5", "D5", "C5")),
    "merged": ('ws.Range("B2:C3").Merge\nws.Range("B2").Value = "m"\nws.Range("D4").Value = 1',
               region("B2", "C3", "D4", "E5", "A1")),
    "merged_far": ('ws.Range("B2:C3").Merge\nws.Range("B2").Value = "m"\nws.Range("E3").Value = 1',
                   region("B2", "E3", "D3")),
    "hidden_row": ('ws.Range("B2:B4").Value = 1\nws.Rows(3).Hidden = True', region("B2", "B4")),
    "corner_start": (fill("A1", "A2", "B1"), region("A1", "B2", "C3")),
    "corner_end": (fill("XFD1048576", "XFC1048575"), region("XFD1048576", "XFC1048576", "XFB1048574")),
    "whole_lines": (fill("B2:D4", "F7"), ["ws.Rows(3).CurrentRegion.Address", 'ws.Columns("C").CurrentRegion.Address',
                                          "ws.Cells.CurrentRegion.Address", 'ws.Rows("5:6").CurrentRegion.Address']),
    "several_areas": (fill("B2:D4", "F7"), region("B2,F7", "F7,B2", "C3,H9")),
    "text_and_numbers": ('ws.Range("B2").Value = "a"\nws.Range("C2").Value = 2\nws.Range("D3").Value = True\n'
                         'ws.Range("E4").Formula = "=1+1"', region("B2", "E4", "F5")),
    # End looks for values too: a cell with only a format is not one.
    "end_formatted": ('ws.Range("B1:B3").Value = 1\nws.Range("B5").Font.Bold = True\nws.Range("B7").Value = 1',
                      ['ws.Range("B1").End(xlDown).Address', 'ws.Range("B3").End(xlDown).Address',
                       'ws.Range("B7").End(xlUp).Address', 'ws.Range("B5").End(xlDown).Address',
                       'ws.Range("B5").End(xlUp).Address']),
    "end_crossing_formats": ('ws.Range("B1:B3").Value = 1\nws.Range("B8").Value = 1\nws.Rows(5).Font.Bold = True\n'
                             'ws.Columns("B").Font.Italic = True',
                             ['ws.Range("B3").End(xlDown).Address', 'ws.Range("B8").End(xlUp).Address',
                              'ws.Range("A5").End(xlToRight).Address']),
    "end_formula_empty": ('ws.Range("B1").Value = 1\nws.Range("B2").Formula = "="""""\nws.Range("B3").Value = 1',
                          ['ws.Range("B1").End(xlDown).Address', 'ws.Range("B3").End(xlUp).Address']),
}

#: Seeded random grids: which cells of A1:L12 hold a value.
GRID_ROWS, GRID_COLUMNS, GRID_SEEDS, GRID_FILL = 12, 12, range(8), 0.3


def column_name(column: int) -> str:
    return chr(ord("A") + column - 1)


def grid(seed: int) -> tuple[str, list[str]]:
    chooser = random.Random(seed)
    filled = [f"{column_name(column)}{row}" for row in range(1, GRID_ROWS + 1)
              for column in range(1, GRID_COLUMNS + 1) if chooser.random() < GRID_FILL]
    starts = [f"{column_name(column)}{row}" for row in range(1, GRID_ROWS + 1)
              for column in range(1, GRID_COLUMNS + 1) if (row + column) % 2 == 0]
    return fill(*filled), region(*starts)


def layouts() -> dict[str, tuple[str, list[str]]]:
    return {**LAYOUTS, **{f"grid_{seed}": grid(seed) for seed in GRID_SEEDS}}


def reads_function(name: str, reads: list[str]) -> list[str]:
    lines = [f"Private Function {name}(ws As Object) As String", "Dim out As String, v As Variant",
             "On Error Resume Next"]
    for expression in reads:
        lines += ["Err.Clear", "v = Empty", f"v = {expression}",
                  'If Err.Number <> 0 Then out = out & "E" & Err.Number & ";" Else out = out & Show(v) & ";"']
    return [*lines, "On Error GoTo 0", f"{name} = out", "End Function"]


def case_code(index: int, setup: str, reads: list[str]) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String", "Dim failed As String", "On Error Resume Next",
             "Err.Clear", *setup.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"',
             "On Error GoTo 0", f"Case{index} = failed & Reads{index}(ws)", "End Function",
             *reads_function(f"Reads{index}", reads)]
    return "\n".join(lines) + "\n"


def main() -> None:
    cases = layouts()
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, (setup, reads)) in enumerate(cases.items()):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{name}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, setup, reads))
    build += ["wb.Close False", "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = str(result.value).split("|")[: len(cases)]
    record = {"helper": HELPER,
              "layouts": [{"name": name, "setup": setup, "reads": reads, "answers": answer}
                          for (name, (setup, reads)), answer in zip(cases.items(), answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, answer in zip(cases, answers, strict=True):
        print(f"{name}: {answer[:300]}")


if __name__ == "__main__":
    main()
