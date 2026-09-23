"""Range.SpecialCells, as Excel's object model answers it.

Every layout sits on a sheet of its own in one new workbook: a setup that
fills some cells, then SpecialCells asked of several ranges for constants
and formulas of each kind, blanks, the last cell and visible cells. Each
answer records the address of what came back and how many areas it has,
or E<number> for an error, since how Excel splits the cells it finds
into rectangular areas is half of what the method answers. Most layouts
ask one question -- a single cell searching the whole used range, a
format-only cell, hidden rows and columns, merged cells, a range outside
the used range -- and eight seeded random grids pin the splitting.

    python scripts/measure_special_cells.py

writes tests/fixtures/special_cells.json, which
tests/test_excel_special_cells.py replays.
"""

from __future__ import annotations

import json
import random
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "special_cells.json"

HELPER = '''Private Function Found(r As Object, kind As Long, Optional which As Variant) As String
    Dim got As Object
    On Error Resume Next
    Err.Clear
    If IsMissing(which) Then
        Set got = r.SpecialCells(kind)
    Else
        Set got = r.SpecialCells(kind, which)
    End If
    If Err.Number <> 0 Then
        Found = "E" & Err.Number
    Else
        Found = got.Address(False, False) & "#" & got.Areas.Count
    End If
    On Error GoTo 0
End Function
'''

#: Every kind a layout is asked for: (label, SpecialCells type, value flags or None).
KINDS = [("constants", 2, None), ("numbers", 2, 1), ("text", 2, 2), ("logical", 2, 4), ("errors", 2, 16),
         ("numbers_text", 2, 3), ("formulas", -4123, None), ("formula_numbers", -4123, 1),
         ("formula_text", -4123, 2), ("formula_logical", -4123, 4), ("formula_errors", -4123, 16),
         ("blanks", 4, None), ("last", 11, None), ("visible", 12, None)]


def reads(*ranges: str) -> list[str]:
    out: list[str] = []
    for reference in ranges:
        target = "ws.Cells" if reference == "*" else f'ws.Range("{reference}")'
        for _, kind, which in KINDS:
            out.append(f"Found({target}, {kind}{'' if which is None else f', {which}'})")
    return out


MIXED = ('ws.Range("A1").Value = 1\nws.Range("A2").Value = "x"\nws.Range("B1").Value = True\n'
         'ws.Range("B3").Value = "#N/A"\nws.Range("C2").Formula = "=1+1"\nws.Range("C3").Formula = "=""a"""\n'
         'ws.Range("D1").Formula = "=1/0"\nws.Range("D2").Formula = "=TRUE"\nws.Range("D4").Font.Bold = True\n'
         'ws.Range("E5").Value = 5')

#: name -> (setup, reads). Every setup works on ws, the layout's own sheet.
LAYOUTS: dict[str, tuple[str, list[str]]] = {
    "mixed": (MIXED, reads("*", "A1", "A1:B2", "B2:E5", "G10:H12", "C:C", "2:2")),
    "empty_sheet": ("", reads("*", "A1")),
    "format_beyond": ('ws.Range("B2").Value = 1\nws.Range("F8").Interior.ColorIndex = 6', reads("*", "B2")),
    "hidden": ('ws.Range("A1:D4").Value = 1\nws.Rows(2).Hidden = True\nws.Columns("C").Hidden = True',
               reads("*", "A1:D4", "B2")),
    "merged": ('ws.Range("B2:C3").Merge\nws.Range("B2").Value = 1\nws.Range("E5").Value = 2', reads("*")),
    "only_formulas": ('ws.Range("A1:B2").Formula = "=1"', reads("*")),
    "rows_of_runs": ('ws.Range("A1:C1").Value = 1\nws.Range("A2:B2").Value = 1\nws.Range("A3:C3").Value = 1\n'
                     'ws.Range("E1:E3").Value = 1', reads("*")),
    "columns_of_runs": ('ws.Range("A1:A3").Value = 1\nws.Range("B1:B2").Value = 1\nws.Range("C1:C3").Value = 1',
                        reads("*")),
    "checker": ('ws.Range("A1,C1,B2,A3,C3").Value = 1', reads("*")),
    "empty_strings": ('ws.Range("A1").Value = "\'"\nws.Range("B1").Formula = "="""""\nws.Range("C1").Value = " "\n'
                      'ws.Range("D2").Value = 1', reads("*")),
}

GRID_ROWS, GRID_COLUMNS, GRID_SEEDS = 10, 8, range(8)


def column_name(column: int) -> str:
    return chr(ord("A") + column - 1)


def grid(seed: int) -> tuple[str, list[str]]:
    """Numbers, text and formulas scattered over A1:H10, with the rest blank."""
    chooser = random.Random(seed)
    lines: list[str] = []
    for row in range(1, GRID_ROWS + 1):
        for column in range(1, GRID_COLUMNS + 1):
            pick = chooser.random()
            reference = f"{column_name(column)}{row}"
            if pick < 0.25:
                lines.append(f'ws.Range("{reference}").Value = {chooser.randint(1, 9)}')
            elif pick < 0.35:
                lines.append(f'ws.Range("{reference}").Value = "t"')
            elif pick < 0.45:
                lines.append(f'ws.Range("{reference}").Formula = "=1+{chooser.randint(1, 9)}"')
    return "\n".join(lines), reads("*", "B2:G9")


def layouts() -> dict[str, tuple[str, list[str]]]:
    return {**LAYOUTS, **{f"grid_{seed}": grid(seed) for seed in GRID_SEEDS}}


def case_code(index: int, setup: str, asked: list[str]) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String", "Dim failed As String, out As String",
             "On Error Resume Next", "Err.Clear", *setup.splitlines(),
             'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0"]
    lines += [f'out = out & {expression} & ";"' for expression in asked]
    return "\n".join([*lines, f"Case{index} = failed & out", "End Function"]) + "\n"


def main() -> None:
    cases = layouts()
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, (setup, asked)) in enumerate(cases.items()):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{name}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, setup, asked))
    build += ["wb.Close False", "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = str(result.value).split("|")[: len(cases)]
    record = {"helper": HELPER, "kinds": [label for label, _, _ in KINDS],
              "layouts": [{"name": name, "setup": setup, "reads": asked, "answers": answer}
                          for (name, (setup, asked)), answer in zip(cases.items(), answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, answer in zip(cases, answers, strict=True):
        print(f"{name}: {answer[:400]}")


if __name__ == "__main__":
    main()
