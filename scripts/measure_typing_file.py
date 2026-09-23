"""What Excel writes to its file for cells a macro typed into.

Typing gives a cell a number format, and an apostrophe a prefix
character, and both live in the cell's xf once the workbook is saved:
which built-in format id, which apply flags, whether quotePrefix is set.
Each case writes into its own cell, one per row of column A, through the
same steps the object-model probes use -- a value, or a statement run as
it stands -- and Excel saves the workbook.

    python scripts/measure_typing_file.py

writes tests/fixtures/typing/typing.xlsx as Excel saved it and
tests/fixtures/typing/typing.json with the cases, which
tests/test_excel_value_typing.py replays.
"""

from __future__ import annotations

import datetime
import json
import shutil
import tempfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "typing"

#: Each case's steps, in the cell the case has to itself.
CASES = [
    *[[text] for text in ('"5"', '"1,000"', '"1,000.5"', '"5%"', '"5.5%"', '"$5"', '"$5.5"', '"1e3"', '"1 1/2"',
                          '"1 3/16"', '"1/2/2020"', '"1/2"', '"Jan 2020"', '"2-Jan-2020"', '"12:30"', '"12:30:45"',
                          '"12:30 PM"', '"12:30:45 AM"', '"25:00"', '"12:30:45.5"', '"1/2/2020 12:30"', '"1:60"',
                          '"1e400"', "#1/2/2020#", "#12:30:00 PM#", "#1/2/2020 12:30:00 PM#", "CCur(5.5)", "5",
                          "True", '"TRUE"', '"#N/A"', '"abc"', '""', '"\'5"', '"\'"')],
    ['"\'5"', "5"],
    ['"\'5"', "c.ClearContents"],
    ['"\'5"', '"abc"'],
    ['c.NumberFormat = "@"', '"5"'],
    ['c.NumberFormat = "@"', "CCur(5.5)"],
    ['c.NumberFormat = "@"', "#1/2/2020#"],
    ['c.NumberFormat = "@"', '""'],
    ['c.NumberFormat = "0.00"', '"5%"'],
    ['c.NumberFormat = "0%"', '"$5"'],
    ['c.NumberFormat = "m/d/yyyy"', '"12:30"'],
    ["c.Value2 = #1/2/2020#"],
    ['c.Formula = "5%"'],
    ["c.Font.Bold = True", '"5%"'],
    ['"\'5"', 'c.NumberFormat = "0.00"'],
]


def statement(step: str) -> str:
    return step if step.startswith("c.") else f"c.Value = {step}"


def main() -> None:
    saved = Path(tempfile.mkdtemp()) / "typing.xlsx"
    lines = ["Public Function Build() As String", "Dim wb As Object, ws As Object, c As Object",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)"]
    for row, steps in enumerate(CASES, start=1):
        lines += [f"Set c = ws.Cells({row}, 1)", *(statement(step) for step in steps)]
    lines += [f'wb.SaveAs Filename:="{saved}", FileFormat:=51', "wb.Close False", 'Build = "ok"', "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(lines) + "\n", "Build", timeout=300.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    OUT.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(saved, OUT / "typing.xlsx")
    # A date typed without a year falls in the year it was typed.
    record = {"year": datetime.date.today().year, "cases": CASES}
    (OUT / "typing.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print(f"saved {len(CASES)} cases")


if __name__ == "__main__":
    main()
