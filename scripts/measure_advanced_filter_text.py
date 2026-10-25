"""Which text AdvancedFilter keeps beyond ASCII: accents, ligatures, a sharp s, Greek, Cyrillic, case and wildcards.

A macro writes each of VALUES as text under a header in column A of a new workbook, built from ChrW since a module's
text cannot hold them, then filters the list in place once by each of CRITERIA, which stands as text under the same
header in the criteria range G1:G2, reading after each which of the values' rows are visible: a 1 for a row shown,
a 0 for one hidden, or E and the error the filter raised. A plain criterion is a beginning, one after = the whole
text. A second macro does the same for a value of each of LETTERS' letters between an x and a y, filtered by =x,
each letter and a star.

    python scripts/measure_advanced_filter_text.py

writes tests/fixtures/advanced_filter_text.json, which tests/test_excel_advanced_filter_text.py replays.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

sys.path.insert(0, str(Path(__file__).resolve().parent))
from measure_filter_text import CRITERIA, LETTERS, VALUES, vba_text  # noqa: E402

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "advanced_filter_text.json"
#: Plain criteria, each the beginning of some of the values in one spelling or another.
BEGINNINGS = ["ae", "æ", "aes", "stras", "straß", "strass", "ré", "re", "co-", "co", "kel",
              "Kel", "é", "œ", "oe", "naï", "nai", "ǆ", "dz", "res"]


def code(values: list[str], criteria: list[str], macro: str) -> str:
    last = len(values) + 1
    lines = ["Private Function Shown(ws As Object) As String",
             "    Dim r As Long, out As String",
             f"    For r = 2 To {last}",
             '        If ws.Rows(r).Hidden Then out = out & "0" Else out = out & "1"',
             "    Next",
             "    Shown = out",
             "End Function",
             "",
             f"Public Function {macro}() As String",
             "    Dim wb As Object, ws As Object, out As String",
             "    Set wb = Workbooks.Add(xlWBATWorksheet)",
             "    Set ws = wb.Worksheets(1)",
             '    ws.Columns("A").NumberFormat = "@"',
             '    ws.Columns("G").NumberFormat = "@"',
             '    ws.Cells(1, 1).Value = "Word"',
             '    ws.Cells(1, 7).Value = "Word"']
    lines += [f"    ws.Cells({row}, 1).Value = {vba_text(text)}" for row, text in enumerate(values, start=2)]
    lines += ["    On Error Resume Next"]
    for criterion in criteria:
        lines += ["    Err.Clear", f"    ws.Cells(2, 7).Value = {vba_text(criterion)}",
                  f'    ws.Range("A1:A{last}").AdvancedFilter xlFilterInPlace, ws.Range("G1:G2")',
                  '    If Err.Number <> 0 Then out = out & "E" & Err.Number & "," Else out = out & Shown(ws) & ","',
                  "    If ws.FilterMode Then ws.ShowAllData"]
    lines += ["    On Error GoTo 0", "    wb.Close False", f"    {macro} = out", "End Function"]
    return "\n".join(lines) + "\n"


def letters() -> tuple[list[str], list[str]]:
    values = [f"x{letter}y" for pair in LETTERS for letter in pair]
    return values, [f"=x{letter}*" for pair in LETTERS for letter in pair]


def main() -> None:
    source = code(VALUES, CRITERIA, "Measure")
    letter_values, letter_criteria = letters()
    lettered_source = code(letter_values, letter_criteria, "Letters")
    begun_source = code(VALUES, BEGINNINGS, "Beginnings")
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, "Measure", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
        lettered = excel.run_vba(lettered_source, "Letters", timeout=600.0)
        assert lettered.ok, f"{lettered.outcome}: {lettered.message} {lettered.error}"
        begun = excel.run_vba(begun_source, "Beginnings", timeout=600.0)
        assert begun.ok, f"{begun.outcome}: {begun.message} {begun.error}"
    record = {"values": VALUES, "criteria": CRITERIA, "code": source, "raw": str(result.value),
              "letters": LETTERS, "letters_code": lettered_source, "letters_raw": str(lettered.value),
              "beginnings": BEGINNINGS, "beginnings_code": begun_source, "beginnings_raw": str(begun.value)}
    OUT.write_text(json.dumps(record, indent=0, ensure_ascii=True) + "\n", encoding="utf-8")
    for criteria, raw in ((CRITERIA, str(result.value)), (BEGINNINGS, str(begun.value))):
        for criterion, shown in zip(criteria, raw.split(",")[:-1], strict=True):
            kept = [VALUES[at] for at, bit in enumerate(shown) if bit == "1"] if not shown.startswith("E") else shown
            print(ascii(criterion), ascii(kept))
    print("letters", lettered.value)


if __name__ == "__main__":
    main()
