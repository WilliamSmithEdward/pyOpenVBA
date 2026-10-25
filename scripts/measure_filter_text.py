"""Which text an AutoFilter keeps beyond ASCII: accents, ligatures, a sharp s, Greek, Cyrillic, case and wildcards.

A macro writes each of VALUES as text under a header in column A of a new workbook, built from ChrW since a module's
text cannot hold them, then filters the column once by each of CRITERIA and once by each of LISTS as
xlFilterValues, reading after each which of the values' rows are visible: a 1 for a row shown, a 0 for one hidden,
or E and the error the filter raised. A second macro does the same for a value of each of LETTERS' letters between
an x and a y, filtered by each letter after an x and before a star; then filters SAVED_LIST by all its values and
saves the workbook, whose sheet lists the values in the order the record keeps.

    python scripts/measure_filter_text.py

writes tests/fixtures/filter_text.json, which tests/test_excel_filter_text.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "filter_text.json"
#: The column's text: spellings of a word that differ in case, accents, ligatures, a sharp s and hyphens.
VALUES = ["aesop", "\u00c6sop", "AESOP", "strasse", "stra\u00dfe", "STRASSE", "resume", "r\u00e9sum\u00e9",
          "R\u00c9SUM\u00c9", "re\u0301sume\u0301", "oeuvre", "\u0153uvre", "kelvin", "\u212aelvin", "coop", "co-op",
          "co\u2010op", "naive", "na\u00efve", "istanbul", "\u0130stanbul", "\u01c5emal", "\u01c6emal",
          "\u03c3\u03bf\u03c6\u03af\u03b1", "\u03a3\u039f\u03a6\u0399\u0391", "\u0440\u0435\u0437\u044e\u043c\u0435",
          "\u0420\u0415\u0417\u042e\u041c\u0415", "caf\u00e9", "CAF\u00c9", "cafe", "m", "n", "z", "\u00e9", "e", "f"]
#: Criteria the column is filtered by, one at a time.
CRITERIA = ["aesop", "=aesop", "=\u00c6sop", "=AESOP", "=strasse", "=stra\u00dfe", "=resume", "=r\u00e9sum\u00e9",
            "=RESUME", "=re\u0301sume\u0301", "=oeuvre", "=\u0153uvre", "=kelvin", "=\u212aelvin", "=coop", "=co-op",
            "=naive", "=na\u00efve", "=istanbul", "=\u0130stanbul", "=\u01c6emal", "=\u03c3\u03bf\u03c6\u03af\u03b1",
            "=\u03a3\u039f\u03a6\u0399\u0391", "=\u0440\u0435\u0437\u044e\u043c\u0435", "<>aesop", "<>caf\u00e9",
            "*sop", "\u00e6*", "ae*", "*\u00e9*", "*e*", "r?sum?", "caf?", "*ss*", "*\u00df*", "co?op", "co*op",
            "\u212a*", "k*", ">m", ">=e", "<f", ">\u00e9", "<\u00e9", ">=caf\u00e9", ">resume", "<r\u00e9sum\u00e9",
            ">=oeuvre", "<\u0153uvre", ">\u03b1", "<\u0430"]
#: Lists of values the column is filtered by, each as xlFilterValues.
LISTS = [["aesop"], ["\u00c6sop"], ["caf\u00e9", "resume"], ["strasse"], ["\u212aelvin"], ["co-op"]]
#: Pairs of letters that one case rule or another takes for one, each between an x and a y in a value of its own,
#: and each looked for by a wildcard pattern x, the letter and a star.
LETTERS = [("I", "\u0131"), ("S", "\u017f"), ("\u03bc", "\u00b5"), ("\u03b8", "\u03d1"), ("\u03c3", "\u03c2"),
           ("\u00df", "\u1e9e"), ("\u01c4", "\u01c6"), ("\u00c5", "\u212b"), ("\u03a9", "\u2126"), ("i", "\u0130"),
           ("K", "\u212a"), ("e", "\u00e9"), ("\u00e6", "a")]
#: A list with every value of it in Excel's word sort order but for case, saved to see the order the file lists it in.
SAVED_LIST = ["\u00e9", "z", "\u00c6sop", "e", "\u0420\u0415\u0417\u042e\u041c\u0415", "strasse", "\u03b1"]


def vba_text(text: str) -> str:
    return " & ".join(f"ChrW({ord(character)})" for character in text) if text else '""'


def code() -> str:
    last = len(VALUES) + 1
    lines = ["Private Function Shown(ws As Object) As String",
             "    Dim r As Long, out As String",
             f"    For r = 2 To {last}",
             '        If ws.Rows(r).Hidden Then out = out & "0" Else out = out & "1"',
             "    Next",
             "    Shown = out",
             "End Function",
             "",
             "Public Function Measure() As String",
             "    Dim wb As Object, ws As Object, out As String",
             "    Set wb = Workbooks.Add(xlWBATWorksheet)",
             "    Set ws = wb.Worksheets(1)",
             '    ws.Columns("A").NumberFormat = "@"',
             '    ws.Cells(1, 1).Value = "Word"']
    lines += [f"    ws.Cells({row}, 1).Value = {vba_text(text)}" for row, text in enumerate(VALUES, start=2)]
    lines += ["    On Error Resume Next"]
    for criterion in CRITERIA:
        lines += ["    Err.Clear", f'    ws.Range("A1:A{last}").AutoFilter Field:=1, Criteria1:={vba_text(criterion)}',
                  '    If Err.Number <> 0 Then out = out & "E" & Err.Number & "," Else out = out & Shown(ws) & ","']
    lines += ['    out = out & "|"']
    for items in LISTS:
        listed = ", ".join(vba_text(item) for item in items)
        lines += ["    Err.Clear",
                  f'    ws.Range("A1:A{last}").AutoFilter Field:=1, Criteria1:=Array({listed}), '
                  "Operator:=xlFilterValues",
                  '    If Err.Number <> 0 Then out = out & "E" & Err.Number & "," Else out = out & Shown(ws) & ","']
    lines += ["    On Error GoTo 0", "    wb.Close False", "    Measure = out", "End Function"]
    return "\n".join(lines) + "\n"


def letters_code() -> str:
    """Each pair's two values, x, the letter and y, in column A under a header, filtered by x, each letter and a
    star in turn; then the saved list's values, filtered by all of them, and the workbook saved to ``path``."""
    values = [f"x{letter}y" for pair in LETTERS for letter in pair]
    last = len(values) + 1
    lines = ["Private Function Shown(ws As Object, last As Long) As String",
             "    Dim r As Long, out As String",
             "    For r = 2 To last",
             '        If ws.Rows(r).Hidden Then out = out & "0" Else out = out & "1"',
             "    Next",
             "    Shown = out",
             "End Function",
             "",
             "Public Function Letters(path As String) As String",
             "    Dim wb As Object, ws As Object, out As String",
             "    Set wb = Workbooks.Add(xlWBATWorksheet)",
             "    Set ws = wb.Worksheets(1)",
             '    ws.Columns("A").NumberFormat = "@"',
             '    ws.Cells(1, 1).Value = "Word"']
    lines += [f"    ws.Cells({row}, 1).Value = {vba_text(text)}" for row, text in enumerate(values, start=2)]
    lines += ["    On Error Resume Next"]
    for pair in LETTERS:
        for letter in pair:
            lines += ["    Err.Clear",
                      f'    ws.Range("A1:A{last}").AutoFilter Field:=1, Criteria1:={vba_text("x" + letter + "*")}',
                      '    If Err.Number <> 0 Then out = out & "E" & Err.Number & "," Else out = out & '
                      f'Shown(ws, {last}) & ","']
    lines += ['    out = out & "|"', "    On Error GoTo 0", "    ws.AutoFilterMode = False", "    ws.Cells.Clear",
              '    ws.Cells(1, 1).Value = "Word"']
    lines += [f"    ws.Cells({row}, 1).Value = {vba_text(text)}" for row, text in enumerate(SAVED_LIST, start=2)]
    listed = ", ".join(vba_text(item) for item in SAVED_LIST)
    lines += [f'    ws.Range("A1:A{len(SAVED_LIST) + 1}").AutoFilter Field:=1, Criteria1:=Array({listed}), '
              "Operator:=xlFilterValues",
              f'    out = out & Shown(ws, {len(SAVED_LIST) + 1})',
              "    Application.DisplayAlerts = False",
              "    wb.SaveAs path, 51",
              "    wb.Close False",
              "    Letters = out",
              "End Function"]
    return "\n".join(lines) + "\n"


def saved_values(path: Path) -> list[str]:
    """The filter values the saved workbook's sheet lists, in its order."""
    import html
    import re
    import zipfile

    with zipfile.ZipFile(path) as package:
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    return [html.unescape(found) for found in re.findall(r'<filter val="([^"]*)"', sheet)]


def main() -> None:
    import tempfile

    source, letters = code(), letters_code()
    with tempfile.TemporaryDirectory() as folder:
        saved = Path(folder) / "filter_text.xlsx"
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(source, "Measure", timeout=600.0)
            assert result.ok, f"{result.outcome}: {result.message} {result.error}"
            lettered = excel.run_vba(letters, "Letters", [str(saved)], timeout=600.0)
            assert lettered.ok, f"{lettered.outcome}: {lettered.message} {lettered.error}"
        written = saved_values(saved)
    criteria, lists = (part.split(",")[:-1] for part in str(result.value).split("|"))
    by_letter, saved_shown = str(lettered.value).split("|")
    record = {"values": VALUES, "criteria": CRITERIA, "lists": LISTS, "code": source, "raw": str(result.value),
              "by_criterion": dict(zip(CRITERIA, criteria, strict=True)),
              "by_list": [[items, shown] for items, shown in zip(LISTS, lists, strict=True)],
              "letters": LETTERS, "letters_code": letters, "letters_raw": str(lettered.value),
              "by_letter": by_letter.split(",")[:-1], "saved_list": SAVED_LIST, "saved_kept": saved_shown,
              "saved_values": written}
    for (first, second), (one, other) in zip(LETTERS, zip(*[iter(record["by_letter"])] * 2), strict=True):
        print(ascii(first), ascii(second), "first finds", one, "second finds", other)
    print("saved", ascii(written), saved_shown)
    OUT.write_text(json.dumps(record, indent=0, ensure_ascii=True) + "\n", encoding="utf-8")
    for criterion, shown in zip(CRITERIA, criteria, strict=True):
        kept = [VALUES[at] for at, bit in enumerate(shown) if bit == "1"] if not shown.startswith("E") else shown
        print(ascii(criterion), ascii(kept))
    for items, shown in zip(LISTS, lists, strict=True):
        kept = [VALUES[at] for at, bit in enumerate(shown) if bit == "1"] if not shown.startswith("E") else shown
        print("list", ascii(items), ascii(kept))


if __name__ == "__main__":
    main()
