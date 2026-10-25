"""How Excel's Range.Sort orders text beyond ASCII: letters with accents, Greek, Cyrillic, combining marks, dashes,
punctuation and symbols, case ignored and matched.

The strings are drawn from a fixed seed, with the variants of a word that differ only in accents, case, hyphens and
apostrophes among them. A macro writes each into a cell of column A as text, built from ChrW since a module's text
cannot hold them, its place beside it in column B, sorts the two columns by A and reads back the places in their new
order, once ignoring case and once matching it.

    python scripts/measure_sort_text.py

writes tests/fixtures/sort_text.json, which tests/test_excel_sort_text.py replays.
"""

from __future__ import annotations

import json
import random
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "sort_text.json"
#: Latin, its extensions, combining marks, Greek, Cyrillic, general punctuation, currency and letterlike symbols.
RANGES = [(0x20, 0x7E), (0xA0, 0x24F), (0x300, 0x36F), (0x370, 0x3FF), (0x400, 0x52F), (0x1E00, 0x1EFF),
          (0x1F00, 0x1FFF), (0x2000, 0x206F), (0x20A0, 0x20CF), (0x2100, 0x214F)]
#: Spellings of one word that differ in accents, case, a hyphen or an apostrophe alone.
VARIANTS = ["resume", "résumé", "Resume", "RÉSUMÉ", "re-sume", "re'sume", "resumé",
            "résume", "Re-sume", "é", "é", "E", "e", "Æsop", "aesop", "Aesop", "straße",
            "strasse", "Strasse", "абв", "АБВ", "αβγ", "ΑΒΓ",
            "co‐op", "co-op", "coop", "co—op", "naïve", "naive", "Naïve"]


def strings() -> list[str]:
    """None begins with a combining mark, which Windows' CompareString orders as no sort key does."""
    import unicodedata

    pick = random.Random(20261025)
    alphabet = [chr(code) for low, high in RANGES for code in range(low, high + 1)
                if code not in (0x385, 0x1FEE) and chr(code).isprintable()]
    marks = {character for character in alphabet if unicodedata.category(character) in ("Mn", "Me")}
    found = list(VARIANTS)
    while len(found) < 400:
        text = "".join(pick.choice(alphabet) for _ in range(pick.randint(1, 6)))
        if text[0] not in marks and text.strip() == text:
            found.append(text)
    return found


def vba_text(text: str) -> str:
    return " & ".join(f"ChrW({ord(character)})" for character in text)


def code(texts: list[str]) -> str:
    lines = ["Private Sub Fill(ws As Object)"]
    for row, text in enumerate(texts, start=1):
        lines += [f"    ws.Cells({row}, 1).Value = {vba_text(text)}", f"    ws.Cells({row}, 2).Value = {row}"]
    lines += ["End Sub", "",
              "Private Function Order(ws As Object, matching As Boolean) As String",
              "    Dim r As Long, out As String",
              f'    ws.Range("A1:B{len(texts)}").Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlNo, '
              "MatchCase:=matching, Orientation:=xlTopToBottom",
              f"    For r = 1 To {len(texts)}", '        out = out & ws.Cells(r, 2).Value & ","', "    Next",
              "    Order = out", "End Function", "",
              "Public Function Measure() As String",
              "    Dim wb As Object, ws As Object, out As String",
              "    Set wb = Workbooks.Add(xlWBATWorksheet)", "    Set ws = wb.Worksheets(1)",
              '    ws.Columns("A").NumberFormat = "@"', "    Fill ws", '    out = Order(ws, False) & "|"',
              "    ws.Cells.Clear", '    ws.Columns("A").NumberFormat = "@"', "    Fill ws", "    out = out & Order(ws, True)",
              "    wb.Close False", "    Measure = out", "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    texts = strings()
    source = code(texts)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, "Measure", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    ignoring, matching = (
        [int(place) for place in part.split(",") if place] for part in str(result.value).split("|"))
    OUT.write_text(json.dumps({"strings": texts, "code": source, "ignoring_case": ignoring,
                               "matching_case": matching}, indent=0, ensure_ascii=True) + "\n", encoding="utf-8")
    print(len(texts), "strings;", "ignoring case:", ignoring[:20], "matching case:", matching[:20])


if __name__ == "__main__":
    main()
