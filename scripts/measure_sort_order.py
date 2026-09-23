"""The order Excel's Sort puts values in, as live Excel answers it.

One column of values, each beside its place in the list, is sorted
ascending and descending, ignoring case and matching it, and the places
are read back in their new order. The values are every printable ASCII
character alone, words that differ only by a hyphen, an apostrophe, a
space or a letter's case, accented letters, digits as text, text with
spaces around it, the empty string, 400 seeded random strings of letters,
digits, spaces, hyphens and apostrophes, numbers, Booleans, errors and
blank cells. Every text is written behind an apostrophe, so that nothing
in it is typed as a number or a formula.

    python scripts/measure_sort_order.py

writes tests/fixtures/sort_order.json, which tests/test_excel_sort.py replays.
"""

from __future__ import annotations

import json
import random
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "sort_order.json"

#: Text, as VBA expressions for the string itself.
TEXTS = [f'ChrW({code})' for code in range(32, 127)] + [
    '"co-op"', '"coop"', '"co op"', '"cop"', '"co"', '"co-"', '"-co"', '"\'co"', '"c\'o"', '"can\'t"', '"cant"',
    '"cans"', '"ca"', '"a-b"', '"ab"', '"a b"', '"a_b"', '"a.b"', '"a"', '"A"', '"aa"', '"Aa"', '"aA"', '"AA"',
    '"b"', '"B"', '"ba"', '"Ba"', '"1"', '"10"', '"2"', '"01"', '"1.5"', '"-1"', '"1a"', '"a1"', '"a10"',
    '"a2"', '" a"', '"a "', '"  a"', '""', 'ChrW(233)', 'ChrW(232)', 'ChrW(234)', 'ChrW(201)', '"e"', '"E"',
    '"f"', 'ChrW(252)', '"u"', '"v"', 'ChrW(223)', '"ss"', '"st"', 'ChrW(230)', '"ae"', '"af"', '"x-y-z"',
    '"xyz"', '"x--y"', '"--"', '"-"', "\"'\"", '"\'\'"', '"a-"', '"a--"', '"-a"', '"--a"',
]
#: Seeded random text over letters of both cases, digits, spaces, hyphens, apostrophes and a few marks.
_ALPHABET = "aAbBzZ09 -'-'.!_"


def _random_texts() -> list[str]:
    chooser = random.Random(20260923)
    out: list[str] = []
    for _ in range(400):
        text = "".join(chooser.choice(_ALPHABET) for _ in range(chooser.randint(0, 6)))
        out.append('"' + text.replace('"', '""') + '"')
    return out


TEXTS += _random_texts()

#: Values that are not text, as VBA expressions written through Value.
OTHERS = ["-5", "0", "0.5", "3", "1E+20", "-1E+20", "True", "False", "CVErr(2042)", "CVErr(2007)",
          "CVErr(2015)", "CVErr(2036)", "CVErr(2000)", "CVErr(2023)", "CVErr(2029)", "#1/2/2020#"]
BLANKS = 3


def write_lines() -> list[str]:
    lines: list[str] = []
    row = 0
    for text in TEXTS:
        row += 1
        lines.append(f'ws.Cells({row}, 1).Value = "\'" & {text}')
    for value in OTHERS:
        row += 1
        lines.append(f"ws.Cells({row}, 1).Value = {value}")
    row += BLANKS
    return lines


def count() -> int:
    return len(TEXTS) + len(OTHERS) + BLANKS


def module() -> str:
    total = count()
    head = f'''Private Function Placed(ws As Object) As String
    Dim out As String, r As Long
    For r = 1 To {total}
        out = out & ws.Cells(r, 2).Value & ","
    Next
    Placed = out
End Function

Private Sub Fill(ws As Object)
    Dim r As Long
    ws.Cells.Clear
    For r = 1 To {total}
        ws.Cells(r, 2).Value = r
    Next
'''
    body = "\n".join(f"    {line}" for line in write_lines())
    tail = f'''
End Sub

Public Function Probe() As String
    Dim wb As Object, ws As Object, out As String, order As Variant, cased As Variant, i As Long, j As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    order = Array(xlAscending, xlDescending)
    cased = Array(False, True)
    For i = 0 To 1
        For j = 0 To 1
            Fill ws
            ws.Range("A1:B{total}").Sort Key1:=ws.Range("A1"), Order1:=order(i), Header:=xlNo, MatchCase:=cased(j)
            out = out & Placed(ws) & "|"
        Next
    Next
    wb.Close False
    Probe = out
End Function
'''
    return head + body + tail


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    runs = [[int(place) for place in part.split(",") if place] for part in str(result.value).split("|")[:4]]
    record = {"texts": TEXTS, "others": OTHERS, "blanks": BLANKS, "module": module(),
              "sorts": [{"order": order, "match_case": cased, "places": places}
                        for (order, cased), places in zip(
                            [("ascending", False), ("ascending", True), ("descending", False), ("descending", True)],
                            runs, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    labels = [*TEXTS, *OTHERS, *(["(blank)"] * BLANKS)]
    for sort in record["sorts"]:
        print(sort["order"], "match case" if sort["match_case"] else "", ":")
        print("  " + " ".join(labels[place - 1] for place in sort["places"]))


if __name__ == "__main__":
    main()
