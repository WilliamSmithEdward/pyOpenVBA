"""How Excel spells a stored number: in Range.Formula, in text a formula makes of it, and under General.

A cell holding 1E+20 reads back through Formula as 100000000000000000000
while ``=A1&""`` makes 1E+20 of it, and neither is what VBA's CStr would
write; the cell itself shows 1E+20 through its General format, which has
a width of its own. The probe stores numbers from 1E-25 to 1E+300 with
from one to fifteen significant digits, negatives among them, and reads
all three spellings of each, the last in a column wide enough that its
width never decides what shows.

    python scripts/measure_number_spelling.py

writes tests/fixtures/number_spelling.json, which
tests/test_excel_number_spelling.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "number_spelling.json"

MANTISSA = "123456789012345"
BATCH = 100


def numbers() -> list[str]:
    """Each exponent with one to fifteen significant digits, and every seventh again as a negative."""
    out: list[str] = []
    for exponent in [*range(-25, 0), *range(10, 31), 100, 300]:
        for digits in (1, 2, 3, 5, 8, 10, 12, 14, 15):
            body = MANTISSA[:digits]
            out.append(f"{body[0]}{'.' + body[1:] if digits > 1 else ''}E{exponent:+d}")
    return out + [f"-{number}" for number in out[::7]]


def main() -> None:
    written = numbers()
    lines = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)",
             "Set ws = wb.Worksheets(1)", 'ws.Range("B1").Formula = "=A1&"""""', "ws.Columns(1).ColumnWidth = 60"]
    batches: list[str] = []
    # A procedure holds only so much code, so the numbers go in batches of their own.
    for start in range(0, len(written), BATCH):
        name = f"Batch{start // BATCH}"
        lines.append(f"out = out & {name}(ws)")
        batches += [f"Private Function {name}(ws As Object) As String", "Dim out As String"]
        for number in written[start:start + BATCH]:
            batches += [f'ws.Range("A1").Value = CDbl("{number}")',
                        'out = out & ws.Range("A1").Formula & "~" & ws.Range("B1").Value & "~" & '
                        'ws.Range("A1").Text & "|"']
        batches += [f"{name} = out", "End Function"]
    lines += ["wb.Close False", "Probe = out", "End Function", *batches]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(lines) + "\n", "Probe", timeout=300.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = [answer.split("~") for answer in str(result.value).split("|")[: len(written)]]
    record = [{"number": number, "formula": formula, "text": text, "general": general}
              for number, (formula, text, general) in zip(written, answers, strict=True)]
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for entry in record:
        print(f"{entry['number']:24} {entry['formula']:34} {entry['text']:24} {entry['general']}")


if __name__ == "__main__":
    main()
