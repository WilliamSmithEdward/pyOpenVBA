"""How Excel and VBA round a serial to the second, as live Excel answers it.

A time on the half second is where ways of rounding part: working in
double arithmetic, exactly, adding half a second first, taking the day
off first. Every value here is one of those, chosen where the ways
disagree -- whole seconds, and tenths, hundredths and thousandths of
one -- with random times beside them. Each is written into cells whose
Text shows it through a date and time, an elapsed time and a time with
fractions of a second, and read through the TEXT function, through VBA's
CStr, Format, Hour, Minute and Second, and as the text Replace edits in
the cell (which the formula bar shows), read by replacing each digit in
turn with q.

Each value goes to Excel exactly: as (high * 2^26 + low) * 2^exponent,
which VBA rebuilds without rounding.

    python scripts/measure_time_rounding.py

writes tests/fixtures/time_rounding.json, which
tests/test_excel_time_rounding.py replays.
"""

from __future__ import annotations

import json
import math
import random
from fractions import Fraction
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "time_rounding.json"

#: Range.Text through each of these, in a column wide enough for any of them.
FORMATS = ["m/d/yyyy h:mm:ss", "[h]:mm:ss", "h:mm:ss.0", "h:mm:ss.00", "h:mm:ss.000", "h:mm", "[s]"]
#: What else reads each value: the TEXT function, VBA's own conversions, and the edited text.
READS = ['Application.WorksheetFunction.Text(v, "h:mm:ss")', "CStr(CDate(v))", 'Format(v, "hh:mm:ss")',
         'Hour(v) & ":" & Minute(v) & ":" & Second(v)']

EDIT_TEXT = '''Private Function EditText(c As Object) As String
    Dim d As Long
    For d = 0 To 9
        c.Replace CStr(d), "q", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        If TypeName(c.Value2) = "String" Then
            EditText = Replace(c.Value2, "q", CStr(d))
            Exit Function
        End If
    Next
    EditText = "?" & c.Formula
End Function
'''


def _ties(places: int, days: tuple[int, ...], count: int, chooser: random.Random) -> list[float]:
    """Times on half a unit of ``places`` decimals of a second, where rounding in doubles and adding half first part."""
    scale = 86400 * 10**places
    found: list[float] = []
    while len(found) < count:
        value = chooser.choice(days) + (chooser.randrange(scale) + 0.5) / scale
        for near in (value, math.nextafter(value, 0), math.nextafter(value, math.inf)):
            half_first = near + 0.5 / scale
            direct = math.floor(near * scale + 0.5)
            added = math.floor(half_first * scale)
            if direct != added or Fraction(near) * scale % 1 == Fraction(1, 2):
                found.append(near)
                break
    return found


def values() -> list[float]:
    chooser = random.Random(20260922)
    out: list[float] = []
    for places in range(4):
        out += _ties(places, (0, 1, 60, 43832, 45000, 2958465), 24, chooser)
    out += [chooser.random() for _ in range(24)]
    out += [chooser.uniform(1, 2958466) for _ in range(24)]
    # Named cases: the half second a digit less than a tie, the last instant of a day, the first and last dates.
    out += [0.500005787037037, 0.999999999, 0.9999999999, 1e-20, 0.0, 2958465.9999999, 60.5, 59.99999999]
    return out


def exact_parts(value: float) -> tuple[int, int, int]:
    """A double as (high, low, exponent), which VBA rebuilds exactly as (high * 2^26 + low) * 2^exponent."""
    mantissa, exponent = math.frexp(value)
    whole = int(mantissa * 2**53)
    assert Fraction(whole) * Fraction(2) ** (exponent - 53) == Fraction(value)
    return whole >> 26, whole & (2**26 - 1), exponent - 53


def module(serials: list[float]) -> str:
    numbers = [str(part) for value in serials for part in exact_parts(value)]
    # A VBA line stops at 1,023 characters, so the numbers go in a few at a time.
    chunks = "\n".join(f'    text = text & "{",".join(numbers[at:at + 30])},"' for at in range(0, len(numbers), 30))
    cells = "\n".join(f'        ws.Cells(r, {column}).Value = v\n        ws.Cells(r, {column}).NumberFormat = "{code}"'
                      for column, code in enumerate(FORMATS, start=2))
    shown = " & \"^\" & ".join(f"ws.Cells(r, {column}).Text" for column in range(2, len(FORMATS) + 2))
    # Each read keeps its own error, as E and the number.
    reads = "\n".join(f'    Err.Clear\n    x = ""\n    x = {read}\n    If Err.Number <> 0 Then x = "E" & Err.Number\n'
                      f'    Reads = Reads & "^" & x' for read in READS)
    return EDIT_TEXT + f'''Private Function Reads(v As Double) As String
    Dim x As String
    On Error Resume Next
{reads}
    On Error GoTo 0
End Function

Public Function Probe() As String
    Dim wb As Object, ws As Object, parts As Variant, i As Long, r As Long, out As String, v As Double
    Dim text As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Columns("A:J").ColumnWidth = 30
{chunks}
    parts = Split(Left(text, Len(text) - 1), ",")
    For i = 0 To UBound(parts) Step 3
        r = i \\ 3 + 1
        v = (CDbl(CLng(parts(i))) * 67108864# + CLng(parts(i + 1))) * (2# ^ CLng(parts(i + 2)))
        ws.Cells(r, 1).Value = v
        ws.Cells(r, 1).NumberFormat = "m/d/yyyy h:mm"
        If ws.Cells(r, 1).Value2 <> v Then out = out & "!"
{cells}
        out = out & {shown} & Reads(v) & "^" & EditText(ws.Cells(r, 1)) & "|"
    Next
    wb.Close False
    Probe = out
End Function
'''


def main() -> None:
    serials = values()
    code = module(serials)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Probe", timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    rows = [row.split("^") for row in str(result.value).split("|")[: len(serials)]]
    assert not any(row[0].startswith("!") for row in rows), "a value did not reach Excel exactly"
    assert all(len(row) == len(FORMATS) + len(READS) + 1 for row in rows)
    record = {"formats": FORMATS, "reads": READS, "code": code,
              "cases": [{"value": value, "shown": row[: len(FORMATS)], "read": row[len(FORMATS):-1],
                         "edited": row[-1]} for value, row in zip(serials, rows, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for value, row in zip(serials, rows, strict=True):
        print(repr(value), " | ".join(row))


if __name__ == "__main__":
    main()
