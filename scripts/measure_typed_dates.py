"""How Excel types the dates and times scripts/measure_value_typing.py leaves out.

A space before a date or a time, a month's name between slashes or dashes
or beside a number that cannot be its day, a date and a time in either
order, what may follow them, and times with a space round a colon or a
colon at their end. Then strings made at random from those parts, which
the rules were not fitted to.

Each string is written through Range.Value into its own cell of a fresh
sheet, a few into a cell given a number format first, and the probe reads
TypeName and CStr of Value, CStr of Value2, NumberFormat and Text. The
record keeps the year the cases ran in, which a date typed without one
falls in.

    python scripts/measure_typed_dates.py

writes tests/fixtures/typed_dates.json, which tests/test_excel_typed_dates.py
replays.
"""

from __future__ import annotations

import datetime
import json
import random
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "typed_dates.json"

#: A space before or after a date, a time or a number.
SPACES = [" 1/2/2020", "  1/2/2020", " 1/2", " 1-2-2020", " 2020-1-2", " 2020/1/2", " 12/2020", " Jan 2",
          " Jan 2, 2020", " 2-Jan", " 2 Jan 2020", " Jan 2020", " May2020", " 5May2020", " January 2", " 12:30",
          " 12:30:45", " 12:30 PM", " 12 PM", " 1:2:3", " 12:30:45.5", " 1/2/2020 12:30", " 1,000", " 1e3", " (5)",
          " -5", " 1 1/2", " .5", " $5", "1/2/2020 ", "1/2 ", "Jan 2 ", "2-Jan ", "12:30 ", "12 PM ", "1/2/2020 12:30 ",
          "Jan 2020 ", "1 1/2 ", "1/2/2020  ", " 1/2/2020 ", " 2-Jan-2020", " 2-Jan-20", " 2/Jan", " 2 Jan", " 2Jan",
          " 2-January", " 12-Jan", " Jan-2", " Jan-2020", " Jan/2", " 2-Jan 12:30", "  2-Jan", " 2-Jan ", " 2 - Jan",
          " 2- Jan", " 2 -Jan", " 5-May", " 2-Jan 2020", " 2-Feb", " 31-Jan", " 32-Jan", " 1-2", " 1-2020", " 12-2020",
          " 2-Jan-2020 12:30", " 2-jan", " 2-Sept", " 2/Jan/2020", " 5May", " May5", " 2-Jan/2020", " 1-Jan", " 01-Jan",
          " 2-Jan-2020 ", " 2-Jan, 2020", " 2-Jan-99", " 45-Jan", " 2-", " -Jan", " 2-Jan-", " 2-Jan 1:00",
          " 2-Jan 0:00", " 3-Jan 12:30", " 2-Feb 12:30", " 2-Jan 12:31", " 2-Jan 13:30", " 2 Jan 12:30", " 2Jan 12:30",
          " 2-Jan 12 PM", " 2-Jan 12:30 PM", " 2-Jan 12:30:45", " 5-May 12:30", " 2/Jan 12:30", " 2-Jan 1:2:3",
          " 2-Jan 0:30", " 2-Jan 0:0:1", " 2-Jan 1:00 AM", " 2-Jan 23:59", " 2-Jan 24:00", " 2-Jan 12:30:45.5",
          " 12:30 2-Jan", " 2-Jan  12:30", "  2-Jan 12:30", " 29-Feb", " 30-Feb", " 2 January", " 2-Jan 5",
          " 2-Jan abc", " 2 -  Jan", " 2-Jan  ", " 2 Jan ", " 28-Feb", " Jan 2 12:30", " Jan 12:30", " 5 12:30",
          " 1/2 12:30", " 2-Jan-20 12:30", " 2-Jan 5:00", " 2-Jan 5 PM", " 2-Jan 12:30 5", " 2-Jan 1/2/2020",
          " 2 12:30", " 12:30 5", " 2-Jan 12:30:45 PM", " 2-Jan 30:00", " 2-Jan 99:30", " 2-Jan 1:1:1:1", " 2-Jan 5 6",
          " 2-Jan -5", " 2-Jan 12:30abc", " 2-Jan 10000:00"]
#: A month's name between slashes or dashes, beside a number that cannot be its day, with a year after it.
NAMES = ["Jan/2", "Jan/ 2", "Jan /2", "Jan / 2", "Jan/2/2020", "Jan/2020", "Jan/20", "2/Jan", "2/Jan/2020", "2/Jan/20",
         "Jan-2", "Jan-2-2020", "Jan-2-20", "Jan-20", "2020/Jan/2", "2020-Jan-2", "2020 Jan 2", "Jan.2", "2.Jan",
         "Jan. 2", "Jan. 2, 2020", "Jan, 2", "Jan,2", "Jan 2,2020", "Jan-02", "Jan/02/2020", "2-Jan 2020", "Jan 2-2020",
         "Jan 2/2020", "Jan-2 2020", "2/Jan 2020", "2 Jan-2020", "2 Jan/2020", "Jan 2 /2020", "2 Jan, 2020",
         "2, Jan 2020", "Jan 2nd", "2nd Jan", "January-2", "Jan 2 20", "Jan 02 2020", "2020 Jan", "2020-Jan",
         "2020/Jan", "Jan- 2", "Jan -2", "Jan - 2", "2 /Jan", "2/ Jan", "Jan/ 2020", "Jan /2020", "Jan-/2", "Jan//2",
         "/Jan 2", "Jan 2/", "Jan-2-", "-Jan-2", "Jan 45", "Jan-45", "Jan/45", "Jan 32", "Feb 30", "Feb 29", "Feb-29",
         "2/29", "2-29", "32-Jan", "32 Jan", "0-Jan", "Jan 0", "Jan-0", "Jan 00", "Jan 99", "Jan 100", "Jan 1899",
         "Jan 1900", "Jan 9999", "Jan 10000", "Jan 202", "Jan/202", "Jan 02020", "45-Jan", "Jan 29", "Jan 30", "2/30",
         "4/31", "Apr 31", "31-Apr", "Jan-2, 2020", "Jan/2, 2020", "Jan - 2, 2020", "January 2, 20", "Jan 2, 20",
         "Jan 2 , 2020", "Jan 2 ,2020", "Jan2,2020", "Jan-2,2020", "Jan 2, 99", "Jan 2, 1899", "Jan 2, 202",
         "Jan 2, 02020", "Jan 45, 2020", "Jan 2, 2020,", "Jan 2, 2020 ", "Jan 2 2020,", "2Jan2020", "2-Jan/2020",
         "2/Jan-2020", "2 Jan 20", "2 Jan 99", "2-Jan-99", "2-Jan-1899", "2-Jan-202", "2 - Jan - 2020",
         "2 / Jan / 2020", "2-January-2020", "32-Jan-2020", "2 Jan2020", "2Jan 2020", "2-Jan-02020", "2-Jan-0",
         "2-Jan-00", "2-Jan-2", "29-Feb", "29 Feb", "30-Feb", "29-Feb-2020", "29-Feb-2021", "29-Feb-20", "31-Jun",
         "Feb 29, 2021", "Feb 29, 2020", "Jan 2, 2", "Jan 2, 0", "Jan 2 ,  2020", "Jun 31", "Feb 28", "29Feb", "Feb29"]
#: A date and a time either way round, and a time AM or PM cannot hold, alone and beside a date.
PAIRS = ["May2020 12:30", "Jan 2 12:30", "Jan 2, 2020 12:30", "2 Jan 2020 12:30", "2-Jan 12:30", "2-Jan-2020 12:30",
         "Jan 2020 12:30", "Jan-2020 12:30", "January 2 12:30", "Jan 2 2020 12:30", "5May2020 12:30", "May5 12:30",
         "May2020 12:30:45", "Jan 2 12 PM", "Jan 2 12:30 PM", "2-Jan-20 12:30", "Jan 2, 2020 12:30:45.5", "1/2 12:30",
         "12/2020 12:30", "1/2020 12:30", "1-2 12:30", "1/2/2020 12 PM", "1/2/2020 12:30:45.5", "1/2/2020 25:00",
         "1/2/2020 1:60", "1/2 25:00", "12:30 1/2/2020", "12:30 Jan 2", "12:30 PM 1/2/2020", "12 PM 1/2/2020",
         "Jan 2020 12 PM", "May2020 25:00", "Jan 2 1:2:3", "2 Jan 12:30", "12:30 1/2", "12:30 12/2020",
         "12:30:45.5 1/2/2020", "25:00 1/2/2020", "12:30 2-Jan-2020", "1/2/2020 12:30 1/2/2020", "12:30 12:30",
         "1/2/2020 5", "1/2/2020 5 PM", "1/2/2020 12:30 5", "Jan 2 25:00", "Jan 2 12:30:45.5", "12:30 Jan 2, 2020",
         "12:30 May2020", "1/2/2020 0:00", "1/2/2020 12:00 AM", "2-Jan 12:30:45.5", "12:30 PM Jan 2",
         "1/2/2020 13:30 PM", "Jan 2 13 PM", "1/2/2020 12:30am", "1/2/2020 12:30:45 PM", "1/2 12:30:45", "12:30:45 1/2",
         "1/2/2020 12", "1/2/2020 1:60 PM", "1/2/2020 99:00", "1/2/2020 9999:00", "1/2/2020 0:60", "1/2/2020 0:0:60",
         "12/31/9999 23:59", "12/31/9999 24:00", "1/2/2020 12:30:45.25", "1/2/2020 1:2.5", "1/2/2020 12:30:45.5 PM",
         "12 PM Jan 2", "1/2/2020 12:30 PM ", "1/0/2020 12:30", "1/2/2020  12:30  PM", "12:30  1/2/2020",
         "1/2/2020 12:30:45.5 ", "13:30 PM", "1:60 PM", "13 PM", "0:30 AM", "13:30 AM", "12:60 AM", "0 PM",
         "1/2/2020 25:00 PM", "Jan 2 1:60 PM", "12/2020 13 PM", "1/2/2020 13:30:45.5 PM", "13:30:45 PM", "1/2 13 PM",
         "13:30 PM 1/2/2020", "13 PM 1/2/2020", "1/2/2020 0 PM", "Jan 2020 13 PM", "2-Jan 13:30 PM"]
#: What may follow a date and a time: how many numbers, of how many digits, apart by what.
AFTER = ["12:30 5", "1/2/2020 12:30 abc", "1/2/2020 12:30 5 6", "1/2/2020 12:30 1/3/2021", "1/2 12:30 5",
         "Jan 2 12:30 5", "1/2/2020 12:30:45 5", "1/2/2020 12:30 PM 5", "1/2/2020 12:30 99", "12:30 1/2/2020 5",
         "1/2/2020 12:30 Jan", "1/2/2020 12:30 12:30", "1/2/2020 12:30 -5", "1/2/2020 12:30 5%", "1/2/2020 12:30 PM PM",
         "12:30 5 1/2/2020", "1/2/2020 5 12:30", "12:30 PM 5", "12:30 1/2/2020 12:30", "1/2/2020 12:30 5:00",
         "1/2/2020 12:30 1234567", "1/2/2020 12:30 5.5", "1/2/2020 12:30 2020", "12:30 abc", "12:30 PM abc",
         "1/2/2020 12:30 05", "5 1/2/2020", "abc 1/2/2020", "1/2/2020 12:30,", "1/2/2020 12:30 ,5",
         "1/2/2020 12:30 12345", "1/2/2020 12:30 123456", "1/2/2020 12:30 123", "1/2/2020 12:30 /5",
         "1/2/2020 12:30 5/", "1/2/2020 12:30 5-", "1/2/2020 12:30 5-6-7-8", "1/2/2020 12:30 5 6 7 8 9",
         "1/2/2020 12:30 +5", "1/2/2020 12:30 (5)", "1/2/2020 12:30 $5", "1/2/2020 12:30 5 PM",
         "1/2/2020 12:30:45 PM 5 6", "Jan 2, 2020 12:30 5", "2-Jan-2020 12:30 5", "May2020 12:30 5",
         "1/2/2020 12:30:45.5 5", "1/2/2020 25:00 5", "1/2/2020 12 PM 5", "Jan 2 12 PM 5", "1/2 12:30 1/2",
         "1/2/2020 12:30 0", "1/2/2020 12:30 00000", "1/2/2020 12:30 5  6", "1/2/2020 12:30 5 ", "1/2/2020 12:30 - 5",
         "1/2/2020 12:30 -", "1/2/2020 12:30 /", "1/2/2020 12:30 5//6", "1/2/2020 12:30 --5", "1/2/2020 12:30 5 -6",
         "1/2/2020 12:30 1/2/2020 12:30", "1/2/2020 12:30 13 PM", "1/2/2020 13:30 PM 5", "1/2/2020 12:30 9999",
         "1/2/2020 12:30 10000", "1/2/2020 12:30 1-2-3-4-5-6", "1/2/2020 12:30 5.", "1/2/2020 12:30 .5",
         "1/2/2020 12:30 1 1/2", "12/2020 12:30 5", "1/2/2020 1:60 5", "1/2/2020 12:30 5-Jan", "1/2/2020 12:30 5 abc",
         "1/2/2020 12:30 5 1234567", "1/2/2020 12:30 1/2/2020 5", "1/2/2020 12:30 5 1/2", "1/2/2020 12:30 0005",
         "1/2 12:30 5 6 7 8 9", "1/2/2020 12:30:45 5 6 7 8", "1/2/2020 12:30 0123", "1/2/2020 12:30 0999",
         "1/2/2020 12:30 1899", "1/2/2020 12:30 1900", "1/2/2020 12:30 1000", "1/2/2020 12:30 0020",
         "1/2/2020 12:30 020", "1/2/2020 12:30 999", "1/2/2020 12:30 000", "1/2/2020 12:30 00", "1/2/2020 12:30 5 0005",
         "1/2/2020 12:30 0005 5", "1/2 12:30 0005", "1/2/2020 12:30 5-0005", "1/2/2020 12:30 3000",
         "1/2/2020 12:30 12 31", "1/2/2020 12:30 99 99 99 99", "1/2/2020 12:30 13/45/2020", "1/2/2020 12:30 9999 9999",
         "1/2/2020 12:30 5 5 5 5", "12:30 1/2/2020 -", "12:30 1/2/2020 /", "1/2 12:30 5 6 7 8",
         "1/2 12:30 1 2 3 4 5 6 7", "12/2020 12:30 5 6 7 8 9", "12/2020 12:30 5 6 7 8", "1/2/2020 12:30 PM 5 6 7 8",
         "1/2/2020 12:30 PM 5 6 7", "1/2/2020 12:30:45 5 6 7", "1/2/2020 12:30 0999 5", "1/2/2020 12:30 1234",
         "1/2/2020 12:30 0100", "1/2/2020 12:30 0001", "1/2/2020 12:30 1/2/0005", "1/2/2020 12:30 1/2/1899",
         "1/2/2020 12:30 13/2", "1/2/2020 12:30:45.5 5 6 7", "1/2/2020 12:30:45.5 5 6", "Jan 2, 2020 12:30 5 6 7 8 9",
         "Jan 2, 2020 12:30 5 6 7 8", "2-Jan-2020 12:30 5 6 7 8 9", "2-Jan-2020 12:30 5 6 7 8",
         "Jan 2 12:30 5 6 7 8 9 1", "Jan 2 12:30 5 6 7 8 9", "1/2/2020 12:30 .5 5 6 7", "1/2/2020 12:30 .5 5 6",
         "May2020 12:30 5 6 7 8 9", "May2020 12:30 5 6 7 8 9 1", "1/2/2020 12:30 5 Jan", "Jan 2 12:30 Feb",
         "12:30 Jan 2 5", "1/2/2020 12:30 Jan 5"]
#: Times with a space round a colon or before the point, a colon or a point at the end, a point after AM or PM,
#: and a slash or a dash before the date that follows a time.
CLOCKS = ["1/2/2020 10000:00", "1/2/2020 12:30:45:5", "1/2/2020 12:", "1/2/2020 :30", "1/2/2020 12:30:",
          "1/2/2020 12::30", "1/2/2020 12:30 P", "1/2/2020 12:30 A", "1/2/2020 12:30 AMPM", "1/2/2020 12.30",
          "1/2/2020 13 AM", "1/2/2020 12:30 pm", "12:", "12:30:", "1:2:", "12: PM", "12:30 .5", "12:30:45 .5",
          "12:30 . 5", "12:30.5 PM", "12 :30", "12: 30", "12 : 30", "12:30 :45", "12:30: 45", "12:30 45", "0:", "25:",
          "12:PM", "12:30:PM", "12:30.", "12:30.5.", ":30", "12:30 .", "12:30 .5 PM", "12:30 PM .5", "99:", "12:30 .50",
          "12:30 .123456", "1/2/2020 12: PM", "1/2/2020 12 :30", "1/2/2020 12: 30", "1/2/2020 12:30 .5 5",
          "1/2/2020 12:30 . 5", "12: 1/2/2020", "12:30 .5 1/2/2020", "1/2/2020 12:30:45 .5", "1/2/2020 25:", "12:  30",
          "12:30:45.5.5", "1/2/2020 12: 5", "12: 5 1/2/2020", "12:30: 45 1/2/2020", "12: 1/2", "12: 30 1/2/2020",
          "1/2/2020 12: 1/2", "12: 12/2020", "12: Jan 2", "12:30: 1/2/2020", "12:30: 1/2", "1/2/2020 12:30: 5",
          "12:30 /1/2/2020", "12:30 -1/2/2020", "12:30 /2/2020", "/1/2/2020 12:30", "12:30 1/2/2020 PM", "12:30 /Jan 2",
          "12:30 -2-Jan", "12:30 AM .5", "1/2/2020 12:30 PM .5", "12 PM .5", "12:30:45 PM .5"]
#: Where Excel's misreadings stop: a point after AM or PM with no digits or with words after them, a space before
#: words that are no date's, a fraction of a second with no colon.
EDGES = ["10 am. Meeting", "12.5 PM", "12 .5 PM", " Meeting on Jan 2 at 10:30", " (Jan 2 12:30)", "5 PM .5 note",
         "12 PM .5 abc", "10 am.", "10 am .", "12:30 PM.", "Jan 2 at 12:30", "12:30 on Jan 2", " Jan 2 at 10:30",
         " on Jan 2 12:30", "12.5 AM", "1/2/2020 12.5 PM", "5.5 PM", "12:30.5", "12 PM.5", "12 PM. 5", "12 PM .",
         " 2-Jan 10 am.", " (2-Jan 12:30)", " 2-Jan.12:30", " -2-Jan 12:30", " 2-Jan 12:30 x", " x 2-Jan 12:30"]
#: Strings a VBA expression makes: a tab or a no-break space beside a number or a date.
EXPRESSIONS = [("\t5", 'vbTab & "5"'), ("5\t", '"5" & vbTab'), ("\xa05", 'Chr(160) & "5"'),
               ("\t1/2/2020", 'vbTab & "1/2/2020"'), ("\xa01/2/2020", 'Chr(160) & "1/2/2020"')]
#: A space before a date or a time typed into a cell that already has a number format.
FORMATTED = [(text, code) for code in ("0.00", "# ?/?", "@", "m/d/yyyy", "h:mm")
             for text in (" 1/2", " 1/2/2020", " 12:30", " 2-Jan", " 1 1/2")]

#: The parts the random strings are made of.
PARTS = {"months": ["Jan", "feb", "MAR", "April", "may", "Jun", "July", "Aug", "Sept", "Sep", "oct", "November", "Dec",
                    "Juneau", "Ma"],
         "separators": ["/", "-", " ", "", " / ", " - ", "- ", " /", ",", ", ", ".", "//"],
         "days": ["1", "2", "09", "15", "28", "29", "30", "31", "0", "32", "7"],
         "years": ["2020", "20", "99", "1999", "5", "0", "2021", "1900", "9999", "202", "02020", "1899", "45"],
         "times": ["12:30", "1:05:09", "23:59:59", "0:00", "7 AM", "11 pm", "12:30 a", "3:4 P", "25:30", "1:75",
                   "12:30:45.25", "9:15.5", "13:00 PM", "12:", "12: 5", "12 : 30", "5:06:07 AM", "10:30 .5", "100:00",
                   "12:00 AM"],
         "tails": ["", "", "", " 7", " 7 8", " 2021", " 0099", " 12/31", " -3", " 99999", " x", " 7:00", " 4/5/6/7",
                   " 7-", " /"]}

READS = ["TypeName(c.Value)", "CStr(c.Value)", "CStr(c.Value2)", "c.NumberFormat", "c.Text"]
#: Strings typed by one procedure; many more make a procedure larger than VBA compiles.
BATCH = 100

HELPER = '''Private Function Typed(ws As Object, first As Long, t As Variant, f As Variant) As String
    Dim out As String, c As Object, i As Long
    For i = LBound(t) To UBound(t)
        Set c = ws.Cells(first + i, 1)
        If f(i) <> "" Then c.NumberFormat = f(i)
        c.Value = t(i)
        out = out & TypeName(c.Value) & "`" & CStr(c.Value) & "`" & CStr(c.Value2) & "`" & c.NumberFormat _
            & "`" & c.Text & "^"
    Next
    Typed = out
End Function
'''


def _date(rng: random.Random) -> str:
    """A date made of random parts: a day and a month's name either way round, a month's name and a year, digits."""
    parts = PARTS
    kind = rng.randrange(6)
    month = rng.choice(parts["months"])
    day, year = rng.choice(parts["days"]), rng.choice(parts["years"])
    first, second = rng.choice(parts["separators"]), rng.choice(parts["separators"])
    if kind == 0:
        return f"{day}{first}{month}{second}{year}" if rng.random() < 0.6 else f"{day}{first}{month}"
    if kind == 1:
        return f"{month}{first}{day}" + (f"{rng.choice([', ', ' , ', ',', ' '])}{year}" if rng.random() < 0.6 else "")
    if kind == 2:
        return f"{month}{first}{year}"
    number = str(rng.randint(1, 13))
    if kind == 3:
        return f"{number}{first}{day}{second}{year}" if rng.random() < 0.6 else f"{number}{first}{day}"
    if kind == 4:
        return f"{year}{first}{number}{second}{day}"
    return f"{number}{first}{year}"


def _fresh(rng: random.Random) -> str:
    """A string made at random: a date, a time, or both either way round, sometimes with a space before or after."""
    shape = rng.randrange(5)
    if shape == 0:
        text = _date(rng)
    elif shape == 1:
        text = rng.choice(PARTS["times"])
    elif shape == 2:
        text = f"{_date(rng)} {rng.choice(PARTS['times'])}{rng.choice(PARTS['tails'])}"
    elif shape == 3:
        text = f"{rng.choice(PARTS['times'])} {rng.choice(['', '/', '- '])}{_date(rng)}"
    else:
        text = f"{_date(rng)}{rng.choice(['', ' '])}{rng.choice(PARTS['times'])}"
    if rng.random() < 0.08:
        text = " " + text
    if rng.random() < 0.08:
        text += " "
    return text


def cases() -> list[dict[str, str]]:
    """Every case once: its text, the VBA expression that makes it, and a format the cell is given first."""
    rng = random.Random(20260926)
    fresh = [_fresh(rng) for _ in range(330)]
    found: dict[str, dict[str, str]] = {}
    for text in SPACES + NAMES + PAIRS + AFTER + CLOCKS + EDGES + fresh:
        found.setdefault(text, {"text": text})
    listed = list(found.values())
    listed += [{"text": text, "expression": expression} for text, expression in EXPRESSIONS]
    listed += [{"text": text, "format": code} for text, code in FORMATTED]
    return listed


def _vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def module(listed: list[dict[str, str]]) -> str:
    """The probe: a procedure per batch of strings, each typed into its own row of column A of a new sheet."""
    head = ["Public Function Probe() As String", "Dim ws As Object, out As String",
            "Set ws = ActiveWorkbook.Worksheets.Add", 'ws.Columns("A").ColumnWidth = 40']
    body: list[str] = []
    for number, start in enumerate(range(0, len(listed), BATCH)):
        batch = listed[start:start + BATCH]
        body += [f"Private Function Batch{number}(ws As Object) As String",
                 f"Dim t(0 To {len(batch) - 1}) As String, f(0 To {len(batch) - 1}) As String"]
        for index, case in enumerate(batch):
            body.append(f"t({index}) = {case.get('expression') or _vba_text(case['text'])}")
            if "format" in case:
                body.append(f"f({index}) = {_vba_text(case['format'])}")
        body += [f"Batch{number} = Typed(ws, {start + 1}, t, f)", "End Function"]
        head.append(f"out = out & Batch{number}(ws)")
    head += ["Probe = out", "End Function"]
    return "\n".join([HELPER, *head, *body]) + "\n"


def main() -> None:
    listed = cases()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(listed), "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = [part.split("`") for part in str(result.value).split("^")[: len(listed)]]
    # A date typed without a year falls in the year it was typed, which a replay has to know.
    record = {"year": datetime.date.today().year, "reads": READS,
              "cases": [{**case, "answer": answer} for case, answer in zip(listed, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for case, answer in zip(listed, answers, strict=True):
        print(f"{case['text']!r:34} {case.get('format', ''):9} {answer}")


if __name__ == "__main__":
    main()
