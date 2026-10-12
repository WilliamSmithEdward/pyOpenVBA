"""Measure the date functions over a matrix: DateAdd and DateDiff for every
interval, either way and around day zero (a Date before 12/30/1899 counts its
day back and its time forward), DatePart, week numbers with each first day and
first week, Weekday, DateSerial and TimeSerial at their edges, the part
functions and MonthName and WeekdayName.

Each probe runs a group of calls and records TypeName and value of each, or e
and the error number. Each probe is a module whose ``Function Probe() As String``
the fixture records (see vba_probe_fixture.py).
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_date_functions.json'

CALLS = {
    'dateadd_plus': [
        'DateAdd("yyyy", 1, #1/31/2024#)',
        'DateAdd("yyyy", 1, #2/29/2024#)',
        'DateAdd("yyyy", 1, #12/31/2023 11:59:59 PM#)',
        'DateAdd("yyyy", 1, #3/10/2024 2:30:00 AM#)',
        'DateAdd("q", 1, #1/31/2024#)',
        'DateAdd("q", 1, #2/29/2024#)',
        'DateAdd("q", 1, #12/31/2023 11:59:59 PM#)',
        'DateAdd("q", 1, #3/10/2024 2:30:00 AM#)',
        'DateAdd("m", 1, #1/31/2024#)',
        'DateAdd("m", 1, #2/29/2024#)',
        'DateAdd("m", 1, #12/31/2023 11:59:59 PM#)',
        'DateAdd("m", 1, #3/10/2024 2:30:00 AM#)',
        'DateAdd("y", 1, #1/31/2024#)',
        'DateAdd("y", 1, #2/29/2024#)',
        'DateAdd("y", 1, #12/31/2023 11:59:59 PM#)',
        'DateAdd("y", 1, #3/10/2024 2:30:00 AM#)',
        'DateAdd("d", 1, #1/31/2024#)',
        'DateAdd("d", 1, #2/29/2024#)',
        'DateAdd("d", 1, #12/31/2023 11:59:59 PM#)',
        'DateAdd("d", 1, #3/10/2024 2:30:00 AM#)',
        'DateAdd("w", 1, #1/31/2024#)',
        'DateAdd("w", 1, #2/29/2024#)',
        'DateAdd("w", 1, #12/31/2023 11:59:59 PM#)',
        'DateAdd("w", 1, #3/10/2024 2:30:00 AM#)',
        'DateAdd("ww", 1, #1/31/2024#)',
        'DateAdd("ww", 1, #2/29/2024#)',
        'DateAdd("ww", 1, #12/31/2023 11:59:59 PM#)',
        'DateAdd("ww", 1, #3/10/2024 2:30:00 AM#)',
        'DateAdd("h", 1, #1/31/2024#)',
        'DateAdd("h", 1, #2/29/2024#)',
        'DateAdd("h", 1, #12/31/2023 11:59:59 PM#)',
        'DateAdd("h", 1, #3/10/2024 2:30:00 AM#)',
        'DateAdd("n", 1, #1/31/2024#)',
        'DateAdd("n", 1, #2/29/2024#)',
        'DateAdd("n", 1, #12/31/2023 11:59:59 PM#)',
        'DateAdd("n", 1, #3/10/2024 2:30:00 AM#)',
        'DateAdd("s", 1, #1/31/2024#)',
        'DateAdd("s", 1, #2/29/2024#)',
        'DateAdd("s", 1, #12/31/2023 11:59:59 PM#)',
        'DateAdd("s", 1, #3/10/2024 2:30:00 AM#)',
    ],
    'dateadd_minus': [
        'DateAdd("yyyy", -1, #1/1/1900#)',
        'DateAdd("yyyy", -1, #12/30/1899#)',
        'DateAdd("yyyy", -1, #12/29/1899 6:00:00 AM#)',
        'DateAdd("q", -1, #1/1/1900#)',
        'DateAdd("q", -1, #12/30/1899#)',
        'DateAdd("q", -1, #12/29/1899 6:00:00 AM#)',
        'DateAdd("m", -1, #1/1/1900#)',
        'DateAdd("m", -1, #12/30/1899#)',
        'DateAdd("m", -1, #12/29/1899 6:00:00 AM#)',
        'DateAdd("y", -1, #1/1/1900#)',
        'DateAdd("y", -1, #12/30/1899#)',
        'DateAdd("y", -1, #12/29/1899 6:00:00 AM#)',
        'DateAdd("d", -1, #1/1/1900#)',
        'DateAdd("d", -1, #12/30/1899#)',
        'DateAdd("d", -1, #12/29/1899 6:00:00 AM#)',
        'DateAdd("w", -1, #1/1/1900#)',
        'DateAdd("w", -1, #12/30/1899#)',
        'DateAdd("w", -1, #12/29/1899 6:00:00 AM#)',
        'DateAdd("ww", -1, #1/1/1900#)',
        'DateAdd("ww", -1, #12/30/1899#)',
        'DateAdd("ww", -1, #12/29/1899 6:00:00 AM#)',
        'DateAdd("h", -1, #1/1/1900#)',
        'DateAdd("h", -1, #12/30/1899#)',
        'DateAdd("h", -1, #12/29/1899 6:00:00 AM#)',
        'DateAdd("n", -1, #1/1/1900#)',
        'DateAdd("n", -1, #12/30/1899#)',
        'DateAdd("n", -1, #12/29/1899 6:00:00 AM#)',
        'DateAdd("s", -1, #1/1/1900#)',
        'DateAdd("s", -1, #12/30/1899#)',
        'DateAdd("s", -1, #12/29/1899 6:00:00 AM#)',
    ],
    'dateadd_fractions': [
        'DateAdd("d", 1.5, #1/1/2024#)',
        'DateAdd("m", 1.5, #1/31/2024#)',
        'DateAdd("h", 0.5, #1/1/2024#)',
        'DateAdd("yyyy", -0.5, #1/1/2024#)',
        'DateAdd("m", 13, #1/31/2024#)',
        'DateAdd("q", 1, #11/30/2024#)',
        'DateAdd("x", 1, #1/1/2024#)',
        'DateAdd("d", Null, #1/1/2024#)',
        'DateAdd("d", 1, Null)',
        'DateAdd("d", 1, "2024-01-01")',
    ],
    'datediff_intervals': [
        'DateDiff("yyyy", #12/31/2023 11:00:00 PM#, #1/2/2024 1:00:00 AM#)',
        'DateDiff("q", #12/31/2023 11:00:00 PM#, #1/2/2024 1:00:00 AM#)',
        'DateDiff("m", #12/31/2023 11:00:00 PM#, #1/2/2024 1:00:00 AM#)',
        'DateDiff("y", #12/31/2023 11:00:00 PM#, #1/2/2024 1:00:00 AM#)',
        'DateDiff("d", #12/31/2023 11:00:00 PM#, #1/2/2024 1:00:00 AM#)',
        'DateDiff("w", #12/31/2023 11:00:00 PM#, #1/2/2024 1:00:00 AM#)',
        'DateDiff("ww", #12/31/2023 11:00:00 PM#, #1/2/2024 1:00:00 AM#)',
        'DateDiff("h", #12/31/2023 11:00:00 PM#, #1/2/2024 1:00:00 AM#)',
        'DateDiff("n", #12/31/2023 11:00:00 PM#, #1/2/2024 1:00:00 AM#)',
        'DateDiff("s", #12/31/2023 11:00:00 PM#, #1/2/2024 1:00:00 AM#)',
    ],
    'datediff_reverse': [
        'DateDiff("yyyy", #3/15/2024#, #1/10/2023#)',
        'DateDiff("q", #3/15/2024#, #1/10/2023#)',
        'DateDiff("m", #3/15/2024#, #1/10/2023#)',
        'DateDiff("y", #3/15/2024#, #1/10/2023#)',
        'DateDiff("d", #3/15/2024#, #1/10/2023#)',
        'DateDiff("w", #3/15/2024#, #1/10/2023#)',
        'DateDiff("ww", #3/15/2024#, #1/10/2023#)',
        'DateDiff("h", #3/15/2024#, #1/10/2023#)',
        'DateDiff("n", #3/15/2024#, #1/10/2023#)',
        'DateDiff("s", #3/15/2024#, #1/10/2023#)',
    ],
    'datediff_week_starts': [
        'DateDiff("ww", #1/1/2024#, #1/31/2024#, 0)',
        'DateDiff("ww", #1/1/2024#, #1/31/2024#, 1)',
        'DateDiff("ww", #1/1/2024#, #1/31/2024#, 2)',
        'DateDiff("ww", #1/1/2024#, #1/31/2024#, 3)',
        'DateDiff("ww", #1/1/2024#, #1/31/2024#, 4)',
        'DateDiff("ww", #1/1/2024#, #1/31/2024#, 5)',
        'DateDiff("ww", #1/1/2024#, #1/31/2024#, 6)',
        'DateDiff("ww", #1/1/2024#, #1/31/2024#, 7)',
        'DateDiff("w", #1/1/2024#, #1/31/2024#, 1)',
        'DateDiff("w", #1/1/2024#, #1/31/2024#, 2)',
    ],
    'datepart_all': [
        'DatePart("yyyy", #12/31/2024 1:02:03 PM#)',
        'DatePart("q", #12/31/2024 1:02:03 PM#)',
        'DatePart("m", #12/31/2024 1:02:03 PM#)',
        'DatePart("y", #12/31/2024 1:02:03 PM#)',
        'DatePart("d", #12/31/2024 1:02:03 PM#)',
        'DatePart("w", #12/31/2024 1:02:03 PM#)',
        'DatePart("ww", #12/31/2024 1:02:03 PM#)',
        'DatePart("h", #12/31/2024 1:02:03 PM#)',
        'DatePart("n", #12/31/2024 1:02:03 PM#)',
        'DatePart("s", #12/31/2024 1:02:03 PM#)',
    ],
    'datepart_ww_years': [
        'DatePart("ww", #1/1/2021#, 1, 1)',
        'DatePart("ww", #1/1/2021#, 2, 2)',
        'DatePart("ww", #1/1/2021#, 2, 3)',
        'DatePart("ww", #1/1/2021#, 1, 3)',
        'DatePart("ww", #12/31/2024#, 1, 1)',
        'DatePart("ww", #12/31/2024#, 2, 2)',
        'DatePart("ww", #12/31/2024#, 2, 3)',
        'DatePart("ww", #12/31/2024#, 1, 3)',
        'DatePart("ww", #1/3/2027#, 1, 1)',
        'DatePart("ww", #1/3/2027#, 2, 2)',
        'DatePart("ww", #1/3/2027#, 2, 3)',
        'DatePart("ww", #1/3/2027#, 1, 3)',
    ],
    'weekday_starts': [
        'Weekday(#3/6/2024#, 0)',
        'Weekday(#3/6/2024#, 1)',
        'Weekday(#3/6/2024#, 2)',
        'Weekday(#3/6/2024#, 3)',
        'Weekday(#3/6/2024#, 4)',
        'Weekday(#3/6/2024#, 5)',
        'Weekday(#3/6/2024#, 6)',
        'Weekday(#3/6/2024#, 7)',
        'Weekday(Null)',
        'Weekday("x")',
        'Weekday(0)',
    ],
    'dateserial_edges': [
        'DateSerial(2024, 0, 1)',
        'DateSerial(2024, 13, 1)',
        'DateSerial(2024, 2, 30)',
        'DateSerial(2024, 1, 0)',
        'DateSerial(2024, -1, -1)',
        'DateSerial(99, 1, 1)',
        'DateSerial(29, 1, 1)',
        'DateSerial(30, 1, 1)',
        'DateSerial(100, 1, 1)',
        'DateSerial(10000, 1, 1)',
        'DateSerial(2024.5, 1.5, 1.5)',
        'DateSerial(0, 1, 1)',
        'DateSerial(-1, 1, 1)',
    ],
    'timeserial_edges': [
        'TimeSerial(24, 0, 0)',
        'TimeSerial(-1, 0, 0)',
        'TimeSerial(0, 0, -1)',
        'TimeSerial(0, 1440, 0)',
        'TimeSerial(1.5, 0, 0)',
        'TimeSerial(0, 0, 86400)',
        'TimeSerial(32767, 0, 0)',
        'TimeSerial(-25, 0, 0)',
    ],
    'parts': [
        'Year(#12/29/1899#)',
        'Hour(#12/29/1899 6:00:00 PM#)',
        'Day(-1.5)',
        'Hour(-1.5)',
        'Minute(0.999999)',
        'Second(0.99999999)',
        'Month(Null)',
        'Year("2024-02-03")',
        'Hour("25:00")',
        'Second(#11:59:59 PM#)',
        'Day(2958465)',
        'Year(2958466)',
        'Month(2958466)',
        'Hour(-657435)',
        'Second(3E+6)',
        'CDate(2958466)',
    ],
    'names': [
        'MonthName(0)',
        'MonthName(13)',
        'MonthName(12, True)',
        'WeekdayName(7)',
        'WeekdayName(1, False, 2)',
        'WeekdayName(8)',
        'WeekdayName(1, True, 0)',
        'MonthName(1.5)',
    ],
    'week_one_mondays': [
        'DatePart("ww", #12/29/2003#, 2, 2)',
        'DatePart("ww", #12/29/2003#, 1, 2)',
        'DatePart("ww", #12/30/2024#, 2, 2)',
        'DatePart("ww", #12/30/2024#, 1, 2)',
        'DatePart("ww", #12/31/2007#, 2, 2)',
        'DatePart("ww", #12/31/2007#, 1, 2)',
        'DatePart("ww", #12/28/2026#, 2, 2)',
        'DatePart("ww", #12/28/2026#, 1, 2)',
        'DatePart("ww", #12/31/2026#, 2, 2)',
        'DatePart("ww", #12/31/2026#, 1, 2)',
        'Format(#12/29/2003#, "ww", 2, 2)',
        'Format(#12/31/2024#, "ww", 2, 2)',
    ],
}


def body(group: list[str]) -> str:
    lines = ['Dim s As String, r As Variant', 'On Error Resume Next']
    for call in group:
        lines += ['Err.Clear', 'r = Empty', f'r = {call}',
                  'If Err.Number Then s = s & "e" & Err.Number & ";" Else s = s & TypeName(r) & ":" & r & ";"']
    return '\n'.join([*lines, 'Probe = s'])


CASES = {f'dates_{name}': body(group) for name, group in CALLS.items()}


def macro(name: str) -> str:
    return '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT)


if __name__ == '__main__':
    main()
