"""The expression functions the executor offers.

Every answer here was measured against DAO (see the function gate in
tests/test_live_access_engine_gate.py); these run the same ground truth
without an engine present.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access._format import format_value, partition
from pyopenvba.access_read import AccessError

TEMPLATE = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"
WHEN = dt.datetime(2024, 3, 14, 15, 9, 26)


def _value(expression: str) -> object:
    db = AccessDatabase(TEMPLATE)
    db.execute("CREATE TABLE One (Id LONG, T TEXT(40), D DATETIME, X DOUBLE)")
    db.execute("INSERT INTO One (Id, T, D, X) VALUES (1, ' Hello World ', #3/14/2024 15:09:26#, 2.5)")
    rows = db.execute(f"SELECT ({expression}) AS V FROM One")
    assert isinstance(rows, list)
    return rows[0]["V"]


@pytest.mark.parametrize(
    ("expression", "expected"),
    [
        ("Replace(T, 'l', 'L')", " HeLLo WorLd "),
        ("Space(3) & 'x'", "   x"),
        ("String(4, 'z')", "zzzz"),
        ("StrComp('a', 'b')", -1),
        ("StrReverse('abc')", "cba"),
        ("Asc('A')", 65),
        ("Chr(66)", "B"),
        ("Sgn(-3)", -1),
        ("Sqr(9)", 3.0),
        ("Fix(-2.7)", -2),
        ("Int(-2.7)", -3),
        ("Val('12.5abc')", 12.5),
        ("Str(17)", " 17"),
        ("Str(0.125)", " .125"),
        ("Hex(255)", "FF"),
        ("Oct(8)", "10"),
        ("CByte(5)", 5),
        ("Choose(2, 'a', 'b', 'c')", "b"),
        ("Choose(9, 'a')", None),
        ("Switch(1 > 2, 'no', 2 > 1, 'yes')", "yes"),
        ("Switch(1 > 2, 'no')", None),
        # A truth value comes back the way Jet writes one.
        ("IsNull(T)", 0),
        ("IsNull(Null)", -1),
        ("IsNumeric('12')", -1),
        ("IsNumeric(T)", 0),
        ("IsDate(D)", -1),
        ("2 > 1", -1),
        ("Not (2 > 1)", 0),
    ],
)
def test_a_function_answers_what_the_engine_answers(expression: str, expected: object) -> None:
    assert _value(expression) == expected


@pytest.mark.parametrize(
    ("expression", "expected"),
    [
        ("DateAdd('d', 5, D)", dt.datetime(2024, 3, 19, 15, 9, 26)),
        ("DateAdd('m', 1, #1/31/2020#)", dt.datetime(2020, 2, 29)),
        ("DateAdd('yyyy', -1, D)", dt.datetime(2023, 3, 14, 15, 9, 26)),
        ("DateDiff('d', #1/1/2024#, D)", 73),
        ("DateDiff('m', #12/31/2023#, D)", 3),
        ("DateDiff('yyyy', #12/31/2023#, D)", 1),
        ("DatePart('q', D)", 1),
        ("DatePart('y', D)", 74),
        ("DateSerial(2024, 2, 30)", dt.datetime(2024, 3, 1)),
        ("TimeSerial(13, 5, 6)", dt.datetime(1899, 12, 30, 13, 5, 6)),
        ("Weekday(D)", 5),
        ("WeekdayName(3)", "Tuesday"),
        ("MonthName(4)", "April"),
        ("DateValue(D)", dt.datetime(2024, 3, 14)),
        ("TimeValue(D)", dt.datetime(1899, 12, 30, 15, 9, 26)),
        ("CDate('2024-03-14')", dt.datetime(2024, 3, 14)),
    ],
)
def test_a_date_function_answers_what_the_engine_answers(expression: str, expected: object) -> None:
    assert _value(expression) == expected


@pytest.mark.parametrize(
    ("value", "pattern", "expected"),
    [
        (WHEN, "yyyy-mm-dd", "2024-03-14"),
        (WHEN, "yyyy-mm-dd hh:nn:ss", "2024-03-14 15:09:26"),
        (WHEN, "Short Date", "3/14/2024"),
        (WHEN, "Long Date", "Thursday, March 14, 2024"),
        (WHEN, "Medium Date", "14-Mar-24"),
        (WHEN, "Short Time", "15:09"),
        (WHEN, "Long Time", "3:09:26 PM"),
        (WHEN, "Medium Time", "03:09 PM"),
        (WHEN, "General Date", "3/14/2024 3:09:26 PM"),
        (dt.datetime(2020, 1, 31), "General Date", "1/31/2020"),
        (WHEN, "dddd", "Thursday"),
        (WHEN, "ddd", "Thu"),
        (WHEN, "mmm", "Mar"),
        (WHEN, "q", "1"),
        (WHEN, "y", "74"),
        (WHEN, "h:nn AM/PM", "3:09 PM"),
        (2.5, "0.00", "2.50"),
        (2.5, "0", "3"),
        (1234567.891, "#,##0.00", "1,234,567.89"),
        (1234567.891, "Standard", "1,234,567.89"),
        (1234567.891, "Currency", "$1,234,567.89"),
        (-5.5, "Currency", "($5.50)"),
        (0.5, "Percent", "50.00%"),
        (0.125, "0.0%", "12.5%"),
        (17, "000", "017"),
        (-4, "000", "-004"),
        (1234.5, "Scientific", "1.23E+03"),
        (17, "General Number", "17"),
        (17, "Yes/No", "Yes"),
        (0, "Yes/No", "No"),
        (0, "True/False", "False"),
        (0, "On/Off", "Off"),
        (-3, "0.00;(0.00)", "(3.00)"),
        (0, "0.00;(0.00);zero", "zero"),
        (" Hello ", ">", " HELLO "),
        (" Hello ", "<", " hello "),
        ("abc", "@", "abc"),
        ("", "@", ""),
        (None, "0.00", ""),
    ],
)
def test_format_writes_what_the_engine_writes(value: object, pattern: str, expected: str) -> None:
    assert format_value(value, pattern) == expected


@pytest.mark.parametrize(
    ("number", "start", "stop", "interval", "expected"),
    [
        (17, 0, 100, 10, " 10: 19"),
        (-4, 0, 100, 10, "   : -1"),
        (0, 0, 100, 10, "  0:  9"),
        (17, 1, 50, 5, "16:20"),
        (5, 1, 50, 5, " 1: 5"),
        (17, 0, 10, 5, "11:  "),
        (120, 0, 100, 10, "101:   "),
    ],
)
def test_partition_pads_its_bounds_as_the_engine_does(
    number: int, start: int, stop: int, interval: int, expected: str
) -> None:
    assert partition(number, start, stop, interval) == expected


def test_partition_needs_an_interval() -> None:
    with pytest.raises(AccessError, match="interval"):
        partition(1, 0, 10, 0)


def test_a_column_of_its_own_keeps_its_type() -> None:
    """A Boolean column read straight out comes back as a Boolean, where a
    computed truth value comes back as -1 or 0."""
    db = AccessDatabase(TEMPLATE)
    db.execute("CREATE TABLE Flags (Id LONG, Yes BIT)")
    db.execute("INSERT INTO Flags (Id, Yes) VALUES (1, True)")
    rows = db.execute("SELECT Flags.Yes AS A, (Flags.Yes = True) AS B FROM Flags")
    assert isinstance(rows, list)
    assert rows[0] == {"A": True, "B": -1}


def test_an_aggregate_over_truth_values_answers_with_numbers() -> None:
    """Max of True, False, True is 0 and Min is -1, because the engine
    holds them as -1 and 0; a bare column still reads as a Boolean and
    still sorts True first."""
    db = AccessDatabase(TEMPLATE)
    db.execute("CREATE TABLE Flags (Id LONG, B BIT)")
    for i, value in ((1, "True"), (2, "False"), (3, "True")):
        db.execute(f"INSERT INTO Flags (Id, B) VALUES ({i}, {value})")
    rows = db.execute("SELECT Max(B) AS A, Min(B) AS C, Sum(B) AS D, Count(B) AS E, First(B) AS G FROM Flags")
    assert rows == [{"A": 0, "C": -1, "D": -2, "E": 3, "G": -1}]
    grouped = db.execute("SELECT Flags.B, Count(*) AS N FROM Flags GROUP BY Flags.B ORDER BY Flags.B")
    assert grouped == [{"B": True, "N": 2}, {"B": False, "N": 1}]


def test_an_unknown_function_says_so() -> None:
    with pytest.raises(AccessError, match="function Frobnicate is not available"):
        _value("Frobnicate(1)")
