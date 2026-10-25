"""A Date's moment, to the second, as oleaut32's VariantTimeToSystemTime gives it.

tests/fixtures/variant_time.json is what scripts/measure_variant_time.py
read from the system's own conversion, which VBA's CStr, Format, Hour,
Minute and Second go through: 7,013 serials across the Date range, ties on
the half second, times either side of midnight and the range's first and
last days. On Windows the same rule is checked against oleaut32 itself on
serials drawn afresh.
"""

from __future__ import annotations

import json
import random
import sys
from pathlib import Path

import pytest

from pyopenvba.exceptions import VBARuntimeError
from pyopenvba.interpreter._values import VBADate

CASES: list[list[object]] = json.loads(
    (Path(__file__).parent / "fixtures" / "variant_time.json").read_text(encoding="utf-8"))["cases"]


def _shown(serial: float) -> str:
    return VBADate(serial).to_datetime().strftime("%Y-%m-%d %H:%M:%S").rjust(19, "0")


def test_every_measured_serial_comes_to_the_systems_moment() -> None:
    missed = [(serial, shown, _shown(float(serial))) for serial, shown in CASES  # type: ignore[arg-type]
              if _shown(float(serial)) != shown]  # type: ignore[arg-type]
    assert not missed, missed[:5]


@pytest.mark.parametrize(("serial", "shown"), [
    (0.5 / 86400, "1899-12-30 00:00:01"),  # a tie rounds up
    (45293 + 1.5 / 86400, "2024-01-02 00:00:01"),  # 1.5 seconds in, as the double holds it, rounds down
    (-0.99999999, "1899-12-31 00:00:00"),  # before day zero the time runs forward, into the next day
    (2958465 + 57599 / 86400, "9999-12-31 15:59:58"),  # the last day's times are cut
    (2958464.9999942128, "9999-12-31 00:00:00"),  # rounded onto the last day's first second
])
def test_named_cases(serial: float, shown: str) -> None:
    assert _shown(serial) == shown


def test_the_first_day_runs_to_just_short_of_minus_657435() -> None:
    # Live VBA: a Date of -657434.99999 is 1/1/100 11:59:59 PM, and -657435 is error 6.
    assert _shown(-657434.99999) == "0100-01-01 23:59:59"
    with pytest.raises(VBARuntimeError) as raised:
        VBADate(-657435.0)
    assert raised.value.number == 6


@pytest.mark.skipif(sys.platform != "win32", reason="oleaut32 is a Windows library")
def test_fresh_serials_against_oleaut32() -> None:
    import ctypes
    from ctypes import wintypes

    class SYSTEMTIME(ctypes.Structure):
        _fields_ = [(name, wintypes.WORD) for name in
                    ("wYear", "wMonth", "wDayOfWeek", "wDay", "wHour", "wMinute", "wSecond", "wMilliseconds")]

    convert = ctypes.WinDLL("oleaut32").VariantTimeToSystemTime  # type: ignore[attr-defined]
    convert.argtypes = [ctypes.c_double, ctypes.POINTER(SYSTEMTIME)]
    pick = random.Random()
    missed: list[tuple[float, str, str]] = []
    for _ in range(20000):
        serial = pick.uniform(-657434.99, 2958465.99)
        if pick.random() < 0.5:
            serial = round(serial * 172800) / 172800  # on a half second, as near as a double comes
        found = SYSTEMTIME()
        assert convert(serial, ctypes.byref(found))
        want = (f"{found.wYear:04}-{found.wMonth:02}-{found.wDay:02} "
                f"{found.wHour:02}:{found.wMinute:02}:{found.wSecond:02}")
        if _shown(serial) != want:
            missed.append((serial, want, _shown(serial)))
    assert not missed, missed[:5]
