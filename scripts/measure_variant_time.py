"""What oleaut32's VariantTimeToSystemTime makes of a spread of Date serials, the conversion VBA's CStr, Format,
Hour, Minute and Second go through.

The serials are drawn from a fixed seed: every part of the Date range, half-second ties at small and large
magnitudes, times a hair either side of midnight before and after day zero, and the first and last days of the
range, where a negative serial's time counts forwards and the last day's times are cut rather than rounded.

    python scripts/measure_variant_time.py

Windows only. Writes tests/fixtures/variant_time.json, which tests/test_vba_date_rounding.py replays.
"""

from __future__ import annotations

import ctypes
import json
import random
from ctypes import wintypes
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "variant_time.json"


class SYSTEMTIME(ctypes.Structure):
    _fields_ = [(name, wintypes.WORD) for name in
                ("wYear", "wMonth", "wDayOfWeek", "wDay", "wHour", "wMinute", "wSecond", "wMilliseconds")]


def native(serial: float) -> str | None:
    """The moment oleaut32 gives, as YYYY-MM-DD HH:MM:SS, or None where it refuses the serial."""
    oleaut = ctypes.WinDLL("oleaut32")
    oleaut.VariantTimeToSystemTime.argtypes = [ctypes.c_double, ctypes.POINTER(SYSTEMTIME)]
    found = SYSTEMTIME()
    if not oleaut.VariantTimeToSystemTime(serial, ctypes.byref(found)):
        return None
    return (f"{found.wYear:04}-{found.wMonth:02}-{found.wDay:02} "
            f"{found.wHour:02}:{found.wMinute:02}:{found.wSecond:02}")


def serials() -> list[float]:
    """The serials measured, the same on every run."""
    pick = random.Random(20261025)
    found = [0.0, -0.0, 1.0, -1.0, 0.5 / 86400, -0.5 / 86400, -0.99999999, 2958465.0, 2958465.99999,
             2958464.9999942128, 2958464.99999537, -657434.0, -657434.99999]
    for _ in range(1500):
        found.append(pick.uniform(-657434.0, 2958465.99))
    for _ in range(1500):
        days = pick.choice([pick.randint(-657433, -1), pick.randint(-60, 60), pick.randint(1, 60000),
                            pick.randint(2900000, 2958464)])
        second = pick.randint(0, 86399)
        fraction = (second + pick.choice([0.5, 0.4999, 0.5001, 0.49999999, 0.50000001])) / 86400.0
        found.append(days - fraction if days < 0 else days + fraction)
    for _ in range(1000):
        days = pick.choice([pick.randint(-657433, -1), pick.randint(-30, 30), pick.randint(1, 2958463)])
        fraction = 1 - pick.choice([pick.uniform(0, 2e-5), pick.uniform(5.7e-6, 5.9e-6), pick.uniform(0, 1e-9)])
        found.append(days - fraction if days < 0 else days + fraction)
    for _ in range(1000):
        found.append(2958465 + pick.random())
        found.append(2958465 - pick.uniform(0, 1e-5))
        found.append(-657434 - pick.random())
    return found


def main() -> None:
    cases = [[serial, shown] for serial in serials() if (shown := native(serial)) is not None]
    # One serial to a line: the file reads, and diffs, a case at a time.
    lines = ",\n".join(json.dumps(case) for case in cases)
    OUT.write_text('{"cases": [\n' + lines + "\n]}\n", encoding="utf-8")
    print(len(cases), "serials written to", OUT)


if __name__ == "__main__":
    main()
