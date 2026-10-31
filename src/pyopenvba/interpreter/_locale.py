"""The Windows Regional format VBA runs under: what CStr, Format and the Format* functions write, and what CDate,
IsDate and the number conversions read.

VBA takes these from the user's Regional format through oleaut32, which is why
oleaut32 called with a locale's LCID answers as VBA does under that format
(scripts/measure_vba_locale.py, tests/fixtures/vba_locale.json). en-US is the
default and what every other measurement in the project was made under.
ja-JP was measured live (tests/fixtures/vba_locale_live.json) and through
oleaut32 with LCID 1041.

The locale in force is a context variable: an interpreter sets it from its
host for the length of each outermost call (``Interpreter.call``), so code
outside a run, and every other host, sees en-US.
"""
from __future__ import annotations

import datetime as _dt
from collections.abc import Generator
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass, field
from typing import Final


@dataclass(frozen=True)
class Era:
    """A Japanese era: its first day, its name, the name's first character and its letter."""

    start: _dt.date
    name: str
    short: str
    letter: str


@dataclass(frozen=True)
class Locale:
    """One Regional format, as VBA sees it."""

    name: str
    lcid: int
    #: Format codes for the named formats and the c, ddddd, dddddd and ttttt tokens.
    short_date: str
    long_date: str
    medium_date: str
    long_time: str
    medium_time: str
    short_time: str
    #: What the AMPM marker writes (AM/PM and A/P write themselves).
    morning: str
    afternoon: str
    #: The named Currency format, FormatCurrency's symbol and digits, and whether
    #: FormatCurrency puts a negative in parentheses by default.
    currency: str
    currency_symbol: str
    currency_digits: int
    currency_parens: bool
    #: Characters CDbl, CCur, IsNumeric and their kin take as a currency sign.
    currency_signs: str
    #: MonthName and WeekdayName (Sunday first).
    months: tuple[str, ...]
    month_abbreviations: tuple[str, ...]
    days: tuple[str, ...]
    day_abbreviations: tuple[str, ...]
    #: Japanese eras, oldest first, for the g, gg, ggg, e and ee codes and for reading era dates.
    eras: tuple[Era, ...] = field(default=())
    #: Whether full-width ASCII forms (U+FF01-U+FF5E) read as their ASCII letters, digits and signs.
    folds_width: bool = False
    #: What Excel's cells show (tests/fixtures/excel_locale_live.json): a built-in format, kept as its number,
    #: shows as the locale's own code for that number, given here by the en-US code (missing: the same).
    builtin_formats: tuple[tuple[str, str], ...] = ()
    #: General as NumberFormatLocal spells it.
    general: str = "General"
    #: What [$-F800] and [$-F400] show a date and a time as: the system's long date and long time.
    system_date: str = "dddd, mmmm d, yyyy"
    system_time: str = "h:mm:ss AM/PM"
    #: The character NumberFormatLocal escapes a literal with.
    escape: str = "\\"
    #: Application.International where the locale answers otherwise than en-US, by index.
    international: tuple[tuple[int, object], ...] = ()


EN_US: Final = Locale(
    name="en-US",
    lcid=1033,
    short_date="m/d/yyyy",
    long_date="dddd, mmmm d, yyyy",
    medium_date="dd-mmm-yy",
    long_time="h:nn:ss AM/PM",
    medium_time="hh:nn AM/PM",
    short_time="hh:nn",
    morning="AM",
    afternoon="PM",
    currency="$#,##0.00;($#,##0.00)",
    currency_symbol="$",
    currency_digits=2,
    currency_parens=True,
    currency_signs="$",
    months=("January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
            "November", "December"),
    month_abbreviations=("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"),
    days=("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"),
    day_abbreviations=("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"),
)

#: ja-JP, from Windows' locale data (GetLocaleInfoEx, LCID 1041) and its era
#: table, except that oleaut32 starts Meiji on 23 October 1868, the day it
#: began, where the table says 1 January (scripts/measure_vba_locale.py).
JA_JP: Final = Locale(
    name="ja-JP",
    lcid=1041,
    short_date="yyyy/mm/dd",
    long_date='yyyy"年"m"月"d"日"',
    medium_date="yy-mm-dd",
    long_time="h:nn:ss",
    medium_time="hh:nn AMPM",
    short_time="hh:nn",
    morning="午前",
    afternoon="午後",
    currency="¥#,##0;-¥#,##0",
    currency_symbol="¥",
    currency_digits=0,
    currency_parens=False,
    # The yen sign, its full-width form, and the backslash, which is the yen sign in Japanese code pages.
    currency_signs="¥￥\\",
    months=tuple(f"{month}月" for month in range(1, 13)),
    month_abbreviations=tuple(str(month) for month in range(1, 13)),
    days=("日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"),
    day_abbreviations=("日", "月", "火", "水", "木", "金", "土"),
    eras=(
        Era(_dt.date(1868, 10, 23), "明治", "明", "M"),
        Era(_dt.date(1912, 7, 30), "大正", "大", "T"),
        Era(_dt.date(1926, 12, 25), "昭和", "昭", "S"),
        Era(_dt.date(1989, 1, 8), "平成", "平", "H"),
        Era(_dt.date(2019, 5, 1), "令和", "令", "R"),
    ),
    folds_width=True,
    builtin_formats=(
        ("#,##0_);(#,##0)", "#,##0;-#,##0"),
        ("#,##0_);[Red](#,##0)", "#,##0;[Red]-#,##0"),
        ("#,##0.00_);(#,##0.00)", "#,##0.00;-#,##0.00"),
        ("#,##0.00_);[Red](#,##0.00)", "#,##0.00;[Red]-#,##0.00"),
        ("$#,##0_);($#,##0)", "$#,##0;-$#,##0"),
        ("$#,##0_);[Red]($#,##0)", "$#,##0;[Red]-$#,##0"),
        ("$#,##0.00_);($#,##0.00)", "$#,##0.00;-$#,##0.00"),
        ("$#,##0.00_);[Red]($#,##0.00)", "$#,##0.00;[Red]-$#,##0.00"),
        ("m/d/yyyy", "yyyy/mm/dd"),
        ("d-mmm-yy", "dd-mmm-yy"),
        ("d-mmm", "dd-mmm"),
        ("m/d/yyyy h:mm", "yyyy/mm/dd h:mm"),
        ('_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'),
        ('_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)', '_-$* #,##0_-;-$* #,##0_-;_-$* "-"_-;_-@_-'),
        ('_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-'),
        ('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
         '_-$* #,##0.00_-;-$* #,##0.00_-;_-$* "-"??_-;_-@_-'),
    ),
    general="G/標準",
    system_date='yyyy"年"m"月"d"日"',
    system_time="h:mm:ss",
    escape="!",
    international=((2, 81.0), (25, "¥"), (26, "G/標準"), (27, 0.0), (28, 1.0), (32, 2.0), (33, True), (35, True),
                   (38, True), (41, True), (42, True)),
)

LOCALES: Final = {locale.name.lower(): locale for locale in (EN_US, JA_JP)}

_CURRENT: ContextVar[Locale] = ContextVar("vba_locale", default=EN_US)


def locale_named(name: str) -> Locale:
    """The locale for a name such as ``"ja-JP"``; any other name is a ValueError."""
    found = LOCALES.get(name.lower())
    if found is None:
        raise ValueError(f"unsupported locale {name!r}; supported: {', '.join(sorted(LOCALES))}")
    return found


def current() -> Locale:
    """The locale VBA runs under now."""
    return _CURRENT.get()


@contextmanager
def active(locale: Locale) -> Generator[None]:
    """Run the block under ``locale``."""
    token = _CURRENT.set(locale)
    try:
        yield
    finally:
        _CURRENT.reset(token)


def era_of(locale: Locale, serial: float, year: int) -> tuple[Era, int] | None:
    """The era a Date serial falls in and the year of it (from the Date's Gregorian ``year``), or None before
    the first era.

    oleaut32 compares serials, and a time on a day before 30 December 1899
    makes the serial smaller: 23 October 1868 at noon is before Meiji began.
    """
    for era in reversed(locale.eras):
        if serial >= (era.start - _EPOCH).days:
            # The era's first year is the Gregorian year it starts in.
            return era, year - era.start.year + 1
    return None


_EPOCH: Final = _dt.date(1899, 12, 30)


def fold_width(text: str) -> str:
    """Full-width ASCII forms and the ideographic space as their ASCII characters. The full-width
    backslash stays itself: it is no currency sign (tests/fixtures/vba_locale.json)."""
    return "".join(chr(ord(char) - 0xFEE0) if "！" <= char <= "～" and char != "＼" else
                   (" " if char == "　" else char) for char in text)
