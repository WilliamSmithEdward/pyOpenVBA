"""Range.AutoFilter and the AutoFilter, Filters and Filter objects, as Excel filters.

Measured in live Excel (scripts/measure_autofilter.py):

- AutoFilter with no field turns the sheet's filter on over the range --
  a single cell's current region, whole columns cut to the used rows --
  or, where the sheet has one, off, showing every row again. With a
  field it filters that column of the sheet's filter, whatever range it
  was called on, turning the filter on first where there is none; a
  field outside the range is error 1004, and nothing changes. The first
  filter on a sheet leaves a hidden name, _FilterDatabase, which stays
  after the filter goes.
- Filtering sets every row under the header, shown where it meets each
  column's criteria and hidden otherwise; editing a cell afterwards
  changes nothing until ApplyFilter. ShowAllData shows the rows and
  clears the criteria, error 1004 when none are set.
- A criterion with = or no operator matches the text a cell shows,
  ignoring case, with * and ? wildcards and ~ before either to mean it:
  "1" matches 1 but not 1 shown as 1.00, and a wildcard matches text
  only. = alone matches blank cells, <> alone the rest. <> with a number
  is not equal in value, and with text the opposite of =. >, >=, < and
  <= with a number compare numbers, dates by their serial, and with text
  compare text as Sort orders it; nothing else passes. Criteria read
  back with the operator in front, a date given with >, <, or <> as its
  serial, and spaces around trimmed.
- xlAnd and xlOr join two criteria; Criteria2 alone joins a blank first
  one with xlAnd. A list with xlFilterValues is one criterion with one
  value, xlOr with two, and otherwise stays a list, sorted, whose values
  match exactly; it returns the field rather than True.
- xlTop10Items and the rest keep the numbers at or past the given one's
  place, reading back as >= or <= that number; a percentage counts
  down to whole items, at least one. xlFilterDynamic with
  xlFilterAboveAverage or xlFilterBelowAverage keeps numbers above or
  below their mean. Both are error 1004 over a column with an error.

The file (scripts/measure_autofilter_file.py) holds a plain = criterion,
two joined by xlOr and a list as a filters list of values -- b~* as the
b* it means, blanks as blank="1" -- and anything with a wildcard, <>,
>, < or xlAnd as customFilters; the threshold of a top filter and the
mean of an average one go with them. Excel reads every one back as it
wrote it.

A filter on a cell's colour or its font's (scripts/measure_colour_filters.py)
shows the rows whose cell shows the colour as a colour sort finds it, a
conditional format's included; xlFilterNoFill the cells with no fill, and
xlFilterAutomaticFontColor those in the automatic font. Criteria1 reads a
cell colour as an Interior, white with no pattern for no fill, and a
font's as its colour. The file keeps the colour in a dxf of its own,
made as the filter is, which a filter on the column again or ShowAllData
lets go.

xlFilterValues with Criteria2 filters by dates (scripts/measure_date_filters.py):
pairs of a level, 0 for a year to 5 for a second, and a date that level
of it names, with the values of an array Criteria1 beside them; a single
Criteria1 is passed over, and a year with no date to name it is error
1004, the filter turned on all the same. A cell passes in a date format,
inside a group. The filter reads back as those values alone do. The file
keeps it in an Excel 2017 extension of the column's filter, the values
then the dates, the latest year first and within one the earliest
moment; a file's filters element of dates reads back as xlFilterValues
and is saved as it was.

xlFilterDynamic with xlFilterToday to xlFilterYearToDate keeps the date
cells in a period counted from today, a week running from Sunday, and
xlFilterAllDatesInPeriodQuarter1 to ...December a quarter or a month of
any year (scripts/measure_dynamic_filters.py). It reads back as
xlFilterDynamic and its Criteria1; the file keeps a period's first serial
and the one after its last. A filter read from a file applies to the
period counted from the model's today; what Excel does with one opened on
another day is not measured.

xlFilterIcon with an Icon as Criteria1 keeps the cells a conditional
format's icon set shows that icon, its set and place there and not one
alike in another set, and xlFilterNoIcon those it shows none
(scripts/measure_icon_filters.py, _icon_sets.shown_icon). It reads back
the Icon, or Empty, and the file keeps the set and the place from 0, or
for no icon the column's set alone.

Text is matched as Excel matches it beyond ASCII too
(scripts/measure_filter_text.py): an = with no wildcard, and each value
of a list, keeps the text Windows' word sort cannot tell from it, case
aside, so =aesop keeps the spelling with the ash and =strasse the one
with a sharp s; > and < order text by word sort; and a wildcard pattern
matches one character at a time, each by its word-sort key, so ae* leaves
out the ash's spelling and k* keeps the Kelvin sign's.

What is not modelled reports itself: AutoFilter.Sort, a hidden drop-down,
and text beyond Latin, Greek and Cyrillic where it is compared.

A table has a filter of its own, which a range inside it reaches; the
sheet's members reach whichever filter the active cell is in. Both are
_table_filters's.
"""

from __future__ import annotations

import datetime as _dt
import math
import re
import uuid
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, parse_reference, quote_sheet
from pyopenvba.apps.excel._model import ExcelObject, NameEntry, Range
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.formula._display import UndisplayableError, shown
from pyopenvba.formula._values import ExcelError, number_text
from pyopenvba.interpreter._objects import VBACollection, member, method
from pyopenvba.interpreter._values import (EMPTY, MISSING, NOTHING, VBAArray, VBADate, VBAInt, error, to_bool,
                                           to_integer, to_number, to_text)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._editing import CellShift
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.apps.excel._tables import Table

AND, OR, TOP_ITEMS, BOTTOM_ITEMS, TOP_PERCENT, BOTTOM_PERCENT, VALUES = 1, 2, 3, 4, 5, 6, 7
CELL_COLOR, FONT_COLOR, ICON, DYNAMIC, NO_FILL, AUTOMATIC_FONT, NO_ICON = 8, 9, 10, 11, 12, 13, 14
ABOVE_AVERAGE, BELOW_AVERAGE = 33, 34
WHITE, PATTERN_SOLID, PATTERN_NONE = 0xFFFFFF, 1, -4142
_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
#: The filters on a cell's colour or its font's, and which of them look at the font.
COLOURS = frozenset({CELL_COLOR, FONT_COLOR, NO_FILL, AUTOMATIC_FONT})
_FONTS = frozenset({FONT_COLOR, AUTOMATIC_FONT})
_OPERATORS = ("<>", ">=", "<=", "=", ">", "<")
#: The file's names for the operators a customFilter can hold; = has none.
_FILE_OPERATORS = {"<>": "notEqual", ">": "greaterThan", ">=": "greaterThanOrEqual", "<": "lessThan",
                   "<=": "lessThanOrEqual"}


# --- what a cell shows the filter --------------------------------------------------------------------


@dataclass(frozen=True)
class Seen:
    """A cell as a criterion looks at it, and where it is, which a colour filter looks at."""

    shown: str
    number: float | None
    text: str | None
    error: bool
    sheet: Worksheet | None = None
    row: int = 0
    column: int = 0


def _seen(sheet: Worksheet, row: int, column: int) -> Seen:
    from dataclasses import replace

    return replace(_value_seen(sheet, row, column), sheet=sheet, row=row, column=column)


def _value_seen(sheet: Worksheet, row: int, column: int) -> Seen:
    cell = sheet.cells_.get((row, column))
    if cell is None:
        return Seen("", None, None, False)
    value = sheet.book.calculator.value_of(sheet.name, row, column) if cell.formula else cell.value
    if isinstance(value, VBADate):
        value = value.serial
    if value is EMPTY:
        return Seen("", None, None, False)
    if isinstance(value, ExcelError):
        return Seen(value.name, None, None, True)
    if isinstance(value, bool):
        return Seen("TRUE" if value else "FALSE", None, None, False)
    if isinstance(value, (int, float)):
        try:
            text, _ = shown(float(value), cell.number_format, epoch_1904=sheet.book.epoch_1904)
        except UndisplayableError:
            text = "#"
        return Seen(text, float(value), None, False)
    text = str(value)
    return Seen(text, None, text, False)


# --- criteria -----------------------------------------------------------------------------------------


@dataclass(frozen=True)
class Criterion:
    """One criterion as typed: its operator, what follows it, that read as a number, and how it reads back."""

    operator: str
    operand: str
    number: float | None
    stored: str

    @property
    def plain(self) -> bool:
        """An = with no wildcard in it, which the file keeps as a value in a list."""
        return self.operator == "=" and not _wild(self.operand)

    def matches(self, seen: Seen) -> bool:
        if self.operator == "=":
            return seen.shown == "" if self.operand == "" else _like(seen, self.operand)
        if self.operator == "<>":
            if self.operand == "":
                return seen.shown != ""
            if self.number is not None:
                return seen.number is None or seen.number != self.number
            return not _like(seen, self.operand)
        if self.number is not None:
            return seen.number is not None and _compare(seen.number, self.operator, self.number)
        if seen.text is None:
            return False
        return _compare(ordered(seen.text), self.operator, ordered(self.operand))


def _criterion(value: object, *, epoch_1904: bool = False) -> Criterion:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        text = number_text(float(value), formula=True)
    else:
        text = to_text(value)
    body = text.strip(" ")
    operator = next((one for one in _OPERATORS if body.startswith(one)), "")
    return made_criterion(operator or "=", body[len(operator):].strip(" "), epoch_1904=epoch_1904)


def made_criterion(operator: str, operand: str, *, epoch_1904: bool = False) -> Criterion:
    """A criterion from its operator and operand, reading the operand as a number where it is one.

    A number reads back spelled afresh, a date or time as its serial on
    the workbook's calendar: >2.50 is >2.5, and >2/15/2024 is >43875 on
    the 1904 one (tests/fixtures/date1904/paths.json).
    """
    from pyopenvba.apps.excel._typing import on_1904, typed_text

    number: float | None = None
    stored = operand
    if operator != "=" and operand:
        typed = typed_text(operand)
        if epoch_1904:
            typed = on_1904(typed, operand)
        if isinstance(typed.value, float):
            number = typed.value
            stored = number_text(number, formula=True)
    return Criterion(operator, operand, number, operator + stored)


def _compare(left: object, operator: str, right: object) -> bool:
    if operator == ">":
        return left > right  # type: ignore[operator]
    if operator == ">=":
        return left >= right  # type: ignore[operator]
    if operator == "<":
        return left < right  # type: ignore[operator]
    return left <= right  # type: ignore[operator]


def ordered(text: str) -> bytes:
    """What a filter tells text by, case aside: Windows' word-sort key, which orders text for > and < and lists the
    file's values, and is one for the text an = with no wildcard keeps, so =aesop keeps Aesop with the ash and
    =strasse keeps the sharp s while =resume leaves out the accented spelling (tests/fixtures/filter_text.json)."""
    from pyopenvba.apps.excel._sort import text_key

    try:
        return text_key(text, False)
    except VBAUnsupportedError:
        raise VBAUnsupportedError("filtering text with characters beyond Latin, Greek and Cyrillic, their marks and "
                                  "punctuation, or that begins with a combining mark, is not implemented") from None


def letter(char: str) -> bytes | str:
    """What a filter's wildcard pattern tells a character by: its word-sort key, case aside, so a Kelvin sign is a K,
    a final sigma a sigma and the capital sharp s a sharp s, while a dotless i is no I and the ash no a
    (tests/fixtures/filter_text.json, letters); a combining accent by itself."""
    from pyopenvba.apps.excel._sort_weights import WEIGHTS

    found = WEIGHTS.get(ord(char))
    if found is not None and not found[0] and found[1]:
        return char
    return ordered(char)


def _wild(pattern: str) -> bool:
    return re.search(r"(?<!~)[*?]", pattern) is not None


def _literal(pattern: str) -> str:
    """What a pattern with no live wildcard means: ~* and ~? are * and ?."""
    return re.sub(r"~([*?])", r"\1", pattern)


def _escaped(value: str) -> str:
    """A value from the file's list as a criterion reads back: its * and ? escaped."""
    return re.sub(r"([*?])", r"~\1", value)


def _like(seen: Seen, pattern: str) -> bool:
    """Whether what a cell shows matches a criterion's text: any case, * and ? wild, ~ before one to mean it.

    A pattern with a wildcard in it matches text only, never a number however it shows, one character at a time
    (:func:`letter`); one without is the text its word-sort key is (:func:`ordered`).
    """
    from pyopenvba.apps.excel._filter_wildcards import matches, matches_by

    if _wild(pattern) and seen.text is None:
        return False
    try:
        if not _wild(pattern):
            return ordered(seen.shown) == ordered(_literal(pattern))
        return matches_by(seen.shown, pattern, letter, escape_tilde=False)
    except VBAUnsupportedError:
        # Text past the word sort's weights, a check mark or an emoji, is matched character for character in any
        # case where no letter past ASCII is in it, as before the weights were measured.
        if any(ord(char) > 126 and char.isalpha() for char in seen.shown + pattern):
            raise
        return matches(seen.shown, pattern, escape_tilde=False)


# --- the file's spelling -------------------------------------------------------------------------------


def xml_attribute(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def file_number(value: float) -> str:
    """A number as the file writes it: as short as reads back exactly."""
    text = repr(value)
    return text[:-2] if text.endswith(".0") else text


def _values_xml(values: list[str], blank: bool) -> str:
    flag = ' blank="1"' if blank else ""
    if not values:
        return f"<filters{flag}/>"
    return f"<filters{flag}>" + "".join(f'<filter val="{xml_attribute(value)}"/>' for value in values) + "</filters>"


def _custom_xml(criterion: Criterion) -> str:
    if criterion.operator == "=":
        return f'<customFilter val="{xml_attribute(criterion.operand)}"/>'
    if criterion.operator == "<>" and criterion.operand == "":
        return '<customFilter operator="notEqual" val=" "/>'
    value = file_number(criterion.number) if criterion.number is not None else criterion.operand
    return f'<customFilter operator="{_FILE_OPERATORS[criterion.operator]}" val="{xml_attribute(value)}"/>'


def _criteria_xml(criteria: list[Criterion], *, joined: int) -> str:
    """The filterColumn content for one criterion, or two joined: plain ones as a list, the rest custom."""
    if joined != AND and all(one.plain for one in criteria):
        values = [_literal(one.operand) for one in criteria if one.operand != ""]
        return _values_xml(values, any(one.operand == "" for one in criteria))
    both = ' and="1"' if joined == AND else ""
    return f"<customFilters{both}>" + "".join(_custom_xml(one) for one in criteria) + "</customFilters>"


# --- a column's filter ---------------------------------------------------------------------------------


@dataclass
class FieldFilter:
    """The criteria on one column: what the Filter object reads back, the test a row's cell must pass, and the
    filterColumn the file holds for them -- its content, and any attributes besides colId."""

    operator: int
    criteria1: object
    criteria2: object
    test: Callable[[Seen], bool]
    count: int = 1
    xml: str = ""
    attributes: str = ""
    #: The dxf a colour filter keeps its colour in, by its place in the stylesheet, which a save renumbers.
    dxf: int | None = None

    def markup(self) -> str:
        """What the filterColumn holds: the criteria as read or made, a colour filter naming its dxf's place."""
        if self.dxf is None:
            return self.xml
        return re.sub(r'\bdxfId="\d+"', f'dxfId="{self.dxf}"', self.xml, count=1)


def _one(criterion: Criterion) -> FieldFilter:
    return FieldFilter(0, criterion.stored, MISSING, criterion.matches, xml=_criteria_xml([criterion], joined=0))


def _two(first: Criterion, second: Criterion, joined: int) -> FieldFilter:
    xml = _criteria_xml([first, second], joined=joined)
    if joined == OR:
        return FieldFilter(OR, first.stored, second.stored, lambda seen: first.matches(seen) or second.matches(seen),
                           2, xml)
    return FieldFilter(AND, first.stored, second.stored, lambda seen: first.matches(seen) and second.matches(seen),
                       2, xml)


def _field(sheet: Worksheet, area: Area, index: int, criteria1: object, operator: object,
           criteria2: object) -> FieldFilter | None:
    """What AutoFilter's arguments set on a column: None where they clear it."""
    kind = 0 if operator is MISSING else int(to_integer(operator, "Long"))
    if criteria1 is MISSING and criteria2 is MISSING and kind in (0, AND, OR):
        return None
    if kind == VALUES:
        if criteria2 is not MISSING:
            return _dated(sheet, criteria1, criteria2)
        return _values(criteria1)
    if kind in (TOP_ITEMS, BOTTOM_ITEMS, TOP_PERCENT, BOTTOM_PERCENT):
        return _top(sheet, area, index, kind, criteria1)
    if kind == DYNAMIC:
        return _dynamic(sheet, area, index, criteria1)
    if kind in COLOURS:
        return _coloured(sheet, kind, criteria1)
    if kind in (ICON, NO_ICON):
        return _iconned(sheet, area, index, kind, criteria1)
    if kind not in (0, AND, OR):
        raise VBAUnsupportedError(f"AutoFilter with Operator {kind} is not implemented")
    epoch_1904 = sheet.book.epoch_1904
    if criteria2 is MISSING:
        return _one(_criterion(criteria1, epoch_1904=epoch_1904))
    first = _criterion("=" if criteria1 is MISSING else criteria1, epoch_1904=epoch_1904)
    return _two(first, _criterion(criteria2, epoch_1904=epoch_1904), OR if kind == OR else AND)


def _values(criteria1: object) -> FieldFilter:
    items = [to_text(item) for item in (criteria1.elements() if isinstance(criteria1, VBAArray) else [criteria1])]
    items.sort(key=ordered)
    if len(items) == 1:
        return _one(_criterion(items[0]))
    if len(items) == 2:
        return _two(_criterion("=" + items[0]), _criterion("=" + items[1]), OR)
    return _listed([item for item in items if item != ""], "" in items)


def _listed(values: list[str], blank: bool) -> FieldFilter:
    """A list of three or more values, blank among them where ``blank``, each matched exactly.

    Criteria1 reads the values back sorted, with the blank after them.
    """
    wanted = {value.lower() for value in values}
    shown_items = values + ([""] if blank else [])
    stored = VBAArray([(1, len(shown_items))], items=["=" + item for item in shown_items])
    return FieldFilter(VALUES, stored, MISSING, lambda seen: seen.shown.lower() in wanted or (blank and seen.shown == ""),
                       len(shown_items), _values_xml(values, blank))


def _numbers(sheet: Worksheet, area: Area, index: int) -> list[float]:
    """The numbers in a column under the header; an error among them is error 1004."""
    found: list[float] = []
    for row in range(area.top + 1, area.bottom + 1):
        seen = _seen(sheet, row, area.left + index - 1)
        if seen.error:
            raise error(1004, "AutoFilter cannot rank or average a column with an error in it")
        if seen.number is not None:
            found.append(seen.number)
    if not found:
        raise VBAUnsupportedError("AutoFilter's top, bottom or average filters on a column with no numbers are not "
                                  "implemented")
    return found


def _top(sheet: Worksheet, area: Area, index: int, kind: int, criteria1: object) -> FieldFilter:
    numbers = _numbers(sheet, area, index)
    percent = kind in (TOP_PERCENT, BOTTOM_PERCENT)
    given = float(to_number(criteria1))
    if percent:
        if given <= 0:
            raise error(1004, "AutoFilter needs a percentage above 0")
        count = max(1, math.floor(len(numbers) * given / 100))
    else:
        count = int(to_integer(criteria1, "Long"))
        if count < 1:
            raise error(1004, "AutoFilter needs at least one item")
        given = float(count)
    top = kind in (TOP_ITEMS, TOP_PERCENT)
    threshold = sorted(numbers, reverse=top)[min(count, len(numbers)) - 1]
    return _ranked(kind, given, threshold)


def _ranked(kind: int, given: float, threshold: float) -> FieldFilter:
    top = kind in (TOP_ITEMS, TOP_PERCENT)
    flags = ("" if top else ' top="0"') + (' percent="1"' if kind in (TOP_PERCENT, BOTTOM_PERCENT) else "")
    xml = f'<top10{flags} val="{file_number(given)}" filterVal="{file_number(threshold)}"/>'
    stored = (">=" if top else "<=") + number_text(threshold, formula=True)
    if top:
        return FieldFilter(kind, stored, MISSING, lambda seen: seen.number is not None and seen.number >= threshold,
                           xml=xml)
    return FieldFilter(kind, stored, MISSING, lambda seen: seen.number is not None and seen.number <= threshold, xml=xml)


def _dynamic(sheet: Worksheet, area: Area, index: int, criteria1: object) -> FieldFilter:
    which = int(to_integer(criteria1, "Long"))
    if 1 <= which <= len(PERIODS):
        return _period(sheet, which)
    if which not in (ABOVE_AVERAGE, BELOW_AVERAGE):
        raise VBAUnsupportedError(f"AutoFilter's dynamic filter {which} is not implemented")
    numbers = _numbers(sheet, area, index)
    return _averaged(which, sum(numbers) / len(numbers))


def _averaged(which: int, mean: float) -> FieldFilter:
    kind = "aboveAverage" if which == ABOVE_AVERAGE else "belowAverage"
    xml = f'<dynamicFilter type="{kind}" val="{file_number(mean)}"/>'
    if which == ABOVE_AVERAGE:
        return FieldFilter(DYNAMIC, VBAInt(which, "Long"), MISSING,
                           lambda seen: seen.number is not None and seen.number > mean, xml=xml)
    return FieldFilter(DYNAMIC, VBAInt(which, "Long"), MISSING,
                       lambda seen: seen.number is not None and seen.number < mean, xml=xml)


#: The parts of a moment a date filter's levels name, from the year at 0 to the second at 5.
_DATE_PARTS = ("year", "month", "day", "hour", "minute", "second")
#: The Excel 2017 extension a filter column keeps its dates in, as Excel writes them, round its filters.
_DATES_HEAD = ('<extLst><ext uri="{1AD28BCE-077C-4C59-8B6E-1921CE8616D4}" '
               'xmlns:xlrd2="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2">'
               '<xlrd2:filterColumn><xlrd2:filters>')
_DATES_TAIL = "</xlrd2:filters></xlrd2:filterColumn></ext></extLst>"


def _dated(sheet: Worksheet, criteria1: object, criteria2: object) -> FieldFilter:
    """xlFilterValues with dates as Criteria2, in pairs of a level, 0 for a year to 5 for a second, and a date
    that level of it names, with the values of an array Criteria1 beside them; a single Criteria1 is passed over
    (tests/fixtures/date_filters/)."""
    from pyopenvba.interpreter._values import parse_date_text

    if not isinstance(criteria2, VBAArray) or len(criteria2.elements()) % 2:
        raise VBAUnsupportedError("AutoFilter with a Criteria2 that is not pairs of a level and a date is not "
                                  "implemented")
    given = criteria2.elements()
    groups: list[tuple[int, ...]] = []
    for level, text in zip(given[::2], given[1::2]):
        depth = int(to_integer(level, "Long"))
        if not 0 <= depth < len(_DATE_PARTS):
            raise VBAUnsupportedError(f"AutoFilter on dates at level {depth} is not implemented")
        # A year needs a date to name it: "2025" alone is error 1004.
        parsed = parse_date_text(to_text(text))
        if parsed is None:
            raise error(1004, "AutoFilter method of Range class failed")
        moment = parsed.to_datetime()
        groups.append((moment.year, moment.month, moment.day, moment.hour, moment.minute, moment.second)[:depth + 1])
    words: list[str] = []
    listed: FieldFilter | None = None
    if isinstance(criteria1, VBAArray):
        words = sorted((to_text(item) for item in criteria1.elements()), key=ordered)
        if "" in words:
            raise VBAUnsupportedError("AutoFilter by dates beside blank cells is not implemented")
        listed = _values(criteria1)
    body = "".join(f'<xlrd2:filter val="{xml_attribute(word)}"/>' for word in words) + "".join(
        "<xlrd2:dateGroupItem" + "".join(f' {name}="{part}"' for name, part in zip(_DATE_PARTS, group))
        + f' dateTimeGrouping="{_DATE_PARTS[len(group) - 1]}"/>' for group in sorted(groups, key=_date_order))
    return _date_filter(groups, listed, _DATES_HEAD + body + _DATES_TAIL)


def _date_order(group: tuple[int, ...]) -> tuple[int, ...]:
    """Where Excel writes a date filter's item: the latest year first, and within one the earliest moment first, a
    group before those inside it (tests/fixtures/date_filters/)."""
    rest = [group[index] if index < len(group) else (0 if index < 3 else -1) for index in range(1, len(_DATE_PARTS))]
    return (-group[0], *rest)


def _date_filter(groups: list[tuple[int, ...]], listed: FieldFilter | None, xml: str, *,
                 classic: bool = False) -> FieldFilter:
    """A filter showing the date cells inside ``groups`` and the cells ``listed`` shows. It reads back as the
    values beside the dates do, Operator 0 and no criteria where there are none, and a filter the file kept the
    classic way as xlFilterValues."""
    def test(seen: Seen) -> bool:
        if listed is not None and listed.test(seen):
            return True
        moment = _date_of(seen)
        if moment is None:
            return False
        parts = (moment.year, moment.month, moment.day, moment.hour, moment.minute, moment.second)
        return any(parts[:len(group)] == group for group in groups)

    if listed is None:
        return FieldFilter(VALUES if classic else 0, MISSING, MISSING, test, 0, xml)
    return FieldFilter(VALUES if classic else listed.operator, listed.criteria1, listed.criteria2, test,
                       listed.count, xml)


def _date_of(seen: Seen) -> _dt.datetime | None:
    """The moment a cell holds as a date filter sees it: a number in a date format; None for anything else."""
    from pyopenvba.apps.excel._typing import is_date_format

    if seen.number is None or seen.sheet is None:
        return None
    cell = seen.sheet.cells_.get((seen.row, seen.column))
    if cell is None or not is_date_format(cell.number_format):
        return None
    try:
        return VBADate(seen.number + (1462 if seen.sheet.book.epoch_1904 else 0)).to_datetime()
    except VBARuntimeError:
        # A number past either end of the calendar names no date.
        return None


#: The dynamic filters on dates, by their Criteria1 from 1, as the file names them: those counted from today, then
#: each quarter and month of any year.
PERIODS = ("today", "yesterday", "tomorrow", "thisWeek", "lastWeek", "nextWeek", "thisMonth", "lastMonth",
            "nextMonth", "thisQuarter", "lastQuarter", "nextQuarter", "thisYear", "lastYear", "nextYear",
            "yearToDate", "Q1", "Q2", "Q3", "Q4", *(f"M{month}" for month in range(1, 13)))


def period_bounds(which: int, today: _dt.date) -> tuple[_dt.date, _dt.date] | None:
    """The days a dynamic filter counted from today keeps, from the first to the one after the last, a week from
    Sunday; None for a quarter or a month of any year (tests/fixtures/dynamic_filters/)."""
    day = _dt.timedelta(days=1)

    def months(year: int, month: int, add: int) -> _dt.date:
        index = year * 12 + month - 1 + add
        return _dt.date(index // 12, index % 12 + 1, 1)

    week = today - day * ((today.weekday() + 1) % 7)
    quarter = (today.month - 1) // 3 * 3 + 1
    found = {1: (today, today + day), 2: (today - day, today), 3: (today + day, today + day * 2),
             4: (week, week + day * 7), 5: (week - day * 7, week), 6: (week + day * 7, week + day * 14),
             7: (months(today.year, today.month, 0), months(today.year, today.month, 1)),
             8: (months(today.year, today.month, -1), months(today.year, today.month, 0)),
             9: (months(today.year, today.month, 1), months(today.year, today.month, 2)),
             10: (months(today.year, quarter, 0), months(today.year, quarter, 3)),
             11: (months(today.year, quarter, -3), months(today.year, quarter, 0)),
             12: (months(today.year, quarter, 3), months(today.year, quarter, 6)),
             13: (_dt.date(today.year, 1, 1), _dt.date(today.year + 1, 1, 1)),
             14: (_dt.date(today.year - 1, 1, 1), _dt.date(today.year, 1, 1)),
             15: (_dt.date(today.year + 1, 1, 1), _dt.date(today.year + 2, 1, 1)),
             16: (_dt.date(today.year, 1, 1), today + day)}
    return found.get(which)


def _period(sheet: Worksheet, which: int, xml: str = "") -> FieldFilter:
    """A dynamic filter on dates: the date cells in a period counted from today, or in a quarter or a month of any
    year. It reads back as xlFilterDynamic and its Criteria1, and the file keeps the period's first serial and the
    one after its last (tests/fixtures/dynamic_filters/)."""
    bounds = period_bounds(which, sheet.book.calculator.now().date())

    def test(seen: Seen) -> bool:
        moment = _date_of(seen)
        if moment is None:
            return False
        if bounds is None:
            return (moment.month - 1) // 3 + 1 == which - 16 if which <= 20 else moment.month == which - 20
        return _dt.datetime.combine(bounds[0], _dt.time()) <= moment < _dt.datetime.combine(bounds[1], _dt.time())

    if not xml:
        kind = PERIODS[which - 1]
        if bounds is None:
            xml = f'<dynamicFilter type="{kind}"/>'
        else:
            epoch = _dt.date(1899, 12, 30) if not sheet.book.epoch_1904 else _dt.date(1904, 1, 1)
            xml = (f'<dynamicFilter type="{kind}" val="{(bounds[0] - epoch).days}" '
                   f'maxVal="{(bounds[1] - epoch).days}"/>')
    return FieldFilter(DYNAMIC, VBAInt(which, "Long"), MISSING, test, xml=xml)


def _read_dates(sheet: Worksheet, xml: str, attributes: str, body: str, prefix: str, *,
                classic: bool) -> FieldFilter | None:
    """A date filter from the file, the classic way or in the Excel 2017 extension; None for one not measured."""
    if _xml_attributes(attributes):
        return None
    words: list[str] = []
    groups: list[tuple[int, ...]] = []
    for match in re.finditer(rf"<{prefix}(filter|dateGroupItem)\b([^>]*)/>|(\S)", body):
        if match.group(3) is not None:
            return None
        found = _xml_attributes(match.group(2))
        if match.group(1) == "filter":
            words.append(found.get("val", ""))
            continue
        grouping = found.get("dateTimeGrouping", "")
        if grouping not in _DATE_PARTS:
            return None
        try:
            groups.append(tuple(int(found[name]) for name in _DATE_PARTS[:_DATE_PARTS.index(grouping) + 1]))
        except (KeyError, ValueError):
            return None
    listed = None
    if words:
        listed = _read_values("", "", "".join(f'<filter val="{xml_attribute(word)}"/>' for word in words))
        if listed.operator < 0:
            return None
    return _date_filter(groups, listed, xml, classic=classic)


def _iconned(sheet: Worksheet, area: Area, index: int, kind: int, criteria1: object) -> FieldFilter:
    """A filter on the icon a conditional format's icon set shows: the cells showing the Icon Criteria1 is, its set
    and place in it and not an icon alike in another set, or with xlFilterNoIcon those showing none
    (tests/fixtures/icon_filters/)."""
    from pyopenvba.apps.excel._icon_sets import SETS, Icon, covering, state_of

    if kind == ICON:
        if not isinstance(criteria1, Icon) or criteria1.owner.identifier not in SETS:
            raise VBAUnsupportedError("AutoFilter by an icon not of one of Excel's sets is not implemented")
        return _icon_filter(sheet, kind, criteria1.owner.identifier, criteria1.index)
    # A filter on no icon names the set of the column's icons.
    column = area.left + index - 1
    rule = next((found for found in (covering(sheet, row, column) for row in range(area.top + 1, area.bottom + 1))
                 if found is not None), None)
    if rule is None:
        raise VBAUnsupportedError("AutoFilter on no icon over a column no icon set covers is not implemented")
    return _icon_filter(sheet, kind, state_of(sheet, rule).set_id, 0)


def _icon_filter(sheet: Worksheet, kind: int, set_id: int, place: int, xml: str = "") -> FieldFilter:
    """The cells showing the icon at ``place`` in the set ``set_id``, or for xlFilterNoIcon those showing none."""
    from pyopenvba.apps.excel._icon_sets import SETS, Icon, IconSet, shown_icon

    wanted = (set_id, place) if kind == ICON else None

    def test(seen: Seen) -> bool:
        assert seen.sheet is not None, "an icon filter looks at a cell where it is"
        return shown_icon(seen.sheet, seen.row, seen.column) == wanted

    name, icons = SETS[set_id]
    if kind == ICON:
        criteria: object = Icon(IconSet(sheet.book, set_id, len(icons)), place)
        xml = xml or f'<iconFilter iconSet="{name}" iconId="{place - 1}"/>'
    else:
        # Criteria1 reads Empty on a filter on no icon.
        criteria = EMPTY
        xml = xml or f'<iconFilter iconSet="{name}"/>'
    return FieldFilter(kind, criteria, MISSING, test, xml=xml)


def _read_icon(sheet: Worksheet, attributes: str, xml: str) -> FieldFilter | None:
    """An icon filter from the file: a set and the place of the icon in it, or a set alone for no icon."""
    from pyopenvba.apps.excel._icon_sets import SETS, set_named

    found = _xml_attributes(attributes)
    set_id = set_named(found.get("iconSet", "3TrafficLights1"))
    if set_id is None:
        return None
    place = found.get("iconId")
    if place is None:
        return _icon_filter(sheet, NO_ICON, set_id, 0, xml)
    if not place.isdecimal() or int(place) >= len(SETS[set_id][1]):
        return None
    return _icon_filter(sheet, ICON, set_id, int(place) + 1, xml)


def _coloured(sheet: Worksheet, kind: int, criteria1: object) -> FieldFilter:
    """A filter on a cell's colour or its font's, keeping its colour in a dxf of its own, made as the filter is
    (tests/fixtures/colour_filters/)."""
    colour: int | None = None
    if kind in (CELL_COLOR, FONT_COLOR):
        if criteria1 is MISSING:
            raise VBAUnsupportedError("AutoFilter on a colour with no Criteria1 is not implemented")
        colour = int(to_integer(criteria1, "Long"))
    return _colour_filter(kind, colour, sheet.book.stylesheet.add_dxf(_colour_dxf(kind, colour)))


def _colour_filter(kind: int, colour: int | None, dxf: int) -> FieldFilter:
    """A colour filter keeping its colour in the dxf at ``dxf``: a cell passes that shows the colour, as a sort on
    it finds it, or no fill, or the automatic font colour."""
    from pyopenvba.apps.excel._sort import ON_CELL_COLOR, ON_FONT_COLOR, Key, shows_automatic_font, shows_color

    font = kind in _FONTS
    key = Key(0, on=ON_FONT_COLOR if font else ON_CELL_COLOR, color=colour)

    def test(seen: Seen) -> bool:
        assert seen.sheet is not None, "a colour filter looks at a cell where it is"
        if kind == AUTOMATIC_FONT:
            return shows_automatic_font(seen.sheet, seen.row, seen.column)
        return shows_color(seen.sheet, seen.row, seen.column, key)

    # Criteria1 reads a cell colour as an Interior, white with no pattern for no fill, and a font's as its colour.
    criteria: object = (VBAInt(colour or 0, "Long") if font
                        else FilterInterior(WHITE if colour is None else colour, PATTERN_NONE if colour is None
                                            else PATTERN_SOLID))
    xml = f'<colorFilter dxfId="{dxf}"' + (' cellColor="0"' if font else "") + "/>"
    return FieldFilter(kind, criteria, MISSING, test, xml=xml, dxf=dxf)


def _colour_dxf(kind: int, colour: int | None) -> XmlElement:
    """The dxf a colour filter keeps its colour in, as Excel writes it: a solid fill of the colour on white, a
    font's colour as well; no fill; or the automatic colour."""
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._styles import Color

    dxf = XmlElement(_MAIN + "dxf")
    fill = XmlElement(_MAIN + "fill")
    pattern = XmlElement(_MAIN + "patternFill", {"patternType": "none" if kind == NO_FILL else "solid"})
    if kind == NO_FILL:
        colours = {"indexed": "64"}, {"indexed": "65"}
    elif kind == AUTOMATIC_FONT:
        colours = {"auto": "1"}, {"indexed": "65"}
    else:
        colours = {"rgb": str(Color.from_bgr(colour or 0).value)}, {"rgb": "FFFFFFFF"}
    pattern.extend([XmlElement(_MAIN + "fgColor", colours[0]), XmlElement(_MAIN + "bgColor", colours[1])])
    fill.append(pattern)
    dxf.append(fill)
    return dxf


def _read_colour(sheet: Worksheet, attributes: str) -> FieldFilter | None:
    """A colour filter from the file, its colour read from the dxf it names; None for one not measured."""
    from pyopenvba.apps.excel._styles import bgr

    found = _xml_attributes(attributes)
    if not found.get("dxfId", "").isdecimal():
        return None
    index = int(found["dxfId"])
    font = found.get("cellColor") in ("0", "false")
    try:
        dxf = sheet.book.stylesheet.dxf(index)
    except IndexError:
        return None
    pattern = dxf.find(f"{_MAIN}fill/{_MAIN}patternFill")
    if pattern is None:
        return None
    foreground = pattern.find(_MAIN + "fgColor")
    if pattern.get("patternType") == "none" and not font:
        return _colour_filter(NO_FILL, None, index)
    if pattern.get("patternType") != "solid" or foreground is None:
        return None
    if font and foreground.get("auto") in ("1", "true"):
        return _colour_filter(AUTOMATIC_FONT, None, index)
    rgb = foreground.get("rgb", "")
    if len(rgb) not in (6, 8):
        return None
    return _colour_filter(FONT_COLOR if font else CELL_COLOR, bgr(rgb[-6:]), index)


def _opaque(xml: str) -> FieldFilter:
    """A column filter read from a file that the model does not follow: kept as it is, and not applied."""

    def refuse(seen: Seen) -> bool:
        raise VBAUnsupportedError("applying a filter of this kind from a file is not implemented")

    return FieldFilter(-1, MISSING, MISSING, refuse, xml=xml)


# --- the sheet's filter -----------------------------------------------------------------------------------


@dataclass
class SheetFilter:
    """A sheet's AutoFilter: the range it covers, header row first, and the criteria on its columns, from 1.

    ``uid`` is the revision id the file gives the filter, and ``tail`` whatever the element holds after its
    columns -- a sort state, say -- written back as it was read.
    """

    area: Area
    fields: dict[int, FieldFilter] = field(default_factory=lambda: {})
    uid: str = ""
    tail: str = ""

    def apply(self, sheet: Worksheet) -> None:
        """Show each row under the header that meets every column's criteria, and hide the rest."""
        for row in range(self.area.top + 1, self.area.bottom + 1):
            shown_row = all(test.test(_seen(sheet, row, self.area.left + index - 1))
                            for index, test in self.fields.items())
            if sheet.dims.row_hidden(row) == shown_row:
                sheet.dims.hide_row(row, not shown_row)
        sheet.touched()

    def show_all(self, sheet: Worksheet) -> None:
        for row in range(self.area.top + 1, self.area.bottom + 1):
            if sheet.dims.row_hidden(row):
                sheet.dims.hide_row(row, False)
        sheet.touched()

    def xml(self) -> str:
        """The autoFilter element the file holds for this filter."""
        uid = f' xr:uid="{self.uid}"' if self.uid else ""
        head = f'<autoFilter ref="{self.area.address(absolute=False)}"{uid}'
        columns = "".join(f'<filterColumn colId="{index - 1}"{one.attributes}>{one.markup()}</filterColumn>'
                          for index, one in sorted(self.fields.items()))
        if not columns and not self.tail:
            return head + "/>"
        return f"{head}>{columns}{self.tail}</autoFilter>"


def _changed(sheet: Worksheet) -> None:
    sheet.filter_changed = True
    sheet.touched()


def _filter_area(target: Range) -> Area:
    from pyopenvba.apps.excel._region import current_region

    if len(target.areas) != 1:
        raise VBAUnsupportedError("AutoFilter on several areas is not implemented")
    area = target.first
    sheet = target.sheet
    if target.single:
        return current_region(sheet, area.top, area.left)
    if area.whole_columns or area.whole_rows:
        used = sheet.used_bounds()
        if used is None:
            raise VBAUnsupportedError("AutoFilter on whole columns of an empty sheet is not implemented")
        return Area(area.top, area.left, min(area.bottom, used[2]), min(area.right, used[3]))
    return area


def _remember_database(sheet: Worksheet, area: Area) -> None:
    """The hidden _FilterDatabase name the sheet's filter leaves, kept pointing at its range."""
    names = sheet.book.names_
    local = quote_sheet(sheet.name) + "!_FilterDatabase"
    refers_to = "=" + quote_sheet(sheet.name) + "!" + area.address()
    for entry in names.entries:
        if entry.name.lower() == local.lower():
            if entry.refers_to != refers_to:
                entry.refers_to = refers_to
                names.changed = True
            return
    names.entries.append(NameEntry(local, refers_to, sheet.book, visible=False))
    names.changed = True


def _new_filter(sheet: Worksheet, area: Area) -> SheetFilter:
    made = SheetFilter(area, uid="{" + str(uuid.uuid4()).upper() + "}" if sheet.declares_revisions() else "")
    sheet.auto_filter = made
    _remember_database(sheet, area)
    return made


# --- the filter as the sheet changes around it ------------------------------------------------------------


def grown(sheet: Worksheet, area: Area) -> Area:
    """A filter's range as Excel reads and applies it: its rows run on to the foot of the data around it.

    Measured (scripts/measure_autofilter_edits.py, measure_autofilter_file.py):
    a filter over A1:C7 reads A1:C8 once anything is written in row 8
    where the block around the range reaches -- D8 and E8 do, Z8 does
    not, as CurrentRegion grows -- and A1:C7 again once it is cleared. A
    new filter over A1:D5 of a table down to row 13 filters and saves
    A1:D13 while its hidden name keeps A1:D5; arrows alone keep A1:D5.
    Filtering again, or ApplyFilter, takes in what has grown; filtering
    again renames it, ApplyFilter does not.
    """
    from pyopenvba.apps.excel._region import region_around

    bottom = region_around(sheet, area).bottom
    return area if bottom <= area.bottom else Area(area.top, area.left, bottom, area.right, area.sheet)


def filter_edited(sheet: Worksheet, *, rows: bool, start: int, count: int, delete: bool) -> None:
    """A sheet's filter after ``count`` whole rows or columns go in or out at ``start``.

    Measured: the range moves and stretches as a reference does. Deleting
    its header row takes the filter and its hidden name away, showing
    what is left of its rows, and so does deleting all its columns.
    Deleting a filtered column drops that column's criteria and filters
    again by the rest; the other columns' criteria follow their columns.
    """
    from pyopenvba.apps.excel._editing import interval

    found = sheet.auto_filter
    if found is None:
        return
    box = found.area
    if rows:
        span = interval(box.top, box.bottom, start, count, delete, MAX_ROWS)
        if span is not None and not (delete and start <= box.top < start + count):
            found.area = Area(span[0], box.left, span[1], box.right, box.sheet)
            _changed(sheet)
            return
        if span is not None:
            for row in range(span[0], span[1] + 1):
                if sheet.dims.row_hidden(row):
                    sheet.dims.hide_row(row, False)
        _drop(sheet)
        return
    span = interval(box.left, box.right, start, count, delete, MAX_COLUMNS)
    if span is None:
        found.show_all(sheet)
        _drop(sheet)
        return
    fields: dict[int, FieldFilter] = {}
    for index, one in found.fields.items():
        column = interval(box.left + index - 1, box.left + index - 1, start, count, delete, MAX_COLUMNS)
        if column is not None:
            fields[column[0] - span[0] + 1] = one
    dropped = len(fields) != len(found.fields)
    found.fields = fields
    found.area = Area(box.top, span[0], box.bottom, span[1], box.sheet)
    if dropped:
        found.apply(sheet)
    _changed(sheet)


def filter_shifted(sheet: Worksheet, shift: CellShift) -> Callable[[], None]:
    """What a cell shift does to the sheet's filter: worked out before any cell moves, done by what it returns.

    Measured (scripts/measure_cell_shifts.py): the range moves and
    stretches as a reference to it does, as its hidden name does, and a
    delete up of its header cells, or of all of it, removes the filter and
    the name. A shift that moves a filter with criteria was not measured.
    """
    from pyopenvba.apps.excel._editing import shifted_box

    found = sheet.auto_filter
    if found is None:
        return lambda: None
    box = found.area
    before = (box.top, box.left, box.bottom, box.right)
    moved = shifted_box(before, shift)
    band, start = shift.band, shift.start
    header = shift.vertical and shift.delete and band[0] <= box.left and box.right <= band[1] and \
        start <= box.top < start + shift.count
    if moved == before and not header:
        return lambda: None
    if found.fields:
        raise VBAUnsupportedError("a cell shift that moves a filter with criteria is not implemented")
    if moved is None or header:
        return lambda: _drop(sheet)
    after = moved

    def follow() -> None:
        found.area = Area(after[0], after[1], after[2], after[3], box.sheet)
        _changed(sheet)

    return follow


def filter_moved(sheet: Worksheet, area: Area, target: Worksheet, down: int, across: int) -> None:
    """A sheet's filter after a cut moves ``area`` by ``down`` and ``across`` to ``target``.

    Measured: a cut of the whole range moves the filter along its sheet
    with its criteria cleared, and to another sheet takes it away; its
    rows show either way, and the hidden name follows the cells. A cut
    of part of the range leaves the filter as it was.
    """
    found = sheet.auto_filter
    box = found.area if found is not None else None
    if found is None or box is None or not (area.top <= box.top and box.bottom <= area.bottom
                                            and area.left <= box.left and box.right <= area.right):
        return
    found.show_all(sheet)
    if target is sheet:
        found.area = Area(box.top + down, box.left + across, box.bottom + down, box.right + across, box.sheet)
        found.fields.clear()
    else:
        sheet.auto_filter = None
    _changed(sheet)


def header_cleared(target: Range) -> None:
    """A Clear, ClearContents or ClearFormats of ``target``, which turns the filter off if it took in the whole
    header row.

    Measured: clearing A1:C1 over a filter on A1:C7, or any range around
    it, turns the filter off as AutoFilterMode = False does, its rows shown
    and its hidden name kept; clearing A1 alone, or writing "" to A1:C1,
    leaves it.
    """
    sheet = target.sheet
    found = sheet.auto_filter
    if found is None:
        return
    box = found.area
    if all(any(area.contains(box.top, column) for area in target.areas) for column in range(box.left, box.right + 1)):
        found.show_all(sheet)
        sheet.auto_filter = None
        _changed(sheet)


def filtered_again(sheet: Worksheet, removed: int) -> None:
    """A sheet's filter after RemoveDuplicates over its range, which gave up ``removed`` rows at its foot.

    Measured: the rows given up show, the hidden name follows the range,
    and a filter with criteria filters again.
    """
    found = sheet.auto_filter
    if found is None:
        return
    if removed:
        box = found.area
        found.area = Area(box.top, box.left, box.bottom - removed, box.right, box.sheet)
        _remember_database(sheet, found.area)
        for row in range(found.area.bottom + 1, box.bottom + 1):
            if sheet.dims.row_hidden(row):
                sheet.dims.hide_row(row, False)
    if found.fields:
        found.apply(sheet)
    _changed(sheet)


def _drop(sheet: Worksheet) -> None:
    """The filter gone, and its hidden name with it."""
    sheet.auto_filter = None
    names = sheet.book.names_
    local = (quote_sheet(sheet.name) + "!_FilterDatabase").lower()
    for entry in [entry for entry in names.entries if entry.name.lower() == local]:
        names.entries.remove(entry)
        names.changed = True
    _changed(sheet)


def range_autofilter(target: Range, field_argument: object, criteria1: object, operator: object, criteria2: object,
                     visible_dropdown: object, sub_field: object) -> object:
    """Range.AutoFilter."""
    from pyopenvba.apps.excel._table_filters import meets_table, refocus, table_of

    sheet = target.sheet
    if sheet.advanced_filter is not None:
        raise VBAUnsupportedError('AutoFilter combined with an active AdvancedFilter is not implemented')
    if sub_field is not MISSING:
        raise VBAUnsupportedError("AutoFilter's SubField is not implemented")
    hidden_button = visible_dropdown is not MISSING and not to_bool(visible_dropdown)
    table = table_of(target)
    if table is not None:
        return _filter_table(table, field_argument, criteria1, operator, criteria2, hidden_button=hidden_button)
    existing = sheet.auto_filter
    if existing is None and meets_table(sheet, _filter_area(target)):
        raise VBAUnsupportedError("a sheet's AutoFilter over a range that takes a table in is not implemented")
    filtered = existing is not None and bool(existing.fields)
    if field_argument is MISSING:
        if hidden_button:
            raise VBAUnsupportedError("hiding every drop-down of a filter is not implemented")
        if existing is not None:
            existing.show_all(sheet)
            sheet.auto_filter = None
        else:
            _new_filter(sheet, _filter_area(target))
        _changed(sheet)
        refocus(sheet)
        if existing is not None and filtered:
            from pyopenvba.apps.excel import _events

            _events.visibility_changed(sheet.book)
        return True
    asked = _filter_area(target) if existing is None else existing.area
    # Filtering takes in the rows of data grown on below the range.
    area = grown(sheet, asked)
    if meets_table(sheet, area):
        raise VBAUnsupportedError("a sheet's AutoFilter grown over a table is not implemented")
    index = int(to_integer(field_argument, "Long"))
    if not 1 <= index <= area.columns:
        raise error(1004, "AutoFilter's field is outside the range")
    fresh = False
    if existing is None and criteria2 is not MISSING and operator is not MISSING \
            and int(to_integer(operator, "Long")) == VALUES:
        # Excel turns the filter on before it reads the dates, so one it cannot read leaves the filter on with
        # nothing in it (tests/fixtures/date_filters/).
        existing, fresh = _new_filter(sheet, asked), True
        _changed(sheet)
    made = _field(sheet, area, index, criteria1, operator, criteria2)
    if made is None and hidden_button:
        raise VBAUnsupportedError("hiding a drop-down without criteria is not implemented")
    if existing is None:
        # A new filter's name keeps the range it was asked for; filtering again names the grown one.
        existing = _new_filter(sheet, asked)
    elif not fresh:
        _remember_database(sheet, area)
    existing.area = area
    if made is None:
        existing.fields.pop(index, None)
    else:
        if hidden_button:
            made.attributes = ' hiddenButton="1"'
        existing.fields[index] = made
    existing.apply(sheet)
    _changed(sheet)
    refocus(sheet)
    from pyopenvba.apps.excel import _events

    _events.visibility_changed(sheet.book)
    if operator is not MISSING and int(to_integer(operator, "Long")) == VALUES:
        return float(index)
    return True


def _filter_table(table: Table, field_argument: object, criteria1: object, operator: object, criteria2: object, *,
                  hidden_button: bool) -> object:
    """Range.AutoFilter on a range of a table: the table's own filter, its field counted from the table's first
    column (see _table_filters)."""
    from pyopenvba.apps.excel import _events
    from pyopenvba.apps.excel._table_filters import changed, filtered_block, refocus, table_filter, turn_off

    sheet = table.sheet
    if not table.headers:
        raise VBAUnsupportedError("filtering a table with no header row, which shows the header row again, is not "
                                  "implemented")
    if field_argument is MISSING:
        if hidden_button:
            raise VBAUnsupportedError("hiding every drop-down of a filter is not implemented")
        if table.auto_filter:
            turn_off(table)
        else:
            table.auto_filter = True
            changed(table)
        refocus(sheet)
        return True
    area = filtered_block(table)
    index = int(to_integer(field_argument, "Long"))
    if not 1 <= index <= area.columns:
        raise error(1004, "AutoFilter's field is outside the table")
    made = _field(sheet, area, index, criteria1, operator, criteria2)
    if made is None and hidden_button:
        raise VBAUnsupportedError("hiding a drop-down without criteria is not implemented")
    # A filter turns on a table's filter it had turned off.
    table.auto_filter = True
    found = table_filter(table)
    assert found is not None
    if made is None:
        found.fields.pop(index, None)
    else:
        if hidden_button:
            made.attributes = ' hiddenButton="1"'
        found.fields[index] = made
    found.apply(sheet)
    changed(table)
    refocus(sheet)
    _events.visibility_changed(sheet.book)
    if operator is not MISSING and int(to_integer(operator, "Long")) == VALUES:
        return float(index)
    return True


def show_all_data(sheet: Worksheet) -> None:
    """Worksheet.ShowAllData: the AdvancedFilter's rows shown, or the criteria of the filter at the active cell
    taken away -- a table's where it is in a table -- error 1004 where it has none."""
    from pyopenvba.apps.excel._table_filters import clear, context_table, has_criteria

    if sheet.advanced_filter is not None:
        area = sheet.advanced_filter
        for row in range(area.top + 1, area.bottom + 1):
            if sheet.dims.row_hidden(row):
                sheet.dims.hide_row(row, False)
        sheet.advanced_filter = None
        sheet.advanced_filter_changed = True
        _changed(sheet)
        from pyopenvba.apps.excel import _events

        _events.visibility_changed(sheet.book)
        return
    table = context_table(sheet)
    if table is not None:
        if not has_criteria(table):
            raise error(1004, "ShowAllData method of Worksheet class failed")
        clear(table)
        return
    show_all_own(sheet)


def show_all_own(sheet: Worksheet) -> None:
    """ShowAllData on the sheet's own filter: its criteria taken away and its rows shown; error 1004 with none."""
    from pyopenvba.apps.excel._table_filters import refocus

    found = sheet.auto_filter
    if found is None or not found.fields:
        raise error(1004, "ShowAllData method of Worksheet class failed")
    found.fields.clear()
    found.show_all(sheet)
    _changed(sheet)
    refocus(sheet)
    from pyopenvba.apps.excel import _events

    _events.visibility_changed(sheet.book)


def set_mode(sheet: Worksheet, value: object) -> None:
    """Worksheet.AutoFilterMode = False turns off the filter at the active cell, a table's where it is in a table,
    but only on a sheet with a filter of its own; True cannot turn one on."""
    from pyopenvba.apps.excel._table_filters import context_table, refocus, turn_off

    if to_bool(value):
        raise VBAUnsupportedError("setting AutoFilterMode to True is not implemented")
    if sheet.auto_filter is None:
        return
    table = context_table(sheet)
    if table is not None:
        if table.auto_filter:
            turn_off(table)
        refocus(sheet)
        return
    filtered = bool(sheet.auto_filter.fields)
    sheet.auto_filter.show_all(sheet)
    sheet.auto_filter = None
    _changed(sheet)
    refocus(sheet)
    if filtered:
        from pyopenvba.apps.excel import _events

        _events.visibility_changed(sheet.book)


# --- reading the file ----------------------------------------------------------------------------------------


_COLUMN = re.compile(r'<filterColumn\b([^>]*?)(?:/>|>(.*?)</filterColumn>)', re.DOTALL)


def read_filter(sheet: Worksheet, element: str) -> SheetFilter | None:
    """A sheet's filter from the autoFilter element of its file."""
    head = re.match(r"<autoFilter\b([^>]*?)(/?)>", element)
    if head is None:
        return None
    attributes = head.group(1)
    reference = re.search(r'\bref="([^"]*)"', attributes)
    if reference is None:
        return None
    areas = parse_reference(reference.group(1), sheet=sheet.name)
    if len(areas) != 1:
        return None
    uid = re.search(r'\bxr:uid="([^"]*)"', attributes)
    found = SheetFilter(areas[0], uid=uid.group(1) if uid else "")
    body = "" if head.group(2) else element[head.end():element.rfind("</autoFilter>")]
    last = 0
    for match in _COLUMN.finditer(body):
        column_attributes = match.group(1)
        colid = re.search(r'\bcolId="(\d+)"', column_attributes)
        if colid is None:
            continue
        others = re.sub(r'\s*\bcolId="\d+"', "", column_attributes)
        read = _read_column(match.group(2) or "", sheet)
        read.attributes = others
        found.fields[int(colid.group(1)) + 1] = read
        last = match.end()
    found.tail = body[last:]
    return found


def _read_column(xml: str, sheet: Worksheet) -> FieldFilter:
    """One column's criteria from the file, as Excel reads them back."""
    colour = re.fullmatch(r"<colorFilter\b([^>]*?)/>", xml)
    if colour is not None:
        found_colour = _read_colour(sheet, colour.group(1))
        if found_colour is not None:
            found_colour.xml = xml
            return found_colour
    icon = re.fullmatch(r"<iconFilter\b([^>]*?)/>", xml)
    if icon is not None:
        found_icon = _read_icon(sheet, icon.group(1), xml)
        return found_icon if found_icon is not None else _opaque(xml)
    dates = re.fullmatch(r"<filters\b([^>]*?)>(.*?)</filters>", xml, re.DOTALL)
    if dates is not None and "<dateGroupItem" in dates.group(2):
        found_dates = _read_dates(sheet, xml, dates.group(1), dates.group(2), "", classic=True)
        return found_dates if found_dates is not None else _opaque(xml)
    extension = re.fullmatch(r'<extLst><ext\b[^>]*?\buri="\{1AD28BCE-077C-4C59-8B6E-1921CE8616D4\}"[^>]*>'
                             r"<xlrd2:filterColumn><xlrd2:filters\b([^>]*?)>(.*?)</xlrd2:filters>"
                             r"</xlrd2:filterColumn></ext></extLst>", xml, re.DOTALL)
    if extension is not None:
        found_dates = _read_dates(sheet, xml, extension.group(1), extension.group(2), "xlrd2:", classic=False)
        return found_dates if found_dates is not None else _opaque(xml)
    values = re.fullmatch(r'<filters\b([^>]*?)(?:/>|>(.*?)</filters>)', xml, re.DOTALL)
    if values is not None:
        return _read_values(xml, values.group(1), values.group(2) or "")
    custom = re.fullmatch(r'<customFilters\b([^>]*?)>(.*?)</customFilters>', xml, re.DOTALL)
    if custom is not None:
        return _read_custom(xml, custom.group(1), custom.group(2))
    top = re.fullmatch(r"<top10\b([^>]*?)/>", xml)
    if top is not None:
        return _read_top(xml, top.group(1))
    dynamic = re.fullmatch(r'<dynamicFilter\b([^>]*?)/>', xml)
    if dynamic is not None:
        found = _xml_attributes(dynamic.group(1))
        if found.get("type", "") in PERIODS:
            return _period(sheet, PERIODS.index(found["type"]) + 1, xml)
        which = {"aboveAverage": ABOVE_AVERAGE, "belowAverage": BELOW_AVERAGE}.get(found.get("type", ""))
        if which is not None and "val" in found:
            read = _averaged(which, float(found["val"]))
            read.xml = xml
            return read
    return _opaque(xml)


def _xml_attributes(text: str) -> dict[str, str]:
    from html import unescape

    return {name: unescape(value) for name, value in re.findall(r'([\w:]+)="([^"]*)"', text)}


def _read_values(xml: str, attributes: str, body: str) -> FieldFilter:
    found = _xml_attributes(attributes)
    if set(found) - {"blank"} or re.search(r"<(?!filter\b)\w", body):
        return _opaque(xml)
    blank = found.get("blank") in ("1", "true")
    values = [_xml_attributes(one).get("val", "") for one in re.findall(r"<filter\b([^>]*)/>", body)]
    items = ([""] if blank else []) + values
    if len(items) >= 3:
        read = _listed(values, blank)
    elif len(items) == 2:
        read = _two(made_criterion("=", _escaped(items[0])), made_criterion("=", _escaped(items[1])), OR)
    elif items:
        read = _one(made_criterion("=", _escaped(items[0])))
    else:
        return _opaque(xml)
    read.xml = xml
    return read


def _read_custom(xml: str, attributes: str, body: str) -> FieldFilter:
    joined = AND if _xml_attributes(attributes).get("and") in ("1", "true") else OR
    criteria: list[Criterion] = []
    for one in re.findall(r"<customFilter\b([^>]*)/>", body):
        found = _xml_attributes(one)
        operator = next((symbol for symbol, name in _FILE_OPERATORS.items() if name == found.get("operator")), "=")
        value = found.get("val", "")
        if operator == "<>" and value == " ":
            criteria.append(made_criterion(operator, ""))
            continue
        if operator != "=":
            try:
                number = float(value)
            except ValueError:
                pass
            else:
                # The file keeps the number in full; it reads back in fifteen digits.
                spelled = number_text(number, formula=True)
                criteria.append(Criterion(operator, spelled, number, operator + spelled))
                continue
        criteria.append(made_criterion(operator, value))
    if len(criteria) == 1:
        read = _one(criteria[0])
    elif len(criteria) == 2:
        read = _two(criteria[0], criteria[1], joined)
    else:
        return _opaque(xml)
    read.xml = xml
    return read


def _read_top(xml: str, attributes: str) -> FieldFilter:
    found = _xml_attributes(attributes)
    if "filterVal" not in found:
        return _opaque(xml)
    top = found.get("top", "1") not in ("0", "false")
    percent = found.get("percent") in ("1", "true")
    kind = (TOP_PERCENT if percent else TOP_ITEMS) if top else (BOTTOM_PERCENT if percent else BOTTOM_ITEMS)
    read = _ranked(kind, float(found.get("val", "10")), float(found["filterVal"]))
    read.xml = xml
    return read


# --- the objects -------------------------------------------------------------------------------------------


class AutoFilter(ExcelObject):
    """The sheet's own filter, or a table's where ``table`` is one (see _table_filters)."""

    vba_type_name = "AutoFilter"

    def __init__(self, sheet: Worksheet, table: Table | None = None) -> None:
        self.sheet = sheet
        self.table = table

    def current(self) -> SheetFilter | None:
        """The filter as it is now, None where it has gone."""
        if self.table is None:
            return self.sheet.auto_filter
        from pyopenvba.apps.excel._table_filters import table_filter

        return table_filter(self.table) if self.table in self.sheet.tables else None

    def state(self) -> SheetFilter:
        found = self.current()
        if found is None:
            raise error(1004, "The sheet has no AutoFilter")
        return found

    @member
    def Parent(self) -> object:
        if self.table is None:
            return self.sheet
        from pyopenvba.apps.excel._tables import view

        return view(self.table)

    @member
    def Range(self) -> object:
        found = self.state()
        if self.table is not None:
            from pyopenvba.apps.excel._table_filters import filter_range

            return Range(self.sheet, [filter_range(self.table)])
        return Range(self.sheet, [grown(self.sheet, found.area)])

    @member
    def Filters(self, Index: object = MISSING) -> object:
        found = Filters(self)
        return found if Index is MISSING else found.vba_lookup(Index, found.vba_items())

    @member
    def FilterMode(self) -> object:
        return bool(self.state().fields)

    @method
    def ApplyFilter(self) -> object:
        from pyopenvba.apps.excel._table_filters import changed, refocus

        found = self.state()
        if self.table is not None:
            found.apply(self.sheet)
            changed(self.table)
        else:
            area = grown(self.sheet, found.area)
            if area != found.area:
                found.area = area
                _changed(self.sheet)
            found.apply(self.sheet)
        refocus(self.sheet)
        from pyopenvba.apps.excel import _events

        _events.visibility_changed(self.sheet.book)
        return EMPTY

    @method
    def ShowAllData(self) -> object:
        self.state()
        if self.table is not None:
            from pyopenvba.apps.excel._table_filters import clear

            # A table's clears with no error when nothing is set (tests/fixtures/table_autofilter/).
            clear(self.table)
        else:
            show_all_own(self.sheet)
        return EMPTY


class FilterInterior(ExcelObject):
    """A cell-colour filter's Criteria1: the colour it keeps, solid, or white with no pattern for no fill
    (tests/fixtures/colour_filters/)."""

    vba_type_name = "Interior"

    def __init__(self, colour: int, pattern: int) -> None:
        self.colour, self.pattern = colour, pattern

    @member
    def Color(self) -> object:
        return float(self.colour)

    @member
    def Pattern(self) -> object:
        return VBAInt(self.pattern, "Long")


class Filters(VBACollection, ExcelObject):
    vba_type_name = "Filters"

    def __init__(self, owner: AutoFilter) -> None:
        self.owner = owner

    def vba_items(self) -> list[object]:
        return [Filter(self.owner, index) for index in range(1, self.owner.state().area.columns + 1)]


class Filter(ExcelObject):
    vba_type_name = "Filter"

    def __init__(self, owner: AutoFilter, index: int) -> None:
        self.owner = owner
        self.index = index

    def _field(self) -> FieldFilter | None:
        found = self.owner.current()
        found_field = None if found is None else found.fields.get(self.index)
        if found_field is not None and found_field.operator < 0:
            raise VBAUnsupportedError("reading a filter of this kind from a file is not implemented")
        return found_field

    @member
    def On(self) -> object:
        found = self.owner.current()
        return found is not None and self.index in found.fields

    @member
    def Criteria1(self) -> object:
        found = self._field()
        # A filter on dates alone has none to read back (tests/fixtures/date_filters/).
        if found is None or found.criteria1 is MISSING:
            raise error(1004, "The filter has no criteria")
        return found.criteria1

    @member
    def Criteria2(self) -> object:
        found = self._field()
        if found is None or found.criteria2 is MISSING:
            raise error(1004, "The filter has no second criterion")
        return found.criteria2

    @member
    def Operator(self) -> object:
        found = self._field()
        return VBAInt(0 if found is None else found.operator, "Long")

    @member
    def Count(self) -> object:
        found = self._field()
        return VBAInt(0 if found is None else found.count, "Long")


def sheet_autofilter(sheet: Worksheet) -> object:
    """Worksheet.AutoFilter: the filter at the active cell -- a table's where it is in a table, else the sheet's
    own -- or Nothing."""
    from pyopenvba.apps.excel._table_filters import context_table, table_filter

    table = context_table(sheet)
    if table is not None:
        return NOTHING if table_filter(table) is None else AutoFilter(sheet, table)
    return NOTHING if sheet.auto_filter is None else AutoFilter(sheet)
