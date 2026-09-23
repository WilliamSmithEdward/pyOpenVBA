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

What is not modelled reports itself: colour, icon and date filters, the
other dynamic filters, AutoFilter.Sort, a hidden drop-down, and text
beyond ASCII where it is compared.
"""

from __future__ import annotations

import math
import re
import uuid
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, parse_reference, quote_sheet
from pyopenvba.apps.excel._model import ExcelObject, NameEntry, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._display import UndisplayableError, shown
from pyopenvba.formula._values import ExcelError, number_text
from pyopenvba.interpreter._objects import VBACollection, member, method
from pyopenvba.interpreter._values import (EMPTY, MISSING, NOTHING, VBAArray, VBADate, VBAInt, error, to_bool,
                                           to_integer, to_number, to_text)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

AND, OR, TOP_ITEMS, BOTTOM_ITEMS, TOP_PERCENT, BOTTOM_PERCENT, VALUES = 1, 2, 3, 4, 5, 6, 7
DYNAMIC = 11
ABOVE_AVERAGE, BELOW_AVERAGE = 33, 34
_OPERATORS = ("<>", ">=", "<=", "=", ">", "<")
#: The file's names for the operators a customFilter can hold; = has none.
_FILE_OPERATORS = {"<>": "notEqual", ">": "greaterThan", ">=": "greaterThanOrEqual", "<": "lessThan",
                   "<=": "lessThanOrEqual"}


# --- what a cell shows the filter --------------------------------------------------------------------


@dataclass(frozen=True)
class _Seen:
    """A cell as a criterion looks at it."""

    shown: str
    number: float | None
    text: str | None
    error: bool


def _seen(sheet: Worksheet, row: int, column: int) -> _Seen:
    cell = sheet.cells_.get((row, column))
    if cell is None:
        return _Seen("", None, None, False)
    value = sheet.book.calculator.value_of(sheet.name, row, column) if cell.formula else cell.value
    if isinstance(value, VBADate):
        value = value.serial
    if value is EMPTY:
        return _Seen("", None, None, False)
    if isinstance(value, ExcelError):
        return _Seen(value.name, None, None, True)
    if isinstance(value, bool):
        return _Seen("TRUE" if value else "FALSE", None, None, False)
    if isinstance(value, (int, float)):
        try:
            text, _ = shown(float(value), cell.number_format)
        except UndisplayableError:
            text = "#"
        return _Seen(text, float(value), None, False)
    text = str(value)
    return _Seen(text, None, text, False)


# --- criteria -----------------------------------------------------------------------------------------


@dataclass(frozen=True)
class _Criterion:
    """One criterion as typed: its operator, what follows it, that read as a number, and how it reads back."""

    operator: str
    operand: str
    number: float | None
    stored: str

    @property
    def plain(self) -> bool:
        """An = with no wildcard in it, which the file keeps as a value in a list."""
        return self.operator == "=" and not _wild(self.operand)

    def matches(self, seen: _Seen) -> bool:
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
        return _compare(_collated(seen.text), self.operator, _collated(self.operand))


def _criterion(value: object) -> _Criterion:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        text = number_text(float(value), formula=True)
    else:
        text = to_text(value)
    body = text.strip(" ")
    operator = next((one for one in _OPERATORS if body.startswith(one)), "")
    return _made(operator or "=", body[len(operator):].strip(" "))


def _made(operator: str, operand: str) -> _Criterion:
    """A criterion from its operator and operand, reading the operand as a number where it is one.

    A number reads back spelled afresh, a date or time as its serial: >2.50 is >2.5.
    """
    from pyopenvba.apps.excel._typing import typed_text

    number: float | None = None
    stored = operand
    if operator != "=" and operand:
        typed = typed_text(operand)
        if isinstance(typed.value, float):
            number = typed.value
            stored = number_text(number, formula=True)
    return _Criterion(operator, operand, number, operator + stored)


def _compare(left: object, operator: str, right: object) -> bool:
    if operator == ">":
        return left > right  # type: ignore[operator]
    if operator == ">=":
        return left >= right  # type: ignore[operator]
    if operator == "<":
        return left < right  # type: ignore[operator]
    return left <= right  # type: ignore[operator]


def _collated(text: str) -> object:
    from pyopenvba.apps.excel._sort import text_key

    _plain(text)
    return text_key(text, False)


def _plain(text: str) -> None:
    if any(ord(char) > 126 and char.isalpha() for char in text):
        raise VBAUnsupportedError("filtering text with letters beyond ASCII is not implemented")


def _wild(pattern: str) -> bool:
    return re.search(r"(?<!~)[*?]", pattern) is not None


def _literal(pattern: str) -> str:
    """What a pattern with no live wildcard means: ~* and ~? are * and ?."""
    return re.sub(r"~([*?])", r"\1", pattern)


def _escaped(value: str) -> str:
    """A value from the file's list as a criterion reads back: its * and ? escaped."""
    return re.sub(r"([*?])", r"~\1", value)


def _like(seen: _Seen, pattern: str) -> bool:
    """Whether what a cell shows matches a criterion's text: any case, * and ? wild, ~ before one to mean it.

    A pattern with a wildcard in it matches text only, never a number however it shows.
    """
    if _wild(pattern) and seen.text is None:
        return False
    _plain(seen.shown)
    _plain(pattern)
    parts: list[str] = []
    index = 0
    while index < len(pattern):
        char = pattern[index]
        if char == "~" and index + 1 < len(pattern) and pattern[index + 1] in "*?":
            parts.append(re.escape(pattern[index + 1]))
            index += 2
            continue
        parts.append(".*" if char == "*" else "." if char == "?" else re.escape(char))
        index += 1
    return re.fullmatch("".join(parts), seen.shown, re.IGNORECASE | re.DOTALL) is not None


# --- the file's spelling -------------------------------------------------------------------------------


def _attribute(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _file_number(value: float) -> str:
    """A number as the file writes it: as short as reads back exactly."""
    text = repr(value)
    return text[:-2] if text.endswith(".0") else text


def _values_xml(values: list[str], blank: bool) -> str:
    flag = ' blank="1"' if blank else ""
    if not values:
        return f"<filters{flag}/>"
    return f"<filters{flag}>" + "".join(f'<filter val="{_attribute(value)}"/>' for value in values) + "</filters>"


def _custom_xml(criterion: _Criterion) -> str:
    if criterion.operator == "=":
        return f'<customFilter val="{_attribute(criterion.operand)}"/>'
    if criterion.operator == "<>" and criterion.operand == "":
        return '<customFilter operator="notEqual" val=" "/>'
    value = _file_number(criterion.number) if criterion.number is not None else criterion.operand
    return f'<customFilter operator="{_FILE_OPERATORS[criterion.operator]}" val="{_attribute(value)}"/>'


def _criteria_xml(criteria: list[_Criterion], *, joined: int) -> str:
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
    test: Callable[[_Seen], bool]
    count: int = 1
    xml: str = ""
    attributes: str = ""


def _one(criterion: _Criterion) -> FieldFilter:
    return FieldFilter(0, criterion.stored, MISSING, criterion.matches, xml=_criteria_xml([criterion], joined=0))


def _two(first: _Criterion, second: _Criterion, joined: int) -> FieldFilter:
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
        return _values(criteria1)
    if kind in (TOP_ITEMS, BOTTOM_ITEMS, TOP_PERCENT, BOTTOM_PERCENT):
        return _top(sheet, area, index, kind, criteria1)
    if kind == DYNAMIC:
        return _dynamic(sheet, area, index, criteria1)
    if kind not in (0, AND, OR):
        raise VBAUnsupportedError(f"AutoFilter with Operator {kind} is not implemented")
    if criteria2 is MISSING:
        return _one(_criterion(criteria1))
    first = _criterion("=" if criteria1 is MISSING else criteria1)
    return _two(first, _criterion(criteria2), OR if kind == OR else AND)


def _values(criteria1: object) -> FieldFilter:
    from pyopenvba.apps.excel._sort import text_key

    items = [to_text(item) for item in (criteria1.elements() if isinstance(criteria1, VBAArray) else [criteria1])]
    for item in items:
        _plain(item)
    items.sort(key=lambda item: text_key(item, False))
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
    xml = f'<top10{flags} val="{_file_number(given)}" filterVal="{_file_number(threshold)}"/>'
    stored = (">=" if top else "<=") + number_text(threshold, formula=True)
    if top:
        return FieldFilter(kind, stored, MISSING, lambda seen: seen.number is not None and seen.number >= threshold,
                           xml=xml)
    return FieldFilter(kind, stored, MISSING, lambda seen: seen.number is not None and seen.number <= threshold, xml=xml)


def _dynamic(sheet: Worksheet, area: Area, index: int, criteria1: object) -> FieldFilter:
    which = int(to_integer(criteria1, "Long"))
    if which not in (ABOVE_AVERAGE, BELOW_AVERAGE):
        raise VBAUnsupportedError("AutoFilter's dynamic filters other than above and below average are not implemented")
    numbers = _numbers(sheet, area, index)
    return _averaged(which, sum(numbers) / len(numbers))


def _averaged(which: int, mean: float) -> FieldFilter:
    kind = "aboveAverage" if which == ABOVE_AVERAGE else "belowAverage"
    xml = f'<dynamicFilter type="{kind}" val="{_file_number(mean)}"/>'
    if which == ABOVE_AVERAGE:
        return FieldFilter(DYNAMIC, VBAInt(which, "Long"), MISSING,
                           lambda seen: seen.number is not None and seen.number > mean, xml=xml)
    return FieldFilter(DYNAMIC, VBAInt(which, "Long"), MISSING,
                       lambda seen: seen.number is not None and seen.number < mean, xml=xml)


def _opaque(xml: str) -> FieldFilter:
    """A column filter read from a file that the model does not follow: kept as it is, and not applied."""

    def refuse(seen: _Seen) -> bool:
        raise VBAUnsupportedError("applying a colour, icon or date filter is not implemented")

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
        columns = "".join(f'<filterColumn colId="{index - 1}"{one.attributes}>{one.xml}</filterColumn>'
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
    sheet = target.sheet
    if sub_field is not MISSING:
        raise VBAUnsupportedError("AutoFilter's SubField is not implemented")
    hidden_button = visible_dropdown is not MISSING and not to_bool(visible_dropdown)
    existing = sheet.auto_filter
    if field_argument is MISSING:
        if hidden_button:
            raise VBAUnsupportedError("hiding every drop-down of a filter is not implemented")
        if existing is not None:
            existing.show_all(sheet)
            sheet.auto_filter = None
        else:
            _new_filter(sheet, _filter_area(target))
        _changed(sheet)
        return True
    asked = _filter_area(target) if existing is None else existing.area
    # Filtering takes in the rows of data grown on below the range.
    area = grown(sheet, asked)
    index = int(to_integer(field_argument, "Long"))
    if not 1 <= index <= area.columns:
        raise error(1004, "AutoFilter's field is outside the range")
    made = _field(sheet, area, index, criteria1, operator, criteria2)
    if made is None and hidden_button:
        raise VBAUnsupportedError("hiding a drop-down without criteria is not implemented")
    if existing is None:
        # A new filter's name keeps the range it was asked for; filtering again names the grown one.
        existing = _new_filter(sheet, asked)
    else:
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
    if operator is not MISSING and int(to_integer(operator, "Long")) == VALUES:
        return float(index)
    return True


def show_all_data(sheet: Worksheet) -> None:
    """Worksheet.ShowAllData and AutoFilter.ShowAllData."""
    found = sheet.auto_filter
    if found is None or not found.fields:
        raise error(1004, "ShowAllData method of Worksheet class failed")
    found.fields.clear()
    found.show_all(sheet)
    _changed(sheet)


def set_mode(sheet: Worksheet, value: object) -> None:
    """Worksheet.AutoFilterMode = False turns the filter off; True cannot turn one on."""
    if to_bool(value):
        raise VBAUnsupportedError("setting AutoFilterMode to True is not implemented")
    if sheet.auto_filter is not None:
        sheet.auto_filter.show_all(sheet)
        sheet.auto_filter = None
        _changed(sheet)


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
        read = _read_column(match.group(2) or "")
        read.attributes = others
        found.fields[int(colid.group(1)) + 1] = read
        last = match.end()
    found.tail = body[last:]
    return found


def _read_column(xml: str) -> FieldFilter:
    """One column's criteria from the file, as Excel reads them back."""
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
        read = _two(_made("=", _escaped(items[0])), _made("=", _escaped(items[1])), OR)
    elif items:
        read = _one(_made("=", _escaped(items[0])))
    else:
        return _opaque(xml)
    read.xml = xml
    return read


def _read_custom(xml: str, attributes: str, body: str) -> FieldFilter:
    joined = AND if _xml_attributes(attributes).get("and") in ("1", "true") else OR
    criteria: list[_Criterion] = []
    for one in re.findall(r"<customFilter\b([^>]*)/>", body):
        found = _xml_attributes(one)
        operator = next((symbol for symbol, name in _FILE_OPERATORS.items() if name == found.get("operator")), "=")
        value = found.get("val", "")
        if operator == "<>" and value == " ":
            criteria.append(_made(operator, ""))
            continue
        if operator != "=":
            try:
                number = float(value)
            except ValueError:
                pass
            else:
                # The file keeps the number in full; it reads back in fifteen digits.
                spelled = number_text(number, formula=True)
                criteria.append(_Criterion(operator, spelled, number, operator + spelled))
                continue
        criteria.append(_made(operator, value))
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
    vba_type_name = "AutoFilter"

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def state(self) -> SheetFilter:
        found = self.sheet.auto_filter
        if found is None:
            raise error(1004, "The sheet has no AutoFilter")
        return found

    @member
    def Range(self) -> object:
        return Range(self.sheet, [grown(self.sheet, self.state().area)])

    @member
    def Filters(self, Index: object = MISSING) -> object:
        found = Filters(self)
        return found if Index is MISSING else found.vba_lookup(Index, found.vba_items())

    @member
    def FilterMode(self) -> object:
        return bool(self.state().fields)

    @method
    def ApplyFilter(self) -> object:
        found = self.state()
        area = grown(self.sheet, found.area)
        if area != found.area:
            found.area = area
            _changed(self.sheet)
        found.apply(self.sheet)
        return EMPTY

    @method
    def ShowAllData(self) -> object:
        show_all_data(self.sheet)
        return EMPTY


class Filters(VBACollection, ExcelObject):
    vba_type_name = "Filters"

    def __init__(self, owner: AutoFilter) -> None:
        self.owner = owner

    def vba_items(self) -> list[object]:
        return [Filter(self.owner.sheet, index) for index in range(1, self.owner.state().area.columns + 1)]


class Filter(ExcelObject):
    vba_type_name = "Filter"

    def __init__(self, sheet: Worksheet, index: int) -> None:
        self.sheet = sheet
        self.index = index

    def _field(self) -> FieldFilter | None:
        found = self.sheet.auto_filter
        found_field = None if found is None else found.fields.get(self.index)
        if found_field is not None and found_field.operator < 0:
            raise VBAUnsupportedError("reading a colour, icon or date filter is not implemented")
        return found_field

    @member
    def On(self) -> object:
        found = self.sheet.auto_filter
        return found is not None and self.index in found.fields

    @member
    def Criteria1(self) -> object:
        found = self._field()
        if found is None:
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
    """Worksheet.AutoFilter: the sheet's AutoFilter object, or Nothing."""
    return NOTHING if sheet.auto_filter is None else AutoFilter(sheet)
