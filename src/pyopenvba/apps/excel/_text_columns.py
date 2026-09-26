"""Range.TextToColumns, as Excel splits a column of text and types what it finds.

Measured in live Excel (scripts/measure_text_to_columns.py,
tests/fixtures/text_to_columns.json):

- Only a range of one column splits; two columns are error 1004. Each
  line is split at any of the delimiters asked for, Other taking only its
  first character; with ConsecutiveDelimiter a run of them is one, but a
  field in quotes, "" among them, is a field of its own. A field in the
  text qualifier, a double quote unless it is a single one or none, keeps
  the delimiters in it and reads a doubled qualifier as one.
- Every row of the block is written across the widest line's width: a
  field that is empty, or that a shorter line lacks, clears its cell. A
  Destination moves the block's corner, on the source's sheet whatever
  sheet it is on.
- A field is typed as a cell typing it types it, spaces round a number
  and all, into a number, a date, a time, TRUE or an error, bringing the
  format typing brings where the cell's own is General; a number format
  of the cell's own stays, Text included, which takes the number all the
  same. A field starting with = is a formula. 5- is text unless
  TrailingMinusNumbers is True; -0 is 0; a number past Excel's largest
  stays text, with the scientific format it brings. DecimalSeparator and
  ThousandsSeparator read numbers written with others.
- FieldInfo gives each column, in order, a kind: text keeps the field as
  it is and formats the cell Text, a date order reads the field as that
  order's date, as _date says, a field it reads no date in being typed as
  a General column types it, and a skipped column is dropped. Pairs that
  do not number the columns 1, 2, 3... in order do something else in
  Excel, not worked out.
- Fixed width cuts each line where FieldInfo's pairs start, counted from
  the line's first character.
"""

from __future__ import annotations

import itertools
import re
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAArray, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._typing import Typed

DELIMITED, FIXED_WIDTH = 1, 2
_QUALIFIERS = {1: '"', 2: "'", -4142: None}
GENERAL, TEXT, MDY, DMY, YMD, MYD, DYM, YDM, SKIP = 1, 2, 3, 4, 5, 6, 7, 8, 9
#: The order of the day, month and year each date kind reads, by the letter each stands for.
_ORDERS = {MDY: "mdy", DMY: "dmy", YMD: "ymd", MYD: "myd", DYM: "dym", YDM: "ydm"}
#: The largest number a cell holds; past it a typed number is text.
_LARGEST = 9.99999999999999e307
_TRAILING_MINUS = re.compile(r"\s*([0-9][0-9.,]*)-\s*")
#: What a date column reads a field as: runs of digits, words, and the separators between them.
_DATE_TOKEN = re.compile(r"[0-9]+|[A-Za-z]+|[/\-. ]")
#: A run of separators, as the words of a field stand for it.
_BREAK = "|"
#: How many digits a date column puts in a day, a month and a year before the rest spill into the next.
_WIDTHS = {"d": 2, "m": 2, "y": 4}
#: Orders that read two parts as day and month, then year and month; the rest read month and day, then month
#: and year. A fixed table, not the letters' order (tests/fixtures/text_to_columns.json, the shapes cases).
_DAY_FIRST = frozenset({"dmy", "ymd", "ydm"})


def text_to_columns(target: Range, destination: object, data_type: object, text_qualifier: object,
                    consecutive: object, tab: object, semicolon: object, comma: object, space: object,
                    other: object, other_char: object, field_info: object, decimal: object, thousands: object,
                    trailing_minus: object) -> object:
    """Range.TextToColumns, its arguments as VBA gives them."""
    from pyopenvba.apps.excel import _events
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._protection import writing

    if len(target.areas) != 1 or target.first.columns != 1:
        raise error(1004, "Text to Columns works on one column at a time")
    sheet = target.sheet
    used = sheet.used_bounds()
    if used is None:
        return True
    area = target.first
    top, bottom = area.top, min(area.bottom, used[2])
    kind = DELIMITED if data_type is MISSING else int(to_integer(data_type, "Long"))
    if kind not in (DELIMITED, FIXED_WIDTH):
        raise error(1004, "Text to Columns has no such DataType")
    info = _field_info(field_info, fixed=kind == FIXED_WIDTH)
    delimiters = {mark for flag, mark in ((tab, "\t"), (semicolon, ";"), (comma, ","), (space, " "))
                  if flag is not MISSING and to_bool(flag)}
    if other is not MISSING and to_bool(other) and other_char is not MISSING and to_text(other_char):
        delimiters.add(to_text(other_char)[0])
    quote = _QUALIFIERS.get(1 if text_qualifier is MISSING else int(to_integer(text_qualifier, "Long")), '"')
    merging = consecutive is not MISSING and to_bool(consecutive)
    separators = ("." if decimal is MISSING else to_text(decimal)[:1] or ".",
                  "," if thousands is MISSING else to_text(thousands)[:1] or ",")
    minus = trailing_minus is not MISSING and to_bool(trailing_minus)
    if sheet.merged_areas:
        raise VBAUnsupportedError("TextToColumns on a sheet with merged cells is not implemented")
    # Each row's fields, or None and the value of a row that is no text to split.
    rows: list[tuple[list[str] | None, object]] = []
    for row in range(top, bottom + 1):
        cell = sheet.cell(row, area.left)
        if cell is not None and cell.formula:
            raise VBAUnsupportedError("TextToColumns of a column holding formulas is not implemented")
        value = EMPTY if cell is None else cell.value
        if not isinstance(value, str):
            # A number, a date or nothing is one field that stays as it is.
            rows.append((None, value))
            continue
        fields = [value[start:end] for start, end in _cuts(value, info)] if kind == FIXED_WIDTH \
            else _fields(value, delimiters, quote, merging)
        rows.append((fields, None))
    kinds = [one for _, one in info] if info else []
    width = max((len(_kept(fields, kinds)) for fields, _ in rows if fields is not None), default=1)
    corner_row, corner_column = top, area.left
    if isinstance(destination, Range):
        # The destination's sheet makes no difference: the block lands on the source's (tests/fixtures/).
        corner_row, corner_column = destination.first.top, destination.first.left
    written = Range(sheet, [Area(corner_row, corner_column, corner_row + len(rows) - 1,
                                 corner_column + width - 1, sheet.name)])
    with writing(sheet):
        for offset, (fields, value) in enumerate(rows):
            row = corner_row + offset
            if fields is None:
                _clear(written, row, corner_column + 1, corner_column + width - 1)
                if (row, corner_column) != (top + offset, area.left):
                    _put_value(written, row, corner_column, value)
                continue
            typed_fields = [(field, kinds[index] if index < len(kinds) else GENERAL)
                            for index, field in enumerate(fields)]
            kept = [(field, field_kind) for field, field_kind in typed_fields if field_kind != SKIP]
            for index in range(width):
                column = corner_column + index
                if index >= len(kept) or kept[index][0] == "":
                    _clear(written, row, column, column)
                    continue
                field, field_kind = kept[index]
                _write(written, row, column, field, field_kind, separators, minus)
    _events.after_edit(written)
    return True


def _field_info(value: object, *, fixed: bool) -> list[tuple[int, int]]:
    """FieldInfo's pairs, each a column's number, or where a fixed-width column starts, and its kind."""
    if value is MISSING:
        return []
    if not isinstance(value, VBAArray):
        raise error(1004, "FieldInfo takes an array of pairs")
    pairs: list[tuple[int, int]] = []
    for one in value.elements():
        if not isinstance(one, VBAArray) or len(one.elements()) != 2:
            raise error(1004, "FieldInfo takes an array of pairs")
        first, second = one.elements()
        pairs.append((int(to_integer(first, "Long")), int(to_integer(second, "Long"))))
    if fixed:
        if not pairs or pairs[0][0] != 0 or any(later[0] <= earlier[0] for earlier, later in itertools.pairwise(pairs)):
            raise VBAUnsupportedError("fixed-width FieldInfo that does not start at 0 and go up is not implemented")
    elif [number for number, _ in pairs] != list(range(1, len(pairs) + 1)):
        # Measured: Array(Array(2, 9)) drops the first column and Array(Array(3, 2)) runs the line into one.
        raise VBAUnsupportedError("FieldInfo whose pairs do not number the columns 1, 2, 3 in order is not "
                                  "implemented")
    if any(kind not in (GENERAL, TEXT, MDY, DMY, YMD, MYD, DYM, YDM, SKIP) for _, kind in pairs):
        raise VBAUnsupportedError("a FieldInfo column kind other than general, text, a date order or skip is not "
                                  "implemented")
    return pairs


def _cuts(text: str, info: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """Where each fixed-width field of ``text`` starts and ends."""
    starts = [start for start, _ in info]
    return [(start, starts[index + 1] if index + 1 < len(starts) else len(text))
            for index, start in enumerate(starts) if start < len(text)]


def _fields(text: str, delimiters: set[str], quote: str | None, merging: bool) -> list[str]:
    """A line split at its delimiters, a field in the qualifier kept whole."""
    fields: list[str] = []
    index, end = 0, len(text)
    while True:
        if quote is not None and index < end and text[index] == quote:
            value, index = _quoted(text, index + 1, quote)
            while index < end and text[index] not in delimiters:
                value += text[index]
                index += 1
            fields.append(value)
        else:
            start = index
            while index < end and text[index] not in delimiters:
                index += 1
            fields.append(text[start:index])
        if index >= end:
            return fields
        index += 1
        if merging:
            while index < end and text[index] in delimiters:
                index += 1
        if index >= end:
            fields.append("")
            return fields


def _quoted(text: str, index: int, quote: str) -> tuple[str, int]:
    """A field in quotes read from just inside them: its text, a doubled quote as one, and where it ends."""
    value = ""
    while index < len(text):
        if text[index] == quote:
            if index + 1 < len(text) and text[index + 1] == quote:
                value += quote
                index += 2
                continue
            return value, index + 1
        value += text[index]
        index += 1
    return value, index


def _kept(fields: list[str], kinds: list[int]) -> list[str]:
    return [field for index, field in enumerate(fields) if (kinds[index] if index < len(kinds) else GENERAL) != SKIP]


def _clear(target: Range, row: int, left: int, right: int) -> None:
    """The value of each cell from ``left`` to ``right`` in ``row`` taken, its format kept."""
    sheet = target.sheet
    for column in range(left, right + 1):
        cell = sheet.cell(row, column)
        if cell is None or (cell.value is EMPTY and not cell.formula):
            continue
        cell.value, cell.formula, cell.stale, cell.shared = EMPTY, "", False, None
        sheet.cell_changed(row, column)


def _put_value(target: Range, row: int, column: int, value: object) -> None:
    """A value a row's source cell held, which is no text to split, put where the block's corner moved it."""
    from pyopenvba.apps.excel._typing import Typed

    target.store_typed(row, column, Typed(value))


def _write(target: Range, row: int, column: int, field: str, kind: int, separators: tuple[str, str],
           minus: bool) -> None:
    """One field written into its cell as a column of ``kind`` takes it."""
    from pyopenvba.apps.excel._typing import Typed, format_after

    sheet = target.sheet
    current = sheet.style_at(row, column).number_format
    if kind == TEXT:
        target.store_typed(row, column, Typed(field, None if current == "@" else "@"))
        return
    if kind in _ORDERS:
        dated = _date(field, _ORDERS[kind])
        if dated is not None:
            target.store_typed(row, column, Typed(dated.value, format_after(current, dated.number_format)))
            return
        # A field the column reads no date in is typed as a General column types it.
    if field.startswith("=") and len(field) > 1:
        try:
            target.put_as_value(row, column, field)
        except Exception as exc:
            raise VBAUnsupportedError(f"TextToColumns of {field!r}, which Excel cannot read as a formula, is not "
                                      "implemented") from exc
        return
    found = _typed(field, current, separators, minus)
    target.store_typed(row, column, Typed(found.value, format_after(current, found.number_format)))


def _typed(field: str, current: str, separators: tuple[str, str], minus: bool) -> Typed:
    """A general field typed as a cell types it, read with the separators given and a trailing minus if asked."""
    from pyopenvba.apps.excel._typing import Typed, reads_fractions, typed_text

    text = field
    decimal, thousands = separators
    if (decimal, thousands) != (".", ","):
        swapped = field.replace(thousands, "\0").replace(decimal, ".").replace("\0", ",")
        if not isinstance(typed_text(swapped).value, str):
            text = swapped
    trailing = _TRAILING_MINUS.fullmatch(text)
    if minus and trailing is not None:
        text = "-" + trailing.group(1)
    found = typed_text(text, fractions=reads_fractions(current))
    if isinstance(found.value, str):
        return Typed(field, found.number_format)
    if isinstance(found.value, float):
        if abs(found.value) > _LARGEST:
            # Past the largest number a cell holds the field stays text, with the format it brought.
            return Typed(field, found.number_format)
        if found.value == 0.0:
            return Typed(0.0, found.number_format)
    return found


def _date(field: str, order: str) -> Typed | None:
    """The date a column of ``order`` reads in ``field``, or None where it reads none and General typing takes it.

    The column fills its day, month and year in its order: each run of digits goes into the part it has reached,
    two digits to a day or a month and four to a year, the rest spilling into the next part, and a separator --
    a slash, dash, point or space -- moves on to the next. A month's name fills a month. A run of digits with no
    separator is cut by its length instead. Three parts make a date if they are one; two go by a fixed table:
    day and month, then year and month, in _DAY_FIRST's orders, month and day, then month and year, in the others;
    a day read so falls in the current year (tests/fixtures/text_to_columns.json).
    """
    body = field.strip(" ")
    parts = _digit_run(body, order) if body.isascii() and body.isdigit() else _parts(body, order)
    if parts is None:
        return None
    if len(parts) == 3:
        return _three_parts(dict(parts))
    return _two_parts(parts[0][1], parts[1][1], day_first=order in _DAY_FIRST) if len(parts) == 2 else None


def _digit_run(body: str, order: str) -> list[tuple[str, str]] | None:
    """A run of digits cut into a column's parts by its length: four digits two and two, five with one for the
    month, six two to each part, eight at each part's width; any other run is no date."""
    widths = {4: {"d": 2, "m": 2, "y": 2}, 5: {"d": 2, "m": 1, "y": 2}, 6: {"d": 2, "m": 2, "y": 2},
              8: _WIDTHS}.get(len(body))
    if widths is None:
        return None
    parts: list[tuple[str, str]] = []
    at = 0
    for letter in order:
        if at >= len(body):
            break
        parts.append((letter, body[at:at + widths[letter]]))
        at += widths[letter]
    return parts


def _parts(body: str, order: str) -> list[tuple[str, str]] | None:
    """The parts a column fills from a field with separators or a month's name, in the order it fills them.

    A run of separators counts as one and a separator after the last part as none; one before the first part,
    a month's name straight after digits, and a name Excel's date columns do not know, Sept, leave the field to
    General typing. Four digits straight after a month's name fill the next two parts, two each: May2020 is 20
    May 2020 where the name is the column's first part (tests/fixtures/text_to_columns.json).
    """
    from pyopenvba.apps.excel._typing import month_number

    tokens = _DATE_TOKEN.findall(body)
    if "".join(tokens) != body:
        return None
    words: list[str] = []
    for token in tokens:
        if token not in "/-. ":
            words.append(token)
        elif not words:
            return None
        elif words[-1] != _BREAK:
            words.append(_BREAK)
    if words and words[-1] == _BREAK:
        words.pop()
    parts: list[tuple[str, str]] = []
    pending = ""
    for index, word in enumerate(words):
        previous = words[index - 1] if index else _BREAK
        if word == _BREAK:
            parts.append((order[len(parts)], pending))
            pending = ""
            if len(parts) == 3:
                return None
        elif word.isalpha():
            if previous != _BREAK or word.lower() == "sept" or month_number(word) is None \
                    or order[len(parts)] != "m":
                return None
            pending = word
        elif previous != _BREAK:
            if len(word) != 4:
                raise VBAUnsupportedError(f"TextToColumns of {body!r} in a date column, a month's name run into "
                                          "other than four digits, is not implemented")
            parts.append((order[len(parts)], pending))
            if len(parts) == 3:
                return None
            parts.append((order[len(parts)], word[:2]))
            pending = word[2:]
        else:
            for digit in word:
                if len(pending) == _WIDTHS[order[len(parts)]]:
                    # The part is full: the rest of the digits spill into the next one.
                    parts.append((order[len(parts)], pending))
                    pending = ""
                    if len(parts) == 3:
                        return None
                pending += digit
    if len(parts) == 3 or not pending:
        return None
    parts.append((order[len(parts)], pending))
    return parts


def _three_parts(parts: dict[str, str]) -> Typed | None:
    """A day, a month and a year as one date: m/d/yyyy, or d-mmm-yy where the month was named."""
    from pyopenvba.apps.excel._typing import Typed, date_serial, month_number, typed_year

    year = parts["y"]
    if len(year) == 3 or not parts["d"].isdigit():
        return None
    month = month_number(parts["m"]) if parts["m"].isalpha() else int(parts["m"])
    serial = date_serial(typed_year(year), month or 0, int(parts["d"])) if year.isdigit() else None
    if serial is None:
        return None
    return Typed(serial, "d-mmm-yy" if parts["m"].isalpha() else "m/d/yyyy")


def _two_parts(first: str, second: str, *, day_first: bool) -> Typed | None:
    """Two parts as a date by the table: a day and its month in the current year (d-mmm), else a month and its
    year (mmm-yy), the day first or the month first as the column's order has it."""
    from pyopenvba.apps.excel._typing import Typed, date_serial, month_number, this_year, typed_year

    def month_of(text: str) -> int | None:
        return month_number(text) if text.isalpha() else int(text)

    if day_first:
        day, month_text, year_text = first, second, first
    else:
        day, month_text, year_text = second, first, second
    month = month_of(month_text)
    if month is None:
        return None
    if day.isdigit():
        serial = date_serial(this_year(), month, int(day))
        if serial is not None:
            return Typed(serial, "d-mmm")
    if year_text.isdigit() and len(year_text) != 3:
        serial = date_serial(typed_year(year_text), month, 1)
        if serial is not None:
            return Typed(serial, "mmm-yy")
    return None


__all__ = ["text_to_columns"]
