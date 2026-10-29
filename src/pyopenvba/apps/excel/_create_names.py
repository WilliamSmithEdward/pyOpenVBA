"""Range.CreateNames: names made from a range's labels, as Excel makes them; and Range.ListNames, which pastes
them (see list_names).

Measured in live Excel (scripts/measure_create_names.py,
tests/fixtures/create_names.json).

- Top, Left, Bottom and Right say which edges of the range hold labels.
  The cells between the label edges are the data. A label on a used top
  or bottom edge names its column's data cells, and one on a used left or
  right edge its row's. A corner where a used row edge meets a used column
  edge names all the data. A side left out, or given Null, is guessed: the
  top edge holds labels when its second cell does, the left when its
  second cell down does, and the bottom and right likewise. In a range one
  column wide only the top and bottom are guessed, each from its own cell;
  in one a row high only the left and right. A number given is True
  unless it is 0; text is error 1004.
- A label is text, a formula's text among them, or a date as it shows;
  numbers, logical values, errors and blank text are none. It is cut to
  255 characters, the characters a name cannot hold are taken off its
  ends and turned into _ inside it, and a label that is still no name
  takes a _ after it, or else before it: A1_, _R1C1, _3rd, _1_1_2026.
- The names are the workbook's, each referring to its cells by their
  absolute address on the sheet. A name made twice, or one the workbook
  has already, is redefined with the later spelling. CreateNames returns
  True, names made or not; a range of one cell or of several areas is
  error 1004.
"""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.apps.excel._names import valid
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, VBADate, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

_NOT_VALID = ("This selection isn't valid. Make sure the copy and paste areas don't overlap unless they are the same "
              "size and shape.")
#: What a name cannot hold anywhere: all but letters, digits, _, ., \ and ?.
_OUTSIDE = r"[^\w.\\?]"


def create_names(target: Range, top: object, left: object, bottom: object, right: object) -> bool:
    if target.single or len(target.areas) != 1:
        raise error(1004, _NOT_VALID)
    sheet, area = target.sheet, target.first
    sides = [_side(given) for given in (top, left, bottom, right)]
    wide, tall = area.columns > 1, area.rows > 1
    second_column = area.left + 1 if wide else area.left
    second_row = area.top + 1 if tall else area.top
    # Each side guessed from the cell beside its corner; a range one column wide has no left or right labels, one
    # row high no top or bottom ones.
    guesses = [tall and _label(sheet, area.top, second_column) is not None,
               wide and _label(sheet, second_row, area.left) is not None,
               tall and _label(sheet, area.bottom, second_column) is not None,
               wide and _label(sheet, second_row, area.right) is not None]
    at_top, at_left, at_bottom, at_right = (guess if side is None else side for side, guess in zip(sides, guesses))
    rows = range(area.top + at_top, area.bottom - at_bottom + 1)
    columns = range(area.left + at_left, area.right - at_right + 1)
    made: list[tuple[int, int, Area]] = []
    for column in columns:
        cells = Area(rows.start, column, rows.stop - 1, column) if rows else None
        made += [(edge, column, cells) for edge, used in ((area.top, at_top), (area.bottom, at_bottom))
                 if used and cells is not None]
    for row in rows:
        cells = Area(row, columns.start, row, columns.stop - 1) if columns else None
        made += [(row, edge, cells) for edge, used in ((area.left, at_left), (area.right, at_right))
                 if used and cells is not None]
    data = Area(rows.start, columns.start, rows.stop - 1, columns.stop - 1) if rows and columns else None
    for row, row_used in ((area.top, at_top), (area.bottom, at_bottom)):
        for column, column_used in ((area.left, at_left), (area.right, at_right)):
            if row_used and column_used and data is not None:
                made.append((row, column, data))
    for row, column, cells in made:
        _name(sheet, row, column, cells)
    return True


def _side(given: object) -> bool | None:
    """A side as CreateNames takes it: None to guess, for a side left out or Null."""
    if given is MISSING or given is NULL:
        return None
    if isinstance(given, bool):
        return given
    if isinstance(given, (int, float)):
        return given != 0
    raise error(1004, "Unable to get the CreateNames property of the Range class")


def _label(sheet: Worksheet, row: int, column: int) -> str | None:
    """A cell's label: its text, or a date as it shows; None for anything else."""
    from pyopenvba.apps.excel._model import Range

    from pyopenvba.apps.excel._typing import is_date_format

    value = sheet.book.calculator.value_of(sheet.name, row, column)
    # A number shown as a date is one, as Value reads it.
    if isinstance(value, VBADate) or isinstance(value, float) \
            and is_date_format(sheet.style_at(row, column).number_format):
        return str(Range(sheet, [Area(row, column, row, column, sheet.name)]).Text())
    if isinstance(value, str) and value.strip():
        return value
    return None


def name_of(label: str) -> str | None:
    """The name a label makes, or None where it makes none."""
    cut = re.sub(rf"^{_OUTSIDE}+|{_OUTSIDE}+$", "", label[:255])
    cleaned = re.sub(_OUTSIDE, "_", cut)
    if not cleaned:
        return None
    return next((name for name in (cleaned, cleaned + "_", "_" + cleaned) if valid(name)), None)


def _name(sheet: Worksheet, row: int, column: int, cells: Area) -> None:
    from pyopenvba.apps.excel._model import Range

    label = _label(sheet, row, column)
    name = name_of(label) if label is not None else None
    if name is not None:
        sheet.book.names_.Add(Name=name, RefersTo=Range(sheet, [Area(cells.top, cells.left, cells.bottom,
                                                                      cells.right, sheet.name)]))


def list_names(target: Range) -> object:
    """Range.ListNames: the workbook's visible names pasted from the range's first cell down, in the Names
    collection's order, each beside its RefersTo as text; a sheet's own name bare, and hidden names left out. Only
    the cells it writes change, whatever the range's size; several areas are error 1004
    (tests/fixtures/range_text_helpers.json)."""
    from pyopenvba._a1 import split_sheet
    from pyopenvba.apps.excel import _events
    from pyopenvba.apps.excel._model import DefinedName
    from pyopenvba.apps.excel._model import Range as Cells
    from pyopenvba.apps.excel._protection import writing
    from pyopenvba.apps.excel._typing import Typed

    if len(target.areas) != 1:
        raise error(1004, "This can't be done on a multiple range selection.\n\nSelect a single range and try again.")
    sheet = target.sheet
    shown: list[tuple[str, str]] = []
    for item in sheet.book.names_.vba_items():
        assert isinstance(item, DefinedName)
        if not item.entry.visible:
            continue
        scope, bare = split_sheet(item.entry.name)
        if scope and scope.casefold() != sheet.name.casefold():
            raise VBAUnsupportedError("ListNames in a workbook with another sheet's own names is not implemented")
        shown.append((bare, item.shown_refers_to()))
    if not shown:
        return EMPTY
    top, left = target.first.top, target.first.left
    written = Cells(sheet, [Area(top, left, top + len(shown) - 1, left + 1, sheet.name)])
    with writing(sheet):
        for offset, (name, refers_to) in enumerate(shown):
            written.store_typed(top + offset, left, Typed(name))
            written.store_typed(top + offset, left + 1, Typed(refers_to))
    _events.after_edit(written)
    return EMPTY
