"""Application's custom lists, as Excel keeps them and as AutoFill and Range.Sort read them.

Measured in live Excel (scripts/measure_custom_lists.py,
tests/fixtures/custom_lists.json).

- The first four lists are Excel's own: the short and long day names from
  Sunday and the short and long month names. They cannot be deleted. The
  lists a macro adds follow them, numbered on from 5, for the session.
- AddCustomList takes an array, its text kept as given, spaces and empty
  text included, and its numbers left out; or a range, its blank cells
  left out, making one list of each row where ByRow is True, of each
  column where it is False, and, left out, of each line along the
  range's longer side. A list with no text is error 1004. A list equal to
  one there already, case aside, is not added: silently from an array,
  error 1004 from a range.
- GetCustomListNum finds a whole list, case aside, or 0; it takes no
  range. GetCustomListContents gives a list from 1, error 13 for a number
  that names none. DeleteCustomList of Excel's own lists is error 1004,
  and of a number past the last does nothing.
- AutoFill steps along a list from a cell holding one of its items, case
  aside, in the case the cell shows; Range.Sort's OrderCustom n orders by
  list n - 1, its items first in the list's order, case aside, then the
  rest as a plain sort orders them, all of it reversed for descending; a
  number past the lists sorts plainly.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAArray, error, to_bool, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application

BUILT_IN: tuple[tuple[str, ...], ...] = (
    ("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"),
    ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"),
    ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"),
    ("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November",
     "December"),
)


def lists(app: Application) -> list[tuple[str, ...]]:
    """Every list, Excel's own first."""
    return [*BUILT_IN, *app.custom_lists]


def added(app: Application) -> list[tuple[str, ...]]:
    """The lists a macro added, which AutoFill reads; Excel's own it fills by their own rules (see _autofill)."""
    return list(app.custom_lists)


def _failed(name: str) -> Exception:
    return error(1004, f"Method '{name}' of object '_Application' failed")


def _same(one: tuple[str, ...], other: tuple[str, ...]) -> bool:
    return len(one) == len(other) and all(a.lower() == b.lower() for a, b in zip(one, other, strict=True))


def number_of(app: Application, items: tuple[str, ...]) -> int:
    return next((index for index, one in enumerate(lists(app), start=1) if _same(one, items)), 0)


def _array_items(value: VBAArray) -> tuple[str, ...]:
    found: list[str] = []
    for item in value.elements():
        if isinstance(item, bool) or isinstance(item, (int, float)):
            continue
        if not isinstance(item, str):
            raise VBAUnsupportedError("AddCustomList of an array holding anything but text and numbers is not "
                                      "implemented")
        found.append(item)
    return tuple(found)


def _range_lists(target: object, by_row: object) -> list[tuple[str, ...]]:
    from pyopenvba.apps.excel._model import Range

    assert isinstance(target, Range)
    if len(target.areas) != 1:
        raise VBAUnsupportedError("AddCustomList of a range of several areas is not implemented")
    area, sheet = target.first, target.sheet
    rows = (to_bool(by_row) if by_row is not MISSING else area.columns > area.rows)
    lines = ([[(row, column) for column in range(area.left, area.right + 1)] for row in range(area.top, area.bottom + 1)]
             if rows else
             [[(row, column) for row in range(area.top, area.bottom + 1)] for column in range(area.left, area.right + 1)])
    made: list[tuple[str, ...]] = []
    for line in lines:
        items: list[str] = []
        for row, column in line:
            value = sheet.book.calculator.value_of(sheet.name, row, column)
            if value is EMPTY or value is None or isinstance(value, float | int) and not isinstance(value, bool):
                # A number is left out of the list (tests/fixtures/custom_orders.json).
                continue
            if not isinstance(value, str):
                raise VBAUnsupportedError("AddCustomList of a range holding logical values or errors is not "
                                          "implemented")
            items.append(value)
        # A line of numbers alone makes no list, which is error 1004 (tests/fixtures/custom_lists.json).
        made.append(tuple(items))
    return made


def add(app: Application, source: object, by_row: object) -> None:
    from pyopenvba.apps.excel._model import Range

    if isinstance(source, Range):
        made = _range_lists(source, by_row)
        if any(not items for items in made):
            raise _failed("AddCustomList")
        fresh = [items for items in made if not number_of(app, items)]
        if not fresh:
            raise _failed("AddCustomList")
        if len(fresh) != len(made):
            raise VBAUnsupportedError("AddCustomList of a range some of whose lists are there already is not "
                                      "implemented")
        app.custom_lists.extend(fresh)
        return
    if not isinstance(source, VBAArray):
        raise VBAUnsupportedError("AddCustomList of anything but an array or a range is not implemented")
    items = _array_items(source)
    if not items:
        raise _failed("AddCustomList")
    if not number_of(app, items):
        app.custom_lists.append(items)


def contents(app: Application, number: object) -> VBAArray:
    index = int(to_integer(number, "Long"))
    known = lists(app)
    if not 1 <= index <= len(known):
        raise error(13)
    items = known[index - 1]
    return VBAArray([(1, len(items))], items=list(items))


def find(app: Application, source: object) -> int:
    from pyopenvba.apps.excel._model import Range

    if isinstance(source, Range):
        raise _failed("GetCustomListNum")
    if not isinstance(source, VBAArray):
        raise VBAUnsupportedError("GetCustomListNum of anything but an array is not implemented")
    if any(not isinstance(item, str) for item in source.elements()):
        raise VBAUnsupportedError("GetCustomListNum of an array holding anything but text is not implemented")
    return number_of(app, tuple(str(item) for item in source.elements()))


def delete(app: Application, number: object) -> None:
    index = int(to_integer(number, "Long"))
    if index < 1:
        raise VBAUnsupportedError("DeleteCustomList of a number below 1 is not implemented")
    if index <= len(BUILT_IN):
        raise _failed("DeleteCustomList")
    if index - len(BUILT_IN) <= len(app.custom_lists):
        del app.custom_lists[index - len(BUILT_IN) - 1]


def sort_order(app: Application, order_custom: object) -> tuple[str, ...] | None:
    """The list Range.Sort's OrderCustom orders by; None for a plain sort."""
    if order_custom is MISSING:
        return None
    index = int(to_integer(order_custom, "Long"))
    if index < 1:
        raise VBAUnsupportedError("Range.Sort's OrderCustom below 1 is not implemented")
    known = lists(app)
    if index == 1 or index - 1 > len(known):
        return None
    return known[index - 2]
