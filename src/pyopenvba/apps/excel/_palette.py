"""Workbook.Colors and Workbook.ResetColors: the workbook's palette of 56 colours, which ColorIndex counts in.

Measured in live Excel (tests/fixtures/workbook_colors.json, scripts/measure_workbook_colors.py):

- Colors with no index is the palette as an array of Doubles from 1 to 56; Colors(n) one of them. An index is cut
  to a whole number, from text too ("3", 3.6 and "4.9" read slots 3, 3 and 4); one outside 1 to 56 is error 1004.
- Colors(n) = value takes a number cut to a whole one and keeps its low 24 bits, so -1 is 16777215; True is 1;
  text, even "255", and Empty are 0. What Null leaves is not measured and reports itself unsupported.
- Colors = array sets every slot from the array's first 56 values; an empty array is error 13 and anything but an
  array error 1004. Fewer than 56 reports itself unsupported: Excel answered error 1004 once and crashed once.
- A tab coloured by a slot keeps Excel's own colour for it. ColorIndex of a colour a changed palette does not
  hold reports itself unsupported: Excel answered 3 for a red near the palette's red, and -5 for a red whose slot
  a macro had changed.
- A cell, a font or a border coloured by a slot shows the slot's new colour; ResetColors puts back Excel's own
  palette. A save writes a changed palette as styles.xml's indexedColors, and none where it is Excel's own.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.apps.excel import _styles as styles
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, VBAArray, error, to_number

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook

_FAILED = "Method 'Colors' of object '_Workbook' failed"


def _slot(index: object) -> int:
    try:
        slot = int(float(to_number(index)))
    except (ValueError, OverflowError):
        raise error(1004, _FAILED) from None
    if not 1 <= slot <= 56:
        raise error(1004, _FAILED)
    return slot


def _bgr(rrggbb: str) -> float:
    return float(styles.bgr(rrggbb))


def colors(book: Workbook, index: object) -> object:
    palette = book.stylesheet.colors.indexed
    if index is MISSING:
        return VBAArray([(1, 56)], items=[_bgr(entry) for entry in palette[8:64]])
    return _bgr(palette[_slot(index) + 7])


def _value(value: object) -> int:
    """A colour as Colors(n) = value takes it."""
    if value is NULL:
        raise VBAUnsupportedError("Workbook.Colors(n) set to Null is not implemented")
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, str) or value is EMPTY:
        return 0
    return int(float(to_number(value))) & 0xFFFFFF


def _rrggbb(value: int) -> str:
    """A Long colour, BGR, as RRGGBB."""
    return f"{value & 0xFF:02X}{(value >> 8) & 0xFF:02X}{(value >> 16) & 0xFF:02X}"


def set_colors(book: Workbook, index: object, value: object) -> None:
    current = book.stylesheet.colors
    palette = list(current.indexed)
    if index is not MISSING:
        palette[_slot(index) + 7] = _rrggbb(_value(value))
    else:
        if not isinstance(value, VBAArray):
            raise error(1004, _FAILED)
        if value.size == 0:
            raise error(13, "Type mismatch")
        items = value.elements()
        if len(items) < 56:
            # Excel answered error 1004 for three colours once, and crashed on them once.
            raise VBAUnsupportedError("Workbook.Colors set to an array of fewer than 56 colours is not implemented")
        if any(isinstance(item, str) or item is NULL for item in items[:56]):
            raise VBAUnsupportedError("Workbook.Colors set to an array holding text or Null is not implemented")
        palette[8:64] = [_rrggbb(_value(item)) for item in items[:56]]
    _change(book, tuple(palette))


def reset_colors(book: Workbook) -> None:
    _change(book, styles.DEFAULT_PALETTE)


def _change(book: Workbook, palette: tuple[str, ...]) -> None:
    current = book.stylesheet.colors
    if palette != current.indexed:
        current.indexed = palette
        current.palette_changed = True
        book.saved = False
