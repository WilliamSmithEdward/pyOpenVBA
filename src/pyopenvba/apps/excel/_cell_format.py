"""Application.FindFormat and ReplaceFormat: the formats Range.Find and Range.Replace look for and put in.

Measured in live Excel (scripts/measure_find_format.py, tests/fixtures/find_format.json):

- Each read hands out a new CellFormat over the one format the application keeps for the session. A property not
  set reads Null, but Interior.Color reads 0; Clear unsets them all.
- A colour and its ColorIndex are one setting: either set, the other reads as it, ColorIndex as the palette slot
  nearest the colour.
- Find with SearchFormat finds a cell whose own format holds every property set, and with nothing sought any such
  cell, blank or not; with no property set it finds as without. FindNext and FindPrevious leave the format out.
- Replace with SearchFormat changes only the cells whose format holds; with nothing sought each such cell's text
  becomes the replacement, or stays where that is empty. With ReplaceFormat each cell it changes takes every
  property set, its text replaced by the same text or not; with nothing sought and no format to find, the blank
  cells take it.
"""
from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import MemberSpec, member, method
from pyopenvba.interpreter._values import EMPTY, NULL, VBAInt, to_bool, to_integer, to_number, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application, Worksheet
    from pyopenvba.apps.excel._styles import Stylesheet

#: The most cells a search for a format with nothing sought walks.
MOST_CELLS = 1 << 20


def _long(value: object) -> object:
    return VBAInt(int(to_integer(value, "Long")), "Long")


def _double(value: object) -> object:
    return float(to_number(value))


class _Part(ExcelObject):
    """A CellFormat, or its Font or Interior: the properties it keeps under ``prefix`` in the application's format."""

    def __init__(self, application: Application, kept: dict[str, object], prefix: str = "") -> None:
        self.application = application
        self.kept = kept
        self.prefix = prefix

    @member
    def Application(self) -> object:
        return self.application

    def _stylesheet(self) -> Stylesheet:
        book = self.application.active_book
        if book is None:
            raise VBAUnsupportedError("a CellFormat colour with no workbook open is not implemented")
        return book.stylesheet

    def color(self, unset: object) -> object:
        from pyopenvba.apps.excel._styles import bgr

        if self.prefix + "Color" in self.kept:
            return self.kept[self.prefix + "Color"]
        if self.prefix + "ColorIndex" in self.kept:
            from pyopenvba.apps.excel._styles import Color

            stylesheet = self._stylesheet()
            index = int(to_integer(self.kept[self.prefix + "ColorIndex"], "Long"))
            return float(bgr(stylesheet.colors.rrggbb(Color("indexed", index + 7), "000000")))
        return unset

    def color_index(self) -> object:
        if self.prefix + "ColorIndex" in self.kept:
            return self.kept[self.prefix + "ColorIndex"]
        if self.prefix + "Color" in self.kept:
            from pyopenvba.apps.excel._styles import Color

            color = Color.from_bgr(int(to_integer(self.kept[self.prefix + "Color"], "Long")))
            return VBAInt(self._stylesheet().colors.color_index(color, "000000"), "Long")
        return NULL


def _setting(name: str, convert: Callable[[object], object]) -> MemberSpec:
    """A property a format keeps as set, Null until then."""

    def get(part: _Part) -> object:
        return part.kept.get(part.prefix + name, NULL)

    def put(part: _Part, value: object) -> None:
        part.kept[part.prefix + name] = convert(value)

    return MemberSpec(name=name, kind="property", getter=get, setter=put)


def _style(name: str, other: str) -> MemberSpec:
    """Bold or Italic, one font style: with the other set it reads False, but finds as not set
    (tests/fixtures/find_format.json)."""
    spec = _setting(name, to_bool)

    def get(part: _Part) -> object:
        if part.prefix + name in part.kept:
            return part.kept[part.prefix + name]
        return False if part.prefix + other in part.kept else NULL

    return MemberSpec(name=name, kind="property", getter=get, setter=spec.setter)


def _unmeasured(name: str) -> MemberSpec:
    """A property that reads Null, whose finding is not measured."""

    def put(part: _Part, value: object) -> None:
        raise VBAUnsupportedError(f"setting CellFormat {name} is not implemented")

    def get(part: _Part) -> object:
        return NULL

    return MemberSpec(name=name, kind="property", getter=get, setter=put)


def _colors(unset: object) -> list[MemberSpec]:
    """Color and ColorIndex, one setting read either way."""

    def put_color(part: _Part, value: object) -> None:
        part.kept.pop(part.prefix + "ColorIndex", None)
        part.kept[part.prefix + "Color"] = _double(value)

    def put_index(part: _Part, value: object) -> None:
        index = int(to_integer(value, "Long"))
        if not 1 <= index <= 56:
            raise VBAUnsupportedError("a CellFormat ColorIndex outside the palette is not implemented")
        part.kept.pop(part.prefix + "Color", None)
        part.kept[part.prefix + "ColorIndex"] = VBAInt(index, "Long")

    def get_color(part: _Part) -> object:
        return part.color(unset)

    def get_index(part: _Part) -> object:
        return part.color_index()

    return [MemberSpec(name="Color", kind="property", getter=get_color, setter=put_color),
            MemberSpec(name="ColorIndex", kind="property", getter=get_index, setter=put_index)]


class CellFormatFont(_Part):
    vba_type_name = "Font"


class CellFormatInterior(_Part):
    vba_type_name = "Interior"


class CellFormat(_Part):
    vba_type_name = "CellFormat"

    @member
    def Parent(self) -> object:
        return self.application

    @member
    def Font(self) -> object:
        return CellFormatFont(self.application, self.kept, "Font.")

    @member
    def Interior(self) -> object:
        return CellFormatInterior(self.application, self.kept, "Interior.")

    @member
    def Borders(self) -> object:
        raise VBAUnsupportedError("CellFormat Borders is not implemented")

    @method
    def Clear(self) -> object:
        self.kept.clear()
        return EMPTY


for _spec in (_style("Bold", "Italic"), _style("Italic", "Bold"), _setting("Name", to_text),
              _setting("Size", _double), _setting("Underline", _long), *_colors(NULL), _unmeasured("Strikethrough")):
    CellFormatFont.vba_add_member(_spec)
for _spec in (*_colors(0.0), _setting("Pattern", _long)):
    CellFormatInterior.vba_add_member(_spec)
for _spec in (_setting("NumberFormat", to_text), _setting("Locked", to_bool), _setting("FormulaHidden", to_bool),
              _setting("HorizontalAlignment", _long), _setting("WrapText", to_bool), _setting("IndentLevel", _long),
              *(_unmeasured(name) for name in ("VerticalAlignment", "Orientation", "ShrinkToFit", "MergeCells",
                                              "AddIndent"))):
    CellFormat.vba_add_member(_spec)


def _cell(sheet: Worksheet, row: int, column: int) -> Range:
    return Range(sheet, [Area(row, column, row, column, sheet.name)])


def _same(actual: object, wanted: object) -> bool:
    if isinstance(wanted, bool) or isinstance(actual, bool):
        return actual is wanted
    if isinstance(wanted, str) or isinstance(actual, str):
        return actual == wanted
    if actual is NULL:
        return False
    return float(to_number(actual)) == float(to_number(wanted))


def holds(sheet: Worksheet, row: int, column: int, kept: dict[str, object]) -> bool:
    """Whether the cell's own format holds every property ``kept`` sets."""
    cell = _cell(sheet, row, column)
    for path, wanted in kept.items():
        found: object = cell
        for name in path.split("."):
            if not isinstance(found, ExcelObject):
                return False
            found = found.vba_get(name)
        if not _same(found, wanted):
            return False
    return True


def put_in(sheet: Worksheet, row: int, column: int, kept: dict[str, object]) -> None:
    """Give the cell every property ``kept`` sets."""
    cell = _cell(sheet, row, column)
    for path, value in kept.items():
        *names, last = path.split(".")
        owner: object = cell
        for name in names:
            owner = owner.vba_get(name) if isinstance(owner, ExcelObject) else owner
        if not isinstance(owner, ExcelObject):
            raise VBAUnsupportedError(f"putting in CellFormat {path} is not implemented")
        owner.vba_set(last, value)


def every_cell(target: Range, traversal: int) -> list[tuple[int, int]]:
    """Each cell of ``target``, area by area, by rows or by columns; refused past MOST_CELLS."""
    if sum(area.rows * area.columns for area in target.areas) > MOST_CELLS:
        raise VBAUnsupportedError("searching more than a million cells for a format is not implemented")
    found: list[tuple[int, int]] = []
    for area in target.areas:
        if traversal == 1:
            found += [(row, column) for row in range(area.top, area.bottom + 1)
                      for column in range(area.left, area.right + 1)]
        else:
            found += [(row, column) for column in range(area.left, area.right + 1)
                      for row in range(area.top, area.bottom + 1)]
    return found
