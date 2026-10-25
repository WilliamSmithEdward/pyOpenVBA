"""Manual page breaks: HPageBreaks, VPageBreaks, Range.PageBreak and Worksheet.ResetAllPageBreaks, as Excel keeps
them on a sheet and saves them as its rowBreaks and colBreaks (tests/fixtures/page_breaks/,
scripts/measure_page_breaks.py).

A break stands before a row or a column, never the first, and is saved as a brk whose id is the line before it. A
collection counts only the breaks before the rows or columns the sheet uses and the one after them, though the file
keeps every one, and in their order; a whole row or column read for its PageBreak answers its own break wherever it
is. Rows or columns going in move the breaks at and after them along; one going out takes the break before it away
and moves those after it back; cells shifting move none.

Excel works out automatic breaks only as it lays pages out, which a macro's session has not done, so the model has
none; a sheet with a print area, whose breaks Excel cuts to it and edges with automatic ones, reports itself
unsupported. Subtotal's breaks between groups are its own (see _subtotals), with the setting it keeps here.
"""
from __future__ import annotations

import re
from copy import deepcopy
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, quote_sheet
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, ref_setter
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAInt, error, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

MANUAL = -4135  # xlPageBreakManual
AUTOMATIC = -4105  # xlPageBreakAutomatic
NONE = -4142  # xlPageBreakNone
FULL = 1  # xlPageBreakFull
DEFINED = "Application-defined or object-defined error"
UNSET = "Unable to set the PageBreak property of the Range class"
UNGOT = "Unable to get the PageBreak property of the Range class"
#: A break's own attributes, after its id, as Excel writes a manual one across the whole sheet.
_ROW, _COLUMN = {"max": "16383", "man": "1"}, {"max": "1048575", "man": "1"}
_FOLLOWING = ("customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
              "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts",
              "extLst")


def _breaks() -> dict[int, dict[str, str]]:
    return {}


@dataclass
class SheetBreaks:
    """A sheet's breaks: the row or column each stands before, and the attributes its brk has besides its id."""

    rows: dict[int, dict[str, str]] = field(default_factory=_breaks)
    columns: dict[int, dict[str, str]] = field(default_factory=_breaks)
    #: True once a break went in, out or along, so a save writes them again.
    changed: bool = False
    #: Subtotal's PageBreaks as the sheet's last call left it, which a call leaving it out takes (see _subtotals).
    subtotal: bool = False


def _read(xml: str) -> SheetBreaks:
    found = SheetBreaks()
    for kind, body in re.findall(r"<(rowBreaks|colBreaks)\b[^>]*?(?:/>|>(.*?)</\1>)", xml, re.DOTALL):
        lines = found.rows if kind == "rowBreaks" else found.columns
        for attributes in re.findall(r"<brk\b([^>]*?)/?>", body):
            values = dict(re.findall(r'(\w+)="([^"]*)"', attributes))
            line = values.pop("id", "")
            if not line.isdigit():
                raise VBAUnsupportedError("a page break without an id is not implemented")
            lines[int(line) + 1] = values
    return found


def state(sheet: Worksheet) -> SheetBreaks:
    """The sheet's breaks, read from its file the first time they are needed."""
    if sheet.page_breaks is None:
        package = sheet.book.package
        xml = ""
        if package is not None and sheet.part_name and package.has(sheet.part_name):
            data = package.read(sheet.part_name)
            if b"Breaks" in data:
                xml = data.decode("utf-8", errors="replace")
        sheet.page_breaks = _read(xml)
    return sheet.page_breaks


def _plain(attributes: dict[str, str], rows: bool) -> bool:
    """Whether a break is a manual one across the whole sheet, the only kind the model works with."""
    return (attributes.get("man") in ("1", "true") and attributes.get("max") == (_ROW if rows else _COLUMN)["max"]
            and attributes.get("min", "0") == "0" and attributes.get("pt", "0") in ("0", "false"))


def admitted(sheet: Worksheet) -> SheetBreaks:
    """The breaks a macro works with: a print area, or a break Excel cut to one, reports itself unsupported."""
    local = (quote_sheet(sheet.name) + "!Print_Area").casefold()
    if any(entry.name.casefold() == local for entry in sheet.book.names_.entries):
        raise VBAUnsupportedError("Page breaks on a sheet with a print area are not implemented")
    found = state(sheet)
    if not all(_plain(attributes, True) for attributes in found.rows.values()) or \
            not all(_plain(attributes, False) for attributes in found.columns.values()):
        raise VBAUnsupportedError("Automatic or partial page breaks are not implemented")
    return found


def add(sheet: Worksheet, rows: bool, line: int) -> None:
    """A manual break before a row or a column; one there already stays as it is."""
    lines = state(sheet).rows if rows else state(sheet).columns
    if line not in lines:
        lines[line] = dict(_ROW if rows else _COLUMN)
        _changed(sheet)


def remove(sheet: Worksheet, rows: bool, line: int) -> None:
    lines = state(sheet).rows if rows else state(sheet).columns
    if lines.pop(line, None) is not None:
        _changed(sheet)


def remove_within(sheet: Worksheet, first: int, last: int) -> None:
    """Every row break before a row from ``first`` to ``last``."""
    for line in [line for line in state(sheet).rows if first <= line <= last]:
        remove(sheet, True, line)


def _changed(sheet: Worksheet) -> None:
    state(sheet).changed = True
    sheet.touched()


def lines_moved(sheet: Worksheet, start: int, count: int, *, delete: bool, rows: bool) -> None:
    """Rows or columns from ``start`` went in or out: the breaks at and after them go along, and one before a line
    that went out goes with it."""
    if sheet.page_breaks is None:
        package = sheet.book.package
        if package is None or not sheet.part_name or not package.has(sheet.part_name) \
                or b"Breaks" not in package.read(sheet.part_name):
            return
    found = state(sheet)
    lines = found.rows if rows else found.columns
    limit = MAX_ROWS if rows else MAX_COLUMNS
    moved: dict[int, dict[str, str]] = {}
    for line, attributes in lines.items():
        if delete:
            if start <= line < start + count:
                continue
            target = line - count if line >= start + count else line
        else:
            target = line + count if line >= start else line
        if 1 < target <= limit:
            moved[target] = attributes
    if moved != lines:
        if rows:
            found.rows = moved
        else:
            found.columns = moved
        _changed(sheet)


def copied(source: Worksheet, destination: Worksheet) -> None:
    """A copied sheet has the breaks the sheet had."""
    found = state(source)
    destination.page_breaks = SheetBreaks(deepcopy(found.rows), deepcopy(found.columns),
                                          changed=bool(found.rows or found.columns))


def with_breaks(sheet: Worksheet, xml: str) -> str:
    """The sheet's XML with its rowBreaks and colBreaks as the model has them, after the page setup and before what
    follows it."""
    found = sheet.page_breaks
    if found is None or not found.changed:
        return xml
    xml = re.sub(r"<(rowBreaks|colBreaks)\b[^>]*?(?:/>|>.*?</\1>)", "", xml, flags=re.DOTALL)
    markup = "".join(_element(kind, lines) for kind, lines in (("rowBreaks", found.rows), ("colBreaks", found.columns))
                     if lines)
    if not markup:
        return xml
    following = re.search(r"<(?:" + "|".join(_FOLLOWING) + r")\b|</worksheet>", xml)
    at = following.start() if following is not None else len(xml)
    return xml[:at] + markup + xml[at:]


def _element(kind: str, lines: dict[int, dict[str, str]]) -> str:
    manual = sum(1 for attributes in lines.values() if attributes.get("man") in ("1", "true"))
    breaks = "".join(f'<brk id="{line - 1}"' + "".join(f' {key}="{value}"' for key, value in attributes.items())
                     + "/>" for line, attributes in sorted(lines.items()))
    return f'<{kind} count="{len(lines)}" manualBreakCount="{manual}">{breaks}</{kind}>'


def _counted(sheet: Worksheet, rows: bool) -> list[int]:
    """The breaks a collection counts: before the lines the sheet uses, and the one after them."""
    found = admitted(sheet)
    used = sheet.used_bounds()
    if used is None:
        return []
    last = used[2] if rows else used[3]
    return sorted(line for line in (found.rows if rows else found.columns) if line <= last + 1)


def _line_of(before: object, rows: bool) -> int:
    """The row or column a break goes before, from a range; the first one takes none."""
    if not isinstance(before, Range):
        raise error(13, "Type mismatch")
    line = before.first.top if rows else before.first.left
    if line <= 1:
        raise error(1004, DEFINED)
    return line


# --- the collections and their breaks ------------------------------------------------------------------------------


class _Breaks(VBACollection, ExcelObject):
    """A sheet's horizontal or vertical breaks, as the collection counts them."""

    #: No type of its own: the base HPageBreaks and VPageBreaks share.
    vba_type_name = ""
    rows = True

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def vba_items(self) -> list[object]:
        made = HPageBreak if self.rows else VPageBreak
        return [made(self.sheet, line) for line in _counted(self.sheet, self.rows)]

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")

    @method
    def Add(self, Before: object = MISSING) -> object:
        if Before is MISSING:
            raise error(449)
        admitted(self.sheet)
        line = _line_of(Before, self.rows)
        add(self.sheet, self.rows, line)
        return (HPageBreak if self.rows else VPageBreak)(self.sheet, line)


class HPageBreaks(_Breaks):
    vba_type_name = "HPageBreaks"


class VPageBreaks(_Breaks):
    vba_type_name = "VPageBreaks"
    rows = False


class _Break(ExcelObject):
    """One manual break, before its row or column."""

    #: No type of its own: the base HPageBreak and VPageBreak share.
    vba_type_name = ""
    rows = True

    def __init__(self, sheet: Worksheet, line: int) -> None:
        self.sheet = sheet
        self.line = line

    @member
    def Location(self) -> object:
        """The cell the break stands before: in column A for a row's, in row 1 for a column's."""
        area = Area(self.line, 1, self.line, 1, self.sheet.name) if self.rows else \
            Area(1, self.line, 1, self.line, self.sheet.name)
        return Range(self.sheet, [area])

    @ref_setter("Location")
    def _set_location(self, value: object) -> None:
        raise error(1004, DEFINED)

    @member
    def Type(self) -> object:
        return VBAInt(MANUAL, "Long")

    @member
    def Extent(self) -> object:
        return VBAInt(FULL, "Long")

    @method
    def Delete(self) -> object:
        remove(self.sheet, self.rows, self.line)
        return EMPTY

    @method
    def DragOff(self, Direction: object = MISSING, RegionIndex: object = MISSING) -> object:
        raise VBAUnsupportedError("DragOff of a page break, which moves it to where a page's region ends, is not "
                                  "implemented")

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")


class HPageBreak(_Break):
    vba_type_name = "HPageBreak"


class VPageBreak(_Break):
    vba_type_name = "VPageBreak"
    rows = False


def reset(sheet: Worksheet) -> None:
    """ResetAllPageBreaks: every manual break goes."""
    found = state(sheet)
    if found.rows or found.columns:
        found.rows, found.columns = {}, {}
        _changed(sheet)


# --- Range.PageBreak -------------------------------------------------------------------------------------------------


def page_break(target: Range) -> object:
    """A whole row's or column's break, as Rows or Columns made the range; any other range is error 1004."""
    if target.lines not in ("rows", "columns"):
        raise error(1004, UNGOT)
    rows = target.lines == "rows"
    found = admitted(target.sheet)
    line = target.first.top if rows else target.first.left
    return VBAInt(MANUAL if line in (found.rows if rows else found.columns) else NONE, "Long")


def set_page_break(target: Range, value: object) -> None:
    """xlPageBreakManual puts a break before a whole row or column, or before a cell's row and its column, those
    that are not the first; xlPageBreakNone takes them away; anything else is error 1004."""
    kind = int(to_integer(value, "Long"))
    if kind not in (MANUAL, NONE):
        raise error(1004, UNSET)
    sheet, area = target.sheet, target.first
    admitted(sheet)
    lines = [(True, area.top)] if target.lines == "rows" else [(False, area.left)] if target.lines == "columns" \
        else [(True, area.top), (False, area.left)]
    lines = [(rows, line) for rows, line in lines if line > 1]
    if not lines:
        raise error(1004, UNSET)
    for rows, line in lines:
        if kind == MANUAL:
            add(sheet, rows, line)
        else:
            remove(sheet, rows, line)
