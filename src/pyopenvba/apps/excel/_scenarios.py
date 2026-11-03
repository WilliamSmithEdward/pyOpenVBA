"""Worksheet.Scenarios: what-if scenarios, values kept for a sheet's changing cells, as Excel keeps and saves them
(tests/fixtures/scenarios.json, scripts/measure_scenarios.py).

- Add(Name, ChangingCells, [Values], [Comment], [Locked], [Hidden]) keeps each value as text, a number as CStr
  writes it and a logical value in capitals; values left out are the cells' own, and values past the cells are
  passed over. A name that is empty, longer than 255 characters or there already, a formula for a value, or a
  value given as one Integer or Long rather than an array, is error 1004 "Add method of Scenarios class failed".
- The comment says who made it and when, "Created by <UserName> on m/d/yyyy"; ChangeScenario and Locked or Hidden
  add "Modified by <UserName> on m/d/yyyy" on a line of its own.
- Show writes the values into the cells as Value would. Scenarios(name) finds one ignoring case; one that is not
  there is error 1004.
- The sheet part keeps them as scenarios, with the one last shown as show.
"""
from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, cast

from pyopenvba._a1 import Area
from pyopenvba._xml import attributes, escape
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, method, setter, VBACollection, VBAObject
from pyopenvba.interpreter._values import MISSING, VBAArray, VBAInt, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

_ADD_FAILED = "Add method of Scenarios class failed"
_ELEMENT = re.compile(r"<scenarios\b.*?</scenarios>|<scenarios\b[^>]*/>", re.DOTALL)
#: What follows scenarios in a worksheet part (ECMA-376 Part 1, 18.3.1.99).
_AFTER = re.compile(r"<(?:autoFilter|sortState|dataConsolidate|customSheetViews|mergeCells|phoneticPr|"
                    r"conditionalFormatting|dataValidations|hyperlinks|printOptions|pageMargins|pageSetup|"
                    r"headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|"
                    r"legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b"
                    r"|</worksheet>")


@dataclass
class Kept:
    name: str
    cells: list[tuple[int, int]]
    values: list[str]
    comment: str
    user: str
    locked: bool = True
    hidden: bool = False


@dataclass
class SheetScenarios:
    """A sheet's scenarios and the one last shown."""

    items: list[Kept] = field(default_factory=list[Kept])
    shown: Kept | None = None


def state(sheet: Worksheet) -> SheetScenarios:
    if sheet.scenarios is None:
        sheet.scenarios = _read(sheet)
    return sheet.scenarios


def _position(reference: str) -> tuple[int, int]:
    from pyopenvba._a1 import parse_area

    area = parse_area(reference)
    return area.top, area.left


def _read(sheet: Worksheet) -> SheetScenarios:
    package = sheet.book.package
    found = SheetScenarios()
    if package is None or not sheet.part_name or not package.has(sheet.part_name):
        return found
    element = _ELEMENT.search(package.read(sheet.part_name).decode("utf-8", errors="replace"))
    if element is None:
        return found
    for opening, body in re.findall(r"<scenario\b([^>]*)>(.*?)</scenario>", element.group(0), re.DOTALL):
        fields = attributes(f"<scenario {opening}>")
        cells = [attributes(one) for one in re.findall(r"<inputCells\b[^>]*/?>", body)]
        found.items.append(Kept(fields.get("name", ""), [_position(one.get("r", "A1")) for one in cells],
                                [one.get("val", "") for one in cells], fields.get("comment", "").replace("_x000a_", "\n"),
                                fields.get("user", ""), fields.get("locked", "0") in ("1", "true"),
                                fields.get("hidden", "0") in ("1", "true")))
    shown = attributes(re.match(r"<scenarios\b[^>]*", element.group(0)).group(0) + ">").get("show", "")  # type: ignore[union-attr]
    if shown.isdecimal() and int(shown) < len(found.items):
        found.shown = found.items[int(shown)]
    return found


def _changed(sheet: Worksheet) -> None:
    sheet.scenarios_changed = True
    sheet.touched()


def _today() -> str:
    day = _dt.date.today()
    return f"{day.month}/{day.day}/{day.year}"


def _stamp(sheet: Worksheet, verb: str) -> str:
    return f"{verb} by {sheet.book.application.user_name} on {_today()}"


def _cells(target: object) -> list[tuple[int, int]]:
    if not isinstance(target, Range):
        raise error(1004, _ADD_FAILED)
    return [(row, column) for area in target.areas for row in range(area.top, area.bottom + 1)
            for column in range(area.left, area.right + 1)]


def _text(value: object) -> str:
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    text = to_text(value)
    if text.startswith("="):
        raise error(1004, _ADD_FAILED)
    return text


def _values(sheet: Worksheet, cells: list[tuple[int, int]], given: object) -> list[str]:
    """The values kept for the cells: those given, the cells' own for the rest."""
    own = [_text(Range(sheet, [Area(row, column, row, column)]).vba_get("Value")) for row, column in cells]
    if given is MISSING:
        return own
    if isinstance(given, VBAArray):
        items = list(given.items)
    elif isinstance(given, VBAInt):
        raise error(1004, _ADD_FAILED)
    else:
        items = [given]
    texts = [_text(item) for item in items[:len(cells)]]
    return texts + own[len(texts):]


class Scenarios(VBACollection, ExcelObject):
    """A sheet's scenarios."""

    vba_type_name = "Scenarios"

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def vba_items(self) -> list[object]:
        return [Scenario(self.sheet, kept) for kept in state(self.sheet).items]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        return find(self.sheet, index)

    @method
    def CreateSummary(self, ReportType: object = MISSING, ResultCells: object = MISSING) -> object:
        return create_summary(self.sheet, ReportType, ResultCells)

    @method
    def Add(self, Name: object = MISSING, ChangingCells: object = MISSING, Values: object = MISSING,
            Comment: object = MISSING, Locked: object = MISSING, Hidden: object = MISSING) -> object:
        if Name is MISSING or ChangingCells is MISSING:
            raise error(449)
        name = to_text(Name)
        kept = state(self.sheet)
        if not name or len(name) > 255 or any(one.name.lower() == name.lower() for one in kept.items):
            raise error(1004, _ADD_FAILED)
        cells = _cells(ChangingCells)
        if len(cells) > 32:
            raise VBAUnsupportedError("A scenario of more than 32 changing cells is not implemented")
        # Added to a protected sheet with Locked left out, a scenario is not locked (tests/fixtures/scenarios.json).
        protection = self.sheet.protection
        locked = not (protection is not None and protection.contents) if Locked is MISSING else to_bool(Locked)
        made = Kept(name, cells, _values(self.sheet, cells, Values),
                    _stamp(self.sheet, "Created") if Comment is MISSING else to_text(Comment),
                    self.sheet.book.application.user_name, locked, False if Hidden is MISSING else to_bool(Hidden))
        kept.items.append(made)
        _changed(self.sheet)
        return Scenario(self.sheet, made)

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")


def find(sheet: Worksheet, index: object) -> Scenario:
    """Scenarios(index), by number or by name ignoring case."""
    items = state(sheet).items
    failed = error(1004, "Method 'Scenarios' of object '_Worksheet' failed")
    if isinstance(index, str):
        found = next((one for one in items if one.name.lower() == index.lower()), None)
        if found is None:
            raise failed
        return Scenario(sheet, found)
    position = int(to_integer(index))
    if not 1 <= position <= len(items):
        raise failed
    return Scenario(sheet, items[position - 1])


_NAVY, _WHITE, _BLACK, _BANNER, _GRAY = 8388608, 16777215, 0, 8388736, 12632256
_LEFT, _RIGHT, _CONTINUOUS, _LINE_NONE = -4131, -4152, 1, -4142
_SUMMARY_FAILED = "CreateSummary method of Scenarios class failed"
_NOTES = ("Notes:  Current Values column represents values of changing cells at",
          "time Scenario Summary Report was created.  Changing cells for each",
          "scenario are highlighted in gray.")


def _named(sheet: Worksheet, row: int, column: int) -> str:
    """A cell as a summary names it: by a defined name for it alone, or its address."""
    from pyopenvba._a1 import column_letter, split_sheet

    address = f"${column_letter(column)}${row}"
    for entry in sheet.book.names_.entries:
        place, area = split_sheet(entry.refers_to.lstrip("="))
        if place.casefold() == sheet.name.casefold() and area == address:
            return split_sheet(entry.name)[1]
    return address


def _part(owner: VBAObject, name: str, *arguments: object) -> VBAObject:
    return cast(VBAObject, owner.vba_get(name, arguments))


def create_summary(sheet: Worksheet, report_type: object, result_cells: object) -> object:
    """Scenarios.CreateSummary: a sheet before this one, with what each scenario's changing cells would hold and what
    its result cell would come to, the sheet's own cells as they were (tests/fixtures/connection_members.json). One
    changing cell and one result cell are measured; a pivot table's report and a guess at the result cells are not."""
    from pyopenvba.apps.excel._model import Range as _Range

    kept = state(sheet)
    if not kept.items:
        raise error(1004, _SUMMARY_FAILED)
    kind = 1 if report_type is MISSING else int(to_integer(report_type, "Long"))
    if kind not in (1, -4148):
        raise error(1004, _SUMMARY_FAILED)
    if kind == -4148:
        raise VBAUnsupportedError("Scenarios.CreateSummary as a pivot table is not implemented")
    if isinstance(result_cells, str):
        raise error(1004, "This formula is missing a range reference or a defined name. ")
    if not isinstance(result_cells, _Range):
        raise VBAUnsupportedError("Scenarios.CreateSummary guessing its result cells is not implemented")
    results = _cells(result_cells)
    if len(results) != 1 or any(len(one.cells) != 1 for one in kept.items) \
            or len({tuple(one.cells) for one in kept.items}) != 1:
        raise VBAUnsupportedError("Scenarios.CreateSummary of more than one changing or result cell is not implemented")
    book = sheet.book
    target, result = kept.items[0].cells[0], results[0]
    place = _Range(sheet, [Area(*target, *target)])
    formula = sheet.cells_[target].formula if target in sheet.cells_ else ""
    before = place.vba_get("Formula") if formula else place.vba_get("Value")
    current = (book.calculator.value_of(sheet.name, *target), book.calculator.value_of(sheet.name, *result))
    answers: list[tuple[str, object]] = []
    for one in kept.items:
        place.put_as_value(*target, one.values[0])
        answers.append((one.values[0], book.calculator.value_of(sheet.name, *result)))
    place.vba_set("Formula" if formula else "Value", before)
    taken = {one.name.casefold() for one in book.sheets_}
    name, number = "Scenario Summary", 1
    while name.casefold() in taken:
        number += 1
        name = f"Scenario Summary {number}"
    made = book.add_sheet(name, at=book.sheets_.index(sheet))
    last = 4 + len(kept.items)

    def cell(row: int, column: int, value: object = None, *, bold: bool = False, colour: int = _BLACK,
             fill: int | None = None, align: int | None = None, bottom: int = _CONTINUOUS,
             top: int = _CONTINUOUS) -> None:
        at = _Range(made, [Area(row, column, row, column)])
        if value is not None:
            at.vba_set("Value", value)
        font = _part(at, "Font")
        font.vba_set("Bold", bold)
        font.vba_set("Color", VBAInt(colour, "Long"))
        if fill is not None:
            _part(at, "Interior").vba_set("Color", VBAInt(fill, "Long"))
        if align is not None:
            at.vba_set("HorizontalAlignment", VBAInt(align, "Long"))
        borders = _part(at, "Borders")
        for edge, style in ((9, bottom), (8, top)):
            if style != _LINE_NONE:
                _part(borders, "Item", VBAInt(edge, "Long")).vba_set("LineStyle", VBAInt(style, "Long"))

    # The empty cell above the banner, whose bottom edge is the banner's top, counts in the sheet's UsedRange.
    cell(1, 2, top=_LINE_NONE)
    for column in range(2, last + 1):
        label = column < 4
        cell(2, column, "Scenario Summary" if column == 2 else None, bold=label, colour=_WHITE, fill=_BANNER,
             align=_LEFT if label else _RIGHT, bottom=_LINE_NONE)
        heading = "Current Values:" if column == 4 else kept.items[column - 5].name if column > 4 else None
        cell(3, column, heading, bold=label, colour=_WHITE, fill=_BANNER, align=_LEFT if label else _RIGHT,
             top=_LINE_NONE)
    for row, title, shown in ((5, "Changing Cells:", None), (7, "Result Cells:", None)):
        for column in (2, 3):
            cell(row, column, title if column == 2 else shown, bold=True, colour=_NAVY, fill=_GRAY, align=_LEFT)
    for row, shown in ((4, None), (6, _named(sheet, *target)), (8, _named(sheet, *result))):
        for column in (2, 3):
            cell(row, column, shown if column == 3 else None, bold=True, fill=_GRAY, align=_LEFT)
    for offset, (one, (changing, comes_to)) in enumerate(zip(kept.items, answers, strict=True)):
        cell(4, 5 + offset, one.comment)
        cell(6, 5 + offset, changing, fill=_GRAY)
        cell(8, 5 + offset, comes_to)
    cell(6, 4, current[0])
    cell(8, 4, current[1])
    for offset, text in enumerate(_NOTES):
        cell(9 + offset, 2, text, bottom=_LINE_NONE, top=_CONTINUOUS if offset == 0 else _LINE_NONE)
    for row in (4, 6, 8):
        _part(_Range(made, [Area(row, 1, row, 1)]), "EntireRow").vba_set("OutlineLevel", VBAInt(2, "Long"))
    _part(made, "Outline").vba_set("SummaryRow", VBAInt(0, "Long"))
    book.application.activate_book(book)
    made.vba_get("Activate")
    book.saved = False
    return made


class Scenario(ExcelObject):
    """One scenario."""

    vba_type_name = "Scenario"

    def __init__(self, sheet: Worksheet, kept: Kept) -> None:
        self.sheet = sheet
        self.kept = kept

    def _modified(self) -> None:
        stamp = _stamp(self.sheet, "Modified")
        lines = self.kept.comment.split("\n")
        if len(lines) > 1 and lines[-1].startswith("Modified by "):
            lines.pop()
        self.kept.comment = "\n".join([*lines, stamp])
        _changed(self.sheet)

    @member
    def Name(self) -> object:
        return self.kept.name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        self.kept.name = to_text(value)
        _changed(self.sheet)

    @member
    def Comment(self) -> object:
        return self.kept.comment

    @setter("Comment")
    def _set_comment(self, value: object) -> None:
        self.kept.comment = to_text(value)
        _changed(self.sheet)

    @member
    def Locked(self) -> object:
        return self.kept.locked

    @setter("Locked")
    def _set_locked(self, value: object) -> None:
        self.kept.locked = to_bool(value)
        self._modified()

    @member
    def Hidden(self) -> object:
        return self.kept.hidden

    @setter("Hidden")
    def _set_hidden(self, value: object) -> None:
        self.kept.hidden = to_bool(value)
        self._modified()

    @member
    def Index(self) -> object:
        return VBAInt(state(self.sheet).items.index(self.kept) + 1, "Long")

    @member
    def ChangingCells(self) -> object:
        return _union(self.sheet, self.kept.cells)

    @member
    def Values(self, Index: object = MISSING) -> object:
        if Index is not MISSING:
            position = int(to_integer(Index))
            if not 1 <= position <= len(self.kept.values):
                raise error(9)
            return self.kept.values[position - 1]
        return VBAArray([(1, len(self.kept.values))], items=list(self.kept.values))

    @method
    def Show(self) -> object:
        from pyopenvba.apps.excel._model import Range as _Range

        for (row, column), value in zip(self.kept.cells, self.kept.values, strict=True):
            _Range(self.sheet, [Area(row, column, row, column)]).put_as_value(row, column, value)
        state(self.sheet).shown = self.kept
        _changed(self.sheet)
        return True

    @method
    def ChangeScenario(self, ChangingCells: object = MISSING, Values: object = MISSING) -> object:
        if ChangingCells is MISSING:
            raise error(449)
        cells = _cells(ChangingCells)
        self.kept.cells, self.kept.values = cells, _values(self.sheet, cells, Values)
        self._modified()
        return True

    @method
    def Delete(self) -> object:
        kept = state(self.sheet)
        kept.items.remove(self.kept)
        _changed(self.sheet)
        return True

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")


def _union(sheet: Worksheet, cells: list[tuple[int, int]]) -> Range:
    from pyopenvba.apps.excel._model import coalesced

    return Range(sheet, coalesced([Area(row, column, row, column) for row, column in cells]))


def copied(source: Worksheet, copy: Worksheet) -> None:
    """A sheet's copy has its scenarios."""
    kept = state(source)
    if kept.items:
        copy.scenarios = SheetScenarios([Kept(one.name, list(one.cells), list(one.values), one.comment, one.user,
                                              one.locked, one.hidden) for one in kept.items])
        copy.scenarios_changed = True


def written(sheet: Worksheet, xml: str) -> str:
    """The sheet part with its scenarios as the model has them."""
    if not sheet.scenarios_changed:
        return xml
    from pyopenvba._a1 import column_letter

    kept = state(sheet)
    xml = _ELEMENT.sub("", xml)
    if kept.items:
        parts: list[str] = []
        for one in kept.items:
            flags = (' locked="1"' if one.locked else "") + (' hidden="1"' if one.hidden else "")
            comment = escape(one.comment).replace("\n", "_x000a_")
            cells = "".join(f'<inputCells r="{column_letter(column)}{row}" val="{escape(value)}"/>'
                            for (row, column), value in zip(one.cells, one.values, strict=True))
            parts.append(f'<scenario name="{escape(one.name)}"{flags} count="{len(one.cells)}" '
                         f'user="{escape(one.user)}" comment="{comment}">{cells}</scenario>')
        shown = kept.items.index(kept.shown) if kept.shown in kept.items else None
        show = f' show="{shown}"' if shown is not None else ""
        element = f'<scenarios current="{shown or 0}"{show}>{"".join(parts)}</scenarios>'
        following = _AFTER.search(xml, max(xml.find("</sheetData>"), 0))
        assert following is not None
        xml = xml[:following.start()] + element + xml[following.start():]
    sheet.scenarios_changed = False
    return xml
