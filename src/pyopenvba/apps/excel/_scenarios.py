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
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba._xml import attributes, escape
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
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
