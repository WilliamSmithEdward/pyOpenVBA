"""Range.Errors: Excel's background error checks of a cell, and which of them a macro has told it to ignore
(tests/fixtures/range_errors.json, scripts/measure_range_errors.py).

Each check answers for one cell, whatever Application.ErrorCheckingOptions says:

1. xlEvaluateToError: a formula that comes to an error other than #N/A.
2. xlTextDate: text that types as a date written m/d/yy.
3. xlNumberAsText: text that types as a number, not a date or a logical value.
4. xlInconsistentFormula: a formula between two formulas, above and below or left and right, the same as each
   other in R1C1 and not the same as it.
5. xlOmittedCells: a formula reading a row or column of cells, with a number in the cell just past either end.
6. xlUnlockedFormulaCells: a formula in an unlocked cell.
7. xlEmptyCellReferences: a formula reading a cell, or a range, with nothing in it.
8 and 9, list validation and table formulas, are False away from validation and tables; 10, misleading formats,
never came up.

Ignore holds for the cell until something is written to it or it is cleared. A save writes each ignored check as
ignoredErrors, cells next to each other in one sqref; a cell ignoring two checks or more is written with none,
as Excel writes it. A range of more than one cell is error 1004, as is a protected sheet.
"""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_reference, split_sheet
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._objects import VBACollection, member, setter
from pyopenvba.interpreter._values import VBAInt, error, to_bool, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

#: The ignoredError attribute of each check, in the order of Errors' indexes.
ATTRIBUTES = ("evalError", "twoDigitTextYear", "numberStoredAsText", "formula", "formulaRange", "unlockedFormula",
              "emptyCellReference", "listDataValidation", "calculatedColumn", "x16r3:misleadingFormat")
_TEXT_DATE = re.compile(r"[0-9]{1,2}[/-][0-9]{1,2}[/-][0-9]{2}")


def _cell(target: Range) -> tuple[int, int]:
    if len(target.areas) != 1 or target.first.rows != 1 or target.first.columns != 1:
        raise error(1004)
    protection = target.sheet.protection
    if protection is not None and protection.contents:
        raise error(1004)
    return target.first.top, target.first.left


def _value(sheet: Worksheet, row: int, column: int) -> object:
    cell = sheet.cells_.get((row, column))
    if cell is None:
        return None
    return sheet.book.calculator.value_of(sheet.name, row, column) if cell.formula else cell.value


def flagged(sheet: Worksheet, row: int, column: int, index: int) -> bool:
    """Whether check ``index`` flags the cell."""
    cell = sheet.cells_.get((row, column))
    if cell is None or index > 10:
        return False
    if index in (2, 3):
        return not cell.formula and isinstance(cell.value, str) and _text_check(sheet, cell.value, index)
    if index == 8:
        if any(area.contains(row, column) for rule in sheet.validations for area in rule.areas):
            raise VBAUnsupportedError("Range.Errors(xlListDataValidation) on validated cells is not implemented")
        return False
    if index == 9:
        if any(table.area.contains(row, column) for table in sheet.tables):
            raise VBAUnsupportedError("Range.Errors(xlInconsistentListFormula) in a table is not implemented")
        return False
    if not cell.formula or index == 10:
        return False
    if index == 1:
        value = _value(sheet, row, column)
        return isinstance(value, ExcelError) and value.name != "#N/A"
    if index == 4:
        return _inconsistent(sheet, row, column)
    if index == 6:
        return not sheet.style_at(row, column).protection.locked
    return _reading(sheet, row, column, index)


def _text_check(sheet: Worksheet, text: str, index: int) -> bool:
    from pyopenvba.apps.excel._typing import is_date_format, typed

    result = typed(text, "General", epoch_1904=sheet.book.epoch_1904)
    if not isinstance(result.value, float):
        return False
    is_date = result.number_format is not None and is_date_format(result.number_format)
    if index == 2:
        return is_date and _TEXT_DATE.fullmatch(text) is not None
    return not is_date


def _r1c1(sheet: Worksheet, row: int, column: int) -> str | None:
    from pyopenvba.formula._r1c1 import from_a1

    cell = sheet.cells_.get((row, column))
    return from_a1(cell.formula, row, column) if cell is not None and cell.formula else None


def _inconsistent(sheet: Worksheet, row: int, column: int) -> bool:
    own = _r1c1(sheet, row, column)
    for before, after in (((row - 1, column), (row + 1, column)), ((row, column - 1), (row, column + 1))):
        if min(before) < 1:
            continue
        first, second = _r1c1(sheet, *before), _r1c1(sheet, *after)
        if first is not None and first == second and first != own:
            return True
    return False


def _references(sheet: Worksheet, formula: str) -> list[tuple[Worksheet, Area]]:
    """The cells a formula reads, by its references; a defined name it reads reports itself unsupported."""
    from pyopenvba.apps.excel import _names
    from pyopenvba.formula._parse import FormulaError, tokenize

    try:
        tokens = tokenize(formula.lstrip("="))
    except FormulaError:
        return []
    out: list[tuple[Worksheet, Area]] = []
    for at, token in enumerate(tokens):
        if token.kind == "name" and tokens[at + 1].kind != "open" and token.text.upper() not in ("TRUE", "FALSE") \
                and any(_names.same_entry(entry.name.split("!")[-1], token.text)
                        for entry in sheet.book.names_.entries):
            raise VBAUnsupportedError("Range.Errors of a formula reading a defined name is not implemented")
        if token.kind != "ref":
            continue
        name, bare = split_sheet(token.text)
        owner = sheet
        if name:
            found = next((one for one in sheet.book.sheets_ if one.name.lower() == name.lower()), None)
            if found is None:
                continue
            owner = found
        try:
            areas = parse_reference(bare.replace("$", ""))
        except ValueError:
            continue
        out.extend((owner, area) for area in areas)
    return out


def _holds(sheet: Worksheet, row: int, column: int) -> bool:
    from pyopenvba.apps.excel._region import holds_content

    return holds_content(sheet, row, column)


def _reading(sheet: Worksheet, row: int, column: int, index: int) -> bool:
    cell = sheet.cells_[(row, column)]
    references = _references(sheet, cell.formula)
    if index == 7:
        for owner, area in references:
            filled = sum(1 for (r, c) in owner.cells_ if area.contains(r, c) and _holds(owner, r, c))
            if filled < area.rows * area.columns:
                return True
        return False
    if index == 5:
        for owner, area in references:
            if area.rows * area.columns < 2:
                continue
            if area.rows > 1 and area.columns > 1:
                raise VBAUnsupportedError("Range.Errors(xlOmittedCells) of a formula reading a block is not "
                                          "implemented")
            ends = [(area.top - 1, area.left), (area.bottom + 1, area.left)] if area.columns == 1 else \
                [(area.top, area.left - 1), (area.top, area.right + 1)]
            for r, c in ends:
                if r < 1 or c < 1 or (owner is sheet and (r, c) == (row, column)):
                    continue
                value = _value(owner, r, c)
                if isinstance(value, float):
                    return True
        return False
    return False


def ignored(sheet: Worksheet, row: int, column: int, index: int) -> bool:
    if not 1 <= index <= 10:
        return False
    _loaded_check(sheet)
    cell = sheet.cells_.get((row, column))
    return cell is not None and bool(cell.ignored_errors & (1 << (index - 1)))


def _loaded_check(sheet: Worksheet) -> None:
    """A workbook read with ignoredErrors keeps them as they are; reading or changing them reports unsupported."""
    package = sheet.book.package
    if sheet.ignored_from_file is None:
        sheet.ignored_from_file = False
        if package is not None and sheet.part_name and package.has(sheet.part_name):
            sheet.ignored_from_file = "<ignoredErrors" in package.read(sheet.part_name).decode("utf-8", "replace")
    if sheet.ignored_from_file:
        raise VBAUnsupportedError("Range.Errors(...).Ignore on a sheet read with ignored errors is not implemented")


def set_ignored(sheet: Worksheet, row: int, column: int, index: int, value: bool) -> None:
    from pyopenvba.apps.excel._model import Cell

    if not 1 <= index <= 10:
        raise VBAUnsupportedError("Range.Errors(...).Ignore past index 10 is not implemented")
    _loaded_check(sheet)
    cell = sheet.cells_.get((row, column))
    if cell is None:
        if not value:
            return
        cell = sheet.cells_[(row, column)] = Cell()
    bit = 1 << (index - 1)
    cell.ignored_errors = cell.ignored_errors | bit if value else cell.ignored_errors & ~bit
    sheet.settle(row, column)
    sheet.ignored_changed = True
    sheet.touched()


def written(sheet: Worksheet, xml: str) -> str:
    """The sheet part with its ignoredErrors as the model has them."""
    if not sheet.ignored_changed:
        return xml
    from pyopenvba.apps.excel._model import coalesced

    singles: dict[int, list[Area]] = {}
    for (row, column), cell in sorted(sheet.cells_.items()):
        bits = cell.ignored_errors
        if bits and bits & (bits - 1) == 0:
            singles.setdefault(bits.bit_length(), []).append(Area(row, column, row, column))
    if len(singles) > 1:
        raise VBAUnsupportedError("Saving more than one kind of ignored error check on a sheet is not implemented")
    xml = re.sub(r"<ignoredErrors>.*?</ignoredErrors>", "", xml, flags=re.DOTALL)
    if singles:
        (index, areas), = singles.items()
        sqref = " ".join(area.address(absolute=False) for area in coalesced(areas))
        attribute = ATTRIBUTES[index - 1]
        namespace = ' xmlns:x16r3="http://schemas.microsoft.com/office/spreadsheetml/2018/08/main"' \
            if attribute.startswith("x16r3:") else ""
        element = f'<ignoredErrors><ignoredError sqref="{sqref}"{namespace} {attribute}="1"/></ignoredErrors>'
        after = re.search(r"<(?:smartTags|drawing|legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|"
                          r"controls|webPublishItems|tableParts|extLst)\b|</worksheet>", xml)
        assert after is not None
        xml = xml[:after.start()] + element + xml[after.start():]
    sheet.ignored_changed = False
    return xml


class Errors(VBACollection, ExcelObject):
    """A cell's error checks."""

    vba_type_name = "Errors"

    def __init__(self, target: Range) -> None:
        self.target = target

    def vba_items(self) -> list[object]:
        return [ErrorCheck(self, index) for index in range(1, 11)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        number = int(to_integer(index, "Long"))
        if number < 1:
            raise error(1004)
        return ErrorCheck(self, number)

    @member
    def Parent(self) -> object:
        return self.target

    @member
    def Application(self) -> object:
        return self.target.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")


class ErrorCheck(ExcelObject):
    """One of a cell's error checks: Excel's Error object."""

    vba_type_name = "Error"

    def __init__(self, errors: Errors, index: int) -> None:
        self.errors = errors
        self.target = errors.target
        self.index = index

    @member(default=True)
    def Value(self) -> object:
        row, column = _cell(self.target)
        return flagged(self.target.sheet, row, column, self.index)

    @member
    def Ignore(self) -> object:
        row, column = _cell(self.target)
        return ignored(self.target.sheet, row, column, self.index)

    @setter("Ignore")
    def _set_ignore(self, value: object) -> None:
        row, column = _cell(self.target)
        set_ignored(self.target.sheet, row, column, self.index, to_bool(value))

    @member
    def Parent(self) -> object:
        return self.errors

    @member
    def Application(self) -> object:
        return self.target.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")
