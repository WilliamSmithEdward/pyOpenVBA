"""Range.ApplyNames: references in formulas turned into the defined names that refer to them.

Measured in live Excel (tests/fixtures/apply_names.json, scripts/measure_apply_names.py):

- A reference written without a sheet that covers the same cells as a name on this sheet becomes the name, in
  every formula of the range, or of the whole sheet where the range is one cell. With IgnoreRelativeAbsolute,
  the default, the dollar signs do not matter; without it, only an absolute reference is replaced. A reference
  with its sheet written out, and text in quotes, stay as they are. A range of one cell, A1:A1, becomes @Rate.
- Names:= limits the names to those given, by one name or an array of them; one that is not a name is passed
  over, and none at all changes nothing.
- A workbook with no names is error 1004; a range with no formula, or with no reference a name covers, is error
  1004, "Microsoft Excel cannot find any references to replace."
- UseRowColumnNames left nothing to replace where a name covers a cell's row and another its column; Order,
  OmitColumn, OmitRow and AppendLast, which bear on such names, changed nothing measured.

Two names covering the same cells report themselves unsupported: Excel took the later of Rate and Rate2, which
alphabetical and defined order both make the later. So does a formula that is a range alone, whose @ was not
measured.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import FormulaError, split_sheet, tokenize
from pyopenvba.interpreter._values import MISSING, VBAArray, error, to_bool, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import NameEntry, Range, Worksheet

_NOTHING = "Microsoft Excel cannot find any references to replace."


def _covered(entry: NameEntry, sheet: Worksheet) -> Area | None:
    """The cells a name covers on ``sheet``, None for a name that is not a range there."""
    if entry.sheet is not None and entry.sheet is not sheet:
        return None
    text = entry.refers_to[1:] if entry.refers_to.startswith("=") else entry.refers_to
    try:
        sheet_name, reference = split_sheet(text)
    except FormulaError:
        return None
    if sheet_name.strip("'").lower() != sheet.name.lower() or "$" not in reference:
        return None
    try:
        return parse_area(reference.replace("$", ""))
    except ValueError:
        return None


def _chosen(target: Range, names: object) -> list[NameEntry]:
    entries = list(target.sheet.book.names_.entries)
    if names is MISSING:
        return entries
    wanted = [to_text(one) for one in names.elements()] if isinstance(names, VBAArray) else [to_text(names)]
    lowered = {one.lower() for one in wanted}
    return [entry for entry in entries if entry.name.lower() in lowered]


def apply_names(target: Range, names: object, ignore_relative_absolute: object) -> None:
    sheet = target.sheet
    if not sheet.book.names_.entries:
        raise error(1004, "ApplyNames method of Range class failed")
    chosen = [(entry, area) for entry in _chosen(target, names) if (area := _covered(entry, sheet)) is not None]
    if names is not MISSING and not chosen:
        return
    loose = True if ignore_relative_absolute is MISSING else to_bool(ignore_relative_absolute)
    one_cell = len(target.areas) == 1 and target.areas[0].rows == 1 and target.areas[0].columns == 1
    cells = [(row, column, cell) for (row, column), cell in sorted(sheet.cells_.items())
             if cell.formula and (one_cell or any(area.contains(row, column) for area in target.areas))]
    rewritten: list[tuple[int, int, str]] = []
    for row, column, cell in cells:
        formula = _applied(cell.formula, chosen, loose)
        if formula != cell.formula:
            rewritten.append((row, column, formula))
    if not rewritten:
        raise error(1004, _NOTHING)
    from pyopenvba.apps.excel import _arrays
    from pyopenvba.apps.excel._model import Range as RangeObject

    for row, column, formula in rewritten:
        found = _arrays.array_at(sheet, row, column) if sheet.array_formulas else None
        if found is not None:
            anchor, area = found
            if anchor == (row, column):
                RangeObject(sheet, [area]).vba_set("FormulaArray", formula)
            continue
        RangeObject(sheet, [Area(row, column, row, column)]).vba_set("Formula", formula)


def _applied(formula: str, chosen: list[tuple[NameEntry, Area]], loose: bool) -> str:
    body = formula[1:]
    pieces: list[tuple[int, int, str]] = []
    try:
        tokens = tokenize(body)
    except FormulaError:
        return formula
    for token in tokens:
        if token.kind != "ref" or "!" in token.text:
            continue
        if not loose and token.text.count("$") != (2 if ":" not in token.text else 4):
            continue
        try:
            area = parse_area(token.text.replace("$", ""))
        except ValueError:
            continue
        matches = [entry for entry, covered in chosen if covered == area]
        if not matches:
            continue
        if len(matches) > 1:
            raise VBAUnsupportedError("ApplyNames where two names cover the same cells is not implemented")
        name = matches[0].name
        if ":" in token.text:
            if area.rows == 1 and area.columns == 1:
                name = "@" + name
            elif token.at == 0 and token.at + len(token.text) == len(body):
                raise VBAUnsupportedError("ApplyNames to a formula that is a range alone is not implemented")
        pieces.append((token.at, token.at + len(token.text), name))
    for start, stop, replacement in reversed(pieces):
        body = body[:start] + replacement + body[stop:]
    return "=" + body
