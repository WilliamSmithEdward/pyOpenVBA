"""Copy the modeled worksheet state without creating incorrect external links."""
from __future__ import annotations
from dataclasses import replace
from typing import TYPE_CHECKING
from pyopenvba._a1 import Area, quote_sheet
from pyopenvba.formula._parse import split_sheet, tokenize
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import MISSING, EMPTY, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet


def formula(text: str, source: Worksheet, name: str, external: bool) -> str:
    for token in reversed(tokenize(text)):
        if token.kind not in ("ref", "name"):
            continue
        sheet, body = split_sheet(token.text)
        if not sheet:
            continue
        if sheet.casefold() == source.name.casefold():
            replacement = quote_sheet(name) + "!" + body
            text = text[:token.at] + replacement + text[token.at + len(token.text):]
        elif external:
            raise VBAUnsupportedError("Copying formulas that require external workbook links is not implemented")
    return text


def copy_sheet(source: Worksheet, before: object, after: object) -> object:
    from pyopenvba.apps.excel._model import Worksheet, NameEntry, Workbook

    if before is not MISSING and after is not MISSING:
        raise error(1004, "Specify Before or After, not both")
    anchor = before if before is not MISSING else after
    if anchor is not MISSING and not isinstance(anchor, Worksheet):
        raise error(1004, "A worksheet is required")
    app = source.book.application
    if isinstance(anchor, Worksheet) and anchor.book.application is not app:
        raise error(1004, "Worksheets must belong to the same Excel application")
    if source.shapes_:
        raise VBAUnsupportedError("Worksheet.Copy with drawings or controls is not implemented")
    if source.tables:
        raise VBAUnsupportedError("Worksheet.Copy with tables is not implemented")
    if source.validations:
        raise VBAUnsupportedError("Worksheet.Copy with data validation is not implemented")
    if source.book.package is not None and source.part_name:
        xml = source.book.package.read(source.part_name).decode("utf-8")
        if any(f"<{tag}" in xml for tag in ("tableParts", "conditionalFormatting", "dataValidations", "hyperlinks", "legacyDrawing")):
            raise VBAUnsupportedError("Worksheet.Copy with tables, validation, comments or conditional formatting is not implemented")
    destination = anchor.book if isinstance(anchor, Worksheet) else Workbook(app, app.workbooks_.next_name())
    name, number = source.name, 2
    while any(sheet.name.casefold() == name.casefold() for sheet in destination.sheets_):
        suffix = f" ({number})"
        name = source.name[:31 - len(suffix)] + suffix
        number += 1
    external = destination is not source.book
    cells = {position: replace(cell) for position, cell in source.cells_.items()}
    for cell in cells.values():
        if cell.formula:
            cell.formula = formula(cell.formula, source, name, external)
            cell.stale = True
    used = {token.text.casefold() for cell in cells.values() if cell.formula
            for token in tokenize(cell.formula) if token.kind == "name"}
    while True:
        expanded = used | {token.text.casefold() for entry in source.book.names_.entries
                          if split_sheet(entry.name)[1].casefold() in used
                          for token in tokenize(entry.refers_to) if token.kind == "name"}
        if expanded == used:
            break
        used = expanded
    names: list[NameEntry] = []
    for entry in source.book.names_.entries:
        scope, bare = split_sheet(entry.name)
        own_reference = any(token.kind == "ref" and split_sheet(token.text)[0].casefold() == source.name.casefold()
                            for token in tokenize(entry.refers_to))
        if scope.casefold() == source.name.casefold():
            wanted = quote_sheet(name) + "!" + bare
        elif external and not scope and (own_reference or bare.casefold() in used):
            wanted = bare
        else:
            continue
        text = formula(entry.refers_to, source, name, external)
        existing = next((one for one in destination.names_.entries if one.name.casefold() == wanted.casefold()), None)
        if existing is not None:
            if existing.refers_to != text:
                raise VBAUnsupportedError("Worksheet.Copy with conflicting defined names is not implemented")
            continue
        names.append(NameEntry(wanted, text, destination, visible=entry.visible, comment=entry.comment))
    if anchor is MISSING:
        app.workbooks_.books.append(destination)
    at = destination.sheets_.index(anchor) + (0 if before is not MISSING else 1) if isinstance(anchor, Worksheet) else None
    copied = destination.add_sheet(name, at=at)
    copied.cells_ = cells
    # The copied cells keep their shared formulas, so the copy keeps the blocks those name, and its arrays.
    copied.shared_groups = dict(source.shared_groups)
    copied.array_formulas = dict(source.array_formulas)
    copied.dims = source.dims.copied(copied)
    copied.merged_areas = [Area(a.top, a.left, a.bottom, a.right, name) for a in source.merged_areas]
    copied.merges_dirty = bool(copied.merged_areas)
    # A protected sheet's copy is protected the same way (tests/fixtures/protection.json).
    copied.protection = source.protection
    copied.protection_allows = source.protection_allows
    copied.enable_selection = source.enable_selection
    # The copy has the view the sheet had, its zoom, panes and selection (tests/fixtures/windows/).
    from pyopenvba.apps.excel._windows import copied as copied_view

    copied_view(source, copied)
    destination.names_.entries.extend(names)
    destination.names_.changed = True
    copied.shape_changed()
    copied.Activate()
    return EMPTY
