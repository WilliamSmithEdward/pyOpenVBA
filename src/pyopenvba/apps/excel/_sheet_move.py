"""Worksheet relocation with external-link checks before mutation."""
from __future__ import annotations
from typing import TYPE_CHECKING
from pyopenvba.formula._parse import split_sheet, tokenize
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import MISSING, EMPTY, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet


def move_sheet(source: Worksheet, before: object, after: object) -> object:
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.apps.excel._sheet_copy import copy_sheet

    if before is not MISSING and after is not MISSING:
        raise error(1004, "Specify Before or After, not both")
    anchor = before if before is not MISSING else after
    if anchor is not MISSING and not isinstance(anchor, Worksheet):
        raise error(1004, "A worksheet is required")
    origin = source.book
    app = origin.application
    if origin not in app.workbooks_.books or source not in origin.sheets_:
        raise error(1004, "The source worksheet is not open")
    if isinstance(anchor, Worksheet):
        if anchor.book.application is not app or anchor.book not in app.workbooks_.books or anchor not in anchor.book.sheets_:
            raise error(1004, "The destination worksheet must be open in the same application")
        if anchor.book is origin:
            if anchor is not source:
                origin.sheets_.remove(source)
                position = origin.sheets_.index(anchor) + (0 if before is not MISSING else 1)
                origin.sheets_.insert(position, source)
                origin.names_.changed = True
                origin.saved = False
                origin.calculator.rebuild()
            source.Activate()
            return EMPTY
    if anchor is MISSING and len(origin.sheets_) == 1:
        raise error(1004, "Cannot move the only worksheet into a new workbook")
    if origin.package is not None:
        from pyopenvba.apps.excel._io import has_reserved_sheet_names

        if has_reserved_sheet_names(source):
            raise VBAUnsupportedError("Cross-workbook Worksheet.Move with reserved sheet names is not implemented")

    def refers_to_source(text: str) -> bool:
        return any(token.kind in {"ref", "name"} and
                   split_sheet(token.text)[0].casefold() == source.name.casefold()
                   for token in tokenize(text))

    # Remaining sheets/names must not silently acquire broken local references.
    for sheet in origin.sheets_:
        if sheet is not source and any(cell.formula and refers_to_source(cell.formula) for cell in sheet.cells_.values()):
            raise VBAUnsupportedError("Worksheet.Move would require external-link support for source formulas")
    for entry in origin.names_.entries:
        scope, _ = split_sheet(entry.name)
        if scope.casefold() != source.name.casefold() and refers_to_source(entry.refers_to):
            raise VBAUnsupportedError("Worksheet.Move would require external-link support for source names")
    old_name = source.name
    local_entries = [entry for entry in origin.names_.entries
                     if split_sheet(entry.name)[0].casefold() == old_name.casefold()]
    was_active = origin.active_sheet
    old_index = origin.sheets_.index(source)
    # Copy performs validation and prepares cells and names before publishing them.
    copy_sheet(source, before, after)
    destination = app.active_book
    assert destination is not None
    copied = destination.active_sheet
    assert copied is not None
    # Excel replaces the COM objects across workbooks; old VBA references fail.
    for entry in local_entries:
        entry.invalidated = True
    origin.sheets_.remove(source)
    origin.names_.entries = [entry for entry in origin.names_.entries
                             if not any(entry is local for local in local_entries)]
    origin.names_.changed = True
    source.invalidated = True
    if origin.sheets_:
        origin.active_sheet_index = (origin.sheets_.index(was_active) if was_active in origin.sheets_
                                     else min(old_index, len(origin.sheets_) - 1))
    else:
        app.workbooks_.remove(origin)
    origin.saved = False
    origin.calculator.rebuild()
    copied.shape_changed()
    copied.Activate()
    return EMPTY
