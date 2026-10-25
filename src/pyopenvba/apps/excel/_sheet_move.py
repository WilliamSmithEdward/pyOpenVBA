"""Worksheet relocation with external-link checks before mutation."""
from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar
from pyopenvba.formula._parse import split_sheet, tokenize
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import MISSING, EMPTY, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

#: A sheet, or what stands for one: a worksheet, or its name where a test lays tabs out by name.
Sheet = TypeVar("Sheet")


def placed(order: list[Sheet], group: list[Sheet], anchor: Sheet, after: bool) -> list[Sheet]:
    """The tab order once ``group``, named in its array's order, moves Before or After ``anchor`` in its own workbook:
    the group, in tab order, before the first sheet outside it that stands at the place the anchor gives or past it,
    or last -- one place further when the array's first sheet stands past that place and some of the group before
    it (tests/fixtures/sheet_group_moves.json, 608 moves)."""
    target = order.index(anchor) + (1 if after else 0)
    if order.index(group[0]) > target and any(order.index(sheet) < target for sheet in group):
        target += 1
    block = [sheet for sheet in order if sheet in group]
    rest = [sheet for sheet in order if sheet not in group]
    stand = next((sheet for sheet in order[target:] if sheet not in group), None)
    at = rest.index(stand) if stand is not None else len(rest)
    return rest[:at] + block + rest[at:]


def move_sheets(sheets: list[Worksheet], before: object, after: object) -> object:
    """Move a group of sheets, named in its array's order, beside a sheet of its workbook or into another workbook,
    a new one when no place is given (tests/fixtures/sheet_groups.json)."""
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.apps.excel._sheet_groups import in_tab_order, survivor

    origin = sheets[0].book
    group = in_tab_order(origin, sheets)
    if len(group) == 1:
        return move_sheet(group[0], before, after)
    if before is not MISSING and after is not MISSING:
        raise error(1004, "Specify Before or After, not both")
    anchor = before if before is not MISSING else after
    if anchor is not MISSING and not isinstance(anchor, Worksheet):
        raise error(1004, "A worksheet is required")
    app = origin.application
    if origin not in app.workbooks_.books or any(sheet not in origin.sheets_ for sheet in group):
        raise error(1004, "The source worksheet is not open")
    if isinstance(anchor, Worksheet):
        if anchor.book.application is not app or anchor.book not in app.workbooks_.books or anchor not in anchor.book.sheets_:
            raise error(1004, "The destination worksheet must be open in the same application")
        if anchor.book is origin:
            # An anchor of the group's own places it by the same rule (tests/fixtures/sheet_groups.json,
            # move_anchor_in_*, all_move_after).
            origin.sheets_[:] = placed(origin.sheets_, sheets, anchor, after is not MISSING)
            origin.names_.changed = True
            origin.saved = False
            origin.calculator.rebuild()
            in_tab_order(origin, group)[0].Activate()
            return EMPTY
    if anchor is MISSING and len(group) == len(origin.sheets_):
        raise error(1004, "Move method of Sheets class failed")
    if origin.package is not None:
        from pyopenvba.apps.excel._io import has_reserved_sheet_names

        if any(has_reserved_sheet_names(sheet) for sheet in group):
            raise VBAUnsupportedError("Cross-workbook Worksheet.Move with reserved sheet names is not implemented")
    names = {sheet.name.casefold() for sheet in group}

    def refers_to_group(text: str) -> bool:
        return any(token.kind in {"ref", "name"} and split_sheet(token.text)[0].casefold() in names
                   for token in tokenize(text))

    # Sheets and names that stay must not silently acquire broken references to the ones that go.
    for sheet in origin.sheets_:
        if not any(one is sheet for one in group) and any(
                cell.formula and refers_to_group(cell.formula) for cell in sheet.cells_.values()):
            raise VBAUnsupportedError("Worksheet.Move would require external-link support for source formulas")
    for entry in origin.names_.entries:
        scope, _ = split_sheet(entry.name)
        if scope.casefold() not in names and refers_to_group(entry.refers_to):
            raise VBAUnsupportedError("Worksheet.Move would require external-link support for source names")
    local_entries = [entry for entry in origin.names_.entries if split_sheet(entry.name)[0].casefold() in names]
    staying = survivor(origin, group)
    from pyopenvba.apps.excel._sheet_copy import copy_sheets

    # Copy performs validation and prepares cells and names before publishing them.
    copy_sheets(group, before, after)
    destination = app.active_book
    assert destination is not None
    first = destination.active_sheet
    assert first is not None
    # Excel replaces the COM objects across workbooks; old VBA references fail.
    for entry in local_entries:
        entry.invalidated = True
    from pyopenvba.apps.excel._documents import remove_document

    for sheet in group:
        remove_document(sheet)
        origin.sheets_.remove(sheet)
        sheet.invalidated = True
    origin.names_.entries = [entry for entry in origin.names_.entries
                             if not any(entry is local for local in local_entries)]
    origin.names_.changed = True
    if origin.sheets_:
        if staying is not None:
            origin.active_sheet_index = origin.sheets_.index(staying)
        origin.selected_sheets = []
    else:
        app.workbooks_.remove(origin)
    origin.saved = False
    origin.calculator.rebuild()
    first.Activate()
    return EMPTY


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
    from pyopenvba.apps.excel._documents import remove_document

    remove_document(source)
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
