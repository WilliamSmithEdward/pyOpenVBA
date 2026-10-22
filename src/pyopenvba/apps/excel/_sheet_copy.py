"""Copy the modeled worksheet state without creating incorrect external links."""
from __future__ import annotations

from copy import deepcopy
from dataclasses import replace
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, quote_sheet
from pyopenvba.apps.excel._charts_file import xml_tree
from pyopenvba.apps.excel._xstring import decode, encode_text
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import split_sheet, tokenize
from pyopenvba.interpreter._values import EMPTY, MISSING, error

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
    from pyopenvba.apps.excel._model import NameEntry, Workbook, Worksheet
    from pyopenvba.apps.excel._names import owned_reserved

    filter_names = {bare for _entry, bare in owned_reserved(source, ('Criteria', 'Extract', '_FilterDatabase'))}
    if (source.advanced_filter is not None or filter_names & {'Criteria', 'Extract'}
            or ('_FilterDatabase' in filter_names and source.auto_filter is None)):
        raise VBAUnsupportedError('Worksheet.Copy with AdvancedFilter state or owned metadata is not implemented')

    if before is not MISSING and after is not MISSING:
        raise error(1004, "Specify Before or After, not both")
    anchor = before if before is not MISSING else after
    if anchor is not MISSING and not isinstance(anchor, Worksheet):
        raise error(1004, "A worksheet is required")
    app = source.book.application
    if isinstance(anchor, Worksheet) and anchor.book.application is not app:
        raise error(1004, "Worksheets must belong to the same Excel application")
    from pyopenvba.apps.excel._page_setup import guard_copy

    guard_copy(source)
    from pyopenvba.apps.excel._outline import admit_sheet_copy

    admit_sheet_copy(source)
    if source.shapes_:
        raise VBAUnsupportedError("Worksheet.Copy with drawings or controls is not implemented")
    if source.tables:
        raise VBAUnsupportedError("Worksheet.Copy with tables is not implemented")
    if source.validations:
        raise VBAUnsupportedError("Worksheet.Copy with data validation is not implemented")
    if source.notes:
        raise VBAUnsupportedError("Worksheet.Copy with notes is not implemented")
    from pyopenvba.apps.excel._conditional_predicates import PREDICATE_TYPES

    supported_conditions = {'cellIs', 'expression', 'top10', 'aboveAverage',
                            'duplicateValues', 'uniqueValues'} | PREDICATE_TYPES
    if any(rule.root.get('type') not in supported_conditions for rule in source.conditional_formats):
        raise VBAUnsupportedError("Worksheet.Copy with this conditional rule type is not implemented")
    if source.book.package is not None and source.part_name:
        xml = source.book.package.read(source.part_name).decode("utf-8")
        if any(f"<{tag}" in xml for tag in ("tableParts", "dataValidations", "legacyDrawing")):
            raise VBAUnsupportedError("Worksheet.Copy with tables, validation, comments or conditional formatting is not implemented")
        root = xml_tree(xml)
        extensions = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}extLst')
        if extensions is not None and any(
            entry.tag.rsplit('}', 1)[-1] in ('conditionalFormatting', 'conditionalFormattings')
            for entry in extensions.iter()
        ):
            raise VBAUnsupportedError("Worksheet.Copy with extended conditional formatting is not implemented")
    destination = anchor.book if isinstance(anchor, Worksheet) else Workbook(app, app.workbooks_.next_name())
    name, number = source.name, 2
    while any(sheet.name.casefold() == name.casefold() for sheet in destination.sheets_):
        suffix = f" ({number})"
        name = source.name[:31 - len(suffix)] + suffix
        number += 1
    external = destination is not source.book
    conditional = deepcopy(source.conditional_formats)
    for priority, rule in enumerate(sorted(conditional, key=lambda rule: int(rule.root.get('priority', '0'))), 1):
        rule.root.set('priority', str(priority))
    conditional_formulas: list[str] = []
    for rule in conditional:
        rule.areas = [Area(a.top, a.left, a.bottom, a.right, name) for a in rule.areas]
        for node in rule.root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}formula'):
            text = decode(node.text or '')
            conditional_formulas.append(text)
            node.text = encode_text(formula(text, source, name, external))
        dxf_id = rule.root.get('dxfId')
        if external and dxf_id is not None:
            rule.dxf = rule.dxf if rule.dxf is not None else source.book.stylesheet.dxf(int(dxf_id))
    cells = {position: replace(cell) for position, cell in source.cells_.items()}
    for cell in cells.values():
        if cell.formula:
            cell.formula = formula(cell.formula, source, name, external)
            cell.stale = True
    used = {token.text.casefold() for cell in cells.values() if cell.formula
            for token in tokenize(cell.formula) if token.kind == "name"}
    used.update(token.text.casefold() for text in conditional_formulas
                for token in tokenize(text) if token.kind == "name")
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
    from pyopenvba.apps.excel._documents import copy_document

    copy_document(source, copied)
    copied.cells_ = cells
    if external:
        from pyopenvba.apps.excel._differential import number_format, set_number_format

        for rule in conditional:
            if rule.root.get('dxfId') is not None and rule.dxf is not None:
                if rule.dxf.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}numFmt') is not None:
                    set_number_format(rule.dxf, number_format(rule.dxf, destination.stylesheet), destination.stylesheet)
                rule.root.set('dxfId', str(destination.stylesheet.ensure_dxf(rule.dxf)))
    copied.conditional_formats = conditional
    from pyopenvba.apps.excel._conditional import rule_groups
    from pyopenvba.apps.excel._conditional_border_context import initialize

    for group in rule_groups(conditional):
        group_slot = group[0].color_slot
        for order, rule in enumerate(sorted(group, key=lambda rule: int(rule.root.get('priority', '0')))):
            rule.color_slot = group_slot
            rule.child_order = order
    initialize(copied)
    copied.conditional_changed = bool(conditional)
    # The copied cells keep their shared formulas, so the copy keeps the blocks those name, its arrays and its data
    # tables, whose input cells are on the sheet.
    copied.shared_groups = dict(source.shared_groups)
    copied.array_formulas = dict(source.array_formulas)
    copied.data_tables = dict(source.data_tables)
    copied.dims = source.dims.copied(copied)
    from pyopenvba.apps.excel._outline import copied as copied_outline

    copied_outline(source, copied)
    copied.merged_areas = [Area(a.top, a.left, a.bottom, a.right, name) for a in source.merged_areas]
    copied.merges_dirty = bool(copied.merged_areas)
    from pyopenvba.apps.excel._hyperlinks import sheet_copied

    # The copy has the sheet's links, in their order (tests/fixtures/hyperlinks/structure.xlsx).
    sheet_copied(source, copied)
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
