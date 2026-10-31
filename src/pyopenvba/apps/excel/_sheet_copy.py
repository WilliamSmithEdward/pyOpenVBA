"""Copy the modeled worksheet state without creating incorrect external links."""
from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Any

from pyopenvba._a1 import Area, quote_sheet
from pyopenvba.apps.excel._charts_file import xml_tree
from pyopenvba.apps.excel._xstring import decode, encode_text
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import split_sheet, tokenize
from pyopenvba.interpreter._values import EMPTY, MISSING, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Cell, NameEntry, Workbook, Worksheet


def formula(text: str, renamed: dict[str, str], external: bool) -> str:
    """A copied formula: a reference to a sheet copied with it names that sheet's copy, ``renamed`` holding each
    copy's name by its sheet's lower-case name; one to another sheet stays, within the workbook, and would need an
    external link, which is not implemented, from another."""
    for token in reversed(tokenize(text)):
        if token.kind not in ("ref", "name"):
            continue
        sheet, body = split_sheet(token.text)
        if not sheet:
            continue
        copy = renamed.get(sheet.casefold())
        if copy is not None:
            replacement = quote_sheet(copy) + "!" + body
            text = text[:token.at] + replacement + text[token.at + len(token.text):]
        elif external:
            raise VBAUnsupportedError("Copying formulas that require external workbook links is not implemented")
    return text


def _admit_filters(source: Worksheet) -> None:
    from pyopenvba.apps.excel._names import owned_reserved

    filter_names = {bare for _entry, bare in owned_reserved(source, ('Criteria', 'Extract', '_FilterDatabase'))}
    if (source.advanced_filter is not None or filter_names & {'Criteria', 'Extract'}
            or ('_FilterDatabase' in filter_names and source.auto_filter is None)):
        raise VBAUnsupportedError('Worksheet.Copy with AdvancedFilter state or owned metadata is not implemented')


def _admit_content(source: Worksheet) -> None:
    """Report what a copy of ``source`` would leave out, before anything is copied."""
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


@dataclass
class _Staged:
    """One sheet's copy, worked out before any copy is made."""

    source: Worksheet
    name: str
    cells: dict[tuple[int, int], Cell]
    conditional: list[Any]


def _staged(source: Worksheet, name: str, renamed: dict[str, str], external: bool) -> _Staged:
    """The copy's cells and conditional formats, their formulas reading the group's copies."""
    conditional = deepcopy(source.conditional_formats)
    for priority, rule in enumerate(sorted(conditional, key=lambda rule: int(rule.root.get('priority', '0'))), 1):
        rule.root.set('priority', str(priority))
    for rule in conditional:
        rule.areas = [Area(a.top, a.left, a.bottom, a.right, name) for a in rule.areas]
        for node in rule.root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}formula'):
            node.text = encode_text(formula(decode(node.text or ''), renamed, external))
        dxf_id = rule.root.get('dxfId')
        if external and dxf_id is not None:
            rule.dxf = rule.dxf if rule.dxf is not None else source.book.stylesheet.dxf(int(dxf_id))
    cells = {position: replace(cell) for position, cell in source.cells_.items()}
    for cell in cells.values():
        if cell.formula:
            cell.formula = formula(cell.formula, renamed, external)
            cell.stale = True
    return _Staged(source, name, cells, conditional)


def _used_names(source: Worksheet) -> set[str]:
    """The names a sheet's formulas and conditional formats call, and the names those call, in lower case."""
    used = {token.text.casefold() for cell in source.cells_.values() if cell.formula
            for token in tokenize(cell.formula) if token.kind == "name"}
    for rule in source.conditional_formats:
        for node in rule.root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}formula'):
            used.update(token.text.casefold() for token in tokenize(decode(node.text or ''))
                        if token.kind == "name")
    while True:
        expanded = used | {token.text.casefold() for entry in source.book.names_.entries
                          if split_sheet(entry.name)[1].casefold() in used
                          for token in tokenize(entry.refers_to) if token.kind == "name"}
        if expanded == used:
            return used
        used = expanded


def _copied_names(sources: list[Worksheet], destination: Workbook, renamed: dict[str, str],
                  external: bool) -> list[NameEntry]:
    """The defined names the copies bring: each sheet's own, as its copy's, and in another workbook the workbook's
    names they read or call."""
    from pyopenvba.apps.excel._model import NameEntry

    names: list[NameEntry] = []
    for source in sources:
        used = _used_names(source)
        for entry in source.book.names_.entries:
            scope, bare = split_sheet(entry.name)
            own_reference = any(token.kind == "ref" and split_sheet(token.text)[0].casefold() == source.name.casefold()
                                for token in tokenize(entry.refers_to))
            if scope.casefold() == source.name.casefold():
                wanted = quote_sheet(renamed[source.name.casefold()]) + "!" + bare
            elif external and not scope and (own_reference or bare.casefold() in used):
                wanted = bare
            else:
                continue
            text = formula(entry.refers_to, renamed, external)
            existing = next((one for one in [*destination.names_.entries, *names]
                             if one.name.casefold() == wanted.casefold()), None)
            if existing is not None:
                if existing.refers_to != text:
                    raise VBAUnsupportedError("Worksheet.Copy with conflicting defined names is not implemented")
                continue
            names.append(NameEntry(wanted, text, destination, visible=entry.visible, comment=entry.comment))
    return names


def copy_sheet(source: Worksheet, before: object, after: object) -> object:
    return copy_sheets([source], before, after)


def copy_sheets(sources: list[Worksheet], before: object, after: object) -> object:
    """Copy ``sources``, in their order, into a new workbook or at Before or After, each reference among them reading
    the copies; the first copy is then active (tests/fixtures/sheet_groups.json)."""
    from pyopenvba.apps.excel._model import Workbook, Worksheet
    from pyopenvba.apps.excel._names import same_sheet_name

    for source in sources:
        _admit_filters(source)
    if before is not MISSING and after is not MISSING:
        raise error(1004, "Specify Before or After, not both")
    anchor = before if before is not MISSING else after
    if anchor is not MISSING and not isinstance(anchor, Worksheet):
        raise error(1004, "A worksheet is required")
    book = sources[0].book
    app = book.application
    if isinstance(anchor, Worksheet) and anchor.book.application is not app:
        raise error(1004, "Worksheets must belong to the same Excel application")
    for source in sources:
        _admit_content(source)
    destination = anchor.book if isinstance(anchor, Worksheet) else Workbook(app, app.workbooks_.next_name())
    renamed: dict[str, str] = {}
    taken = [sheet.name for sheet in destination.sheets_]
    for source in sources:
        name, number = source.name, 2
        while any(same_sheet_name(one, name) for one in taken):
            suffix = f" ({number})"
            name = source.name[:31 - len(suffix)] + suffix
            number += 1
        taken.append(name)
        renamed[source.name.casefold()] = name
    external = destination is not book
    staged = [_staged(source, renamed[source.name.casefold()], renamed, external) for source in sources]
    names = _copied_names(sources, destination, renamed, external)
    if anchor is MISSING:
        app.workbooks_.books.append(destination)
    at = destination.sheets_.index(anchor) + (0 if before is not MISSING else 1) if isinstance(anchor, Worksheet) else None
    copies = [_published(stage, destination, None if at is None else at + offset, external)
              for offset, stage in enumerate(staged)]
    destination.names_.entries.extend(names)
    destination.names_.changed = True
    for copied in copies:
        copied.shape_changed()
    copies[0].Activate()
    return EMPTY


def _published(stage: _Staged, destination: Workbook, at: int | None, external: bool) -> Worksheet:
    """Make one staged copy in ``destination`` at ``at``, the end with None."""
    source, name, conditional = stage.source, stage.name, stage.conditional
    copied = destination.add_sheet(name, at=at)
    from pyopenvba.apps.excel._documents import copy_document

    copy_document(source, copied)
    copied.cells_ = stage.cells
    from pyopenvba.apps.excel._differential import renumber

    # The copy's rules take dxfs of their own, as Excel makes for each rule it makes (tests/fixtures/dxf_table/).
    for rule in conditional:
        dxf_id = rule.root.get('dxfId')
        if dxf_id is None:
            continue
        own = rule.dxf if rule.dxf is not None else source.book.stylesheet.dxf(int(dxf_id))
        if external:
            renumber(own, destination.stylesheet)
        rule.root.set('dxfId', str(destination.stylesheet.add_dxf(own)))
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
    # The copy's tab has the sheet's colour (tests/fixtures/sheet_tabs/), and its page breaks (page_breaks/).
    copied.tab_color = source.tab_color
    copied.tab_changed = source.tab_color is not None
    from pyopenvba.apps.excel._page_breaks import copied as copied_breaks

    copied_breaks(source, copied)
    from pyopenvba.apps.excel._custom_properties import copied as copied_properties

    # Not its custom properties (tests/fixtures/custom_properties.json).
    copied_properties(source, copied)
    from pyopenvba.apps.excel._page_setup import copied as copied_headers

    # And its headers and footers (tests/fixtures/page_headers.json, copied).
    copied_headers(source, copied)
    from pyopenvba.apps.excel._worksheet_settings import copied as copied_settings

    # And the settings Excel copies with it (tests/fixtures/worksheet_settings.json, copy).
    copied_settings(source, copied)
    # A protected sheet's copy is protected the same way (tests/fixtures/protection.json).
    copied.protection = source.protection
    copied.protection_allows = source.protection_allows
    copied.enable_selection = source.enable_selection
    # The copy has the view the sheet had, its zoom, panes and selection (tests/fixtures/windows/).
    from pyopenvba.apps.excel._windows import copied as copied_view

    copied_view(source, copied)
    return copied
