"""Bounded shared/CSE storage ownership for worksheet rename staging."""
from __future__ import annotations

import re
from dataclasses import replace
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, parse_area
from pyopenvba._xml import attributes, tag_attributes
from pyopenvba.apps.excel._charts_file import xml_parser
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import FormulaError, shift_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook, Worksheet
    from pyopenvba.apps.excel._worksheet_rename import Budget

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
MAX_REGIONS = 1024
MAX_REGION_CELLS = 250_000
MAX_STORED_NODES = 1_000_000


def refuse(detail: str) -> VBAUnsupportedError:
    return VBAUnsupportedError('Renaming a worksheet with ' + detail + ' is not implemented')


def _area(area: Area, owner: str) -> None:
    if (not 1 <= area.top <= area.bottom <= MAX_ROWS or not 1 <= area.left <= area.right <= MAX_COLUMNS
            or area.sheet not in ('', owner)):
        raise refuse('unowned formula storage bounds')


def admit(book: Workbook, target: Worksheet, wanted: str, budget: Budget) -> list[tuple[int, Area]]:
    """Check fixed ownership without evaluating or materializing array members."""
    from pyopenvba.apps.excel._worksheet_rename import MAX_EXPRESSION_TEXT

    # Whole-population preflight precedes *every* shared-formula lexer call,
    # even when this helper is called independently of the rename coordinator.
    for sheet in book.sheets_:
        for cell in sheet.cells_.values():
            if len(cell.formula) > MAX_EXPRESSION_TEXT:
                raise refuse('formula storage beyond the supported expression text length')
    count = sum(len(sheet.shared_groups) + len(sheet.array_formulas) for sheet in book.sheets_)
    if count > MAX_REGIONS:
        raise refuse('formula storage beyond the supported region budget')
    population = sum(area.rows * area.columns for sheet in book.sheets_
                     for area in (*sheet.shared_groups.values(), *sheet.array_formulas.values()))
    if population > MAX_REGION_CELLS:
        raise refuse('formula storage beyond the supported region or cell budget')
    from pyopenvba.apps.excel._shared import names_a_sheet

    updates: list[tuple[int, Area]] = []
    for sheet in book.sheets_:
        occupied: dict[tuple[int, int], tuple[str, object]] = {}
        members: dict[int, list[tuple[tuple[int, int], str]]] = {key: [] for key in sheet.shared_groups}
        for key, area in sheet.shared_groups.items():
            if type(key) is not int or key < 0:
                raise refuse('unrecognized shared formula identifiers')
            _area(area, sheet.name)
            if sheet is target and area.sheet:
                updates.append((key, replace(area, sheet=wanted)))
        for anchor, area in sheet.array_formulas.items():
            _area(area, sheet.name)
            if area.sheet or anchor != (area.top, area.left):
                raise refuse('unowned fixed array anchors')
            cell = sheet.cells_.get(anchor)
            if cell is None or not cell.formula or cell.shared is not None:
                raise refuse('missing or shared fixed array anchors')
        for kind, regions in (('shared', sheet.shared_groups), ('array', sheet.array_formulas)):
            for key, area in regions.items():
                for row in range(area.top, area.bottom + 1):
                    for column in range(area.left, area.right + 1):
                        position = (row, column)
                        if position in occupied:
                            raise refuse('overlapping formula storage regions')
                        occupied[position] = (kind, key)
        for position, cell in sheet.cells_.items():
            assigned = occupied.get(position)
            if cell.shared is not None:
                if (type(cell.shared) is not int or cell.shared not in members or not cell.formula
                        or assigned != ('shared', cell.shared)):
                    raise refuse('orphan shared formula members')
                members[cell.shared].append((position, cell.formula))
            if assigned is not None and assigned[0] == 'array' and position != assigned[1] and cell.formula:
                raise refuse('independent formulas inside fixed array blocks')
        for cells in members.values():
            if not cells:
                raise refuse('shared formula groups without members')
            anchor, formula = min(cells)
            budget.text(formula)
            if names_a_sheet(formula):
                raise refuse('sheet-qualified shared formulas')
            for position, text in cells:
                # shift_text lexes the master for each follower: charge before
                # that work, independently of the ordinary formula plan.
                budget.text(formula)
                try:
                    if shift_text(formula, position[0] - anchor[0], position[1] - anchor[1]) != text:
                        raise refuse('inconsistent expanded shared formulas')
                except FormulaError:
                    raise refuse('unrecognized shared formulas') from None
    return updates


def stored(raw: bytes) -> None:
    """Admit actual formula elements only where the lexical cell reader owns them."""
    parser = xml_parser()
    path: list[str] = []
    actual: list[tuple[int, int]] = []
    data_start = 0
    data_open_end = 0
    data_span: tuple[int, int] | None = None
    actual_rows: list[tuple[int, int]] = []
    actual_cells: list[tuple[int, int]] = []
    row_start = 0
    row_open_end = 0
    cell_start = 0
    cell_open_end = 0
    formula_start = 0
    formula_open_end = 0
    positions: set[str] = set()
    nodes = 0
    formula_cells: set[str] = set()
    current_cell = ''
    lexical_data = False
    lexical_row = False
    lexical_cell = False
    formula_attributes: dict[str, str] | None = None
    has_body = False
    masters: set[str] = set()
    shared_ids: set[str] = set()

    def plain(local: bytes) -> bool:
        at = parser.CurrentByteIndex
        return (raw[at:at + len(local)] == local
                and raw[at + len(local):at + len(local) + 1] in (b' ', b'\t', b'\r', b'\n', b'>', b'/'))

    def opening_end(at: int) -> int:
        quote = 0
        for index in range(at, len(raw)):
            value = raw[index]
            if quote:
                if value == quote:
                    quote = 0
            elif value in (34, 39):
                quote = value
            elif value == 62:
                return index + 1
        raise refuse('unterminated stored formula tags')

    def owned_attributes(attrs: dict[str, str], keys: tuple[str, ...], *, formula: bool = False) -> None:
        at = parser.CurrentByteIndex
        # Match the existing loader's actual attribute grammar, independently
        # of mutable modeled cells and of unrelated opaque attributes.
        text = raw[at:opening_end(at)].decode('utf8')
        legacy = attributes(text) if formula else tag_attributes(text)
        if any(attrs.get(key) != legacy.get(key) for key in keys):
            raise refuse('unowned stored formula key attributes')

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal nodes, current_cell, lexical_data, lexical_row, lexical_cell, formula_attributes, has_body
        nonlocal formula_start, formula_open_end, data_start, data_open_end
        nonlocal row_start, row_open_end, cell_start, cell_open_end
        nodes += 1
        if nodes > MAX_STORED_NODES or len(path) >= 64:
            raise refuse('stored formulas beyond the metadata node or depth budget')
        path.append(name)
        local = name.rsplit('}', 1)[-1]
        if path == [NS + '}worksheet', NS + '}sheetData']:
            if data_span is not None:
                raise refuse('multiple stored worksheet data containers')
            data_start = parser.CurrentByteIndex
            data_open_end = opening_end(data_start)
            lexical_data = plain(b'<sheetData')
        if path == [NS + '}worksheet', NS + '}sheetData', NS + '}row']:
            owned_attributes(attrs, ('r',))
            lexical_row = plain(b'<row')
            row_start = parser.CurrentByteIndex
            row_open_end = opening_end(row_start)
        if len(path) == 4 and path == [NS + '}worksheet', NS + '}sheetData', NS + '}row', NS + '}c']:
            owned_attributes(attrs, ('r', 't', 's', 'cm', 'vm'))
            lexical_cell = plain(b'<c')
            cell_start = parser.CurrentByteIndex
            cell_open_end = opening_end(cell_start)
            if 'cm' in attrs or 'vm' in attrs:
                lexical_cell = False
            current_cell = attrs.get('r', '')
            if current_cell in positions:
                raise refuse('duplicate stored formula cells')
            positions.add(current_cell)
        in_data = path[:2] == [NS + '}worksheet', NS + '}sheetData']
        if local == 'f' and not in_data:
            x14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}'
            xm = 'http://schemas.microsoft.com/office/excel/2006/main}f'
            prefix = [NS + '}worksheet', NS + '}extLst', NS + '}ext']
            validation = (path[:3] == prefix and path[3:5] == [x14 + 'dataValidations', x14 + 'dataValidation']
                          and path[5:] in ([x14 + 'formula1', xm], [x14 + 'formula2', xm]))
            conditional = (path[:3] == prefix and path[3:] == [x14 + 'conditionalFormattings',
                           x14 + 'conditionalFormatting', x14 + 'cfRule', xm])
            if not (validation or conditional):
                raise refuse('unrecognized formula carriers outside cell storage')
        if local == 'f' and in_data:
            if path != [NS + '}worksheet', NS + '}sheetData', NS + '}row', NS + '}c', NS + '}f']:
                raise refuse('nested or foreign formula elements')
            if not lexical_data or not lexical_row or not lexical_cell:
                raise refuse('qualified or rich formula storage')
            owned_attributes(attrs, ('t', 'si', 'ref', 'ca', 'aca'), formula=True)
            formula_attributes = attrs
            has_body = False
            if current_cell in formula_cells:
                raise refuse('multiple formulas in one stored cell')
            formula_cells.add(current_cell)
            try:
                area = parse_area(current_cell)
            except ValueError:
                raise refuse('unrecognized stored formula coordinates') from None
            _area(area, '')
            if area.rows != 1 or area.columns != 1:
                raise refuse('unrecognized stored formula coordinates')
            formula_start = parser.CurrentByteIndex
            formula_open_end = opening_end(formula_start)
            kind = attrs.get('t', 'normal')
            allowed = {'t', 'ca'}
            if kind == 'shared':
                allowed |= {'si', 'ref'}
                shared_ids.add(attrs.get('si', ''))
                if 'ref' in attrs:
                    if attrs.get('si', '') in masters:
                        raise refuse('duplicate stored shared masters')
                    masters.add(attrs.get('si', ''))
                if re.fullmatch(r'[0-9]{1,10}', attrs.get('si', '')) is None:
                    raise refuse('unrecognized stored shared identifiers')
            elif kind == 'array':
                allowed |= {'ref', 'aca'}
                if not attrs.get('ref'):
                    raise refuse('missing stored array bounds')
            elif kind != 'normal':
                raise refuse('unmeasured stored formula types')
            if (set(attrs) - allowed or any(attrs.get(flag, '0') not in ('0', '1', 'true', 'false')
                                            for flag in ('ca', 'aca'))):
                raise refuse('unmeasured formula attributes')
            if 'ref' in attrs:
                try:
                    block = parse_area(attrs['ref'])
                except ValueError:
                    raise refuse('unrecognized stored formula bounds') from None
                _area(block, '')
                if not block.contains(area.top, area.left):
                    raise refuse('stored formula anchors outside their bounds')
        elif formula_attributes is not None:
            raise refuse('child elements inside stored formulas')

    def characters(text: str) -> None:
        nonlocal has_body
        if formula_attributes is not None and text.strip():
            has_body = True

    def append(target: list[tuple[int, int]], span: tuple[int, int]) -> None:
        if len(actual_rows) + len(actual_cells) + len(actual) >= 350_000:
            raise refuse('stored cells beyond the metadata record budget')
        target.append(span)

    def finish(open_end: int) -> int:
        return (open_end if raw[open_end - 2:open_end] == b'/>'
                else raw.index(b'>', parser.CurrentByteIndex) + 1)

    def end(_name: str) -> None:
        nonlocal current_cell, formula_attributes, data_span
        if path[-1] == NS + '}f' and formula_attributes is not None:
            append(actual, (formula_start, finish(formula_open_end)))
            kind = formula_attributes.get('t', 'normal')
            if (((kind == 'array' or (kind == 'shared' and 'ref' in formula_attributes)) and not has_body)
                    or (kind == 'shared' and 'ref' not in formula_attributes and has_body)):
                raise refuse('unowned stored array or shared formula bodies')
            formula_attributes = None
        if path == [NS + '}worksheet', NS + '}sheetData']:
            data_span = (data_start, finish(data_open_end))
        if path == [NS + '}worksheet', NS + '}sheetData', NS + '}row']:
            append(actual_rows, (row_start, finish(row_open_end)))
        if path == [NS + '}worksheet', NS + '}sheetData', NS + '}row', NS + '}c']:
            append(actual_cells, (cell_start, finish(cell_open_end)))
        if len(path) == 4 and path[-1] == NS + '}c':
            current_cell = ''
        path.pop()

    def namespace(_prefix: str | None, uri: str | None) -> None:
        if uri is not None and len(uri) > 4096:
            raise refuse('oversized stored formula namespaces')

    parser.StartNamespaceDeclHandler = namespace
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = characters
    parser.Parse(raw, True)
    if shared_ids != masters:
        raise refuse('stored shared followers without masters')
    def opening_ownership(local: bytes, spans: list[tuple[int, int]], start: int, end: int) -> None:
        count = 0
        # Simple opening-name matches are linear; establish every candidate's
        # real ownership before running the legacy nested closing-tag regexes.
        for match in re.compile(b'<' + local + rb'\b').finditer(raw, start, end):
            if count >= len(spans) or match.start() != spans[count][0]:
                raise refuse('unowned lexical cell storage openings')
            count += 1
        if count != len(spans):
            raise refuse('unowned lexical cell storage openings')

    opening_ownership(b'sheetData', [] if data_span is None else [data_span], 0, len(raw))
    if data_span is not None:
        for local, spans in ((b'row', actual_rows), (b'c', actual_cells), (b'f', actual)):
            opening_ownership(local, spans, *data_span)
    from pyopenvba.apps.excel._io import formula_spans

    lexical = formula_spans(raw)
    if (data_span != lexical.data or actual_rows != lexical.rows
            or actual_cells != lexical.cells or actual != lexical.formulas):
        raise refuse('unowned lexical formula elements')
