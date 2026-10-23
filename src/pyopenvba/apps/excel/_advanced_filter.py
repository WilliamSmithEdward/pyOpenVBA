"""Measured rectangular AdvancedFilter criteria and extracted values.

Formula criteria, shared filters and feature-bearing output mappings are refused
before mutation. This is not a general worksheet Copy implementation.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from functools import lru_cache
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_ROWS, Area, quote_sheet
from pyopenvba._xml import escape, unescape
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import EMPTY, MISSING, VBADate, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from collections.abc import Callable

    from pyopenvba.apps.excel._model import Range, Worksheet


def _unsupported(detail: str) -> VBAUnsupportedError:
    return VBAUnsupportedError('AdvancedFilter ' + detail + ' is not implemented')


_TEXT_WORK_LIMIT = 4_000_000


@dataclass
class _Work:
    remaining: int = _TEXT_WORK_LIMIT

    def text(self, value: str) -> None:
        if len(value) > self.remaining:
            raise _unsupported('requiring more than four million text characters of comparison work')
        self.remaining -= len(value)


def _value(sheet: Worksheet, row: int, column: int) -> object:
    cell = sheet.cells_.get((row, column))
    if cell is None:
        return EMPTY
    if cell.formula:
        return sheet.book.calculator.value_of(sheet.name, row, column)
    return cell.value


def _plain(text: str) -> None:
    if not text.isascii():
        raise _unsupported('with non-ASCII comparison text')


def _compare(left: float | str, operator: str, right: float | str) -> bool:
    if isinstance(left, str):
        if not isinstance(right, str):
            return False
        order = (left > right) - (left < right)
    else:
        if isinstance(right, str):
            return False
        order = (left > right) - (left < right)
    return {'>': order > 0, '>=': order >= 0, '<': order < 0, '<=': order <= 0}[operator]


def _number(value: object) -> float | None:
    if isinstance(value, VBADate):
        return value.serial
    return float(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else None


@lru_cache(maxsize=64)
def _pattern(pattern: str) -> tuple[int, int, dict[str, int], int]:
    """A bounded bitset NFA; wildcard input never becomes a regular expression."""
    if len(pattern) > 255:
        raise _unsupported('with criteria longer than 255 characters')
    stars, questions, count = 0, 0, 0
    literals: dict[str, int] = {}
    index = 0
    while index < len(pattern):
        char = pattern[index]
        literal = char == '~' and index + 1 < len(pattern) and pattern[index + 1] in '*?~'
        if literal:
            index += 1
            char = pattern[index]
        if not literal and char == '*':
            if count and stars & (1 << (count - 1)):
                index += 1
                continue
            stars |= 1 << count
        elif not literal and char == '?':
            questions |= 1 << count
        else:
            literals[char.casefold()] = literals.get(char.casefold(), 0) | (1 << count)
        count += 1
        index += 1
    return stars, questions, literals, 1 << count


def _matches(text: str, pattern: str, *, is_text: bool) -> bool:
    stars, questions, literals, final = _pattern(pattern)
    if (stars or questions) and not is_text:
        return False
    states = 1
    states |= (states & stars) << 1
    for char in text.casefold():
        states = ((states & (questions | literals.get(char, 0))) << 1) | (states & stars)
        states |= (states & stars) << 1
        if not states:
            return False
    return bool(states & final)


@dataclass(frozen=True)
class Criterion:
    column: int
    operator: str
    operand: str
    number: float | None
    prefix: bool = False
    error_only: bool = False

    def matches(self, value: object, work: _Work) -> bool:
        if isinstance(value, str):
            work.text(value)
        if self.error_only:
            return isinstance(value, ExcelError) and value.name == self.operand
        number = _number(value)
        if isinstance(value, str) and self.number is not None and self.operator == '=':
            from pyopenvba.apps.excel._typing import typed_text

            number = _number(typed_text(value).value)
        if self.number is not None:
            if self.operator == '=':
                return number is not None and number == self.number
            if self.operator == '<>':
                return number is None or number != self.number
            return number is not None and _compare(number, self.operator, self.number)
        if value is EMPTY:
            text = ''
        elif isinstance(value, ExcelError):
            text = value.name
        elif isinstance(value, bool):
            text = 'TRUE' if value else 'FALSE'
        else:
            text = to_text(value)
        _plain(text)
        pattern = self.operand + ('*' if self.prefix else '')
        if self.operator in ('=', '<>'):
            found = _matches(text, pattern, is_text=isinstance(value, str))
            return found if self.operator == '=' else not found
        return isinstance(value, str) and _compare(text.casefold(), self.operator, self.operand.casefold())


def _criterion(column: int, value: object, work: _Work) -> Criterion:
    if isinstance(value, str):
        work.text(value)
    if isinstance(value, bool):
        return Criterion(column, '=', 'TRUE' if value else 'FALSE', None)
    if isinstance(value, ExcelError):
        return Criterion(column, '=', value.name, None, error_only=True)
    numeric = _number(value)
    if numeric is not None:
        return Criterion(column, '=', '', numeric)
    text = to_text(value)
    _plain(text)
    if len(text) > 254:
        raise _unsupported('with criteria longer than 254 characters')
    operator = next((operator for operator in ('<>', '>=', '<=', '=', '>', '<') if text.startswith(operator)), '')
    operand = text[len(operator):]
    from pyopenvba.apps.excel._typing import typed_text

    typed = typed_text(operand)
    numeric = _number(typed.value)
    if operator not in ('', '=', '<>') and numeric is None:
        raise _unsupported('with text-order criteria')
    return Criterion(column, operator or '=', operand, numeric,
                     prefix=not operator and numeric is None and not re.search(r'(?<!~)[*?]', operand))


def _range(value: object, sheet: Worksheet, label: str) -> Range:
    from pyopenvba.apps.excel._model import Range

    if not isinstance(value, Range):
        raise error(1004)
    if value.sheet is not sheet or len(value.areas) != 1:
        raise _unsupported('with a nonlocal or multiple-area ' + label)
    if value.first.whole_rows or value.first.whole_columns:
        raise _unsupported('with whole-row or whole-column ' + label)
    return value


def _headers(sheet: Worksheet, area: Area, work: _Work) -> dict[str, int]:
    result: list[str] = []
    for column in range(area.left, area.right + 1):
        value = _value(sheet, area.top, column)
        if not isinstance(value, str) or not value:
            raise _unsupported('with blank or nontext list headers')
        work.text(value)
        _plain(value)
        result.append(value.casefold())
    if len(set(result)) != len(result):
        raise _unsupported('with duplicate source headers')
    return {name: area.left + index for index, name in enumerate(result)}


def _criteria(sheet: Worksheet, area: Area, headers: dict[str, int], source: Area, work: _Work) -> list[list[Criterion]]:
    if area.rows * area.columns > 4096:
        raise _unsupported('with more than 4096 criteria cells')
    columns: list[int] = []
    for column in range(area.left, area.right + 1):
        header = _value(sheet, area.top, column)
        if isinstance(header, str):
            work.text(header)
        if header is EMPTY and all(_value(sheet, row, column) is EMPTY for row in range(area.top + 1, area.bottom + 1)):
            columns.append(source.left)
            continue
        if not isinstance(header, str) or header.casefold() not in headers:
            raise _unsupported('with formula criteria or unmatched criteria headers')
        columns.append(headers[header.casefold()])
    return [[_criterion(columns[column - area.left], value, work)
             for column in range(area.left, area.right + 1)
             if (value := _value(sheet, row, column)) is not EMPTY and value != '']
            for row in range(area.top + 1, area.bottom + 1)]


def _key(value: object, work: _Work) -> tuple[object, ...]:
    if value is EMPTY:
        return ('blank',)
    if isinstance(value, str):
        work.text(value)
        _plain(value)
        return 'text', value.casefold()
    if isinstance(value, ExcelError):
        return 'error', value.name
    if isinstance(value, bool):
        return 'bool', value
    number = _number(value)
    return ('number', number) if number is not None else ('other', to_text(value))


def attribute_spans(fragment: str) -> list[tuple[str, int, int, str]]:
    """Walk complete quoted attributes, never names embedded in their values."""
    spans: list[tuple[str, int, int, str]] = []
    index = 0
    while index < len(fragment):
        while index < len(fragment) and fragment[index] in ' \t\r\n':
            index += 1
        if index == len(fragment):
            break
        start = index
        while index < len(fragment) and fragment[index] not in ' \t\r\n=':
            if fragment[index] in '<>\"\'':
                raise _unsupported('with malformed attribute metadata')
            index += 1
        name = fragment[start:index]
        while index < len(fragment) and fragment[index] in ' \t\r\n':
            index += 1
        if not name or index == len(fragment) or fragment[index] != '=':
            raise _unsupported('with malformed attribute metadata')
        index += 1
        while index < len(fragment) and fragment[index] in ' \t\r\n':
            index += 1
        if index == len(fragment) or fragment[index] not in '\"\'':
            raise _unsupported('with unquoted attribute metadata')
        quote = fragment[index]
        value_start = index + 1
        end = fragment.find(quote, value_start)
        if end < 0:
            raise _unsupported('with unterminated attribute metadata')
        index = end + 1
        if index < len(fragment) and fragment[index] not in ' \t\r\n':
            raise _unsupported('with malformed attribute metadata')
        spans.append((name, start, index, fragment[value_start:end]))
    return spans


def attribute_values(fragment: str) -> dict[str, str]:
    return {name: unescape(value) for name, _start, _end, value in attribute_spans(fragment)}


def patch_attribute(fragment: str, owned: str, value: str | None) -> str:
    spans = attribute_spans(fragment)
    for name, start, end, _value in reversed(spans):
        if name == owned:
            fragment = fragment[:start] + fragment[end:]
    return fragment + (' ' + owned + '=\"' + escape(value) + '\"' if value is not None else '')


def _prepare_name(sheet: Worksheet, name: str, area: Area, *, visible: bool = True) -> Callable[[], None]:
    from pyopenvba.apps.excel._model import NameEntry

    local = quote_sheet(sheet.name) + '!' + name
    refers = '=' + quote_sheet(sheet.name) + '!' + area.address()
    names = sheet.book.names_
    entry = next((entry for entry in names.entries if entry.name.casefold() == local.casefold()), None)
    attributes = patch_attribute(entry.attributes if entry is not None else '', 'name', '_xlnm.' + name)

    def commit() -> None:
        if entry is None:
            names.entries.append(NameEntry(local, refers, sheet.book, visible=visible, attributes=attributes))
        else:
            entry.refers_to, entry.visible, entry.attributes = refers, visible, attributes
        names.changed = True

    return commit


def _admit(sheet: Worksheet, area: Area) -> None:
    from pyopenvba.apps.excel._conditional import meets
    from pyopenvba.apps.excel._conditional_transfer import guard_extensions
    from pyopenvba.apps.excel._hyperlinks import touches

    guard_extensions(sheet)

    if (any(area.contains(*key) and cell.dynamic for key, cell in sheet.cells_.items())
            or any(meets(area, box) for box in [*sheet.merged_areas, *sheet.array_formulas.values(),
                                        *(spill.area for spill in sheet.spills.values()),
                                        *(table.area for table in sheet.tables),
                                        *(table.area for table in sheet.data_tables.values())])
            or any(area.contains(*key) for key in sheet.notes) or touches(sheet, [area])
            or any(meets(area, part) for rule in sheet.validations for part in rule.areas)
            or any(meets(area, part) for rule in sheet.conditional_formats for part in rule.areas)):
        raise _unsupported('over cells with arrays, merges, tables, annotations or rules')


def advanced_filter(target: Range, action: object, criteria_argument: object, copy_argument: object,
                    unique_argument: object) -> object:
    from pyopenvba.apps.excel._conditional import meets
    from pyopenvba.apps.excel._model import Cell
    from pyopenvba.apps.excel._protection import enforced

    sheet = target.sheet
    if sheet.book.application.udf_callers:
        return False
    kind = int(to_integer(action, 'Long'))
    if kind not in (1, 2):
        raise error(1004)
    if len(target.areas) != 1:
        raise _unsupported('on multiple source areas')
    source = target.first
    if source.rows < 2:
        raise error(1004)
    if source.rows * source.columns > 1_000_000:
        raise _unsupported('over more than one million list cells')
    if source.whole_rows or source.whole_columns:
        raise _unsupported('over whole rows or columns')
    if sheet.auto_filter is not None:
        raise _unsupported('combined with an AutoFilter')
    criteria_range = None if criteria_argument is MISSING else _range(criteria_argument, sheet, 'criteria range')
    output = None if kind == 1 else _range(copy_argument, sheet, 'extract range')
    if criteria_range is not None and criteria_range.first.rows * criteria_range.first.columns > 4096:
        raise _unsupported('with more than 4096 criteria cells')
    projection_width = output.first.columns if output is not None else source.columns
    if source.rows * projection_width > 1_000_000:
        raise _unsupported('staging more than one million cells for extraction or uniqueness')
    if criteria_range is not None and meets(source, criteria_range.first):
        raise _unsupported('with overlapping source and criteria cells')
    if kind == 2 and enforced(sheet) is not None:
        raise _unsupported('extracting on a protected sheet (native partial writes)')
    _admit(sheet, source)
    work = _Work()
    headers = _headers(sheet, source, work)
    clauses = _criteria(sheet, criteria_range.first, headers, source, work) if criteria_range is not None else []
    if source.rows * max(1, sum(map(len, clauses))) > 4_000_000:
        raise _unsupported('requiring more than four million row/criterion comparisons')
    unique = unique_argument is not MISSING and to_bool(unique_argument)
    projection: list[int] = list(range(source.left, source.right + 1))
    extract = output.first if output is not None else None
    if output is not None:
        given = [_value(sheet, output.first.top, column) for column in range(output.first.left, output.first.right + 1)]
        blank = [value is EMPTY or value == '' for value in given]
        if len(given) == 1 and output.first.rows == 1 and blank[0]:
            # A single blank cell takes every column under its headers, and Excel's Extract name spans that header
            # row (tests/fixtures/excel_everyday_2.json, advanced_filter_blank_target).
            first = output.first
            extract = Area(first.top, first.left, first.top, first.left + len(projection) - 1, first.sheet)
        elif all(blank):
            # A row of blank headers names no column (advanced_filter_blank_row).
            raise error(1004)
        else:
            projection = []
            for value in given:
                if isinstance(value, str):
                    work.text(value)
                if not isinstance(value, str) or value.casefold() not in headers:
                    raise _unsupported('with unmatched extract headers (native name side effects)')
                projection.append(headers[value.casefold()])
    rows: list[int] = []
    seen: set[tuple[tuple[object, ...], ...]] = set()
    for row in range(source.top + 1, source.bottom + 1):
        if clauses and not any(all(criterion.matches(_value(sheet, row, criterion.column), work) for criterion in clause)
                               for clause in clauses):
            continue
        key = tuple(_key(_value(sheet, row, column), work) for column in projection) if unique else ()
        if not unique or key not in seen:
            rows.append(row)
            seen.add(key)
    replacements: dict[tuple[int, int], Cell] = {}
    footprint = None
    if output is not None:
        if (len(rows) + 1) * len(projection) > 1_000_000:
            raise _unsupported('extracting more than one million cells')
        first = output.first
        right = max(first.right, first.left + len(projection) - 1)
        last = max((row for row, column in sheet.cells_ if first.left <= column <= right
                    and row >= first.top), default=first.top)
        last = max(last, first.top + len(rows))
        if last > MAX_ROWS:
            raise error(1004)
        footprint = Area(first.top, first.left, last, right)
        if meets(source, footprint) or (criteria_range is not None and meets(criteria_range.first, footprint)):
            raise _unsupported('with overlapping source, criteria and extract cells')
        _admit(sheet, footprint)
        for offset, source_row in enumerate([source.top, *rows]):
            for across, column in enumerate(projection):
                replacements[(first.top + offset, first.left + across)] = Cell(
                    value=_value(sheet, source_row, column), style=sheet.style_at(source_row, column))
    # Every criterion, style/value read and admission check completed before mutation.
    name_updates = [_prepare_name(sheet, '_FilterDatabase', source, visible=False)]
    if criteria_range is not None:
        name_updates.append(_prepare_name(sheet, 'Criteria', criteria_range.first))
    if extract is not None:
        name_updates.append(_prepare_name(sheet, 'Extract', extract))
    for update in name_updates:
        update()
    if output is not None:
        assert footprint is not None
        for key in list(sheet.cells_):
            if footprint.contains(*key):
                old = sheet.cells_[key]
                sheet.cells_[key] = Cell(style=old.style)
                sheet.cell_changed(*key)
                sheet.book.calculator.entered(sheet.name, *key)
        for key, cell in replacements.items():
            sheet.cells_[key] = cell
            sheet.cell_changed(*key)
            sheet.book.calculator.entered(sheet.name, *key)
    else:
        selected = set(rows)
        for row in range(source.top + 1, source.bottom + 1):
            hidden = row not in selected
            if sheet.dims.row_hidden(row) != hidden:
                sheet.dims.hide_row(row, hidden)
        sheet.advanced_filter = source if any(clauses) or unique else None
        sheet.filter_changed = True
        sheet.advanced_filter_changed = True
    sheet.touched()
    if output is None:
        from pyopenvba.apps.excel._events import visibility_changed

        visibility_changed(sheet.book)
    return True


def read_state(sheet: Worksheet, xml: bytes) -> None:
    """Recover persisted in-place filter state from its flag and local list name."""
    from pyopenvba._a1 import parse_reference, split_sheet
    from pyopenvba._opc_xml import opc_part

    namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    part = opc_part(xml, namespace, 'worksheet')
    properties = next((attrs for name, attrs, _start, _end in part.children
                       if name == namespace + '}sheetPr'), None)
    if sheet.auto_filter is not None or properties is None or properties.get('filterMode') not in ('1', 'true'):
        return
    local = (quote_sheet(sheet.name) + '!_FilterDatabase').casefold()
    entry = next((entry for entry in sheet.book.names_.entries if entry.name.casefold() == local), None)
    if entry is None:
        raise _unsupported('loading filter mode without its list name')
    owner, reference = split_sheet(entry.refers_to.lstrip('='))
    if owner and owner.casefold() != sheet.name.casefold():
        raise _unsupported('loading an advanced filter whose list name targets another sheet')
    try:
        areas = parse_reference(reference)
    except ValueError:
        raise _unsupported('loading a nonrectangular advanced-filter list name') from None
    if len(areas) != 1:
        raise _unsupported('loading a multiple-area advanced-filter list name')
    sheet.advanced_filter = areas[0]


def with_filter_mode(xml: str, filtered: bool) -> str:
    """Patch a direct sheetPr flag without dropping namespace scopes or other bytes."""
    from pyopenvba._opc_xml import opc_part

    namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    part = opc_part(xml.encode('utf-8'), namespace, 'worksheet')
    found = next(((start, end) for name, _attrs, start, end in part.children
                  if name == namespace + '}sheetPr'), None)
    if found is not None:
        start, _end = found
        quote = 0
        opening_end = start
        for at in range(start, len(part.data)):
            char = part.data[at]
            if quote:
                if char == quote:
                    quote = 0
            elif char in (34, 39):
                quote = char
            elif char == 62:
                opening_end = at + 1
                break
        tag = part.data[start:opening_end].decode('utf-8')
        attributes_end = len(tag) - (2 if tag.endswith('/>') else 1)
        attributes_start = next((at for at, char in enumerate(tag) if char in ' \t\r\n'), attributes_end)
        fragment = patch_attribute(tag[attributes_start:attributes_end], 'filterMode', '1' if filtered else None)
        tag = (tag[:attributes_start] + fragment + tag[attributes_end:]).encode('utf-8')
        return (part.data[:start] + tag + part.data[opening_end:]).decode('utf-8')
    if not filtered:
        return xml
    return part.insert(part.opening_end, part.child('sheetPr', 'filterMode="1"')).decode('utf-8')
