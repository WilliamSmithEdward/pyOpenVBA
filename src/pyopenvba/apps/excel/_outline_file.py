"""Lossless, bounded outline edits to row/column openings and sheet metadata."""
from __future__ import annotations

import re
from bisect import bisect_left, bisect_right
from collections.abc import Mapping
from dataclasses import dataclass
from itertools import pairwise
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS
from pyopenvba._opc_xml import OpcPart, opc_part
from pyopenvba.apps.excel._advanced_filter import attribute_values, patch_attribute
from pyopenvba.apps.excel._dimensions import width_text
from pyopenvba.apps.excel._page_setup import configuration_tag_end, patch_configuration
from pyopenvba.exceptions import VBAUnsupportedError

if TYPE_CHECKING:
    from pyopenvba.apps.excel._dimensions import ColumnRecord, RowRecord
    from pyopenvba.apps.excel._model import Worksheet

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
MAX_BYTES = 16 * 1024 * 1024
KEYS = ('outlineLevel', 'collapsed', 'hidden')
_QNAME = re.compile(rb'<([^\s/>]+)')


@dataclass
class _Axes:
    rows: Mapping[int, RowRecord]
    columns: Mapping[int, ColumnRecord]
    format: Mapping[str, str]


def _unsupported() -> VBAUnsupportedError:
    return VBAUnsupportedError('Worksheet outlines on unmeasured or oversized row/column XML are not implemented')


def _index(value: str | None, limit: int) -> int:
    if value is None or not re.fullmatch(r'[0-9]{1,10}', value) or not 1 <= int(value) <= limit:
        raise _unsupported()
    return int(value)


def _qname(part: OpcPart, start: int) -> str:
    found = _QNAME.match(part.data, start)
    assert found is not None
    return found[1].decode()


def _parts(xml: str) -> tuple[OpcPart, OpcPart, OpcPart | None]:
    data = xml.encode()
    if len(data) > MAX_BYTES:
        raise _unsupported()
    try:
        root = opc_part(data, NS, 'worksheet', max_children=256)
        for name, _attrs, start, _end in root.children:
            local = name.rsplit('}', 1)[-1]
            if local in ('sheetData', 'cols', 'sheetFormatPr') and (name != NS + '}' + local or _qname(root, start) != local):
                raise _unsupported()
        rows = opc_part(data, NS, 'worksheet', container_path=('sheetData',), max_children=50_000)
        cols = opc_part(data, NS, 'worksheet', container_path=('cols',), max_children=MAX_COLUMNS)
    except ValueError:
        raise _unsupported() from None
    if rows is None:
        raise _unsupported()
    last = 0
    for name, attrs, start, _end in rows.children:
        if name != NS + '}row' or _qname(rows, start) != 'row':
            raise _unsupported()
        number = _index(attrs.get('r'), MAX_ROWS)
        if number <= last:
            raise _unsupported()
        last = number
        _validate(attrs)
    last = 0
    if cols is not None:
        for name, attrs, start, end in cols.children:
            if name != NS + '}col' or _qname(cols, start) != 'col' or not cols.data[start:end].rstrip().endswith(b'/>'):
                raise _unsupported()
            first, stop = _index(attrs.get('min'), MAX_COLUMNS), _index(attrs.get('max'), MAX_COLUMNS)
            if first <= last or stop < first:
                raise _unsupported()
            last = stop
            _validate(attrs)
    return root, rows, cols


def _validate(attrs: dict[str, str]) -> None:
    if attrs.get('outlineLevel', '0') not in ('0', '1', '2', '3', '4', '5', '6', '7'):
        raise _unsupported()
    if any(attrs.get(key, '0') not in ('0', '1', 'false', 'true') for key in ('hidden', 'collapsed')):
        raise _unsupported()


def _outline_values(attrs: Mapping[str, str]) -> tuple[str, bool]:
    return attrs.get('outlineLevel', '0'), attrs.get('collapsed', '0') in ('1', 'true')


def admit(sheet: Worksheet) -> set[int]:
    package = sheet.book.package
    if package is not None and sheet.part_name and package.has(sheet.part_name):
        _root, rows, cols = _parts(package.read(sheet.part_name).decode())
        actual_rows = {_index(attrs.get('r'), MAX_ROWS): attrs for _name, attrs, _start, _end in rows.children}
        actual_columns: dict[int, dict[str, str]] = {}
        if cols is not None:
            for _name, attrs, _start, _end in cols.children:
                for number in range(_index(attrs.get('min'), MAX_COLUMNS), _index(attrs.get('max'), MAX_COLUMNS) + 1):
                    actual_columns[number] = attrs
        # Legacy dimension loading is lexical. Its nondefault outline state must
        # originate in actual direct axis elements, never comments or CDATA.
        # Hidden is deliberately excluded: ordinary prior Hidden edits are valid.
        for records, actual in ((sheet.dims.rows, actual_rows), (sheet.dims.columns, actual_columns)):
            loaded = {number: _outline_values(record.extra) for number, record in records.items()
                      if _outline_values(record.extra) != ('0', False)}
            observed = {number: _outline_values(attrs) for number, attrs in actual.items()
                        if _outline_values(attrs) != ('0', False)}
            if loaded != observed:
                raise _unsupported()
        return set(actual_rows)
    return set()


def _opening(tag: bytes, values: dict[str, str | None]) -> bytes:
    stop = configuration_tag_end(tag)
    text = tag[:stop].decode()
    first = re.match(r'<[^\s/>]+', text)
    assert first is not None
    tail = len(text) - (2 if text.endswith('/>') else 1)
    fragment = text[first.end():tail]
    for key, value in values.items():
        old = attribute_values(fragment)
        if (value is None and key not in old) or old.get(key) == value:
            continue
        fragment = patch_attribute(fragment, key, value)
    return (text[:first.end()] + fragment + text[tail:]).encode() + tag[stop:]


def _apply(data: bytes, edits: list[tuple[int, int, bytes]]) -> bytes:
    if len(data) + sum(len(text) - (end - start) for start, end, text in edits) > MAX_BYTES:
        raise _unsupported()
    pieces: list[bytes] = []
    position = 0
    for start, end, replacement in sorted(edits, key=lambda e: (e[0], e[1])):
        if start < position:
            raise _unsupported()
        pieces.extend((data[position:start], replacement))
        position = end
    pieces.append(data[position:])
    return b''.join(pieces)


def _values(axes: _Axes, rows: bool, number: int) -> dict[str, str | None]:
    record = (axes.rows if rows else axes.columns).get(number)
    return {key: ('1' if record is not None and record.hidden else None) if key == 'hidden'
            else record.extra.get(key) if record is not None else None for key in KEYS}


def _row_changes(axes: _Axes, xml: str, numbers: set[int]) -> str:
    _root, rows, _cols = _parts(xml)
    found = {_index(attrs.get('r'), MAX_ROWS): (start, end) for _name, attrs, start, end in rows.children}
    edits: list[tuple[int, int, bytes]] = []
    inserted: dict[int, list[bytes]] = {}
    keys = sorted(found)
    for number in sorted(numbers):
        values = _values(axes, True, number)
        if number in found:
            start, end = found[number]
            original = rows.data[start:end]
            patched = _opening(original, values)
            if patched != original:
                edits.append((start, end, patched))
        elif any(value is not None for value in values.values()):
            where = bisect_right(keys, number)
            at = found[keys[where]][0] if where < len(keys) else rows.closing_start
            markup = _opening(rows.child('row', f'r="{number}"'), values)
            inserted.setdefault(at, []).append(markup)
    edits.extend((at, at, b''.join(markup)) for at, markup in inserted.items())
    if rows.data[rows.opening_end - 2:rows.opening_end] == b'/>':
        return rows.insert(rows.closing_start, b''.join(markup for _at, _end, markup in edits)).decode()
    return _apply(rows.data, edits).decode()


def _column_changes(axes: _Axes, xml: str, numbers: set[int]) -> str:
    root, _rows, cols = _parts(xml)
    if cols is None:
        if not any(any(v is not None for v in _values(axes, False, n).values()) for n in numbers):
            return xml
        at = next(start for name, _a, start, _end in root.children if name == NS + '}sheetData')
        xml = root.insert(at, root.child('cols', '')).decode()
        _root, _rows, cols = _parts(xml)
        assert cols is not None
    edits: list[tuple[int, int, bytes]] = []
    ranges: list[tuple[int, int, int]] = []
    covered: set[int] = set()
    ordered = sorted(numbers)
    delta = 0
    for _name, attrs, start, end in cols.children:
        lo, hi = _index(attrs.get('min'), MAX_COLUMNS), _index(attrs.get('max'), MAX_COLUMNS)
        ranges.append((lo, hi, start))
        reached = set(ordered[bisect_left(ordered, lo):bisect_right(ordered, hi)])
        covered.update(reached)
        if not reached:
            continue
        boundaries = sorted({lo, hi + 1, *(point for n in reached for point in (n, n + 1))})
        runs: list[tuple[int, int, dict[str, str | None] | None]] = []
        for left, stop in pairwise(boundaries):
            values = _values(axes, False, left) if left in reached else None
            if runs and runs[-1][1] == left - 1 and runs[-1][2] == values:
                runs[-1] = (runs[-1][0], stop - 1, values)
            else:
                runs.append((left, stop - 1, values))
        pieces: list[bytes] = []
        length = len(cols.data) - (end - start) + delta
        for left, right, values in runs:
            attributes = {'min': str(left), 'max': str(right), **(values or {})}
            piece = _opening(cols.data[start:end], attributes)
            length += len(piece)
            if length > MAX_BYTES:
                raise _unsupported()
            pieces.append(piece)
        replacement = b''.join(pieces)
        edits.append((start, end, replacement))
        delta += len(replacement) - (end - start)
    inserted: dict[int, list[bytes]] = {}
    starts = [lo for lo, _hi, _start in ranges]
    for number in sorted(numbers - covered):
        values = _values(axes, False, number)
        if not any(value is not None for value in values.values()):
            continue
        where = bisect_right(starts, number)
        at = ranges[where][2] if where < len(ranges) else cols.closing_start
        record = axes.columns.get(number)
        if record is None or record.width is None:
            raise _unsupported()
        attributes = (f'min="{number}" max="{number}" width="{record.width_as_read or width_text(record.width)}"'
                      + (' customWidth="1"' if record.custom else ''))
        markup = _opening(cols.child('col', attributes), values)
        inserted.setdefault(at, []).append(markup)
    edits.extend((at, at, b''.join(markup)) for at, markup in inserted.items())
    if cols.data[cols.opening_end - 2:cols.opening_end] == b'/>':
        return cols.insert(cols.closing_start, b''.join(markup for _at, _end, markup in edits)).decode()
    return _apply(cols.data, edits).decode()


def patch(sheet: Worksheet, xml: str, rows: set[int], columns: set[int], *,
          prospective_rows: Mapping[int, RowRecord] | None = None,
          prospective_columns: Mapping[int, ColumnRecord] | None = None,
          prospective_format: Mapping[str, str] | None = None) -> str:
    axes = _Axes(sheet.dims.rows if prospective_rows is None else prospective_rows,
                 sheet.dims.columns if prospective_columns is None else prospective_columns,
                 sheet.dims.format if prospective_format is None else prospective_format)
    if rows:
        xml = _row_changes(axes, xml, rows)
    if columns:
        xml = _column_changes(axes, xml, columns)
    root, _rows, _cols = _parts(xml)
    if not any(name == NS + '}sheetFormatPr' for name, _a, _start, _end in root.children):
        at = next(start for name, _a, start, _end in root.children if name in (NS + '}cols', NS + '}sheetData'))
        xml = root.insert(at, root.child('sheetFormatPr', 'defaultRowHeight="15"')).decode()
    values = {key: axes.format.get(key) for key in ('outlineLevelRow', 'outlineLevelCol')}
    return patch_configuration(xml, 'worksheet', ('sheetFormatPr',), values)
