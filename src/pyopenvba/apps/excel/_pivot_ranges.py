"""Pivot field ranges selected from persisted report-line metadata."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba.apps.excel._model import Range
from pyopenvba.apps.excel._pivots import NS
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivots import PivotField, PivotTable
    from pyopenvba.apps.excel._pivot_reports import PivotLine

VALUES = -2


def page_range(table: PivotTable) -> Range:
    pages = table.root.findall(f'{NS}pageFields/{NS}pageField')
    location = table.root.find(NS + 'location')
    if not pages or location is None:
        raise error(1004)
    area = parse_area(location.get('ref', 'A1'))
    count = int(location.get('rowPageCount', '0'))
    columns = int(location.get('colPageCount', '0'))
    if count < 1 or columns < 1 or area.top - count - 1 < 1:
        raise error(1004)
    return Range(table.sheet, [Area(area.top - count - 1, area.left,
                                     area.top - 2, area.left + 3 * columns - 2, table.sheet.name)])


def _length(lines: tuple[PivotLine, ...]) -> int:
    return next((index for index, line in enumerate(lines) if line.kind == 'grand'), len(lines))


def _measure(lines: tuple[PivotLine, ...], axis: tuple[int, ...], wanted: int) -> list[int]:
    position = axis.index(VALUES)
    return [index for index, line in enumerate(lines)
            if line.kind not in ('grand', 'blank') and position < len(line.items)
            and line.items[position] == wanted]


def field_range(field: PivotField, label: bool) -> Range:
    field.check_alive()
    table = field.table
    root = table.root
    rows = tuple(int(entry.get('x', '0')) for entry in root.findall(f'{NS}rowFields/{NS}field'))
    columns = tuple(int(entry.get('x', '0')) for entry in root.findall(f'{NS}colFields/{NS}field'))
    pages = root.findall(f'{NS}pageFields/{NS}pageField')
    real_columns = any(index != VALUES for index in columns)
    location = root.find(NS + 'location')
    if location is None:
        raise error(1004)
    area = parse_area(location.get('ref', 'A1'))
    top, left = area.top, area.left
    dr, dc = int(location.get('firstDataRow', '0')), int(location.get('firstDataCol', '0'))

    def block(r1: int, c1: int, r2: int, c2: int) -> Range:
        if r1 < 1 or c1 < 1 or r2 < r1 or c2 < c1:
            raise error(1004)
        return Range(table.sheet, [Area(r1, c1, r2, c2, table.sheet.name)])

    if not rows and not pages and not real_columns:
        if field.index == VALUES:
            if int(field.Orientation()) == 0:
                raise error(1004)
            row, column = top - 1 if label else top + dr, left
        elif field.data:
            count = len(root.findall(f'{NS}dataFields/{NS}dataField'))
            row = top if label and count > 1 else top + dr
            column = left + int(field.Position()) - 1
        else:
            raise error(1004)
        return block(row, column, row, column)

    if not field.data:
        for position, page in enumerate(pages):
            if int(page.get('fld', '-1')) == field.index:
                page_area = page_range(table).first
                if root.get('pageOverThenDown', '0') in ('1', 'true'):
                    page_columns = int(location.get('colPageCount', '1'))
                    row_offset, column_offset = divmod(position, page_columns)
                else:
                    column_offset, row_offset = divmod(position, page_area.rows)
                row = page_area.top + row_offset
                column = left + 3 * column_offset + int(not label)
                return block(row, column, row, column)
        if field.index not in rows and field.index not in columns:
            raise error(1004)
    report = table.report()
    if report is None or not report.readable:
        raise VBAUnsupportedError('This PivotTable field-range layout is not readable')
    row_count, column_count = _length(report.row_lines), _length(report.column_lines)
    if not field.data and field.index in rows:
        column = left
        definitions = root.findall(f'{NS}pivotFields/{NS}pivotField')
        for index in rows[:rows.index(field.index)]:
            compact = root.get('compactData', '1') if index == VALUES else definitions[index].get('compact', root.get('compact', '1'))
            column += int(compact in ('0', 'false'))
        return block(top + dr - 1 if label else top + dr, column,
                     top + dr - 1 if label else top + dr + row_count - 1, column)
    if not field.data:
        position = columns.index(field.index)
        if field.index == VALUES and not real_columns:
            width = len(report.column_lines) if len(rows) == 1 else 1
            return block(top - 1 if label else top + dr, left + dc,
                         top - 1 if label else top + dr, left + dc if label else left + dc + width - 1)
        # Native multi-value reports with exactly one row field include total
        # columns here. Zero or multiple row fields exclude the grand columns.
        width = len(report.column_lines) if len(rows) == 1 and VALUES in columns else column_count
        return block(top if label else top + position + 1, left + dc + position if label else left + dc,
                     top if label else top + position + 1, left + dc + position if label else left + dc + width - 1)
    measure = int(field.Position()) - 1
    if VALUES in rows:
        indices = _measure(report.row_lines, rows, measure)
        column = left if label else left + dc
        areas = [Area(top + dr + index, column, top + dr + index,
                      column if label else left + dc + column_count - 1, table.sheet.name) for index in indices]
    elif VALUES in columns:
        indices = _measure(report.column_lines, columns, measure)
        row = top + dr - 1
        areas = [Area(row if label else row + 1, left + dc + index,
                      row if label else top + dr + row_count - 1, left + dc + index, table.sheet.name) for index in indices]
    elif real_columns:
        if label:
            return block(top if rows else top + dr, left, top if rows else top + dr, left)
        return block(top + dr, left + dc, top + dr + row_count - 1, left + dc + column_count - 1)
    else:
        row = top + dr - 1 if rows else top + dr
        return block(row if label else top + dr, left + dc,
                     row if label else top + dr + row_count - 1, left + dc)
    if not areas:
        raise error(1004)
    return Range(table.sheet, areas)
