"""Pivot field ranges selected from persisted report-line metadata."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba.apps.excel._model import Range
from pyopenvba.apps.excel._pivots import NS
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivots import PivotField, PivotItem, PivotTable
    from pyopenvba.apps.excel._pivot_reports import PivotLine, PivotReport

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


def data_label_range(table: PivotTable) -> Range:
    """DataLabelRange: the corner cell over a lone data field where the table has column fields, the data fields'
    captions where the Values field is on the columns, and the Values field's column where it is on the rows;
    error 1004 otherwise (tests/fixtures/pivot_layouts/)."""
    from pyopenvba.apps.excel._pivot_forms import label_columns

    root = table.entry.root
    rows = [int(entry.get('x', '0')) for entry in root.findall(f'{NS}rowFields/{NS}field')]
    columns = [int(entry.get('x', '0')) for entry in root.findall(f'{NS}colFields/{NS}field')]
    location = root.find(NS + 'location')
    if location is None or not root.findall(f'{NS}dataFields/{NS}dataField'):
        raise error(1004)
    area = parse_area(location.get('ref', 'A1'))
    dr, dc = int(location.get('firstDataRow', '0')), int(location.get('firstDataCol', '0'))
    real_columns = any(index != VALUES for index in columns)
    if VALUES in rows:
        column = area.left + label_columns(root, rows)[0][rows.index(VALUES)]
        found = Area(area.top + dr - 1, column, area.bottom, column, table.sheet.name)
    elif VALUES in columns:
        row = area.top + 1 + columns.index(VALUES) if real_columns else area.top + dr - 1
        found = Area(row, area.left + dc, row, area.right, table.sheet.name)
    elif real_columns:
        found = Area(area.top, area.left, area.top, area.left, table.sheet.name)
    else:
        raise error(1004)
    return Range(table.sheet, [found])


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
        from pyopenvba.apps.excel._pivot_forms import label_columns

        column = left + label_columns(root, list(rows))[0][rows.index(field.index)]
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


def _readable(table: PivotTable) -> PivotReport:
    report = table.report()
    if report is None or not report.readable:
        raise VBAUnsupportedError('This PivotTable item-range layout is not readable')
    return report


def _place(report: PivotReport, field: int, x: int) -> int | None:
    """Where in its field's list an item stands, by its place in the cache."""
    return next((position for position, item in enumerate(report.fields[field].items)
                 if item is not None and item.index == x), None)


def _runs(indices: list[int]) -> list[tuple[int, int]]:
    """Consecutive line numbers, as the first and last of each run."""
    runs: list[tuple[int, int]] = []
    for index in indices:
        if runs and runs[-1][1] == index - 1:
            runs[-1] = (runs[-1][0], index)
        else:
            runs.append((index, index))
    return runs


def item_range(item: PivotItem, label: bool) -> Range:
    """PivotItem.LabelRange and DataRange (tests/fixtures/pivot_grouping/ ranges_*): every line an item stands at,
    its own lines alone for a row field in compact form, whose column the fields inside it share, else each run
    of the lines inside it, their subtotals left out; the label cells in its field's column or row, the value
    cells across every line of the other axis but its grand total. A page field's item has no label cells and
    every value cell but the grand totals' for its data; a hidden item, or one no line shows, error 1004."""
    from pyopenvba.apps.excel._pivot_forms import label_columns, layout_of

    field = item.field
    field.check_alive()
    table = field.table.entry
    root = table.root
    if field.data or field.index < 0 or item.x < 0:
        raise error(1004)
    report = _readable(table)
    location = root.find(NS + 'location')
    if location is None:
        raise error(1004)
    area = parse_area(location.get('ref', 'A1'))
    top, left = area.top, area.left
    dr, dc = int(location.get('firstDataRow', '0')), int(location.get('firstDataCol', '0'))
    rows, columns = list(report.rows), list(report.columns)
    row_count = _length(report.row_lines) if rows else 1
    column_count = _length(report.column_lines)
    position = _place(report, field.index, item.x)
    areas: list[Area] = []
    if field.index in rows and position is not None:
        depth = rows.index(field.index)
        own = depth < len(rows) - 1 and layout_of(root, field.index).mode == 'compact'
        found = [index for index, line in enumerate(report.row_lines)
                 if line.kind == 'data' and len(line.items) > depth and line.items[depth] == position
                 and (not own or len(line.items) == depth + 1)]
        column = left + label_columns(root, rows)[0][depth]
        for first, last in _runs(found):
            areas.append(Area(top + dr + first, column if label else left + dc, top + dr + last,
                              column if label else left + dc + column_count - 1, table.sheet.name))
    elif field.index in columns and position is not None:
        depth = columns.index(field.index)
        found = [index for index, line in enumerate(report.column_lines)
                 if line.kind == 'data' and len(line.items) > depth and line.items[depth] == position]
        for first, last in _runs(found):
            areas.append(Area(top + 1 + depth if label else top + dr, left + dc + first,
                              top + 1 + depth if label else top + dr + row_count - 1, left + dc + last,
                              table.sheet.name))
    elif not label and any(int(page.get('fld', '-1')) == field.index
                           for page in root.findall(f'{NS}pageFields/{NS}pageField')):
        areas.append(Area(top + dr, left + dc, top + dr + row_count - 1, left + dc + column_count - 1,
                          table.sheet.name))
    if not areas:
        raise error(1004)
    return Range(table.sheet, areas)


def cell_item(table: PivotTable, row: int, column: int) -> tuple[int, int] | None:
    """The field and the item, by its place in the cache, whose label a cell of a table's rows or columns shows,
    the cells an item's LabelRange covers; None for any other cell, a total's or a data cell."""
    from pyopenvba.apps.excel._pivot_forms import label_columns

    report = _readable(table)
    root = table.root
    location = root.find(NS + 'location')
    if location is None:
        return None
    area = parse_area(location.get('ref', 'A1'))
    top, left = area.top, area.left
    dr, dc = int(location.get('firstDataRow', '0')), int(location.get('firstDataCol', '0'))
    rows, columns = list(report.rows), list(report.columns)
    line_at = row - top - dr
    if rows and 0 <= line_at < len(report.row_lines) and left <= column < left + dc:
        line = report.row_lines[line_at]
        places = label_columns(root, rows)[0]
        depths = [depth for depth in range(min(len(rows), len(line.items))) if places[depth] == column - left]
        if line.kind != 'data' or not depths:
            return None
        depth = max(depths)
        chosen = rows[depth]
    else:
        line_at, depth = column - left - dc, row - top - 1
        if not (columns and 0 <= line_at < len(report.column_lines) and 0 <= depth < len(columns) and row < top + dr):
            return None
        line = report.column_lines[line_at]
        if line.kind != 'data' or depth >= len(line.items):
            return None
        chosen = columns[depth]
    if chosen == VALUES:
        return None
    position = line.items[depth]
    items = report.fields[chosen].items
    entry = items[position] if 0 <= position < len(items) else None
    return (chosen, entry.index) if entry is not None else None
