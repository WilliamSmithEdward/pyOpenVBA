"""SeriesCollection.Add and Extend, measured in live Excel (scripts/measure_series_add.py,
tests/fixtures/series_add.json).

Add on a chart with series takes the source a column at a time, or a row at a time for xlRows, each line a new series.
Its first cell names the series only if SeriesLabels says so, and its first line gives the categories only if
CategoryLabels does; otherwise the new series take the chart's first series' categories. On an empty chart it works out
what it is not told as SetSourceData does.

Extend adds each line of the source to the series in its place: a line that carries on below the series' last area joins
it, and any other is a further area. A first line that starts with text is labels, not values. Cells on another sheet
change nothing."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_reference
from pyopenvba.apps.excel._charts import Chart, SeriesState, default_marker, label_block, reference
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, MISSING, error, to_bool, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet


def _source(chart: Chart, source: object) -> Range:
    """What Source names: a range, or the text of one in A1 or R1C1."""
    from pyopenvba.apps.excel._convert_formula import convert_formula
    from pyopenvba.apps.excel._model import Range

    if source is MISSING:
        raise error(450, 'Wrong number of arguments or invalid property assignment')
    if isinstance(source, Range):
        return source
    if isinstance(source, str):
        text = source if source.startswith('=') else f'={source}'
        if 'R' in text and 'C' in text and '$' not in text:
            text = str(convert_formula(chart.sheet.book.application, text, -4150, 1, MISSING, MISSING))
        try:
            area = parse_reference(text[1:])[0]
            owner = chart.sheet.book.sheet_named(area.sheet) if area.sheet else chart.sheet
        except (ValueError, KeyError):
            raise error(5, 'Invalid procedure call or argument') from None
        return Range(owner, [Area(area.top, area.left, area.bottom, area.right)])
    raise error(5, 'Invalid procedure call or argument')


def _clipped(sheet: Worksheet, area: Area) -> Area:
    """A whole row or column, as far as the sheet is used."""
    used = sheet.used_bounds()
    if used is None:
        return area
    top, left, bottom, right = used
    return Area(area.top if area.bottom < 1048576 else top, area.left if area.right < 16384 else left,
                min(area.bottom, bottom) if area.bottom == 1048576 else area.bottom,
                min(area.right, right) if area.right == 16384 else area.right)


def _lines(area: Area, by_rows: bool) -> list[Area]:
    if by_rows:
        return [Area(row, area.left, row, area.right) for row in range(area.top, area.bottom + 1)]
    return [Area(area.top, column, area.bottom, column) for column in range(area.left, area.right + 1)]


def _without_first(line: Area, by_rows: bool) -> Area:
    if by_rows:
        return Area(line.top, line.left + 1, line.bottom, line.right)
    return Area(line.top + 1, line.left, line.bottom, line.right)


def _first_of(line: Area) -> Area:
    return Area(line.top, line.left, line.top, line.left)


def _made(chart: Chart, entries: list[SeriesState]) -> object:
    from pyopenvba.apps.excel._charts import Series

    state = chart.state
    chart.changed()
    identity = max((one.series_id for one in state.series), default=-1)
    for entry in entries:
        identity += 1
        entry.series_id = identity
        entry.unique, state.next_unique = state.next_unique, state.next_unique + 1
        entry.marker = default_marker(state)
        state.series.append(entry)
    return Series(chart, entries[-1])


def add(chart: Chart, source: object, rowcol: object, labels: object, categories: object) -> object:
    chart.admit_change()
    chart.nothing_filtered('SeriesCollection.Add')
    target = _source(chart, source)
    sheet = target.sheet
    if sheet.book is not chart.sheet.book:
        raise VBAUnsupportedError('SeriesCollection.Add from another workbook is not implemented')
    by_rows = rowcol is not MISSING and int(to_integer(rowcol, 'Long')) == 1
    named = labels is not MISSING and to_bool(labels)
    labelled = categories is not MISSING and to_bool(categories)
    if chart.state.series:
        return _added_beside(chart, target, by_rows, named, labelled)
    if len(target.areas) != 1:
        raise VBAUnsupportedError('SeriesCollection.Add of several areas on an empty chart is not implemented')
    area = _clipped(sheet, target.first)
    rows, columns = label_block(sheet, area)
    if named or labelled:
        rows = 1 if named else rows if labels is MISSING else 0
        columns = 1 if labelled else columns if categories is MISSING else 0
    return _made(chart, _planned(sheet, area, by_rows, rows, columns))


def _planned(sheet: Worksheet, area: Area, by_rows: bool, rows: int, columns: int) -> list[SeriesState]:
    """The series of a block whose first rows and columns are its labels."""
    entries: list[SeriesState] = []
    top, left = area.top + rows, area.left + columns
    for line in _lines(Area(top, left, area.bottom, area.right), by_rows):
        entry = SeriesState()
        entry.value_reference = reference(sheet, line)
        if by_rows:
            if columns:
                entry.name_reference = reference(sheet, Area(line.top, area.left, line.top, left - 1))
            if rows:
                entry.category_reference = reference(sheet, Area(area.top, left, top - 1, area.right))
        else:
            if rows:
                entry.name_reference = reference(sheet, Area(area.top, line.left, top - 1, line.left))
            if columns:
                entry.category_reference = reference(sheet, Area(top, area.left, area.bottom, left - 1))
        entries.append(entry)
    return entries


def _added_beside(chart: Chart, target: Range, by_rows: bool, named: bool, labelled: bool) -> object:
    sheet = target.sheet
    shared = chart.state.series[0].category_reference
    entries: list[SeriesState] = []
    for area in target.areas:
        lines = _lines(_clipped(sheet, area), by_rows)
        category = shared
        if labelled and len(lines) > 1:
            first = lines.pop(0)
            category = reference(sheet, _without_first(first, by_rows) if named else first)
        for line in lines:
            entry = SeriesState()
            entry.category_reference = category
            if named:
                entry.name_reference = reference(sheet, _first_of(line))
                entry.value_reference = reference(sheet, _without_first(line, by_rows))
            else:
                entry.value_reference = reference(sheet, line)
            entries.append(entry)
    return _made(chart, entries)


def _text_at(sheet: Worksheet, area: Area) -> bool:
    cell = sheet.cells_.get((area.top, area.left))
    return cell is not None and isinstance(cell.value, str)


def _joined(sheet: Worksheet, reference_text: str, line: Area, by_rows: bool) -> str:
    """A series' reference with this line added: into its last area if it carries it on, else as an area of its own."""
    addition = reference(sheet, line)
    if not reference_text:
        return addition
    areas = parse_reference(reference_text)
    last = areas[-1]
    if ((last.left, last.right) == (line.left, line.right) and last.bottom + 1 == line.top) or (
            by_rows and (last.top, last.bottom) == (line.top, line.bottom) and last.right + 1 == line.left):
        merged = Area(last.top, last.left, line.bottom, line.right)
        pieces = [reference(sheet, Area(one.top, one.left, one.bottom, one.right)) for one in areas[:-1]]
        pieces.append(reference(sheet, merged))
    else:
        pieces = [reference(sheet, Area(one.top, one.left, one.bottom, one.right)) for one in areas]
        pieces.append(addition)
    return pieces[0] if len(pieces) == 1 else f'({",".join(pieces)})'


def extend(chart: Chart, source: object, rowcol: object, categories: object) -> object:
    chart.admit_change()
    chart.nothing_filtered('SeriesCollection.Extend')
    target = _source(chart, source)
    sheet = target.sheet
    if len(target.areas) != 1:
        raise VBAUnsupportedError('SeriesCollection.Extend of several areas is not implemented')
    area = _clipped(sheet, target.first)
    by_rows = rowcol is not MISSING and int(to_integer(rowcol, 'Long')) == 1
    state = chart.state
    if not state.series:
        if rowcol is not MISSING or categories is not MISSING:
            raise error(-2147467259, 'Automation error\nUnspecified error ')
        _made(chart, _planned(sheet, area, False, *label_block(sheet, area)))
        return EMPTY
    if sheet is not chart.sheet and not any(sheet.name in entry.value_reference for entry in state.series):
        return EMPTY
    lines = _lines(area, by_rows)
    labels = lines.pop(0) if len(lines) > 1 and _text_at(sheet, lines[0]) else None
    chart.changed()
    for entry, line in zip(state.series, lines, strict=False):
        entry.value_reference = _joined(sheet, entry.value_reference, line, by_rows)
    if labels is not None and categories is not MISSING and to_bool(categories):
        for entry in state.series:
            entry.category_reference = _joined(sheet, entry.category_reference, labels, by_rows)
    return EMPTY

