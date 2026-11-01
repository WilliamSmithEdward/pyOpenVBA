"""Bounded ordinary worksheet outlines and persisted outline configuration."""
# ruff: noqa: N803 -- Excel public named arguments retain their reference spelling.
from __future__ import annotations

import math
from copy import deepcopy
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba._opc_xml import opc_part
from pyopenvba.apps.excel._dimensions import ColumnRecord, RowRecord
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.apps.excel._page_setup import patch_configuration
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, VBAInt, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
LIMIT = 50_000


def unsupported(detail: str) -> VBAUnsupportedError:
    return VBAUnsupportedError('Worksheet outlines ' + detail + ' are not implemented')


def _attributes() -> dict[str, str]:
    return {}


def _changes() -> dict[str, str | None]:
    return {}


def _indices() -> set[int]:
    return set()


@dataclass
class OutlineState:
    attributes: dict[str, str] = field(default_factory=_attributes)
    changes: dict[str, str | None] = field(default_factory=_changes)
    changed_rows: set[int] = field(default_factory=_indices)
    changed_columns: set[int] = field(default_factory=_indices)
    axes_checked: bool = False
    loaded_rows: set[int] = field(default_factory=_indices)


def state(sheet: Worksheet) -> OutlineState:
    if sheet.outline_state is None:
        attrs: dict[str, str] = {}
        package = sheet.book.package
        if package is not None and sheet.part_name and package.has(sheet.part_name):
            data = package.read(sheet.part_name)
            if len(data) > 16 * 1024 * 1024:
                raise unsupported('with more than 16 MiB of worksheet XML')
            try:
                part = opc_part(data, NS, 'worksheet', container_path=('sheetPr',), max_children=32)
            except ValueError:
                raise unsupported('with invalid or oversized outline metadata') from None
            if part is not None:
                found = [a for name, a, _start, _end in part.children if name == NS + '}outlinePr']
                if len(found) > 1:
                    raise unsupported('with duplicate outline properties')
                if found:
                    attrs = dict(found[0])
        sheet.outline_state = OutlineState(attrs)
    return sheet.outline_state


def with_outline(sheet: Worksheet, xml: str) -> str:
    current = sheet.outline_state
    if current is None:
        return xml
    if current.changed_rows or current.changed_columns:
        from pyopenvba.apps.excel._outline_file import patch

        xml = patch(sheet, xml, current.changed_rows, current.changed_columns)
    if not current.changes:
        return xml
    part = opc_part(xml.encode(), NS, 'worksheet', container_path=('sheetPr',))
    if part is None:
        xml = patch_configuration(xml, 'worksheet', ('sheetPr',), {})
        part = opc_part(xml.encode(), NS, 'worksheet', container_path=('sheetPr',))
        assert part is not None
    found = [1 for name, _a, _s, _e in part.children if name == NS + '}outlinePr']
    if not found:
        # outlinePr precedes pageSetUpPr in sheetPr, including prefixed originals.
        at = next((start for name, _a, start, _end in part.children if name == NS + '}pageSetUpPr'), part.closing_start)
        xml = part.insert(at, part.child('outlinePr', '')).decode()
    return patch_configuration(xml, 'worksheet', ('sheetPr', 'outlinePr'), current.changes)


def _alive(sheet: Worksheet) -> None:
    if sheet not in sheet.book.sheets_:
        raise unsupported('on a removed worksheet')


def _admit_axes(sheet: Worksheet) -> None:
    _alive(sheet)
    current = state(sheet)
    if not current.axes_checked:
        from pyopenvba.apps.excel._outline_file import admit

        current.loaded_rows = admit(sheet)
        current.axes_checked = True


def _guard(sheet: Worksheet, *, configuration: bool = False) -> bool:
    from pyopenvba.apps.excel._protection import FILTERING, enforced

    _alive(sheet)
    if sheet.book.application.udf_callers:
        return False
    if enforced(sheet) is not None:
        raise error(1004) if configuration else error(1004, FILTERING)
    if not configuration:
        _admit_axes(sheet)
    if not configuration and state(sheet).attributes.get('applyStyles') in ('1', 'true'):
        raise unsupported('with automatic summary styles')
    return True


def _flag(sheet: Worksheet, key: str, default: bool) -> bool:
    value = state(sheet).attributes.get(key)
    if value is None:
        return default
    if value not in ('0', '1', 'false', 'true'):
        raise unsupported('with unrecognized stored outline flags')
    return value in ('1', 'true')


def _number(value: object) -> int:
    if value is EMPTY:
        return 0
    if value is NULL or isinstance(value, bool):
        raise error(1004)
    if not isinstance(value, (int, float, str)):
        raise unsupported('with unmeasured level arguments')
    try:
        number = float(value)
    except ValueError:
        raise error(1004) from None
    if not math.isfinite(number):
        raise unsupported('with nonfinite level arguments')
    return int(number)


def _change_flag(sheet: Worksheet, key: str, value: str) -> None:
    if not _guard(sheet, configuration=True):
        return
    prepared = deepcopy(state(sheet))
    prepared.attributes[key] = value
    prepared.changes[key] = value
    package = sheet.book.package
    xml = package.read(sheet.part_name).decode() if package is not None and sheet.part_name else f'<worksheet xmlns="{NS}"><sheetData/></worksheet>'
    # Validate duplicate paths/namespace admission before publishing state.
    old = sheet.outline_state
    try:
        sheet.outline_state = prepared
        with_outline(sheet, xml)
    finally:
        sheet.outline_state = old
    sheet.outline_state = prepared
    sheet.book.saved = False


def axis(target: Range) -> tuple[bool, int, int]:
    _alive(target.sheet)
    if len(target.areas) != 1:
        raise unsupported('on multiple areas')
    area = target.first
    rows = area.left == 1 and area.right == MAX_COLUMNS
    columns = area.top == 1 and area.bottom == MAX_ROWS
    if rows == columns:
        # The whole sheet, or a block of it: rows or columns as Rows or Columns made it.
        rows, columns = target.lines == 'rows', target.lines == 'columns'
    if rows == columns:
        raise unsupported('on ranges other than whole rows or columns')
    return (True, area.top, area.bottom) if rows else (False, area.left, area.right)


def _records(sheet: Worksheet, rows: bool) -> dict[int, RowRecord] | dict[int, ColumnRecord]:
    return sheet.dims.rows if rows else sheet.dims.columns


def _depth(record: RowRecord | ColumnRecord | None) -> int:
    raw = record.extra.get('outlineLevel', '0') if record else '0'
    if raw not in tuple(str(n) for n in range(8)):
        raise unsupported('with invalid stored outline depths')
    return int(raw)


def _folded(record: RowRecord | ColumnRecord | None) -> bool:
    return record is not None and record.extra.get('collapsed') in ('1', 'true')


def _cap(count: int) -> None:
    if count > LIMIT:
        raise unsupported('over more than 50000 rows or columns per operation')


def _commit_plans(sheet: Worksheet, plans: list[tuple[bool, dict[int, RowRecord | ColumnRecord]]], *, notify: bool = True) -> bool:
    current = state(sheet)
    if len(sheet.cells_) > 250_000:
        raise unsupported('requiring more than 250000 existing cell records of row-budget work')
    row_records, column_records = dict(sheet.dims.rows), dict(sheet.dims.columns)
    pending_rows, pending_columns = set(current.changed_rows), set(current.changed_columns)
    for rows, changed in plans:
        for number, record in changed.items():
            if rows:
                assert isinstance(record, RowRecord)
                row_records[number] = record
                pending_rows.add(number)
            else:
                assert isinstance(record, ColumnRecord)
                column_records[number] = record
                pending_columns.add(number)
    _cap(len(current.loaded_rows | set(row_records) | {r for r, _c in sheet.cells_}))
    formatting = dict(sheet.dims.format)
    for key, records in (('outlineLevelRow', row_records), ('outlineLevelCol', column_records)):
        depth = max((_depth(r) for r in records.values()), default=0)
        if depth:
            formatting[key] = str(depth)
        else:
            formatting.pop(key, None)
    from pyopenvba.apps.excel._outline_file import patch

    package = sheet.book.package
    xml = package.read(sheet.part_name).decode() if package is not None and sheet.part_name else f'<worksheet xmlns="{NS}"><sheetData/></worksheet>'
    # Validate both axes together before publishing records or notifying calculation.
    patch(sheet, xml, pending_rows, pending_columns, prospective_rows=row_records,
          prospective_columns=column_records, prospective_format=formatting)
    visible = False
    for rows, changed in plans:
        for number, record in changed.items():
            if rows:
                assert isinstance(record, RowRecord)
                visible |= sheet.dims.row_hidden(number) != record.hidden
                sheet.dims.rows[number] = record
                sheet.dims.settle_row(number)
            else:
                assert isinstance(record, ColumnRecord)
                visible |= sheet.dims.column_hidden(number) != record.hidden
                sheet.dims.columns[number] = record
                sheet.dims.columns_changed = True
                if record.empty():
                    sheet.dims.columns.pop(number, None)
        (current.changed_rows if rows else current.changed_columns).update(changed)
    sheet.dims.format.clear()
    sheet.dims.format.update(formatting)
    sheet.dims.format_changed = True
    sheet.book.saved = False
    if visible and notify:
        from pyopenvba.apps.excel._events import visibility_changed

        visibility_changed(sheet.book)
    return visible


def _commit(sheet: Worksheet, rows: bool, changed: dict[int, RowRecord | ColumnRecord]) -> bool:
    return _commit_plans(sheet, [(rows, changed)])


def _clone(sheet: Worksheet, rows: bool, number: int) -> RowRecord | ColumnRecord:
    record = _records(sheet, rows).get(number)
    copied = deepcopy(record) if record is not None else RowRecord() if rows else ColumnRecord()
    if isinstance(copied, ColumnRecord) and copied.width is None:
        copied.width, copied.custom = sheet.dims.standard_width(), True
    return copied


def group(target: Range, operation: str) -> bool:
    rows, first, last = axis(target)
    _cap(last - first + 1)
    if not _guard(target.sheet):
        return False
    records = _records(target.sheet, rows)
    # Validate every existing depth before any record/style/visibility changes.
    old = {n: _depth(records.get(n)) for n in range(first, last + 1)}
    if operation == 'Group' and any(depth == 7 for depth in old.values()):
        raise error(1004, 'Unable to get the Group property of the Range class')
    if operation == 'Ungroup' and not any(old.values()):
        raise error(1004, 'Unable to get the Ungroup property of the Range class')
    changed: dict[int, RowRecord | ColumnRecord] = {}
    for number, depth in old.items():
        record = _clone(target.sheet, rows, number)
        new = depth + 1 if operation == 'Group' else max(0, depth - 1) if operation == 'Ungroup' else 0
        if new:
            record.extra['outlineLevel'] = str(new)
        else:
            record.extra.pop('outlineLevel', None)
        if operation == 'ClearOutline':
            record.hidden = False
            record.extra.pop('collapsed', None)
        changed[number] = record
    if operation == 'ClearOutline':
        after = _flag(target.sheet, 'summaryBelow' if rows else 'summaryRight', True)
        summary, neighbor = (last + 1, last) if after else (first - 1, first)
        limit = MAX_ROWS if rows else MAX_COLUMNS
        if 1 <= summary <= limit and _folded(records.get(summary)) and _depth(changed[neighbor]) <= _depth(records.get(summary)):
            summary_record = _clone(target.sheet, rows, summary)
            summary_record.extra.pop('collapsed', None)
            changed[summary] = summary_record
    _commit(target.sheet, rows, changed)
    return True


def read_level(target: Range) -> object:
    _admit_axes(target.sheet)
    rows, first, last = axis(target)
    if first != last:
        records = _records(target.sheet, rows)
        found = {_depth(records.get(first)), *(_depth(r) for n, r in records.items() if first <= n <= last)}
        if len(found) > 1 or (next(iter(found)) and last - first + 1 > sum(first <= n <= last for n in records)):
            raise unsupported('with heterogeneous multi-row or multi-column outline level reads')
    return float(_depth(_records(target.sheet, rows).get(first)) + 1)


def set_level(target: Range, value: object) -> None:
    rows, first, last = axis(target)
    _cap(last - first + 1)
    number = int(value) if isinstance(value, bool) else _number(value)
    if number < 1 or number > 8:
        raise error(1004)
    if not _guard(target.sheet):
        raise unsupported('setting OutlineLevel from worksheet functions')
    changed: dict[int, RowRecord | ColumnRecord] = {}
    for n in range(first, last + 1):
        record = _clone(target.sheet, rows, n)
        if number == 1:
            record.extra.pop('outlineLevel', None)
        else:
            record.extra['outlineLevel'] = str(number - 1)
        changed[n] = record
    _commit(target.sheet, rows, changed)


def read_detail(target: Range) -> bool:
    _admit_axes(target.sheet)
    rows, first, last = axis(target)
    if first != last:
        raise unsupported('with multiple summary rows or columns')
    return not _folded(_records(target.sheet, rows).get(first))


def set_detail(target: Range, value: object) -> None:
    rows, first, last = axis(target)
    if first != last:
        raise unsupported('with multiple summary rows or columns')
    if not isinstance(value, bool):
        raise unsupported('with unmeasured ShowDetail coercions')
    if not _guard(target.sheet):
        raise unsupported('setting ShowDetail from worksheet functions')
    records = _records(target.sheet, rows)
    level = _depth(records.get(first))
    step = -1 if _flag(target.sheet, 'summaryBelow' if rows else 'summaryRight', True) else 1
    position = first + step
    detail: list[int] = []
    limit = MAX_ROWS if rows else MAX_COLUMNS
    while 1 <= position <= limit and _depth(records.get(position)) > level:
        _cap(len(detail) + 1)
        detail.append(position)
        position += step
    if not detail:
        return
    if not value and _folded(records.get(first)):
        raise error(1004)
    changed = {n: _clone(target.sheet, rows, n) for n in [*detail, first]}
    for n in detail:
        changed[n].hidden = not value
    changed[first].extra.pop('collapsed', None)
    if not value:
        changed[first].extra['collapsed'] = '1'
    _commit(target.sheet, rows, changed)


class Outline(ExcelObject):
    vba_type_name = 'Outline'

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet
        self._vba_application = sheet.book.application

    def _check_alive(self) -> None:
        super()._check_alive()
        _alive(self.sheet)

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, 'Long')

    @member
    def AutomaticStyles(self) -> bool:
        return _flag(self.sheet, 'applyStyles', False)

    @setter('AutomaticStyles')
    def _set_automatic_styles(self, value: object) -> None:
        if value is not False:
            raise unsupported('enabling automatic summary styles')
        _change_flag(self.sheet, 'applyStyles', '0')

    @member
    def SummaryRow(self) -> object:
        return VBAInt(int(_flag(self.sheet, 'summaryBelow', True)), 'Long')

    @setter('SummaryRow')
    def _set_summary_row(self, value: object) -> None:
        if isinstance(value, str):
            raise error(1004)
        if value is EMPTY:
            below = False
        elif value is NULL:
            below = True
        elif isinstance(value, bool):
            below = value
        elif isinstance(value, (int, float)):
            number = float(value)
            if not math.isfinite(number):
                raise unsupported('with nonfinite summary row values')
            if number.is_integer() and number not in (0, 1):
                raise unsupported('with unmeasured integer summary row values')
            below = number != 0
        else:
            raise unsupported('with unmeasured summary row values')
        _change_flag(self.sheet, 'summaryBelow', '1' if below else '0')

    @member
    def SummaryColumn(self) -> object:
        return VBAInt(-4152 if _flag(self.sheet, 'summaryRight', True) else -4131, 'Long')

    @setter('SummaryColumn')
    def _set_summary_column(self, value: object) -> None:
        if isinstance(value, str):
            raise error(1004)
        if value is EMPTY:
            right = False
        elif value is NULL:
            right = True
        elif isinstance(value, bool):
            right = value
        elif isinstance(value, (int, float)):
            number = float(value)
            if not math.isfinite(number):
                raise unsupported('with nonfinite summary column values')
            if number.is_integer() and number not in (-4152, -4131, 3):
                raise unsupported('with unmeasured integer summary column values')
            right = number != -4131
        else:
            raise unsupported('with unmeasured summary column values')
        _change_flag(self.sheet, 'summaryRight', '1' if right else '0')

    @method
    def ShowLevels(self, RowLevels: object = MISSING, ColumnLevels: object = MISSING) -> object:
        values: list[int] = []
        for value in (RowLevels, ColumnLevels):
            number = 0 if value is MISSING or value is NULL else int(value) if isinstance(value, bool) else _number(value)
            if not 0 <= number <= 8:
                raise error(1004, 'Unable to get the ShowLevels property of the Outline class')
            values.append(number)
        if not _guard(self.sheet):
            return False
        plans: list[tuple[bool, dict[int, RowRecord | ColumnRecord]]] = []
        for rows, number in zip((True, False), values, strict=True):
            if not number:
                continue
            records = _records(self.sheet, rows)
            _cap(len(records))
            levels = {n: _depth(r) for n, r in records.items()}
            positions = {*records}
            below = _flag(self.sheet, 'summaryBelow' if rows else 'summaryRight', True)
            step = -1 if below else 1
            limit = MAX_ROWS if rows else MAX_COLUMNS
            positions.update(n - step for n in records if 1 <= n - step <= limit)
            _cap(len(positions))
            changed = {n: _clone(self.sheet, rows, n) for n in positions}
            for n, record in changed.items():
                record.hidden = levels.get(n, 0) >= number
                if levels.get(n, 0) < number:
                    record.extra.pop('collapsed', None)
                if levels.get(n + step, 0) > levels.get(n, 0) and levels.get(n, 0) == number - 1:
                    record.extra['collapsed'] = '1'
            plans.append((rows, changed))
        _commit_plans(self.sheet, plans)
        return True


def admit_layout(sheet: Worksheet) -> None:
    """Admit an outline a command lays out line by line itself, as Subtotal does: the stored one read and checked."""
    _alive(sheet)
    _admit_axes(sheet)
    if state(sheet).attributes.get('applyStyles') in ('1', 'true'):
        raise unsupported('with automatic summary styles')


def depth(sheet: Worksheet, rows: bool, number: int) -> int:
    """A row's or a column's outline depth as the file numbers it, 0 outside every group."""
    return _depth(_records(sheet, rows).get(number))


def relevel(sheet: Worksheet, rows: dict[int, int], columns: dict[int, int], shown_rows: set[int],
            shown_columns: set[int], *, folded: dict[int, bool] | None = None) -> bool:
    """Rows and columns laid at the depths given, the ``shown`` ones made visible with their collapsed marks taken
    off and rows ``folded`` marked collapsed or not, in one commit that raises no event; whether a line's visibility
    changed."""
    plans: list[tuple[bool, dict[int, RowRecord | ColumnRecord]]] = []
    for lines, depths, shown in ((True, rows, shown_rows), (False, columns, shown_columns)):
        marks = (folded or {}) if lines else {}
        changed: dict[int, RowRecord | ColumnRecord] = {}
        for number in sorted({*depths, *shown, *marks}):
            record = _clone(sheet, lines, number)
            if number in depths:
                if depths[number]:
                    record.extra['outlineLevel'] = str(depths[number])
                else:
                    record.extra.pop('outlineLevel', None)
            if number in marks:
                if marks[number]:
                    record.extra['collapsed'] = '1'
                else:
                    record.extra.pop('collapsed', None)
            if number in shown:
                record.hidden = False
                record.extra.pop('collapsed', None)
            changed[number] = record
        if changed:
            plans.append((lines, changed))
    return bool(plans) and _commit_plans(sheet, plans, notify=False)


def apply_outline_styles(target: Range) -> None:
    """Range.ApplyOutlineStyles: each summary row in the range takes RowLevel_n and each summary column ColLevel_n,
    n being its own outline level. A summary is the row (column) beside a deeper one on the side its detail lies,
    below (right of) it unless the outline says otherwise. The styles are made in level order and given top to
    bottom, as Excel makes and gives them (tests/fixtures/outline_styles.json)."""
    from pyopenvba.apps.excel._cell_styles import find
    from pyopenvba.apps.excel._model import Range

    sheet = target.sheet
    found: list[tuple[bool, int, int]] = []
    for rows in (True, False):
        records = _records(sheet, rows)
        step = -1 if _flag(sheet, 'summaryBelow' if rows else 'summaryRight', True) else 1
        for area in target.areas:
            first, last = (area.top, area.bottom) if rows else (area.left, area.right)
            for number in range(first, last + 1):
                own = _depth(records.get(number))
                if _depth(records.get(number + step)) > own:
                    found.append((rows, number, own + 1))
    if not found:
        return
    names = {(rows, level): f"{'Row' if rows else 'Col'}Level_{level}" for rows, _number, level in found}
    for key in sorted(names, key=lambda key: (not key[0], key[1])):
        find(sheet.book, names[key])
    for rows, number, level in sorted(found, key=lambda item: (not item[0], item[1])):
        for area in target.areas:
            cells = (Area(number, area.left, number, area.right, sheet.name) if rows
                     else Area(area.top, number, area.bottom, number, sheet.name))
            if (area.top <= number <= area.bottom) if rows else (area.left <= number <= area.right):
                Range(sheet, [cells]).vba_set('Style', names[(rows, level)])


def summary_below(sheet: Worksheet) -> bool:
    """Whether the sheet's summary rows go below their detail, as Outline.SummaryRow has it."""
    return _flag(sheet, 'summaryBelow', True)


def set_summary_below(sheet: Worksheet, below: bool) -> None:
    _change_flag(sheet, 'summaryBelow', '1' if below else '0')


def set_summary_right(sheet: Worksheet, right: bool) -> None:
    _change_flag(sheet, 'summaryRight', '1' if right else '0')


def outlined(sheet: Worksheet, rows: bool) -> list[int]:
    """The rows, or columns, in a group."""
    return sorted(number for number, record in _records(sheet, rows).items() if _depth(record))


def _admit_stored_axes(sheet: Worksheet) -> None:
    package = sheet.book.package
    if package is not None and sheet.part_name and package.has(sheet.part_name):
        data = package.read(sheet.part_name)
        if b'outlineLevel' in data or b'collapsed' in data:
            # Unparsed qualified axis records must not disappear through a legacy
            # structural serializer simply because dimensions did not see them.
            _admit_axes(sheet)


def guard_axis_edit(sheet: Worksheet, rows: bool, start: int) -> None:
    _admit_stored_axes(sheet)
    if any(n >= max(1, start - 1) and (_depth(record) or _folded(record))
           for n, record in _records(sheet, rows).items()):
        raise unsupported('for structural edits moving or intersecting existing outline metadata')


def guard_axis_copy(source: Worksheet, area: Area, destination: Worksheet, target: Area) -> None:
    if area.whole_rows or area.whole_columns:
        _admit_stored_axes(source)
        _admit_stored_axes(destination)
    for rows in (True, False):
        if not (area.whole_rows if rows else area.whole_columns):
            continue
        for sheet, box in ((source, area), (destination, target)):
            first, last = (box.top, box.bottom) if rows else (box.left, box.right)
            if any(first <= n <= last and (_depth(record) or _folded(record))
                   for n, record in _records(sheet, rows).items()):
                raise unsupported('for whole-axis range copies of outline metadata')


def admit_sheet_copy(source: Worksheet) -> None:
    _admit_stored_axes(source)
    current = state(source)
    if current.attributes or any(_depth(r) or _folded(r) for records in (source.dims.rows, source.dims.columns) for r in records.values()):
        _admit_axes(source)
        _cap(len(source.dims.rows))
        _cap(len(source.dims.columns))
    if set(current.attributes) - {'summaryBelow', 'summaryRight', 'applyStyles', 'showOutlineSymbols'}:
        raise unsupported('copying unmodeled outline properties')
    if current.attributes.get('applyStyles') in ('1', 'true'):
        raise unsupported('copying automatic outline styles')


def copied(source: Worksheet, destination: Worksheet) -> None:
    current = deepcopy(state(source))
    current.changes.update(current.attributes)
    current.changed_rows.update(destination.dims.rows)
    current.changed_columns.update(destination.dims.columns)
    current.loaded_rows = set()
    current.axes_checked = True
    destination.outline_state = current
