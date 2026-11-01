"""What a cell of a pivot table stands for, as Range.PivotCell, PivotTable, PivotField and PivotItem answer it,
worked out from the table's lines and the layout of its fields (tests/fixtures/pivot_cells/: every cell of 595
tables in 23 workbooks Excel saved).

A cell of the body stands at a line of the rows and one of the columns.
Where both are data, it is a value of the data field its lines show; a
label cell names the field and item the column or row it sits in stands
for there; the header names the fields; the page area names each page
field and the item it shows. The row between the page fields and the
body, the column between two page fields and the cells beside a body
narrower than its page fields are not the table's.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba.apps.excel import _model, _pivots as pivots
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.apps.excel._pivot_forms import KINDS, compact_captions, label_columns, layout_of
from pyopenvba.apps.excel._pivot_reports import EVERY, GAP, PivotLine, axis_lines
from pyopenvba.apps.excel._pivots import NS
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._objects import member, method
from pyopenvba.interpreter._values import NULL, VBAInt, error, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Worksheet

VALUES = -2
#: A role's field where the cell stands for the data field it shows, Role.data.
DATA = -3

#: PivotCellType, xlPivotCellValue to xlPivotCellBlankCell.
VALUE, ITEM, SUBTOTAL, GRAND_TOTAL, DATA_FIELD, PIVOT_FIELD, PAGE_ITEM, CUSTOM_SUBTOTAL, DATA_PIVOT_FIELD, BLANK = \
    range(10)

#: CustomSubtotalFunction: the XlConsolidationFunction of each subtotal's type, in Subtotals' order.
FUNCTIONS = dict(zip(KINDS, (-4157, -4112, -4106, -4136, -4139, -4149, -4113, -4155, -4156, -4164, -4165),
                     strict=True))

#: The line a body cell stands at where its axis has no field: one line, at no item.
LONE = PivotLine('data', (), 0)


@dataclass(frozen=True)
class Role:
    """What a cell stands for: its PivotCellType; the field it names, a pivot field's index, VALUES for the Values
    field or DATA for the data field it shows; that data field's place among the table's; the item it names, by
    its field and place in the field's items, -1 for a page field's (All); the items its row and column stand
    at; and the function a custom subtotal works out. None where PivotCell has none to give."""

    kind: int
    field: int | None = None
    data: int | None = None
    item: tuple[int, int] | None = None
    rows: tuple[tuple[int, int], ...] | None = None
    columns: tuple[tuple[int, int], ...] | None = None
    custom: int | None = None


def _tabular(fields: list[XmlElement], index: int) -> bool:
    return index != VALUES and fields[index].get('outline', '1') in ('0', 'false')


def _page_role(table: pivots.PivotTable, row: int, column: int) -> Role | None:
    """A cell above the body: a page field's name, or the item it shows."""
    from pyopenvba.apps.excel._pivot_ranges import field_range

    fields = table.root.findall(f'{NS}pivotFields/{NS}pivotField')
    for page in table.root.findall(f'{NS}pageFields/{NS}pageField'):
        index = int(page.get('fld', '0'))
        view = pivots.PivotField(table, index, fields[index])
        for label in (True, False):
            found = field_range(view, label).first
            if (found.top, found.left) == (row, column):
                if label:
                    return Role(PIVOT_FIELD, index)
                # A page field showing several items still names (All) (tests/fixtures/pivot_cells/
                # pivot_options.json hidden_page_items).
                chosen = page.get('item')
                return Role(PAGE_ITEM, index, None, (index, -1 if chosen is None else int(chosen)))
    return None


def role(table: pivots.PivotTable, row: int, column: int) -> Role | None:
    """What the cell at ``row``, ``column`` stands for in ``table``, or None where it is not the table's."""
    root = table.root
    location = root.find(NS + 'location')
    if location is None:
        return None
    area = parse_area(location.get('ref', 'A1'))
    pages = int(location.get('rowPageCount', '0'))
    if not (area.top - (pages + 1 if pages else 0) <= row <= area.bottom and column >= area.left):
        return None
    if row < area.top:
        return _page_role(table, row, column)
    if column > area.right:
        return None
    if root.find(NS + 'cacheHierarchies') is not None or table.cache.root.find(NS + 'cacheHierarchies') is not None:
        raise VBAUnsupportedError('PivotCell on a pivot table of a cube is not implemented')
    header, top, left = (int(location.get(name, '0')) for name in ('firstHeaderRow', 'firstDataRow', 'firstDataCol'))
    rows = [int(entry.get('x', '0')) for entry in root.findall(f'{NS}rowFields/{NS}field')]
    columns = [int(entry.get('x', '0')) for entry in root.findall(f'{NS}colFields/{NS}field')]
    row_lines, column_lines = axis_lines(root, 'rowItems'), axis_lines(root, 'colItems')
    data_count = len(root.findall(f'{NS}dataFields/{NS}dataField'))
    fields = root.findall(f'{NS}pivotFields/{NS}pivotField')
    down, across = row - area.top, column - area.left
    if not rows and not columns and not data_count:
        # The empty table ClearTable leaves.
        return Role(BLANK)

    def items(line: PivotLine, axis: list[int]) -> tuple[tuple[int, int], ...]:
        return tuple((axis[place], item) for place, item in enumerate(line.items)
                     if axis[place] != VALUES and item not in (EVERY, GAP))

    def shows(line: PivotLine, axis: list[int]) -> bool:
        """Whether a line shows values: an item's own line, with fields inside it, only where the item is
        collapsed or shows its subtotal there, which a lone subtotal of an outline field does at the top."""
        depth = len(line.items) - 1
        if line.kind != 'data' or depth < 0 or depth >= len(axis) - 1 or axis[depth + 1] == VALUES:
            return True
        listed = fields[axis[depth]].findall(f'{NS}items/{NS}item')
        if 0 <= line.items[depth] < len(listed) and listed[line.items[depth]].get('sd') in ('0', 'false'):
            return True
        layout = layout_of(root, axis[depth])
        return len(layout.functions) == 1 and layout.outline and layout.top

    def label(line: PivotLine, axis: list[int], place: int, on_rows: bool) -> Role:
        """A label cell of a line, in the column or row of the field at ``place`` of the axis."""
        depth = len(line.items) - 1
        custom = FUNCTIONS.get(line.kind)
        if line.items[place] == EVERY:
            # An inner field's subtotal over every item of the fields outside it (tests/fixtures/pivot_cells/
            # pivot_layouts.json inner_sum_count_tabular, column_inner_sum_count).
            if on_rows:
                return Role(BLANK, axis[place])
            return Role(ITEM, axis[depth], None, (axis[depth], line.items[depth]), custom=custom)
        if line.kind not in ('data', 'grand', 'blank') and place < depth and not _tabular(fields, axis[place]):
            # An outline field's cell on a subtotal's line stands for the subtotal.
            place = depth
        field = axis[place]
        if place < depth and _tabular(fields, field):
            return Role(ITEM, field, None, (field, line.items[place]))
        if field == VALUES:
            return Role(DATA_FIELD, DATA, line.items[place])
        if axis[depth] == VALUES and on_rows:
            return Role(DATA_FIELD, field, None, (field, line.items[place]))
        if line.kind == 'data':
            return Role(ITEM, field, None, (field, line.items[place]))
        return Role(SUBTOTAL if custom is None else CUSTOM_SUBTOTAL, field, None, (field, line.items[place]),
                    custom=custom)

    if down >= top and across >= left:
        row_line = row_lines[down - top] if down - top < len(row_lines) else LONE
        column_line = column_lines[across - left] if across - left < len(column_lines) else LONE
        on_rows, on_columns = items(row_line, rows), items(column_line, columns)
        showing = shows(row_line, rows) and shows(column_line, columns)
        data = 0
        for line, axis in ((row_line, rows), (column_line, columns)):
            if VALUES in axis:
                place = axis.index(VALUES)
                if line.kind == 'grand' or (line.kind != 'data' and place >= len(line.items)):
                    data = line.data
                elif place < len(line.items):
                    data = line.items[place]
                else:
                    # A line the Values field's place is beyond names the Values field and no data field.
                    return Role(VALUE if showing else BLANK, VALUES, None, None, on_rows, on_columns)
                break
        if 'blank' in (row_line.kind, column_line.kind):
            return Role(BLANK, DATA, data)
        # A value on a line of custom subtotals names its function where the values are on that axis.
        custom = next((FUNCTIONS[line.kind] for line, axis in ((row_line, rows), (column_line, columns))
                       if VALUES in axis and line.kind in FUNCTIONS), None)
        return Role(VALUE if showing else BLANK, DATA, data, None, on_rows, on_columns, custom)
    if down >= top:
        line = row_lines[down - top] if down - top < len(row_lines) else LONE
        if line.kind == 'blank':
            return Role(BLANK)
        if line.kind == 'grand':
            return Role(GRAND_TOTAL, DATA, line.data if VALUES in rows else 0)
        if not rows:
            return Role(DATA_FIELD, DATA, 0)
        depth = len(line.items) - 1
        if depth < 0:
            return Role(BLANK)
        places = label_columns(root, rows)[0]
        within = [place for place in range(depth + 1) if places[place] == across] if across <= places[depth] else []
        return label(line, rows, within[-1] if within else depth, True)
    if header <= down < top and across >= left and columns:
        line = column_lines[across - left] if across - left < len(column_lines) else LONE
        if line.kind == 'blank':
            return Role(BLANK)
        if line.kind == 'grand':
            return Role(GRAND_TOTAL, DATA, line.data if VALUES in columns else 0)
        depth = len(line.items) - 1
        if depth < 0:
            return Role(BLANK)
        return label(line, columns, min(down - header, depth), False)
    # The header: the fields' names or buttons, the data field's caption, and blanks.
    captions = root.get('showHeaders', '1') not in ('0', 'false')
    compact = compact_captions(root, rows + columns)
    if across < left and down == top - 1 and rows:
        named = [place for place, at in enumerate(label_columns(root, rows)[0]) if at == across]
        if named:
            field = rows[named[0]]
            if not captions:
                return Role(BLANK)
            if field == VALUES:
                return Role(DATA_PIVOT_FIELD, VALUES)
            return Role(BLANK if compact else PIVOT_FIELD, field)
    if not columns and down == top - 1 and across >= left:
        return Role(GRAND_TOTAL) if VALUES in rows else Role(DATA_FIELD, DATA, 0)
    if down == 0 and across == 0 and data_count == 1 and columns:
        return Role(DATA_FIELD, DATA, 0)
    if down < header and across >= left and columns:
        place = 0 if compact else across - left
        if not captions or (compact and across != left) or place >= len(columns):
            return Role(BLANK)
        if columns[place] == VALUES:
            return Role(DATA_PIVOT_FIELD, VALUES)
        return Role(BLANK if compact else PIVOT_FIELD, columns[place])
    if across < left and (down < top - 1 or not rows):
        return Role(BLANK, DATA, 0) if data_count == 1 else Role(BLANK, VALUES)
    return Role(BLANK)


def owner(sheet: Worksheet, row: int, column: int) -> tuple[pivots.PivotTable, Role] | None:
    """The pivot table a cell is part of, and what it stands for there."""
    for table in sheet.pivot_tables:
        found = role(table, row, column)
        if found is not None:
            return table, found
    return None


def _field(table: pivots.PivotTable, found: Role) -> pivots.PivotField | None:
    root = table.root
    if found.field == DATA:
        reports = root.findall(f'{NS}dataFields/{NS}dataField')
        if found.data is None or found.data >= len(reports):
            return None
        return pivots.PivotField(table, int(reports[found.data].get('fld', '0')), reports[found.data], True)
    if found.field == VALUES:
        return table.values_field()
    if found.field is None:
        return None
    return pivots.PivotField(table, found.field, root.findall(f'{NS}pivotFields/{NS}pivotField')[found.field])


def _item(table: pivots.PivotTable, field: int, place: int) -> pivots.PivotItem:
    from pyopenvba.apps.excel._pivot_filters import all_item

    if field == VALUES:
        return table.values_field().listed_items()[place]
    view = pivots.PivotField(table, field, table.root.findall(f'{NS}pivotFields/{NS}pivotField')[field])
    if place < 0:
        return all_item(view)
    return view.listed_items()[place]


def _data_item(table: pivots.PivotTable, found: Role) -> pivots.PivotItem | None:
    """The item a cell showing a data field names: the Values field's item of that data field."""
    if found.data is None:
        return None
    listed = table.values_field().listed_items()
    return listed[found.data] if found.data < len(listed) else None


def _unable(what: str, of: str) -> VBARuntimeError:
    return error(1004, f'Unable to get the {what} property of the {of} class')


class PivotCell(ExcelObject):
    """A cell of a pivot table, as Range.PivotCell gives it."""

    vba_type_name = 'PivotCell'

    # What-if changes belong to an OLAP table; a cell of a table over a list refuses both with 1004
    # (tests/fixtures/pivot_olap_members.json).

    @method
    def AllocateChange(self) -> object:
        raise error(1004, 'AllocateChange method of PivotCell class failed')

    @method
    def DiscardChange(self) -> object:
        raise error(1004, 'DiscardChange method of PivotCell class failed')

    def __init__(self, table: pivots.PivotTable, found: Role, cell: _model.Range) -> None:
        self.table, self.found, self.cell = table, found, cell

    @member
    def Application(self) -> object:
        return self.table.sheet.book.application

    @member
    def PivotCellType(self) -> object:
        return self.found.kind

    @member
    def PivotTable(self) -> object:
        return self.table

    @member
    def PivotField(self) -> object:
        field = _field(self.table, self.found)
        if field is None:
            raise _unable('PivotField', 'PivotCell')
        return field

    @member
    def DataField(self) -> object:
        field = _field(self.table, self.found) if self.found.data is not None else None
        if field is None:
            raise _unable('DataField', 'PivotCell')
        return field

    @member
    def PivotItem(self) -> object:
        if self.found.item is None:
            raise _unable('PivotItem', 'PivotCell')
        return _item(self.table, *self.found.item)

    def _items(self, entries: tuple[tuple[int, int], ...] | None, what: str) -> object:
        if entries is None:
            raise _unable(what, 'PivotCell')
        return pivots.PivotItemList(self, [_item(self.table, field, place) for field, place in entries])

    @member
    def RowItems(self) -> object:
        return self._items(self.found.rows, 'RowItems')

    @member
    def ColumnItems(self) -> object:
        return self._items(self.found.columns, 'ColumnItems')

    @member
    def CustomSubtotalFunction(self) -> object:
        if self.found.custom is None:
            raise _unable('CustomSubtotalFunction', 'PivotCell')
        return self.found.custom

    @member
    def Range(self) -> object:
        return self.cell


def _cell(target: _model.Range) -> tuple[pivots.PivotTable, Role, _model.Range] | None:
    """The range's first cell, with the table it is part of and what it stands for there."""
    first = target.first
    found = owner(target.sheet, first.top, first.left)
    if found is None:
        return None
    cell = _model.Range(target.sheet, [Area(first.top, first.left, first.top, first.left, target.sheet.name)])
    return found[0], found[1], cell


def pivot_cell(target: _model.Range) -> PivotCell:
    """Range.PivotCell: the pivot cell its first cell is, error 1004 outside a pivot table."""
    found = _cell(target)
    if found is None:
        raise _unable('PivotCell', 'Range')
    return PivotCell(*found)


#: LocationInTable of a pivot field's own button by the field's orientation: xlRowHeader, xlColumnHeader, xlPageHeader.
_HEADERS = {1: -4153, 2: -4110, 3: 2}
#: LocationInTable of an item, subtotal or custom subtotal by its field's orientation: xlRowItem, xlColumnItem.
_ITEMS = {1: 4, 2: 5}
#: LocationInTable of a blank cell by its field's orientation, inside that field's area and outside it.
_BLANKS = {1: ('RowRange', -4153, 3), 2: ('ColumnRange', -4110, 3)}


def _within(table: pivots.PivotTable, name: str, row: int, column: int) -> bool:
    try:
        found = table.vba_get(name)
    except VBARuntimeError:
        return False
    return isinstance(found, _model.Range) and any(area.contains(row, column) for area in found.areas)


def _location(sheet: Worksheet, row: int, column: int) -> int:
    """LocationInTable of one cell: by what it stands for and its field's orientation, the values area's body
    first (tests/fixtures/location_in_table/)."""
    found = owner(sheet, row, column)
    field = _field(*found) if found is not None else None
    if found is None or field is None:
        raise _unable('LocationInTable', 'Range')
    table, place = found
    kind = int(place.kind)
    orientation = int(to_integer(field.vba_get('Orientation'), 'Long'))
    if kind == 6:
        return 6
    if kind == 5 and orientation in _HEADERS:
        return _HEADERS[orientation]
    if kind == 8:
        return 3
    if _within(table, 'DataBodyRange', row, column) or kind == 0:
        return 8
    if kind in (1, 2, 7) and orientation in _ITEMS:
        return _ITEMS[orientation]
    if kind == 3 or (kind in (4, 9) and orientation == 4):
        return 7
    if kind == 4 and orientation == 1:
        return 4
    if kind == 9 and orientation in _BLANKS:
        area, inside, outside = _BLANKS[orientation]
        return inside if _within(table, area, row, column) else outside
    raise VBAUnsupportedError('Range.LocationInTable of a pivot cell of this kind is not implemented')


def location_in_table(target: _model.Range) -> object:
    """Range.LocationInTable: where in a pivot table its cells are; Null where they differ, a cell that is no pivot
    cell with a field among them, and error 1004 where none is one (tests/fixtures/location_in_table/)."""
    if len(target.areas) == 1 and target.single:
        return VBAInt(_location(target.sheet, target.first.top, target.first.left), 'Long')
    if sum(area.rows * area.columns for area in target.areas) > 1 << 20:
        raise VBAUnsupportedError('Range.LocationInTable of more than a million cells is not implemented')
    found: set[int | None] = set()
    for area in target.areas:
        for row in range(area.top, area.bottom + 1):
            for column in range(area.left, area.right + 1):
                try:
                    found.add(_location(target.sheet, row, column))
                except VBARuntimeError:
                    found.add(None)
    if found == {None}:
        raise _unable('LocationInTable', 'Range')
    only = found.pop() if len(found) == 1 else None
    return NULL if only is None else VBAInt(only, 'Long')


def pivot_table(target: _model.Range) -> pivots.PivotTable:
    """Range.PivotTable: the pivot table its first cell is part of, its page fields too, error 1004 elsewhere."""
    found = _cell(target)
    if found is None:
        raise _unable('PivotTable', 'Range')
    return found[0]


def pivot_field(target: _model.Range) -> pivots.PivotField:
    """Range.PivotField: the field its first cell names, a value's data field among them."""
    found = _cell(target)
    field = _field(found[0], found[1]) if found is not None else None
    if field is None:
        raise _unable('PivotField', 'Range')
    return field


def pivot_item(target: _model.Range) -> pivots.PivotItem:
    """Range.PivotItem: the item its first cell names, and for a cell showing a data field, the Values field's
    item of that data field (tests/fixtures/pivot_cells/)."""
    found = _cell(target)
    if found is None:
        raise _unable('PivotItem', 'Range')
    table, place, _ = found
    item = _item(table, *place.item) if place.item is not None else _data_item(table, place)
    if item is None:
        raise _unable('PivotItem', 'Range')
    return item
