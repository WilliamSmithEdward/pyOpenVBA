"""What-if data tables: the cells Range.Table makes, the edits Excel lets through them, and what a file keeps.

Range.Table over a block makes a table of the block's cells but its first
row and column. With a column input alone, each value down the first
column is tried in the input cell against each formula along the first
row; with a row input alone, each value along the first row against each
formula down the first column; with both, each pair against the formula
at the block's top left. A cell is worked out with the value tried put in
the input cell and everything that reads the input, on any sheet, worked
out again; then all of it is put back (see _calc). A blank formula comes
to 0 there, and a constant to itself.

A table's cell says =TABLE(row input,column input), =TABLE(,A1) for a
column input alone, its HasArray is True and its CurrentArray the whole
block; it holds 0 until it is first worked out. Excel guards a table as
it guards an array formula: a value or formula written over part of it,
or over part of it and more, changes nothing at all, and over all of it
takes the table away for what was written; ClearContents over part of it
is error 1004, over all of it takes it away; rows, columns and shifted
cells cannot go through its block, which moves whole with them, its
input cells following as references do, an input cell taken away being
#REF!. A copy of its cells is their values. In semi-automatic
calculation a table is worked out only by Application.Calculate or
Worksheet.Calculate, never Range.Calculate, or by calculation going back
to automatic. A file keeps a table on its first cell, <f t="dataTable"
ref dt2D dtr del1 del2 r1 r2/>, with each cell's value
(tests/fixtures/data_tables.json).
"""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, column_letter, parse_area
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import MISSING, NOTHING, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Cell, Range, Worksheet

#: What Excel says when an edit would change part of a table (tests/fixtures/data_tables.json, refusal_messages).
PART_OF_A_TABLE = "Can't change part of a data table."
#: What Range.Table says of a block without a row and a column besides its first, and of input cells it cannot take.
_NOT_A_BLOCK = ("This selection isn't valid. Make sure the copy and paste areas don't overlap unless they are the same "
                "size and shape.")
_NOT_AN_INPUT = "Input cell reference is not valid."
Position = tuple[int, int]


@dataclass(frozen=True)
class DataTable:
    """One data table on a sheet."""

    #: The cells it works out, as a file's ref has them: the block Range.Table was given but its first row and column.
    area: Area
    #: The cell each value along the first row is tried in, and the one each value down the first column is; None for
    #: an input the table has not got.
    row_input: Position | None = None
    column_input: Position | None = None
    #: An input cell taken away, which the table spells #REF!; its position is where it was, as a file keeps it.
    row_deleted: bool = False
    column_deleted: bool = False
    #: The file's ca, which the table keeps through a save; Excel's own saves set it on some tables and not others.
    always: bool = False

    @property
    def block(self) -> Area:
        """The block Range.Table was given: the table with the row above it and the column to its left."""
        return Area(self.area.top - 1, self.area.left - 1, self.area.bottom, self.area.right)

    @property
    def broken(self) -> bool:
        """Whether an input cell was taken away, which makes every cell #REF!."""
        return self.row_deleted or self.column_deleted

    def formula(self, *, r1c1: bool = False) -> str:
        """What its cells say: =TABLE(A2,A1), =TABLE(,A1) or =TABLE(A2,), in R1C1 =TABLE(R2C1,R1C1)."""

        def spell(position: Position | None, deleted: bool) -> str:
            if position is None:
                return ""
            if deleted:
                return "#REF!"
            row, column = position
            return f"R{row}C{column}" if r1c1 else _address(position)

        return f"=TABLE({spell(self.row_input, self.row_deleted)},{spell(self.column_input, self.column_deleted)})"

    def trial(self, row: int, column: int) -> tuple[Position, list[tuple[Position, Position]]]:
        """The cell whose formula one of the table's cells works out, and each input cell with the cell holding the
        value tried in it there."""
        top, left = self.area.top - 1, self.area.left - 1
        tried: list[tuple[Position, Position]] = []
        if self.row_input is not None:
            tried.append((self.row_input, (top, column)))
        if self.column_input is not None:
            tried.append((self.column_input, (row, left)))
        if self.row_input is not None and self.column_input is not None:
            return (top, left), tried
        return ((top, column) if self.column_input is not None else (row, left)), tried

    def element(self, *, always: bool = False) -> str:
        """The <f> a file keeps on the table's first cell, its attributes in Excel's order: r1 is the row input of a
        table with both, r2 its column input, and del1 or del2 marks one taken away."""
        two = self.row_input is not None and self.column_input is not None
        first = self.row_input if self.row_input is not None else self.column_input
        assert first is not None
        attributes = [f'ref="{self.area.address(absolute=False)}"', f'dt2D="{int(two)}"',
                      f'dtr="{int(self.row_input is not None)}"']
        if self.row_deleted if self.row_input is not None else self.column_deleted:
            attributes.append('del1="1"')
        if two and self.column_deleted:
            attributes.append('del2="1"')
        attributes.append(f'r1="{_address(first)}"')
        if two:
            assert self.column_input is not None
            attributes.append(f'r2="{_address(self.column_input)}"')
        if always or self.always:
            attributes.append('ca="1"')
        return f'<f t="dataTable" {" ".join(attributes)}/>'


def _address(position: Position) -> str:
    return f"{column_letter(position[1])}{position[0]}"


def _position(text: str) -> Position | None:
    if not text:
        return None
    found = parse_area(text, sheet="")
    return found.top, found.left


def read(attributes: dict[str, str], reference: str) -> DataTable | None:
    """The table a file keeps on its first cell's <f t="dataTable">; None for one the element does not describe."""

    def flag(name: str) -> bool:
        return attributes.get(name, "0").lower() in ("1", "true")

    try:
        area = parse_area(attributes.get("ref") or reference, sheet="")
        first = _position(attributes.get("r1", ""))
        second = _position(attributes.get("r2", ""))
    except ValueError:
        return None
    if first is None:
        return None
    if flag("dt2D"):
        if second is None:
            return None
        return DataTable(area, first, second, flag("del1"), flag("del2"), flag("ca"))
    if flag("dtr"):
        return DataTable(area, first, None, flag("del1"), False, flag("ca"))
    return DataTable(area, None, first, False, flag("del1"), flag("ca"))


def table_at(sheet: Worksheet, row: int, column: int) -> DataTable | None:
    """The table a cell is one of the worked-out cells of, if any."""
    for table in sheet.data_tables.values():
        if table.area.contains(row, column):
            return table
    return None


def _meets(first: Area, second: Area) -> bool:
    return not (first.bottom < second.top or second.bottom < first.top
                or first.right < second.left or second.right < first.left)


def _covers(outer: Area, inner: Area) -> bool:
    return outer.top <= inner.top and outer.left <= inner.left and inner.bottom <= outer.bottom \
        and inner.right <= outer.right


def met(sheet: Worksheet, areas: list[Area]) -> tuple[list[Position], list[Position]]:
    """The tables whose cells ``areas`` cover whole, and those whose cells they meet in part, each by its first cell."""
    whole: list[Position] = []
    part: list[Position] = []
    for anchor, table in sheet.data_tables.items():
        if any(_meets(area, table.area) for area in areas):
            (whole if any(_covers(area, table.area) for area in areas) else part).append(anchor)
    return whole, part


def blocks(sheet: Worksheet) -> list[Area]:
    """Each table's whole block, its first row and column included."""
    return [table.block for table in sheet.data_tables.values()]


def refuse(sheet: Worksheet, areas: list[Area], doing: str, *, whole: bool = False) -> None:
    """Report an edit the model does not carry through a table's cells, or with ``whole`` through any of its block."""
    found = blocks(sheet) if whole else [table.area for table in sheet.data_tables.values()]
    if any(_meets(area, one) for area in areas for one in found):
        raise VBAUnsupportedError(f"{doing} a data table's cells is not implemented")


def make(target: Range, row_input: object, column_input: object) -> None:
    """Range.Table: a table over each of the target's blocks, its cells 0 until they are worked out (see _calc). A
    block needs a row and a column besides its first; an input cell is one cell of the block's sheet, outside the
    block or at its top left, and two are two cells; a table over part of another table, or of an array formula, is
    error 1004, one over all of it takes its place (tests/fixtures/data_tables.json)."""
    from pyopenvba.apps.excel._protection import enforced

    sheet = target.sheet
    if enforced(sheet) is not None:
        raise VBAUnsupportedError("Range.Table on a protected sheet is not implemented")
    if any(block.rows < 2 or block.columns < 2 for block in target.areas):
        raise error(1004, _NOT_A_BLOCK)
    for block in target.areas:
        inputs = (_input(sheet, block, row_input), _input(sheet, block, column_input))
        if inputs == (None, None) or inputs[0] is not None and inputs[0] == inputs[1]:
            raise error(1004, _NOT_AN_INPUT)
    for block in target.areas:
        _make(sheet, block, _input(sheet, block, row_input), _input(sheet, block, column_input))


def _make(sheet: Worksheet, block: Area, row_input: Position | None, column_input: Position | None) -> None:
    """One table, over one block."""
    from pyopenvba.apps.excel import _arrays

    area = Area(block.top + 1, block.left + 1, block.bottom, block.right)
    whole, part = met(sheet, [area])
    if part:
        raise error(1004, PART_OF_A_TABLE)
    arrays, split = _arrays.met(sheet, [area])
    if split:
        raise error(1004, _arrays.PART_OF_AN_ARRAY)
    if any(_meets(area, spill.area) for spill in sheet.spills.values()) \
            or any(cell.dynamic and area.contains(*position) for position, cell in sheet.cells_.items()):
        raise VBAUnsupportedError("a data table over a dynamic-array formula's cells is not implemented")
    if any(_meets(block, merged) for merged in sheet.merged_areas):
        raise VBAUnsupportedError("a data table over merged cells is not implemented")
    if any(_meets(block, one.area) for one in sheet.tables):
        raise VBAUnsupportedError("a data table over a table's cells is not implemented")
    for anchor in whole:
        del sheet.data_tables[anchor]
    for anchor in arrays:
        del sheet.array_formulas[anchor]
    table = DataTable(area, row_input, column_input)
    sheet.data_tables[(area.top, area.left)] = table
    text = table.formula()
    for row in range(area.top, area.bottom + 1):
        for column in range(area.left, area.right + 1):
            cell = sheet.cell(row, column, create=True)
            assert cell is not None
            cell.formula, cell.value, cell.stale, cell.shared, cell.save_dirty = text, 0.0, True, None, None
            sheet.cell_changed(row, column)


def _input(sheet: Worksheet, block: Area, value: object) -> Position | None:
    """An input cell Range.Table is given: one cell of the table's sheet, outside the block or at its top left."""
    from pyopenvba.apps.excel._model import Range

    if value is MISSING or value is NOTHING:
        return None
    if not isinstance(value, Range) or value.sheet is not sheet or not value.single:
        raise error(1004, _NOT_AN_INPUT)
    position = value.first.top, value.first.left
    if block.contains(*position) and position != (block.top, block.left):
        raise error(1004, _NOT_AN_INPUT)
    return position


def as_values(sheet: Worksheet, cells: dict[Position, Cell], origin: Position) -> None:
    """A copy's cells, ``cells`` by their place from ``origin``, with each of a table's turned into the value it
    shows, as Excel copies one."""
    calculator = sheet.book.calculator
    for (down, across), cell in list(cells.items()):
        row, column = origin[0] + down, origin[1] + across
        if table_at(sheet, row, column) is not None:
            value = calculator.value_of(sheet.name, row, column)
            cells[(down, across)] = replace(cell, value=value, formula="", stale=False, shared=None)


def edit_admitted(sheet: Worksheet, *, rows: bool, start: int, count: int, delete: bool) -> None:
    """Refuse putting in or taking out rows (or columns) through a table's block before anything moves; one taking
    out the whole block takes the table."""
    for table in sheet.data_tables.values():
        block = table.block
        low, high = (block.top, block.bottom) if rows else (block.left, block.right)
        if delete:
            end = start + count - 1
            if start <= high and low <= end and not (start <= low and high <= end):
                raise error(1004, PART_OF_A_TABLE)
        elif low < start <= high:
            raise error(1004, PART_OF_A_TABLE)


def shift_admitted(sheet: Worksheet, area: Area, region: Area, *, delete: bool) -> None:
    """A cell shift may carry a table's whole block or delete it, but not split it."""
    for table in sheet.data_tables.values():
        block = table.block
        if not _meets(block, region):
            continue
        if delete and _covers(area, block):
            continue
        if not _covers(region, block) or delete and _meets(area, block):
            raise error(1004, PART_OF_A_TABLE)


def moved(sheet: Worksheet, rewrite: Callable[[str], str]) -> None:
    """Carry each table and its input cells through an edit, as ``rewrite`` carries a reference: a table whose block
    went is gone, and an input cell that went is #REF!, keeping where it was."""
    carried: dict[Position, DataTable] = {}
    for table in sheet.data_tables.values():
        try:
            block = parse_area(rewrite("=" + table.block.address(absolute=False)).removeprefix("="), sheet="")
        except ValueError:
            continue
        area = Area(block.top + 1, block.left + 1, block.bottom, block.right)
        row_input, row_deleted = _carried(rewrite, table.row_input, table.row_deleted)
        column_input, column_deleted = _carried(rewrite, table.column_input, table.column_deleted)
        table = replace(table, area=area, row_input=row_input, column_input=column_input,
                        row_deleted=row_deleted, column_deleted=column_deleted)
        carried[(area.top, area.left)] = table
        text = table.formula()
        for row in range(area.top, area.bottom + 1):
            for column in range(area.left, area.right + 1):
                cell = sheet.cells_.get((row, column))
                if cell is not None:
                    cell.formula = text
    sheet.data_tables = carried


def _carried(rewrite: Callable[[str], str], position: Position | None, deleted: bool) -> tuple[Position | None, bool]:
    if position is None or deleted:
        return position, deleted
    try:
        found = _position(rewrite("=" + _address(position)).removeprefix("="))
    except ValueError:
        return position, True
    return (position, True) if found is None else (found, False)


def elements(sheet: Worksheet, dirty: Callable[[DataTable], bool]) -> dict[Position, str]:
    """The <f> each table's first cell carries in a file, and none for its other cells; ``dirty`` says which tables
    to mark to be worked out again, as a save in manual calculation marks what it has not worked out."""
    out: dict[Position, str] = {}
    for anchor, table in sheet.data_tables.items():
        area = table.area
        for row in range(area.top, area.bottom + 1):
            for column in range(area.left, area.right + 1):
                out[(row, column)] = ""
        out[anchor] = table.element(always=dirty(table))
    return out
