"""Range.Subtotal and Range.RemoveSubtotal, as Excel's Data > Subtotal puts subtotal rows into a list and takes them
out again (tests/fixtures/subtotals/, scripts/measure_subtotals.py).

The list is the region around a single cell, or a range of several cells cut to the block the sheet uses; its
first row is the header. Down its data rows, a run of rows whose GroupBy cells show the same text, spaces at either
end and case aside, is a group, and each group gets a row of its own, put in under it, or over it where
SummaryBelowData is xlSummaryAbove: the text the group's first row shows and the function's word, "East Total", in
the GroupBy column and bold, and =SUBTOTAL(the function's number, the group's cells) in each column TotalList
names. A Grand row, "Grand Total", closes the list over every data row. Under the data a run of blank keys is a
group that gets no row; over it every blank row is a group of its own, " Total". Where TotalList totals the GroupBy
column, the labels go in the nearest column to its left that it does not total, or in a column put in before the
list.

The rows go in whole, one at a time as Excel puts them in, each taking the formats of the row it follows, or above
the data of the row it goes before. The first label a call writes, where that row gives it General, takes instead
the number format a macro last met in a cell (Application.recent_format, see Range.note_seen), a leftover of
Excel's that the labels after it do not see. The list is outlined: its data and subtotal rows one level in, each group's data rows one level further, a
Grand row at the header's level; and Outline.SummaryRow takes SummaryBelowData, which defaults to it.

Replace, the default, first takes the list's subtotals out as RemoveSubtotal does: every row of the list a cell of
which holds one SUBTOTAL call and nothing else goes, and the list's rows and columns leave the outline, those in a
group shown again. Replace:=False keeps them and subtotals each run of data rows between them, in the summary place
the outline has whatever SummaryBelowData says, with each new row nearest its group. Under the data a Grand row
goes in there only where the first one there totals another function, with a formula only in the columns where no
Grand row totals its function; above the data none does.

Neither raises Change; Subtotal raises Calculate, RemoveSubtotal only where a formula reads what went.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, column_letter
from pyopenvba.apps.excel import _events, _outline
from pyopenvba.apps.excel._model import Range
from pyopenvba.apps.excel._styles import applying
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._display import UndisplayableError, format_value
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAArray, VBADate, error, to_bool, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

FAILED = "Subtotal method of Range class failed"
REMOVE_FAILED = "RemoveSubtotal method of Range class failed"
HEADER_ONLY = "This can't be applied to the selected range. Select a single cell in a range and try again."

#: Function's XlConsolidationFunction constants and the numbers 1 to 11 the Subtotal dialog lists them by, each to
#: its place in that list.
_FUNCTIONS = {-4157: 1, -4112: 2, -4106: 3, -4136: 4, -4139: 5, -4149: 6, -4113: 7, -4155: 8, -4156: 9, -4164: 10,
              -4165: 11, **{number: number for number in range(1, 12)}}
#: Each function's word on its labels and its SUBTOTAL number, in the dialog's order: Sum, Count, Average, Max, Min,
#: Product, Count Numbers, StdDev, StdDevp, Var and Varp.
_KINDS = (("Total", 9), ("Count", 3), ("Average", 1), ("Max", 4), ("Min", 5), ("Product", 6), ("Count", 2),
          ("StdDev", 7), ("StdDevp", 8), ("Var", 10), ("Varp", 11))
#: The deepest an outline goes, as the file numbers it.
_DEEPEST = 7


# --- the list ------------------------------------------------------------------------------------------------------


def _blank(sheet: Worksheet, row: int, column: int) -> bool:
    cell = sheet.cells_.get((row, column))
    return cell is None or (cell.value is EMPTY and not cell.formula)


def _list(target: Range, failed: str) -> Area:
    """The cells a call works on: the region around a single cell, or the range cut to the block the sheet uses."""
    from pyopenvba.apps.excel._region import current_region

    sheet = target.sheet
    first = target.first
    if target.single:
        area = current_region(sheet, first.top, first.left)
        if area.rows == 1 and area.columns == 1 and _blank(sheet, area.top, area.left):
            raise error(1004, failed)
        return area
    used = sheet.used_bounds()
    if used is None:
        raise error(1004, failed)
    top, left = max(first.top, used[0]), max(first.left, used[1])
    bottom, right = min(first.bottom, used[2]), min(first.right, used[3])
    if top > bottom or left > right:
        raise error(1004, failed)
    return Area(top, left, bottom, right, sheet.name)


def _admitted(sheet: Worksheet) -> None:
    """Refuse a sheet a call cannot work on, as Excel does or as the model does not yet."""
    from pyopenvba.apps.excel._protection import FILTERING, refuse
    from pyopenvba.apps.excel._visible import filtering

    refuse(sheet, FILTERING)
    if filtering(sheet):
        raise VBAUnsupportedError("Subtotals on a filtered sheet are not implemented")
    package = sheet.book.package
    if package is not None and sheet.part_name and package.has(sheet.part_name) \
            and b"<brk " in package.read(sheet.part_name):
        raise VBAUnsupportedError("Subtotals on a sheet with manual page breaks are not implemented")
    _outline.admit_layout(sheet)


def single_subtotal(formula: str) -> bool:
    """Whether a formula is one SUBTOTAL call and nothing else, the formulas RemoveSubtotal takes out."""
    if not formula.upper().startswith("=SUBTOTAL("):
        return False
    depth, quoted = 0, False
    for index in range(len("=SUBTOTAL"), len(formula)):
        char = formula[index]
        if char == '"':
            quoted = not quoted
        elif quoted:
            continue
        elif char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
            if depth == 0:
                return index == len(formula) - 1
    return False


def _function_number(sheet: Worksheet, row: int, column: int) -> int | None:
    """The function number a cell's single SUBTOTAL call totals with."""
    cell = sheet.cells_.get((row, column))
    if cell is None or not single_subtotal(cell.formula):
        return None
    head = cell.formula[len("=SUBTOTAL("):].split(",", 1)[0].strip()
    return int(head) if head.isdigit() else None


def _subtotal_row(sheet: Worksheet, row: int, area: Area) -> bool:
    return any(single_subtotal(cell.formula) for column in range(area.left, area.right + 1)
               if (cell := sheet.cells_.get((row, column))) is not None and cell.formula)


def _key(sheet: Worksheet, row: int, column: int) -> str:
    """What a GroupBy cell shows, which its group's label starts with."""
    cell = sheet.cells_.get((row, column))
    if cell is None:
        return ""
    value = sheet.book.calculator.value_of(sheet.name, row, column) if cell.formula else cell.value
    if isinstance(value, VBADate):
        value = value.serial
    if value is EMPTY:
        return ""
    if isinstance(value, ExcelError):
        return value.name
    try:
        return format_value(value, sheet.style_at(row, column).number_format, epoch_1904=sheet.book.epoch_1904)
    except UndisplayableError:
        raise VBAUnsupportedError("Subtotals over a value its format cannot show are not implemented") from None


def _joined(text: str) -> str:
    """A key as rows join a group by it: spaces off either end, in capitals."""
    return "".join(char.upper() if len(char.upper()) == 1 else char for char in text.strip(" "))


@dataclass
class _Group:
    first: int
    last: int
    #: What the group's first GroupBy cell shows, and the same spaces off and in capitals, by which rows join it.
    text: str
    key: str


def _groups(sheet: Worksheet, rows: list[int], column: int, below: bool) -> list[_Group]:
    """Runs of rows that show the same text in ``column``, in rows that run on one from the next. Under the data a
    run of blank keys is one group, which gets no row; over it each blank row is a group of its own."""
    groups: list[_Group] = []
    for row in rows:
        text = _key(sheet, row, column)
        key = _joined(text)
        last = groups[-1] if groups and groups[-1].last == row - 1 else None
        if last is not None and last.key == key and (key or below):
            last.last = row
        else:
            groups.append(_Group(row, row, text, key))
    return groups


# --- arguments ---------------------------------------------------------------------------------------------------


def _whole_number(value: object) -> int:
    return int(to_integer(value, "Long"))


def _totals(value: object) -> list[int]:
    """TotalList's columns, counted from 1: an array of them, or one; an empty array is a type mismatch."""
    if isinstance(value, VBAArray):
        items = list(value.elements())
        if not items:
            raise error(13)
        return [_whole_number(item) for item in items]
    return [_whole_number(value)]


def _below(value: object) -> bool | None:
    """SummaryBelowData: 0 above, any other number below, None where left out; text is refused."""
    if value is MISSING:
        return None
    if isinstance(value, str):
        raise error(1004, FAILED)
    return _whole_number(value) != 0


def _label_column(group_by: int, totals: set[int]) -> int | None:
    """The column the labels go in, counted from the list's first: the GroupBy column where TotalList does not
    total it, else the nearest to its left that it does not; None where it totals every one."""
    if group_by not in totals:
        return group_by
    return next((column for column in range(group_by - 1, 0, -1) if column not in totals), None)


# --- the work ----------------------------------------------------------------------------------------------------


def _line(sheet: Worksheet, rows: bool, start: int, *, delete: bool, below: bool = False) -> None:
    """One whole row or column in or out, the outline marks moving with the lines while the caller lays it out."""
    from pyopenvba.apps.excel._editing import edit

    area = Area(start, 1, start, MAX_COLUMNS, sheet.name) if rows else Area(1, start, MAX_ROWS, start, sheet.name)
    edit(Range(sheet, [area], whole="rows" if rows else "columns"), delete=delete, below=below, laid=True)


def _remove(sheet: Worksheet, area: Area) -> Area:
    """Take the list's subtotal rows out and its rows and columns out of the outline; the list that is left."""
    gone = [row for row in range(area.top, area.bottom + 1) if _subtotal_row(sheet, row, area)]
    grouped = [row for row in range(area.top, area.bottom + 1) if _outline.depth(sheet, True, row)]
    for row in reversed(gone):
        _line(sheet, True, row, delete=True)

    def now(row: int) -> int:
        return row - sum(1 for one in gone if one < row)

    rows = {now(row) for row in grouped if row not in gone}
    columns = {column for column in range(area.left, area.right + 1) if _outline.depth(sheet, False, column)}
    if _outline.relevel(sheet, dict.fromkeys(rows, 0), dict.fromkeys(columns, 0), rows, columns):
        sheet.book.calculator.invalidate_visibility()
    return Area(area.top, area.left, max(area.top, area.bottom - len(gone)), area.right, sheet.name)


def remove_subtotal(target: Range) -> bool:
    """RemoveSubtotal: the subtotal rows out of the list around or in the range, and the list out of the outline."""
    sheet = target.sheet
    _admitted(sheet)
    _remove(sheet, _list(target, REMOVE_FAILED))
    sheet.touched()
    _events.quietly_edited(sheet.book)
    return True


def subtotal(target: Range, group_by: object, function: object, total_list: object, replace_them: object,
             page_breaks: object, summary_below_data: object) -> bool:
    """Range.Subtotal over the list around or in the range."""
    sheet, book = target.sheet, target.sheet.book
    if group_by is MISSING or function is MISSING or total_list is MISSING:
        raise error(449)
    _admitted(sheet)
    totals = _totals(total_list)
    area = _list(target, FAILED)
    if area.rows < 2:
        raise error(1004, HEADER_ONLY)
    column_of = _whole_number(group_by)
    kind = _FUNCTIONS.get(_whole_number(function))
    wanted = _below(summary_below_data)
    replacing = True if replace_them is MISSING else to_bool(replace_them)
    width = area.columns
    if kind is None or not 1 <= column_of <= width or any(not 1 <= column <= width for column in totals):
        raise error(1004, FAILED)
    if any(table.area.top <= area.bottom and area.top <= table.area.bottom and table.area.left <= area.right
           and area.left <= table.area.right for table in sheet.tables):
        raise error(1004, FAILED)
    if page_breaks is not MISSING and to_bool(page_breaks):
        raise VBAUnsupportedError("Subtotal's page breaks between groups are not implemented")
    current = _outline.summary_below(sheet)
    nested = not replacing and any(_subtotal_row(sheet, row, area) for row in range(area.top + 1, area.bottom + 1))
    below = current if nested or wanted is None else wanted
    if replacing:
        area = _remove(sheet, area)
    at = _label_column(column_of, set(totals))
    if at is None:
        # Every column left of the GroupBy column is totalled: the labels get a column of their own before the list.
        _line(sheet, False, area.left, delete=False)
        area = Area(area.top, area.left + 1, area.bottom, area.right + 1, sheet.name)
    word, number = _KINDS[kind - 1]
    layout = _Layout(sheet, area, area.left + column_of - 1, area.left - 1 if at is None else area.left + at - 1,
                     sorted({area.left + column - 1 for column in totals}), word, number, below=below, nested=nested)
    layout.groups_in()
    layout.grand_in()
    layout.written()
    layout.outlined()
    if not nested and below != current:
        _outline.set_summary_below(sheet, below)
    book.application.recent_format = "General"
    sheet.touched()
    _events.quietly_edited(book)
    return True


@dataclass
class _Line:
    """One row of the list as Subtotal lays it out, in order from the header."""

    #: "header", "data", "old" for a subtotal row already there, "new" for a group's row put in, "grand".
    kind: str
    #: For a data row, whether its group got a row.
    grouped: bool = False
    #: For a group's row put in, the depth its group's data rows had.
    depth: int = 0


class _Layout:
    """A list being subtotalled: its rows as they now stand, and what each is."""

    def __init__(self, sheet: Worksheet, area: Area, key: int, label: int, columns: list[int], word: str,
                 number: int, *, below: bool, nested: bool) -> None:
        self.sheet, self.area, self.key, self.label, self.columns = sheet, area, key, label, columns
        self.word, self.number, self.below, self.nested = word, number, below, nested
        self.lines = [_Line("header")] + [_Line("old" if _subtotal_row(sheet, row, area) else "data")
                                          for row in range(area.top + 1, area.bottom + 1)]
        #: Whether no label has been written yet, the first of which takes the leftover number format.
        self.first = True

    def _row(self, index: int) -> int:
        return self.area.top + index

    def _depth(self, row: int) -> int:
        return _outline.depth(self.sheet, True, row)

    def _put(self, index: int, kind: str, *, depth: int = 0) -> int:
        """A row put in at a line of the list, taking the formats of the row before it, or above the data of the one
        after it; its row."""
        row = self._row(index)
        _line(self.sheet, True, row, delete=False, below=not self.below)
        self.lines.insert(index, _Line(kind, depth=depth))
        return row

    def _write_label(self, row: int, text: str) -> None:
        """A label, bold. The first a call writes takes the leftover format, unless both it and the format its row
        gave the label are other than General (tests/fixtures/subtotals/, register_* and values_narrow)."""
        sheet, column = self.sheet, self.label
        style = sheet.style_at(row, column)
        recent = sheet.book.application.recent_format
        if self.first and "General" in (style.number_format, recent) and style.number_format != recent:
            style = applying(style, "number_format", number_format=recent)
        self.first = False
        sheet.restyle(row, column, applying(style, "font", font=replace(style.font, bold=True)))
        cell = sheet.cell(row, column, create=True)
        assert cell is not None
        cell.value, cell.formula, cell.stale, cell.shared, cell.dynamic, cell.spilled_from = (
            text, "", False, None, False, None)

    def _write_total(self, row: int, column: int, first: int, last: int) -> None:
        """=SUBTOTAL over the rows of a column, written as Subtotal writes it, with no number format it brings."""
        letter = column_letter(column)
        cell = self.sheet.cell(row, column, create=True)
        assert cell is not None
        cell.value, cell.formula, cell.stale, cell.save_dirty, cell.shared, cell.dynamic, cell.spilled_from = (
            EMPTY, f"=SUBTOTAL({self.number},{letter}{first}:{letter}{last})", True, None, None, False, None)

    def written(self) -> None:
        """Tell the calculator about the cells written, once every row is in: each row put in reads every formula
        of the workbook again, which a cell written in between would have it build for nothing."""
        sheet = self.sheet
        for index, line in enumerate(self.lines):
            if line.kind not in ("new", "grand"):
                continue
            row = self._row(index)
            for column in (self.label, *self.columns):
                cell = sheet.cells_.get((row, column))
                if cell is None:
                    continue
                sheet.cell_changed(row, column)
                if cell.formula:
                    sheet.book.calculator.entered(sheet.name, row, column)

    def groups_in(self) -> None:
        """A row for each group: under the data top down, after its last row; above it bottom up, before its first."""
        rows = [self._row(index) for index, line in enumerate(self.lines) if line.kind == "data"]
        runs: list[list[int]] = []
        for row in rows:
            if runs and runs[-1][-1] == row - 1:
                runs[-1].append(row)
            else:
                runs.append([row])
        groups = [group for run in runs for group in _groups(self.sheet, run, self.key, self.below)]
        if self.below:
            added = 0
            for group in groups:
                first, last = group.first + added, group.last + added
                if not group.key:
                    continue
                row = self._put(last + 1 - self.area.top, "new", depth=self._depth(last))
                for data in range(first, last + 1):
                    self.lines[data - self.area.top].grouped = True
                self._write(row, f"{group.text} {self.word}", first, last)
                added += 1
            return
        for group in reversed(groups):
            row = self._put(group.first - self.area.top, "new", depth=self._depth(group.first))
            for data in range(group.first + 1, group.last + 2):
                self.lines[data - self.area.top].grouped = True
            self._write(row, f"{group.text} {self.word}", group.first + 1, group.last + 1)

    def _write(self, row: int, text: str, first: int, last: int, *, skip: frozenset[int] = frozenset()) -> None:
        self._write_label(row, text)
        for column in self.columns:
            if column not in skip:
                self._write_total(row, column, first, last)

    def grand_in(self) -> None:
        """The Grand row over every data row: after the list, or above the data before it. With other subtotals kept,
        only under the data, and only where the first Grand row already there totals another function; a formula
        only in a column where none of them totals this one."""
        if self.nested and not self.below:
            return
        end = len(self.lines)
        if self.nested:
            header = self._depth(self.area.top)
            while end > 1 and self.lines[end - 1].kind == "old" and self._depth(self._row(end - 1)) == header:
                end -= 1
        kept = [self._row(index) for index in range(end, len(self.lines))]
        if kept and any(_function_number(self.sheet, kept[0], column) == self.number
                        for column in range(self.area.left, self.area.right + 1)):
            return
        row = self._put(end if self.below else 1, "grand")
        skip = frozenset(column for column in self.columns
                         if any(_function_number(self.sheet, one + 1, column) == self.number for one in kept))
        data = [self._row(index) for index, line in enumerate(self.lines) if line.kind == "data"]
        if data:
            self._write(row, f"Grand {self.word}", data[0], data[-1], skip=skip)
        else:
            self._write_label(row, f"Grand {self.word}")

    def outlined(self) -> None:
        """The list's rows at their outline depths: a group's data rows one deeper than its row; with the list
        subtotalled afresh, every row between the header and the Grand row one deeper again, the Grand row too
        under data no group's row was put in for."""
        lines = self.lines
        header = self._depth(self.area.top)
        inner = range(0)
        if not self.nested:
            grand = next(index for index, line in enumerate(lines) if line.kind == "grand")
            if not self.below:
                inner = range(grand + 1, len(lines))
            elif any(line.kind == "new" for line in lines):
                inner = range(1, grand)
            else:
                inner = range(1, len(lines))
        depths: dict[int, int] = {}
        for index, line in enumerate(lines):
            row = self._row(index)
            if line.kind == "header":
                continue
            base = header if line.kind == "grand" else line.depth if line.kind == "new" else self._depth(row)
            depth = base + (index in inner) + (line.kind == "data" and line.grouped)
            if depth > _DEEPEST:
                raise VBAUnsupportedError("Subtotals deeper than the outline's eight levels are not implemented")
            depths[row] = depth
        # A row is collapsed where every row of the deeper run on its detail side is hidden, as the outline now runs:
        # a hidden group's mark moves to the subtotal row put in nearest it, and one hidden row of a group shown
        # marks nothing (tests/fixtures/subtotals/, collapsed_then_nested and hidden_row).
        dims = self.sheet.dims
        step = -1 if self.below else 1
        folded: dict[int, bool] = {}
        for row, depth in depths.items():
            detail = row + step
            run: list[int] = []
            while detail in depths and depths[detail] > depth:
                run.append(detail)
                detail += step
            folded[row] = bool(run) and all(dims.row_hidden(one) for one in run)
        _outline.relevel(self.sheet, depths, {}, set(), set(), folded=folded)
