"""Excel tables: Worksheet.ListObjects, the objects under it, and the part a file keeps each table in.

Measured in live Excel (scripts/measure_tables.py, tests/fixtures/tables/).
A table is a block of a sheet: a header row naming its columns, its data
rows, and a totals row it may show under them. ListObjects.Add makes one
over a range:

* with headers (xlYes), the first row's cells name the columns and
  become text, 2020 the text "2020"; a blank one is the first free
  ColumnN and a repeated one takes the first free number after it,
  Name2;
* without (xlNo), a header row is put in above the range, the cells
  under it moving down, and names Column1, Column2 and on;
* left to guess, a first row of text over data that is not all text is
  taken as headers, and otherwise a header row is put in;
* over a range that meets another table it is error 1004.

The new table is the first free TableN in the workbook, TableStyleMedium2
with row stripes and an AutoFilter. Its name is not a defined name. A
file keeps each table in its own part, xl/tables/tableN.xml, which the
sheet names in a tableParts element; a table read from a file keeps its
part as it was until it changes.

A header names its column, and writing over one renames the column
through every formula; the totals row under the data shows and hides as
:func:`show_totals` and :func:`hide_totals` set out
(scripts/measure_structured_references.py, scripts/measure_totals_row.py).
"""

from __future__ import annotations

import re
import uuid
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba._xml import attributes, escape, escape_text, unescape
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._structured import TableShape, from_file, in_file, renamed_columns, renamed_table
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import (EMPTY, ERR_SUBSCRIPT_OUT_OF_RANGE, MISSING, NOTHING, VBADate, VBAInt, error,
                                           to_bool, to_integer, to_text)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook, Worksheet

#: ListObjects.Add's SourceType for a range of the sheet, and its answers to whether the range has headers.
SOURCE_RANGE = 1
GUESS, YES, NO = 0, 1, 2
DEFAULT_STYLE = "TableStyleMedium2"

_NAMESPACES = ('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" '
               'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" '
               'xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"')
_TABLE = re.compile(r"<table\b[^>]*>")
_COLUMN = re.compile(r"<tableColumn\b[^>]*?(?:/>|>.*?</tableColumn>)", re.DOTALL)
_STYLE = re.compile(r"<tableStyleInfo\b[^>]*/>")
_FILTER = re.compile(r"<autoFilter\b[^>]*?(?:/>|>.*?</autoFilter>)", re.DOTALL)
_TOTALS_FORMULA = re.compile(r"<totalsRowFormula\b[^>]*>(.*?)</totalsRowFormula>", re.DOTALL)


@dataclass
class TableColumn:
    name: str
    id: int
    uid: str = ""
    #: What the totals row holds under the column, kept while the row is hidden: a label, or a function by its
    #: name in the part -- sum, count, custom and the rest -- with a custom one's formula.
    totals_label: str = ""
    totals_function: str = ""
    totals_formula: str = ""


@dataclass
class Table:
    """One table of a sheet, as the model holds it."""

    sheet: Worksheet
    id: int
    name: str
    #: The whole table, its header and totals rows included.
    area: Area
    columns: list[TableColumn]
    headers: bool = True
    totals: bool = False
    #: Whether the totals row has ever been shown, which a part says by leaving out totalsRowShown="0".
    totals_shown: bool = False
    style: str = DEFAULT_STYLE
    first_column: bool = False
    last_column: bool = False
    row_stripes: bool = True
    column_stripes: bool = False
    auto_filter: bool = True
    uid: str = ""
    #: The part the table was read from or written to, and the sheet relationship naming it; "" for a new table.
    part: str = ""
    relationship: str = ""
    #: The part as the file had it, kept while nothing about the table changes.
    xml: str = ""
    changed: bool = False
    view: ListObject | None = field(default=None, repr=False)

    @property
    def header_row(self) -> int:
        return self.area.top if self.headers else 0

    @property
    def data(self) -> Area | None:
        """The data rows, or None for a table with none."""
        top = self.area.top + (1 if self.headers else 0)
        bottom = self.area.bottom - (1 if self.totals else 0)
        return Area(top, self.area.left, bottom, self.area.right) if top <= bottom else None

    def rows(self) -> int:
        data = self.data
        return 0 if data is None else data.rows


# --- the file ------------------------------------------------------------------------------------


def read_table(sheet: Worksheet, part: str, relationship: str, xml: str) -> Table:
    """A table as its part spells it."""
    found = _TABLE.search(xml)
    head = attributes(found.group(0)) if found else {}
    style = _STYLE.search(xml)
    shown = attributes(style.group(0)) if style else {}
    columns: list[TableColumn] = []
    for match in _COLUMN.finditer(xml):
        one = attributes(_opening(match.group(0)))
        formula = _TOTALS_FORMULA.search(match.group(0))
        columns.append(TableColumn(
            name=_column_name(one.get("name", "")), id=int(one.get("id", "0") or 0), uid=one.get("xr3:uid", ""),
            totals_label=_column_name(one.get("totalsRowLabel", "")), totals_function=one.get("totalsRowFunction", ""),
            totals_formula=from_file("=" + unescape(formula.group(1))) if formula else ""))
    return Table(
        sheet=sheet, id=int(head.get("id", "0") or 0), name=head.get("displayName") or head.get("name", ""),
        area=parse_area(head.get("ref", "A1"), sheet=""), columns=columns,
        headers=head.get("headerRowCount", "1") != "0", totals=int(head.get("totalsRowCount", "0") or 0) > 0,
        totals_shown=head.get("totalsRowShown", "1") != "0",
        style=shown.get("name", ""), first_column=shown.get("showFirstColumn") == "1",
        last_column=shown.get("showLastColumn") == "1", row_stripes=shown.get("showRowStripes") == "1",
        column_stripes=shown.get("showColumnStripes") == "1", auto_filter=_FILTER.search(xml) is not None,
        uid=head.get("xr:uid", ""), part=part, relationship=relationship, xml=xml)


def _opening(element: str) -> str:
    stop = element.find(">")
    return element if stop < 0 else element[: stop + 1]


def _column_attribute(name: str) -> str:
    """A column's name as a table part spells it: a control character as _xHHHH_, in small letters as Excel
    writes it there, a tab _x0009_ (tests/fixtures/structured_references/)."""
    return re.sub(r"[\x00-\x1f]", lambda found: f"_x{ord(found.group(0)):04x}_", escape(name))


def _column_name(text: str) -> str:
    """A column's name from its table part, each _xHHHH_ the character it stands for."""
    return re.sub(r"_x([0-9A-Fa-f]{4})_", lambda found: chr(int(found.group(1), 16)), text)


def table_xml(table: Table) -> str:
    """The part for a table the model made, as Excel writes one."""
    reference = table.area.address(absolute=False)
    header = "" if table.headers else ' headerRowCount="0"'
    data = table.data
    filtered = Area(table.area.top, table.area.left, data.bottom if data is not None else table.area.top,
                    table.area.right).address(absolute=False)
    auto_filter = f'<autoFilter ref="{filtered}" xr:uid="{table.uid}"/>' if table.auto_filter and table.headers \
        else ""
    names = [one.name for one in all_tables(table.sheet)]
    columns = "".join(_patched_column(f'<tableColumn id="{column.id}" xr3:uid="{column.uid}" name=""/>', column, names)
                      for column in table.columns)
    counts = _totals_counts(table)
    style = (f'<tableStyleInfo name="{escape(table.style)}" showFirstColumn="{int(table.first_column)}" '
             f'showLastColumn="{int(table.last_column)}" showRowStripes="{int(table.row_stripes)}" '
             f'showColumnStripes="{int(table.column_stripes)}"/>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<table {_NAMESPACES} id="{table.id}" xr:uid="{table.uid}" name="{escape(table.name)}" '
            f'displayName="{escape(table.name)}" ref="{reference}"{header}{counts}>{auto_filter}'
            f'<tableColumns count="{len(table.columns)}">{columns}</tableColumns>{style}</table>')


def patched_xml(table: Table) -> str:
    """A table read from a file, its part brought up to date: name, block, AutoFilter and style changed in
    place, everything else -- totals, calculated columns, a query's fields -- as the file had it."""
    xml = table.xml
    found = _TABLE.search(xml)
    if found is None:
        return table_xml(table)
    head = found.group(0)
    for attribute, value in (("name", escape(table.name)), ("displayName", escape(table.name)),
                             ("ref", table.area.address(absolute=False))):
        head = re.sub(rf'(\s{attribute}=")[^"]*(")', lambda match: match.group(1) + value + match.group(2), head,
                      count=1)
    head = re.sub(r'\s(?:totalsRowCount|totalsRowShown)="[^"]*"', "", head)
    after = max((match.end() for name in ("ref", "tableType", "headerRowCount", "insertRow", "insertRowShift")
                 for match in re.finditer(rf'\s{name}="[^"]*"', head)), default=len(head) - 1)
    head = head[:after] + _totals_counts(table) + head[after:]
    xml = xml[: found.start()] + head + xml[found.end():]
    elements = list(_COLUMN.finditer(xml))
    if len(elements) == len(table.columns):
        # Each column under the name it has now and with what its totals row holds.
        names = [one.name for one in all_tables(table.sheet)]
        for element, column in reversed(list(zip(elements, table.columns))):
            xml = xml[: element.start()] + _patched_column(element.group(0), column, names) + xml[element.end():]
    data = table.data
    filtered = Area(table.area.top, table.area.left, data.bottom if data is not None else table.area.top,
                    table.area.right).address(absolute=False)
    existing = _FILTER.search(xml)
    if table.auto_filter and table.headers:
        if existing is not None:
            element = re.sub(r'(\sref=")[^"]*(")', rf"\g<1>{filtered}\2", existing.group(0), count=1)
            xml = xml[: existing.start()] + element + xml[existing.end():]
        else:
            opened = _TABLE.search(xml)
            assert opened is not None
            xml = xml[: opened.end()] + f'<autoFilter ref="{filtered}" xr:uid="{table.uid}"/>' + xml[opened.end():]
    elif existing is not None:
        xml = xml[: existing.start()] + xml[existing.end():]
    style = (f'<tableStyleInfo name="{escape(table.style)}" showFirstColumn="{int(table.first_column)}" '
             f'showLastColumn="{int(table.last_column)}" showRowStripes="{int(table.row_stripes)}" '
             f'showColumnStripes="{int(table.column_stripes)}"/>')
    shown = _STYLE.search(xml)
    if shown is not None:
        return xml[: shown.start()] + style + xml[shown.end():]
    return xml.replace("</table>", style + "</table>", 1)


def _totals_counts(table: Table) -> str:
    """The table element's word on its totals row: shown, never shown, or neither for one shown and hidden again."""
    if table.totals:
        return ' totalsRowCount="1"'
    return "" if table.totals_shown else ' totalsRowShown="0"'


def _patched_column(element: str, column: TableColumn, tables: list[str]) -> str:
    """A tableColumn element with the column's name and what its totals row holds, anything else kept: a label
    or a function as an attribute after the name, and a custom function's formula as the file spells it."""
    opening = _opening(element)
    closed = opening.endswith("/>")
    inside = opening[len("<tableColumn"):-2 if closed else -1]
    body = "" if closed else element[len(opening):element.rindex("</tableColumn>")]
    inside = re.sub(r'\s(?:totalsRowFunction|totalsRowLabel)="[^"]*"', "", inside)
    totals = f' totalsRowFunction="{column.totals_function}"' if column.totals_function else \
        f' totalsRowLabel="{_column_attribute(column.totals_label)}"' if column.totals_label else ""
    inside = re.sub(r'(\sname=")[^"]*(")', lambda match: match.group(1) + _column_attribute(column.name)
                    + match.group(2) + totals, inside, count=1)
    body = _TOTALS_FORMULA.sub("", body)
    if column.totals_function == "custom" and column.totals_formula:
        formula = f"<totalsRowFormula>{escape_text(in_file(column.totals_formula, tables)[1:])}</totalsRowFormula>"
        calculated = body.find("</calculatedColumnFormula>")
        at = calculated + len("</calculatedColumnFormula>") if calculated >= 0 else 0
        body = body[:at] + formula + body[at:]
    return f"<tableColumn{inside}>{body}</tableColumn>" if body else f"<tableColumn{inside}/>"


def _guid() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"


def shape(table: Table) -> TableShape:
    """What a structured reference needs to know of a table."""
    area = table.area
    return TableShape(table.name, Area(area.top, area.left, area.bottom, area.right, table.sheet.name), table.headers,
                      table.totals, tuple(column.name for column in table.columns))


# --- names ---------------------------------------------------------------------------------------


def rewrite_formulas(book: Workbook, change: Callable[[str], str]) -> None:
    """Put every formula of the workbook -- its cells', its defined names' and its tables' totals -- through
    ``change``."""
    for sheet in book.sheets_:
        for cell in sheet.cells_.values():
            if cell.formula:
                updated = change(cell.formula)
                if updated != cell.formula:
                    cell.formula = updated
                    sheet.touched()
        for table in sheet.tables:
            for column in table.columns:
                if column.totals_formula:
                    updated = change(column.totals_formula)
                    if updated != column.totals_formula:
                        column.totals_formula, table.changed = updated, True
    for entry in book.names_.entries:
        updated = change(entry.refers_to)
        if updated != entry.refers_to:
            entry.refers_to = updated
            book.names_.changed = True
    book.calculator.rebuild()


def rename(table: Table, wanted: str) -> None:
    """ListObject.Name: the table's new name, and every formula naming it follows."""
    old = table.name
    table.name = wanted
    table.changed = True
    table.sheet.touched()
    rewrite_formulas(table.sheet.book, lambda text: renamed_table(text, old, wanted))


#: Set while the model writes a table's header or totals cells itself, which is an edit of them too.
_naming: set[int] = set()


def edited(sheet: Worksheet, row: int, column: int) -> None:
    """A cell changed: where it is a table's header, the column takes the name it gives; where it is in the totals
    row, the column totals as it now says."""
    headers_changed(sheet, row, column)
    totals_changed(sheet, row, column)


def headers_changed(sheet: Worksheet, row: int, column: int) -> None:
    """A cell changed: where it is a table's header, the columns take the names its cells now give.

    Measured (tests/fixtures/structured_references/): a header cell names
    its column with the text it shows -- 1.5, TRUE, 1/2/2020 -- and holds
    that text afterwards; one left empty is the first free ColumnN, and of
    two alike the one further left keeps the name and the other is
    numbered, whichever was written. Every formula follows its column.
    """
    table = next((one for one in sheet.tables if one.headers and one.area.top == row
                  and one.area.left <= column <= one.area.right), None)
    if table is None or id(table) in _naming:
        return
    texts = [_cell_text(sheet, row, one) for one in range(table.area.left, table.area.right + 1)]
    name_columns(table, _column_names(texts))


def name_columns(table: Table, names: list[str]) -> None:
    """Give a table's columns these names: in its header cells, as text, and in every formula naming them."""
    renames = {column.name.lower(): name for column, name in zip(table.columns, names) if column.name != name}
    for column, name in zip(table.columns, names):
        column.name = name
    sheet = table.sheet
    if table.headers:
        _naming.add(id(table))
        try:
            for offset, name in enumerate(names):
                cell = sheet.cell(table.area.top, table.area.left + offset, create=True)
                assert cell is not None
                if cell.value != name or cell.formula:
                    cell.value, cell.formula, cell.stale, cell.shared = name, "", False, None
                    sheet.cell_changed(table.area.top, table.area.left + offset)
        finally:
            _naming.discard(id(table))
    if renames:
        table.changed = True
        sheet.touched()
        rewrite_formulas(sheet.book, lambda text: renamed_columns(text, table.name, renames))


#: TotalsCalculation's numbers, each the function's name in a table part, and the SUBTOTAL each function writes.
CALCULATIONS = ("", "sum", "average", "count", "countNums", "min", "max", "stdDev", "var", "custom")
_SUBTOTAL = {"sum": 109, "average": 101, "count": 103, "countNums": 102, "min": 105, "max": 104, "stdDev": 107,
             "var": 110}


def show_totals(table: Table) -> None:
    """ShowTotals = True, as measured (tests/fixtures/tables/totals_row/).

    Cells go in under the table across its columns, those below moving
    down, even where they are empty. The first time the row is shown the
    first column says Total, unless it is the only one or totals already,
    and the last adds itself up, or counts where its values are not all
    numbers. Each column's label, SUBTOTAL or formula is written in.
    """
    from pyopenvba.apps.excel._editing import shift_cells

    sheet, area = table.sheet, table.area
    if not table.totals_shown:
        last, first = table.columns[-1], table.columns[0]
        if not last.totals_function:
            last.totals_function = "sum" if _numbers(table, len(table.columns) - 1) else "count"
        if len(table.columns) > 1 and not first.totals_function and not first.totals_label:
            first.totals_label = "Total"
    below = Area(area.bottom + 1, area.left, area.bottom + 1, area.right, sheet.name)
    shift_cells(Range(sheet, [below]), below, delete=False, vertical=True)
    table.area = Area(area.top, area.left, area.bottom + 1, area.right)
    table.totals = table.totals_shown = table.changed = True
    for index in range(len(table.columns)):
        _write_total(table, index)
    sheet.shape_changed()


def hide_totals(table: Table) -> None:
    """ShowTotals = False: the row's cells are deleted, those below moving up; what each column totals stays for
    the next time it is shown."""
    from pyopenvba.apps.excel._editing import shift_cells

    sheet, area = table.sheet, table.area
    table.area = Area(area.top, area.left, area.bottom - 1, area.right)
    table.totals, table.changed = False, True
    row = Area(area.bottom, area.left, area.bottom, area.right, sheet.name)
    shift_cells(Range(sheet, [row]), row, delete=True, vertical=True)
    sheet.shape_changed()


def _numbers(table: Table, index: int) -> bool:
    """Whether a column's data is all numbers, and there is some."""
    data = table.data
    if data is None:
        return False
    calculator = table.sheet.book.calculator
    values = [calculator.value_of(table.sheet.name, row, data.left + index) for row in range(data.top, data.bottom + 1)]
    present = [value for value in values if value is not EMPTY]
    return bool(present) and all(isinstance(value, (int, float)) and not isinstance(value, bool) for value in present)


def _write_total(table: Table, index: int) -> None:
    """Write a column's totals cell as the column says: its SUBTOTAL, its own formula, its label or nothing."""
    from pyopenvba.formula._parse import Structured
    from pyopenvba.formula._structured import spelled

    sheet, column = table.sheet, table.columns[index]
    row, at = table.area.bottom, table.area.left + index
    formula, value = "", EMPTY
    code = _SUBTOTAL.get(column.totals_function)
    if code is not None:
        formula = f"=SUBTOTAL({code},{spelled(Structured(table=table.name, first=column.name, last=column.name))})"
    elif column.totals_function == "custom":
        formula = column.totals_formula
    elif column.totals_label:
        value = column.totals_label
    cell = sheet.cell(row, at, create=bool(formula) or value is not EMPTY)
    if cell is None:
        return
    _naming.add(id(table))
    try:
        cell.value, cell.formula, cell.stale, cell.shared = value, formula, bool(formula), None
        sheet.cell_changed(row, at)
    finally:
        _naming.discard(id(table))


def totals_changed(sheet: Worksheet, row: int, column: int) -> None:
    """A cell of a totals row written: the column totals as the cell now says. A SUBTOTAL of the column is that
    function, TotalsCalculation reading it back; any other formula is the column's own, custom; a value is its
    label, a number written there becoming its text; an empty cell totals nothing."""
    table = next((one for one in sheet.tables if one.totals and one.area.bottom == row
                  and one.area.left <= column <= one.area.right), None)
    if table is None or id(table) in _naming:
        return
    entry = table.columns[column - table.area.left]
    cell = sheet.cell(row, column)
    entry.totals_label = entry.totals_function = entry.totals_formula = ""
    table.changed = True
    if cell is not None and cell.formula:
        entry.totals_function = _function_of(cell.formula, table, entry) or "custom"
        if entry.totals_function == "custom":
            entry.totals_formula = cell.formula
    elif cell is not None and cell.value is not EMPTY:
        entry.totals_label = _cell_text(sheet, row, column)
        if cell.value != entry.totals_label:
            _naming.add(id(table))
            try:
                cell.value = entry.totals_label
                sheet.cell_changed(row, column)
            finally:
                _naming.discard(id(table))


def _function_of(formula: str, table: Table, column: TableColumn) -> str | None:
    """The totals function a formula is: SUBTOTAL with one of the codes over this column's data, or None."""
    from pyopenvba.formula._parse import Call, FormulaError, Literal, Structured, parse

    try:
        node = parse(formula)
    except FormulaError:
        return None
    if not (isinstance(node, Call) and node.name.upper() == "SUBTOTAL" and len(node.args) == 2):
        return None
    code, reference = node.args
    if not (isinstance(code, Literal) and isinstance(reference, Structured)) or reference.items or reference.span \
            or reference.table.lower() != table.name.lower() or (reference.first or "").lower() != column.name.lower():
        return None
    return next((name for name, number in _SUBTOTAL.items() if code.value == number), None)


def _cell_text(sheet: Worksheet, row: int, column: int) -> str:
    """What a header cell names its column with: the text it shows, a number or a date through its format."""
    from pyopenvba.formula._display import UndisplayableError, shown

    cell = sheet.cell(row, column)
    if cell is None:
        return ""
    value = sheet.book.calculator.value_of(sheet.name, row, column)
    if isinstance(value, VBADate):
        value = value.serial
    if value is EMPTY or value is None:
        return ""
    if isinstance(value, (bool, int, float)):
        try:
            return shown(value, cell.number_format)[0]
        except UndisplayableError:
            pass
    return to_text(value)


def range_areas(sheet: Worksheet, text: str) -> list[Area] | None:
    """What Range reads a table's name or a structured reference as, None for text that is neither.

    Measured: Range("Table1") is the table's data and Range("Table1[Qty]")
    a column of it; this row, a totals row the table does not show, a
    column it has not got, or a table on another sheet is error 1004.
    """
    from pyopenvba.formula._parse import STRUCTURED, THIS_ROW, FormulaError, Structured, read_structured, split_sheet
    from pyopenvba.formula._structured import area
    from pyopenvba.formula._values import ExcelError

    if re.fullmatch(STRUCTURED, text):
        try:
            node = read_structured(text)
        except FormulaError:
            raise error(1004, f"{text!r} is not a reference Excel reads") from None
    else:
        owner, bare = split_sheet(text)
        node = Structured(table=bare, sheet=owner)
    table = next((one for one in all_tables(sheet) if one.name.lower() == node.table.lower()), None)
    if table is None:
        if "[" in text:
            raise error(1004, f"no table for {text!r}")
        return None
    on = table.sheet.name.lower()
    if (node.sheet.lower() != on if node.sheet else table.sheet is not sheet) or THIS_ROW in node.items:
        raise error(1004, f"{text!r} is not a range of this sheet")
    try:
        return [area(node, shape(table), 0)]
    except ExcelError:
        raise error(1004, f"{text!r} names no cells") from None


# --- making one ----------------------------------------------------------------------------------


def all_tables(sheet: Worksheet) -> list[Table]:
    return [table for one in sheet.book.sheets_ for table in one.tables]


def table_at(sheet: Worksheet, row: int, column: int) -> Table | None:
    return next((table for table in sheet.tables if table.area.contains(row, column)), None)


def _free_name(sheet: Worksheet) -> str:
    taken = {table.name.lower() for table in all_tables(sheet)}
    number = 1
    while f"table{number}" in taken:
        number += 1
    return f"Table{number}"


def _column_names(texts: list[str]) -> list[str]:
    """Header texts as Excel makes column names of them: a blank one ColumnN, a repeat numbered after itself."""
    names: list[str] = []
    for text in texts:
        taken = {name.lower() for name in names} | {one.lower() for one in texts if one}
        if not text:
            number = 1
            while f"column{number}" in taken:
                number += 1
            names.append(f"Column{number}")
        elif text.lower() in {name.lower() for name in names}:
            number = 2
            while f"{text}{number}".lower() in taken or f"{text}{number}".lower() in {n.lower() for n in names}:
                number += 1
            names.append(f"{text}{number}")
        else:
            names.append(text)
    return names


def add(sheet: Worksheet, source: Range, headers: int, style: str) -> Table:
    """ListObjects.Add over a range of the sheet."""
    if len(source.areas) != 1:
        raise VBAUnsupportedError("a table over several areas is not implemented")
    area = source.first
    area = Area(area.top, area.left, area.bottom, area.right)
    if any(_meets(table.area, area) for table in sheet.tables):
        raise error(1004, "A table cannot overlap another table.")
    if any(_meets(one, area) for one in [*sheet.array_formulas.values(), *sheet.merged_areas]):
        raise VBAUnsupportedError("a table over array formulas or merged cells is not implemented")
    first_row = [sheet.book.calculator.value_of(sheet.name, area.top, column)
                 for column in range(area.left, area.right + 1)]
    if headers == GUESS:
        rest = [sheet.book.calculator.value_of(sheet.name, row, column)
                for row in range(area.top + 1, area.bottom + 1) for column in range(area.left, area.right + 1)]
        headers = YES if all(isinstance(one, str) and one for one in first_row) \
            and not all(isinstance(one, str) for one in rest) else NO
    if headers == NO:
        from pyopenvba.apps.excel._editing import shift_cells

        # A header row goes in above the range, the cells under it moving down.
        top = Area(area.top, area.left, area.top, area.right, sheet.name)
        shift_cells(Range(sheet, [top]), top, delete=False, vertical=True)
        area = Area(area.top, area.left, area.bottom + 1, area.right)
        names = _column_names([""] * area.columns)
    else:
        names = _column_names([_cell_text(sheet, area.top, column) for column in range(area.left, area.right + 1)])
    for offset, name in enumerate(names):
        # Header cells hold their column's name as text: 2020 becomes "2020".
        cell = sheet.cell(area.top, area.left + offset, create=True)
        assert cell is not None
        if cell.value != name or cell.formula:
            cell.value, cell.formula, cell.stale, cell.shared = name, "", False, None
            sheet.cell_changed(area.top, area.left + offset)
    uid = _guid()
    table = Table(sheet=sheet, id=max((one.id for one in all_tables(sheet)), default=0) + 1, name=_free_name(sheet),
                  area=area, columns=[TableColumn(name, index + 1, _guid()) for index, name in enumerate(names)],
                  style=style, uid=uid, changed=True)
    sheet.tables.append(table)
    sheet.touched()
    return table


def _meets(first: Area, second: Area) -> bool:
    return not (first.bottom < second.top or second.bottom < first.top
                or first.right < second.left or second.right < first.left)


def moved(sheet: Worksheet, rewrite: Callable[[str], str], *, rows: bool) -> None:
    """Carry each table through an insert or a delete of whole rows, as a reference is carried."""
    kept: list[Table] = []
    for table in sheet.tables:
        text = rewrite("=" + table.area.address(absolute=False))
        try:
            area = parse_area(text.removeprefix("="), sheet="")
        except ValueError:
            # Deleting every row of a table takes the table away.
            raise VBAUnsupportedError("deleting all of a table's rows is not implemented") from None
        if (area.rows, area.columns) != (table.area.rows, table.area.columns):
            if not rows or area.columns != table.area.columns:
                raise VBAUnsupportedError("inserting or deleting columns through a table is not implemented")
            table.changed = True
        elif area != table.area:
            table.changed = True
        table.area = area
        kept.append(table)
    sheet.tables = kept


def edit_admitted(sheet: Worksheet, *, rows: bool, start: int, count: int, delete: bool) -> None:
    """Report, before anything moves, an insert or delete of whole rows or columns the model does not carry
    through a table: any through its columns, which Excel makes table columns of, and a delete of its header row
    or of all of it."""
    end = start + count - 1
    for table in sheet.tables:
        area = table.area
        if not rows:
            if (delete and start <= area.right and area.left <= end) or (not delete and area.left < start <= area.right):
                raise VBAUnsupportedError("inserting or deleting columns through a table is not implemented")
        elif delete and start <= area.top <= end:
            raise VBAUnsupportedError("deleting a table's header row is not implemented")


def refuse(sheet: Worksheet, area: Area, doing: str) -> None:
    """Report an edit the model does not carry through tables."""
    if any(_meets(table.area, area) for table in sheet.tables):
        raise VBAUnsupportedError(f"{doing} cells of a table is not implemented")


# --- VBA -----------------------------------------------------------------------------------------


def view(table: Table) -> ListObject:
    """The one ListObject a table answers with, so that Is compares tables."""
    if table.view is None:
        table.view = ListObject(table)
    return table.view


class ListObjects(VBACollection, ExcelObject):
    vba_type_name = "ListObjects"

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def vba_items(self) -> list[object]:
        return [view(table) for table in self.sheet.tables]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for table in self.sheet.tables:
                if table.name.lower() == index.lower():
                    return view(table)
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        return super().vba_lookup(index, items)

    @member
    def Parent(self) -> object:
        return self.sheet

    @method
    def Add(self, SourceType: object = MISSING, Source: object = MISSING, LinkSource: object = MISSING,
            XlListObjectHasHeaders: object = MISSING, Destination: object = MISSING,
            TableStyleName: object = MISSING) -> object:
        kind = SOURCE_RANGE if SourceType is MISSING else int(to_integer(SourceType, "Long"))
        if kind != SOURCE_RANGE or not isinstance(Source, Range):
            raise VBAUnsupportedError("ListObjects.Add from anything but a range of the sheet is not implemented")
        if Source.sheet is not self.sheet:
            raise error(1004, "The table's range must be on this sheet")
        headers = GUESS if XlListObjectHasHeaders is MISSING else int(to_integer(XlListObjectHasHeaders, "Long"))
        style = DEFAULT_STYLE if TableStyleName is MISSING else to_text(TableStyleName)
        return view(add(self.sheet, Source, headers, style))


class ListObject(ExcelObject):
    vba_type_name = "ListObject"

    def __init__(self, table: Table) -> None:
        self.table = table
        self.sheet = table.sheet

    def _range(self, area: Area | None) -> object:
        return NOTHING if area is None else Range(self.sheet, [Area(area.top, area.left, area.bottom, area.right,
                                                                    self.sheet.name)])

    def _changed(self) -> None:
        self.table.changed = True
        self.sheet.touched()

    @member(default=True)
    def _Default(self) -> object:
        return self.table.name

    @member
    def Name(self) -> object:
        return self.table.name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        wanted = to_text(value)
        if not re.fullmatch(r"[A-Za-z_\\][A-Za-z0-9_.\\]*", wanted) or any(
                one is not self.table and one.name.lower() == wanted.lower() for one in all_tables(self.sheet)):
            raise error(1004, "That table name is not valid or is taken")
        rename(self.table, wanted)

    @member
    def DisplayName(self) -> object:
        return self.table.name

    @setter("DisplayName")
    def _set_display_name(self, value: object) -> None:
        self._set_name(value)

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Range(self) -> object:
        return self._range(self.table.area)

    @member
    def DataBodyRange(self) -> object:
        return self._range(self.table.data)

    @member
    def HeaderRowRange(self) -> object:
        area = self.table.area
        return self._range(Area(area.top, area.left, area.top, area.right) if self.table.headers else None)

    @member
    def TotalsRowRange(self) -> object:
        area = self.table.area
        return self._range(Area(area.bottom, area.left, area.bottom, area.right) if self.table.totals else None)

    @member
    def ListColumns(self) -> object:
        return ListColumns(self.table)

    @member
    def ListRows(self) -> object:
        return ListRows(self.table)

    @member
    def ShowTotals(self) -> object:
        return self.table.totals

    @setter("ShowTotals")
    def _set_show_totals(self, value: object) -> None:
        wanted = to_bool(value)
        if wanted != self.table.totals:
            if wanted:
                show_totals(self.table)
            else:
                hide_totals(self.table)

    @member
    def ShowHeaders(self) -> object:
        return self.table.headers

    @setter("ShowHeaders")
    def _set_show_headers(self, value: object) -> None:
        if to_bool(value) != self.table.headers:
            raise VBAUnsupportedError("showing or hiding a table's header row is not implemented")

    @member
    def TableStyle(self) -> object:
        return self.table.style

    @setter("TableStyle")
    def _set_table_style(self, value: object) -> None:
        self.table.style = to_text(value.vba_get("Name") if isinstance(value, ExcelObject) else value)
        self._changed()

    @member
    def ShowAutoFilter(self) -> object:
        return self.table.auto_filter

    @setter("ShowAutoFilter")
    def _set_show_auto_filter(self, value: object) -> None:
        self.table.auto_filter = to_bool(value)
        self._changed()

    @member
    def ShowTableStyleRowStripes(self) -> object:
        return self.table.row_stripes

    @setter("ShowTableStyleRowStripes")
    def _set_row_stripes(self, value: object) -> None:
        self.table.row_stripes = to_bool(value)
        self._changed()

    @member
    def ShowTableStyleColumnStripes(self) -> object:
        return self.table.column_stripes

    @setter("ShowTableStyleColumnStripes")
    def _set_column_stripes(self, value: object) -> None:
        self.table.column_stripes = to_bool(value)
        self._changed()

    @member
    def ShowTableStyleFirstColumn(self) -> object:
        return self.table.first_column

    @setter("ShowTableStyleFirstColumn")
    def _set_first_column(self, value: object) -> None:
        self.table.first_column = to_bool(value)
        self._changed()

    @member
    def ShowTableStyleLastColumn(self) -> object:
        return self.table.last_column

    @setter("ShowTableStyleLastColumn")
    def _set_last_column(self, value: object) -> None:
        self.table.last_column = to_bool(value)
        self._changed()


class ListColumns(VBACollection, ExcelObject):
    vba_type_name = "ListColumns"

    def __init__(self, table: Table) -> None:
        self.table = table
        self.sheet = table.sheet

    def vba_items(self) -> list[object]:
        return [ListColumn(self.table, index) for index in range(1, len(self.table.columns) + 1)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for position, column in enumerate(self.table.columns, 1):
                if column.name.lower() == index.lower():
                    return ListColumn(self.table, position)
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        return super().vba_lookup(index, items)

    @member
    def Parent(self) -> object:
        return view(self.table)


class ListColumn(ExcelObject):
    vba_type_name = "ListColumn"

    def __init__(self, table: Table, index: int) -> None:
        self.table = table
        self.sheet = table.sheet
        self.index = index

    def _column(self) -> int:
        return self.table.area.left + self.index - 1

    @member(default=True)
    def _Default(self) -> object:
        return self.table.columns[self.index - 1].name

    @member
    def Name(self) -> object:
        return self.table.columns[self.index - 1].name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        """Rename the column, its header and every formula naming it with it. Measured: a name another column has,
        in any case, leaves the column as it was without an error; an empty one is the first free ColumnN."""
        wanted = to_text(value)
        names = [column.name for column in self.table.columns]
        if any(index != self.index - 1 and name.lower() == wanted.lower() for index, name in enumerate(names)):
            return
        names[self.index - 1] = wanted
        name_columns(self.table, _column_names(names))

    @member
    def Index(self) -> object:
        return VBAInt(self.index, "Long")

    @member
    def TotalsCalculation(self) -> object:
        function = self.table.columns[self.index - 1].totals_function
        return VBAInt(CALCULATIONS.index(function) if function in CALCULATIONS else 0, "Long")

    @setter("TotalsCalculation")
    def _set_totals_calculation(self, value: object) -> None:
        """What the column's totals row works out, shown or not: a SUBTOTAL, nothing, or custom, which leaves the
        cell empty until a formula is written there (tests/fixtures/tables/totals_row/)."""
        number = int(to_integer(value, "Long"))
        if not 0 <= number < len(CALCULATIONS):
            raise VBAUnsupportedError(f"TotalsCalculation {number} is not implemented")
        column = self.table.columns[self.index - 1]
        column.totals_function, column.totals_label, column.totals_formula = CALCULATIONS[number], "", ""
        self.table.changed = True
        self.sheet.touched()
        if self.table.totals:
            _write_total(self.table, self.index - 1)

    @member
    def Parent(self) -> object:
        return view(self.table)

    @member
    def Range(self) -> object:
        area = self.table.area
        return Range(self.sheet, [Area(area.top, self._column(), area.bottom, self._column(), self.sheet.name)])

    @member
    def DataBodyRange(self) -> object:
        data = self.table.data
        if data is None:
            return NOTHING
        return Range(self.sheet, [Area(data.top, self._column(), data.bottom, self._column(), self.sheet.name)])


class ListRows(VBACollection, ExcelObject):
    vba_type_name = "ListRows"

    def __init__(self, table: Table) -> None:
        self.table = table
        self.sheet = table.sheet

    def vba_items(self) -> list[object]:
        return [ListRow(self.table, index) for index in range(1, self.table.rows() + 1)]

    @member
    def Parent(self) -> object:
        return view(self.table)


class ListRow(ExcelObject):
    vba_type_name = "ListRow"

    def __init__(self, table: Table, index: int) -> None:
        self.table = table
        self.sheet = table.sheet
        self.index = index

    @member
    def Index(self) -> object:
        return VBAInt(self.index, "Long")

    @member
    def Parent(self) -> object:
        return view(self.table)

    @member
    def Range(self) -> object:
        data = self.table.data
        assert data is not None
        row = data.top + self.index - 1
        return Range(self.sheet, [Area(row, data.left, row, data.right, self.sheet.name)])
