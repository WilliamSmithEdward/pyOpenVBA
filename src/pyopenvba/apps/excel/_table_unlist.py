"""ListObject.Unlist: the table goes, its cells stay, its style becomes each cell's own format and references to it
become references to its cells (scripts/measure_list_unlist.py, tests/fixtures/list_unlist.json)."""
from __future__ import annotations

from dataclasses import replace
from typing import TYPE_CHECKING, Any

from pyopenvba._a1 import Area, column_letter
from pyopenvba.apps.excel import _styles as S
from pyopenvba.apps.excel._table_style_table import table_styles
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._calc.host import quote_sheet_name
from pyopenvba.formula._parse import THIS_ROW, Token, read_structured
from pyopenvba.formula._structured import area as structured_area
from pyopenvba.formula._structured import rewritten
from pyopenvba.formula._values import ExcelError

if TYPE_CHECKING:
    from pyopenvba.apps.excel._tables import Table

#: TableStyleElements' numbers.
_WHOLE, _HEADER, _TOTAL, _FIRST_COLUMN, _LAST_COLUMN, _ROW_ODD, _ROW_EVEN, _COLUMN_ODD, _COLUMN_EVEN = range(9)
_FIRST_HEADER, _LAST_HEADER, _FIRST_TOTAL, _LAST_TOTAL = range(9, 13)


def _colour(held: list[Any] | None) -> S.Color | None:
    return None if held is None else S.Color(held[0], held[1], tint=held[2])


def _layers(table: Table, row: int, column: int) -> list[tuple[int, Area]]:
    """The elements that reach a cell, each after those it overrides, with the block of cells each covers."""
    whole = table.area
    rows, columns = whole.rows - 1, whole.columns - 1
    header = table.headers and row == 0
    total = table.totals and row == rows
    layers = [(_WHOLE, Area(0, 0, rows, columns))]
    if table.column_stripes and not header and not total:
        top, bottom = (1 if table.headers else 0), rows - (1 if table.totals else 0)
        layers.append((_COLUMN_ODD if column % 2 == 0 else _COLUMN_EVEN, Area(top, column, bottom, column)))
    if table.row_stripes and not header and not total:
        place = row - (1 if table.headers else 0)
        layers.append((_ROW_ODD if place % 2 == 0 else _ROW_EVEN, Area(row, 0, row, columns)))
    if table.first_column and column == 0:
        layers.append((_FIRST_COLUMN, Area(0, 0, rows, 0)))
    if table.last_column and column == columns:
        layers.append((_LAST_COLUMN, Area(0, columns, rows, columns)))
    if header:
        layers.append((_HEADER, Area(0, 0, 0, columns)))
        if column == 0:
            layers.append((_FIRST_HEADER, Area(0, 0, 0, 0)))
        if column == columns:
            layers.append((_LAST_HEADER, Area(0, columns, 0, columns)))
    if total:
        layers.append((_TOTAL, Area(rows, 0, rows, columns)))
        if column == 0:
            layers.append((_FIRST_TOTAL, Area(rows, 0, rows, 0)))
        if column == columns:
            layers.append((_LAST_TOTAL, Area(rows, columns, rows, columns)))
    return layers


def _composed(table: Table, elements: dict[str, list[Any]], row: int, column: int) -> tuple[Any, ...]:
    """A cell's bold, italic, font colour, fill colour and sides from the elements that reach it, and whether they say
    anything of its font and its border. Later ones win, side by side; a block's edges are its outer sides and the
    insides fall between its cells."""
    bold, italic = False, False
    font: S.Color | None = None
    fill: S.Color | None = None
    sides: dict[str, tuple[int, S.Side]] = {}
    said: set[str] = set()
    for strength, (number, block) in enumerate(_layers(table, row, column)):
        element = elements.get(str(number))
        if element is None:
            continue
        said_bold, said_italic, said_font, said_fill, said_sides = element
        if said_bold is not None or said_italic is not None or said_font is not None:
            said.add('font')
        bold = bold if said_bold is None else said_bold
        italic = italic if said_italic is None else said_italic
        font = _colour(said_font) or font
        fill = _colour(said_fill) or fill
        for name, (style, colour) in said_sides.items():
            side = strength, S.Side(style, _colour(colour))
            if name == 'left' and column == block.left:
                sides['left'] = side
            elif name == 'right' and column == block.right:
                sides['right'] = side
            elif name == 'top' and row == block.top:
                sides['top'] = side
            elif name == 'bottom' and row == block.bottom:
                sides['bottom'] = side
            elif name == 'vertical' and column != block.left:
                sides['left'] = side
            elif name == 'horizontal' and row != block.top:
                sides['top'] = side
    if sides:
        said.add('border')
    return bold, italic, font, fill, sides, said


def _baked(table: Table) -> None:
    """Each cell takes the style's format beneath its own. An edge between two cells belongs to the later one, and beats
    what a weaker element says of that cell's side."""
    if not table.style:
        return
    elements = table_styles().get(table.style)
    if elements is None:
        raise VBAUnsupportedError(f'ListObject.Unlist of a table of the style {table.style} is not implemented')
    whole = table.area
    composed = {(row, column): _composed(table, elements, row, column)
                for row in range(whole.rows) for column in range(whole.columns)}
    for (row, column), (_, _, _, _, sides, _) in composed.items():
        if 'bottom' in sides and row < whole.rows - 1:
            _moved(sides, 'bottom', composed[row + 1, column][4], 'top')
        if 'right' in sides and column < whole.columns - 1:
            _moved(sides, 'right', composed[row, column + 1][4], 'left')
    # Cells of one format are one format in the file, so what any of them applies they all do. A double line along a
    # totals row's top counts as a side of the row above too, where the style has sides at the table's left: the 240
    # swept cases show that, not why (tests/fixtures/list_unlist.json).
    boxed = 'left' in elements[str(_WHOLE)][4] if str(_WHOLE) in elements else False
    alike: dict[tuple[object, ...], set[str]] = {}
    for (row, column), (bold, italic, font, fill, sides, said) in composed.items():
        below = composed[row + 1, column][4].get('top') if table.totals and row == whole.rows - 2 else None
        if sides or (boxed and below is not None and below[1].style == 'double'):
            said.add('border')
        final = tuple(sorted((name, side) for name, (_, side) in sides.items()))
        alike.setdefault((bold, italic, font, fill, final), set()).update(said)
    sheet = table.sheet
    normal = sheet.book.stylesheet.default.font
    for (row, column), (bold, italic, font, fill, sides, _) in composed.items():
        place = (whole.top + row, whole.left + column)
        final = tuple(sorted((name, side) for name, (_, side) in sides.items()))
        sheet.restyle(*place, _beneath(normal, sheet.style_at(*place), bold, italic, font, fill, dict(final),
                                       alike[bold, italic, font, fill, final]))


def _moved(sides: dict[str, tuple[int, S.Side]], name: str, others: dict[str, tuple[int, S.Side]],
           other: str) -> None:
    held = sides.pop(name)
    if other not in others or held[0] > others[other][0]:
        others[other] = held


def _beneath(normal: S.Font, own: S.Style, bold: bool, italic: bool, font: S.Color | None, fill: S.Color | None,
             sides: dict[str, S.Side], said: set[str]) -> S.Style:
    """The cell's format with the style's beneath what was set of it."""
    mine = own.font
    made = replace(mine, bold=mine.bold or bold, italic=mine.italic or italic,
                   color=mine.color if mine.color != normal.color or font is None else font)
    colour = fill if fill is not None else None
    filled = own.fill if own.fill.pattern != 'none' or colour is None else S.Fill('solid', colour, colour)
    border = own.border
    border = replace(border, **{name: side for name, side in sides.items() if getattr(border, name) == S.NO_SIDE})
    applied = (own.applied - {'fill'}) | said | ({'fill'} if filled.pattern != 'none' else set[str]())
    return replace(own, font=made, fill=filled, border=border, applied=applied)


def _reference(table: Table, area: Area, here: bool) -> str:
    """Cells as a formula spells them once the table has gone: on its sheet, in fixed columns, in fixed rows unless it
    was the formula's own row."""
    sheet = quote_sheet_name(table.sheet.name)
    mark = '' if here else '$'
    first = f'${column_letter(area.left)}{mark}{area.top}'
    last = f'${column_letter(area.right)}{mark}{area.bottom}'
    return f'{sheet}!{first}' if first == last else f'{sheet}!{first}:{last}'


def _spelled_out(table: Table) -> None:
    """Every formula that names the table names its cells instead."""
    from pyopenvba.apps.excel._tables import rewrite_formulas, shape

    held = shape(table)
    wanted = table.name.lower()

    def changed(row: int) -> Any:
        def one(token: Token) -> str | None:
            if token.kind == 'name' and token.text.lower() == wanted and table.data is not None:
                return _reference(table, table.data, False)
            if token.kind != 'structured':
                return None
            node = read_structured(token.text)
            if node.table.lower() != wanted:
                return None
            try:
                found = structured_area(node, held, row)
            except ExcelError as failed:
                return str(failed)
            return _reference(table, found, node.items == (THIS_ROW,))

        return one

    for sheet in table.sheet.book.sheets_:
        for (row, _), cell in sheet.cells_.items():
            if cell.formula:
                updated = rewritten(cell.formula, changed(row))
                if updated != cell.formula:
                    cell.formula = updated
                    sheet.touched()
    rewrite_formulas(table.sheet.book, lambda text: rewritten(text, changed(0)))


def unlist(table: Table) -> None:
    from pyopenvba.apps.excel._table_filters import deleted
    from pyopenvba.apps.excel._table_sort import guard_edit

    guard_edit(table, 'unlisting')
    _baked(table)
    _spelled_out(table)
    deleted(table)
    table.sheet.tables.remove(table)
    table.sheet.book.calculator.rebuild()
    table.sheet.touched()
