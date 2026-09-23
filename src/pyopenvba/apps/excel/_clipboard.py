"""Copy and Cut without a destination, Worksheet.Paste, Range.PasteSpecial, and Cut with one.

Measured in live Excel (scripts/measure_paste.py):

- Copy and Cut put the range on the clipboard, and CutCopyMode reads 1
  or 2. What is pasted is what the cells hold when it is pasted, not when
  they were copied, and writing to a cell does not end the copy. A copy
  lasts through any number of pastes; a cut ends with its paste, and
  cannot be pasted special. Pasting with nothing copied is error 1004.
- A paste fills the destination with copies of the source when it is a
  whole number of source sizes, and otherwise takes the source's size
  from its top-left cell.
- PasteSpecial takes along everything, the values the cells show, their
  formulas, their formats, number formats with values or formulas, all
  but borders, or column widths alone. A blank source cell pastes as a
  blank, which clears what is under it, unless blanks are skipped.
- An operation combines what is pasted with what is under it. Blanks on
  either side count as 0, and text on either side leaves the cell as it
  was. Two numbers give a number, division by 0 #DIV/0!; a formula on
  either side gives a formula: 10 under =D1*2 multiplied is =10*(D1*2),
  =5+5 under 1 added is =(5+5)+1.
- Transposed, a formula's references that are relative in both row and
  column keep their distance with the rows and columns swapped; the
  rest stay as they are.
- Cut moves the cells, formats and all, and leaves blanks behind. Every
  reference in the workbook that lies wholly inside the moved block,
  names included, follows it, to another sheet with the sheet's name;
  one wholly inside the cells the block lands on becomes #REF!.
  (scripts/measure_cut_references.py, 24 layouts:) one that loses a
  whole edge to the block keeps the rest when the block goes to another
  sheet, and on its own sheet takes the edge along when the block moves
  straight along it; any other overlap stays as it is.

Measured in live Excel (scripts/measure_autofilter_edits.py):

- A copy of several areas -- a filtered range's visible cells among
  them -- closes them up, one under another when they share their
  columns and side by side when they share their rows; areas that share
  neither are error 1004. Copy with a destination and Worksheet.Paste
  paste their formulas as the values they show; PasteSpecial keeps a
  formula, moved as far as its own cell was.
- Cutting a filter's whole range moves the filter with it on its sheet,
  its criteria cleared and its rows shown; to another sheet the filter
  goes, its rows shown, and its hidden name follows the cells.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, column_letter, column_number, quote_sheet
from pyopenvba.apps.excel._protection import whole
from pyopenvba.apps.excel._visible import unmeasured, visible, visible_areas
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import shift_text, split_sheet, tokenize, transpose_text
from pyopenvba.formula._values import ERRORS, ExcelError, number_text
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAInt, error, to_bool, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Cell, Range, Worksheet
    from pyopenvba.apps.excel._styles import Style

ALL, VALUES, FORMULAS, FORMATS = -4104, -4163, -4123, -4122
ALL_EXCEPT_BORDERS, COLUMN_WIDTHS, FORMULAS_AND_NUMBER_FORMATS, VALUES_AND_NUMBER_FORMATS = 7, 8, 11, 12
#: Pastes that take everything: the theme and merging variants are the same thing to the model.
_EVERYTHING = {ALL, ALL_EXCEPT_BORDERS, 13, 14}
_CONTENT_FORMULAS = _EVERYTHING | {FORMULAS, FORMULAS_AND_NUMBER_FORMATS}
_CONTENT_VALUES = {VALUES, VALUES_AND_NUMBER_FORMATS}
_UNMODELLED = {-4144: "comments", 6: "data validation"}
_OPERATIONS = {-4142: "", 2: "+", 3: "-", 4: "*", 5: "/"}


@dataclass(frozen=True)
class Block:
    """What a copy covers: its areas, closed up into one block, each with the row and column it starts at."""

    parts: tuple[tuple[Area, int, int], ...]
    rows: int
    columns: int

    def source(self, down: int, across: int) -> tuple[int, int]:
        """The cell a position of the block comes from."""
        for area, top, left in self.parts:
            if top <= down < top + area.rows and left <= across < left + area.columns:
                return area.top + down - top, area.left + across - left
        raise AssertionError("a position outside the block")

    def covers(self, row: int, column: int) -> bool:
        return any(area.contains(row, column) for area, _, _ in self.parts)


def block_of(areas: list[Area]) -> Block:
    """The block copied areas close up into: one under another when they share their columns, side by side
    when they share their rows, and error 1004 when they share neither."""
    first = areas[0]
    if len(areas) == 1:
        return Block(((first, 0, 0),), first.rows, first.columns)
    stacked = all((area.left, area.right) == (first.left, first.right) for area in areas)
    beside = all((area.top, area.bottom) == (first.top, first.bottom) for area in areas)
    if not stacked and not beside:
        raise error(1004, "That command cannot be used on multiple selections")
    starts = [area.top if stacked else area.left for area in areas]
    ends = [area.bottom if stacked else area.right for area in areas]
    if any(start <= end for end, start in zip(ends, starts[1:])):
        raise VBAUnsupportedError("copying areas out of order, or overlapping, is not implemented")
    parts: list[tuple[Area, int, int]] = []
    offset = 0
    for area in areas:
        parts.append((area, offset, 0) if stacked else (area, 0, offset))
        offset += area.rows if stacked else area.columns
    return Block(tuple(parts), offset if stacked else first.rows, first.columns if stacked else offset)


@dataclass
class Clip:
    """What Copy or Cut put on the clipboard: the range itself, read when it is pasted."""

    source: Range
    block: Block
    cut: bool

    @property
    def area(self) -> Area:
        """What a cut moves: its one area."""
        return self.block.parts[0][0]


def mode(clip: Clip | None) -> VBAInt:
    return VBAInt(0 if clip is None else 2 if clip.cut else 1, "Long")


def copy(target: Range) -> object:
    """Range.Copy with no destination: on a filtered sheet the visible cells, as _visible has it."""
    source = visible(target)
    target.sheet.book.application.clipboard = Clip(source, block_of(source.areas), cut=False)
    return True


def copy_range(target: Range, destination: object) -> object:
    """Range.Copy with a destination: on a filtered sheet, the visible cells into the visible cells.

    A destination with hidden cells takes the copy in each visible block,
    one after the other, a later one pasting over what an earlier one ran into.
    """
    from pyopenvba.apps.excel._model import Range

    if not isinstance(destination, Range):
        raise error(1004, "Copy needs a range to copy to")
    source = visible(target)
    made = block_of(source.areas)
    whole(destination.sheet, [landing_of(made, destination)], "Copying")
    landing = visible_areas(destination)
    if landing is None:
        return paste_block(target.sheet, made, destination)
    first = made.parts[0][0]
    if len(made.parts) > 1:
        raise VBAUnsupportedError("copying several areas into a filtered range is not implemented")
    for area in landing:
        Range(target.sheet, [first]).copy_to(Range(destination.sheet, [area]))
    return True


def landing_of(made: Block, destination: Range) -> Area:
    """The cells a paste lands on: whole copies tiled over the destination when it holds them, else one at its corner."""
    corner = destination.first
    tiled = corner.rows % made.rows == 0 and corner.columns % made.columns == 0
    height, width = (corner.rows, corner.columns) if tiled else (made.rows, made.columns)
    return Area(corner.top, corner.left, min(corner.top + height - 1, MAX_ROWS),
                min(corner.left + width - 1, MAX_COLUMNS), destination.sheet.name)


def paste_block(sheet: Worksheet, made: Block, destination: Range) -> object:
    """A copy pasted as Copy and Paste paste it: one area as a copy of it, several closed up at the destination's
    top-left with their formulas turned into the values they show."""
    from pyopenvba.apps.excel import _merges
    from pyopenvba.apps.excel._model import Range

    if len(made.parts) == 1:
        return Range(sheet, [made.parts[0][0]]).copy_to(destination)
    corner = destination.first
    if len(destination.areas) != 1 or not (destination.single or (corner.rows, corner.columns) == (made.rows,
                                                                                                    made.columns)):
        raise VBAUnsupportedError("pasting several copied areas over a range of another size is not implemented")
    landing = Area(corner.top, corner.left, corner.top + made.rows - 1, corner.left + made.columns - 1)
    if landing.bottom > MAX_ROWS or landing.right > MAX_COLUMNS:
        raise error(1004, "The paste would run past the edge of the worksheet")
    if destination.sheet is sheet and any(_merges.intersects(landing, area) for area, _, _ in made.parts):
        raise VBAUnsupportedError("pasting copied areas over themselves is not implemented")
    # The values first: a formula shows what it gives where it is.
    calculator = sheet.book.calculator
    values = {position: calculator.value_of(sheet.name, *position) for position, cell in sheet.cells_.items()
              if cell.formula and made.covers(*position)}
    target = destination.sheet
    for area, down, across in made.parts:
        top, left = corner.top + down, corner.left + across
        Range(sheet, [area]).copy_to(Range(target, [Area(top, left, top + area.rows - 1, left + area.columns - 1)]))
        for (row, column), value in values.items():
            if not area.contains(row, column):
                continue
            at = (row - area.top + top, column - area.left + left)
            cell = target.cell(*at)
            if cell is not None:
                cell.formula, cell.value, cell.stale = "", value, False
                target.settle(*at)
                target.cell_changed(*at)
    return True


def cut(target: Range, destination: object) -> object:
    """Range.Cut: onto the clipboard, or straight to ``destination``."""
    from pyopenvba.apps.excel._model import Range

    if len(target.areas) != 1:
        raise error(1004, "That command cannot be used on multiple selections")
    if destination is MISSING:
        target.sheet.book.application.clipboard = Clip(target, block_of([target.first]), cut=True)
        return True
    if not isinstance(destination, Range):
        raise error(1004, "Cut needs a range to move to")
    whole(target.sheet, [target.first], "Cut")
    whole(destination.sheet, [landing_of(block_of([target.first]), destination)], "Cut")
    move(target, target.first, destination)
    return True


def _clip(target: Range, empty: str = "There is nothing to paste") -> Clip:
    clip = target.sheet.book.application.clipboard
    if clip is None:
        raise error(1004, empty)
    return clip


def paste(sheet: Worksheet, destination: object, link: object) -> object:
    """Worksheet.Paste: what was copied or cut, at ``destination`` or the selection."""
    from pyopenvba.apps.excel._model import Range

    if link is not MISSING and to_bool(link):
        raise VBAUnsupportedError("Worksheet.Paste Link:=True is not implemented")
    if destination is MISSING:
        destination = sheet.selection
        if not isinstance(destination, Range):
            raise error(1004, "Paste needs a selection on this sheet")
    if not isinstance(destination, Range):
        raise error(1004, "Paste needs a range to paste into")
    clip = _clip(destination)
    if clip.cut:
        whole(clip.source.sheet, [clip.area], "Cut")
    whole(destination.sheet, [landing_of(clip.block, destination)], "Pasting")
    if clip.cut:
        move(clip.source, clip.area, destination)
        return True
    unmeasured(destination, "Worksheet.Paste")
    return paste_block(clip.source.sheet, clip.block, destination)


def paste_special(target: Range, what: object, operation: object, skip_blanks: object, transpose: object) -> object:
    """Range.PasteSpecial."""
    # Measured with the clipboard emptied, as Protect empties it.
    clip = _clip(target, "PasteSpecial method of Range class failed")
    if clip.cut:
        raise error(1004, "PasteSpecial cannot paste what was cut")
    kind = ALL if what is MISSING else int(to_integer(what, "Long"))
    if kind in _UNMODELLED:
        raise VBAUnsupportedError(f"PasteSpecial of {_UNMODELLED[kind]} is not implemented")
    if kind not in _EVERYTHING | _CONTENT_FORMULAS | _CONTENT_VALUES | {FORMATS, COLUMN_WIDTHS}:
        raise error(1004, "PasteSpecial has no such paste type")
    op = _OPERATIONS.get(-4142 if operation is MISSING else int(to_integer(operation, "Long")))
    if op is None:
        raise error(1004, "PasteSpecial has no such operation")
    skipping = skip_blanks is not MISSING and to_bool(skip_blanks)
    turned = transpose is not MISSING and to_bool(transpose)
    source_sheet, made = clip.source.sheet, clip.block
    if kind == COLUMN_WIDTHS:
        _column_widths(source_sheet, made, target)
        return True
    rows, columns = (made.columns, made.rows) if turned else (made.rows, made.columns)
    corner = target.first
    tiled = corner.rows % rows == 0 and corner.columns % columns == 0
    height, width = (corner.rows, corner.columns) if tiled else (rows, columns)
    if corner.top + height - 1 > MAX_ROWS or corner.left + width - 1 > MAX_COLUMNS:
        raise error(1004, "The paste would run past the edge of the worksheet")
    sheet = target.sheet
    whole(sheet, [Area(corner.top, corner.left, corner.top + height - 1, corner.left + width - 1, sheet.name)],
          "PasteSpecial")
    if (kind in _CONTENT_FORMULAS or op) and sheet.book is not source_sheet.book:
        raise VBAUnsupportedError("PasteSpecial of formulas into another workbook is not implemented")
    source_cells = {position: _snapshot(cell) for position, cell in source_sheet.cells_.items()
                    if made.covers(*position)}
    values = {position: source_sheet.book.calculator.value_of(source_sheet.name, *position)
              for position, cell in source_cells.items() if cell.formula} if kind in _CONTENT_VALUES else {}
    from pyopenvba.apps.excel import _merges

    for down in range(height):
        for across in range(width):
            row, column = corner.top + down, corner.left + across
            if not _merges.writable(sheet, row, column):
                continue
            inner_down, inner_across = down % rows, across % columns
            source = made.source(inner_across, inner_down) if turned else made.source(inner_down, inner_across)
            cell = source_cells.get(source)
            blank = cell is None or (cell.value is EMPTY and not cell.formula)
            if blank and skipping:
                continue
            _paste_one(target, kind, op, (row, column), source, cell, values.get(source), turned)
    return True


def _snapshot(cell: Cell) -> Cell:
    from pyopenvba.apps.excel._model import Cell

    return Cell(value=cell.value, formula=cell.formula, style=cell.style)


def _paste_one(target: Range, kind: int, op: str, at: tuple[int, int], source: tuple[int, int], cell: Cell | None,
               value: object, turned: bool) -> None:
    """Paste one source cell over the cell at ``at``."""
    sheet = target.sheet
    row, column = at
    under = sheet.cell(row, column)
    default = sheet.book.stylesheet.default
    source_style: Style = (cell.style if cell is not None and cell.style is not None else default)
    style = under.style if under is not None and under.style is not None else sheet.style_at(row, column)
    if kind in _EVERYTHING or kind == FORMATS:
        style = replace(source_style, border=style.border) if kind == ALL_EXCEPT_BORDERS else source_style
    elif kind in (VALUES_AND_NUMBER_FORMATS, FORMULAS_AND_NUMBER_FORMATS):
        style = replace(style, number_format=source_style.number_format,
                        applied=style.applied | {"number_format"})
    formula, content = "", EMPTY
    if kind == FORMATS:
        formula = under.formula if under is not None else ""
        content = under.value if under is not None else EMPTY
    elif cell is not None and cell.formula and kind in _CONTENT_FORMULAS:
        formula = transpose_text(cell.formula, source, at) if turned \
            else shift_text(cell.formula, row - source[0], column - source[1])
    elif cell is not None:
        content = value if cell.formula else cell.value
    if op and kind != FORMATS:
        formula, content = _combined(op, under, formula, content)
    fresh = sheet.cell(row, column, create=True)
    assert fresh is not None
    fresh.value, fresh.formula, fresh.stale = (EMPTY, formula, True) if formula else (content, "", False)
    fresh.style, fresh.xf = (None if style == default else style), -1
    sheet.settle(row, column)
    sheet.cell_changed(row, column)


def _combined(op: str, under: Cell | None, formula: str, content: object) -> tuple[str, object]:
    """What an operation leaves: the pasted formula or value combined with what was under it."""
    under_formula = under.formula if under is not None else ""
    under_value = under.value if under is not None else EMPTY
    if not formula and not under_formula:
        first, second = _number(under_value), _number(content)
        if first is None or second is None:
            return "", under_value
        if op == "/" and second == 0:
            return "", ERRORS.get("#DIV/0!", ExcelError("#DIV/0!"))
        return "", {"+": first + second, "-": first - second, "*": first * second, "/": first / (second or 1)}[op]
    left = f"({under_formula[1:]})" if under_formula else _operand(under_value)
    right = f"({formula[1:]})" if formula else _operand(content)
    if left is None or right is None:
        return under_formula, under_value
    return f"={left}{op}{right}", EMPTY


def _number(value: object) -> float | None:
    """A value an operation can use: a number, or 0 for a blank; None for text and the rest."""
    if value is EMPTY:
        return 0.0
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return float(value)


def _operand(value: object) -> str | None:
    number = _number(value)
    return None if number is None else number_text(number, formula=True)


def _column_widths(source_sheet: Worksheet, made: Block, target: Range) -> None:
    dims, widths = target.sheet.dims, source_sheet.dims
    for offset in range(made.columns):
        column = target.first.left + offset
        if column <= MAX_COLUMNS:
            dims.set_column_width(column, widths.column_pixels(made.source(0, offset)[1]))
    target.sheet.touched()


# --- moving ----------------------------------------------------------------------------------------


def move(source: Range, area: Area, destination: Range) -> None:
    """Cut ``area`` to ``destination``'s top-left cell, taking every reference to it along."""
    from pyopenvba.apps.excel import _merges
    from pyopenvba.apps.excel._model import Range

    sheet, target = source.sheet, destination.sheet
    if sheet.book is not target.book:
        raise VBAUnsupportedError("cutting cells to another workbook is not implemented")
    top, left = destination.first.top, destination.first.left
    landing = Area(top, left, top + area.rows - 1, left + area.columns - 1, target.name)
    if landing.bottom > MAX_ROWS or landing.right > MAX_COLUMNS:
        raise error(1004, "The cells would move past the edge of the worksheet")
    if any(_merges.intersects(area, one) for one in sheet.merged_areas) or any(
            _merges.intersects(landing, one) for one in target.merged_areas):
        raise VBAUnsupportedError("cutting merged cells is not implemented")
    if _sized(sheet, area) or _sized(target, landing):
        raise VBAUnsupportedError("cutting whole rows or columns that carry sizes or formats of their own is not "
                                  "implemented")
    unmeasured(Range(target, [landing]), "Cutting cells")
    down, across = top - area.top, left - area.left
    moving = {position: cell for position, cell in sheet.cells_.items() if area.contains(*position)}
    book = sheet.book
    # Every formula and name follows the cells before any cell moves.
    for owner in book.sheets_:
        for position, cell in owner.cells_.items():
            if cell.formula:
                cell.formula = moved_references(cell.formula, owner.name, sheet.name, area, target.name, landing,
                                                down, across, moving=owner is sheet and position in moving)
    from pyopenvba.apps.excel._editing import name_scope

    for entry in book.names_.entries:
        text = moved_references(entry.refers_to, name_scope(book, entry) or sheet.name, sheet.name, area,
                                target.name, landing, down, across)
        if text != entry.refers_to:
            entry.refers_to = text
            book.names_.changed = True
    for position in moving:
        del sheet.cells_[position]
    for position in [position for position in target.cells_ if landing.contains(*position)]:
        del target.cells_[position]
    for (row, column), cell in moving.items():
        cell.stale = bool(cell.formula)
        if cell.formula:
            cell.value = EMPTY
        target.cells_[(row + down, column + across)] = cell
    from pyopenvba.apps.excel._autofilter import filter_moved

    filter_moved(sheet, area, target, down, across)
    for owner in {sheet, target}:
        owner.touched()
        owner.book.calculator.rebuild()
    book.application.clipboard = None


def _sized(sheet: Worksheet, area: Area) -> bool:
    """Whether whole rows or columns of ``area`` keep a height, width, hiding or format of their own."""
    if area.whole_columns:
        return any(area.left <= column <= area.right for column in sheet.dims.columns)
    if area.whole_rows:
        return any(area.top <= row <= area.bottom for row in sheet.dims.rows)
    return False


def moved_references(formula: str, owner: str, sheet: str, area: Area, target: str, landing: Area,
                     down: int, across: int, *, moving: bool = False) -> str:
    """``formula``, on ``owner``, with its references into a moved block following it.

    A reference wholly inside ``area`` on ``sheet`` moves by ``down`` and
    ``across`` to ``target``; one wholly inside ``landing`` on ``target``,
    whose cells the block replaces, becomes #REF!. One that loses a whole
    edge to the block -- its full width at the top or bottom, its full
    height at a side -- keeps the rest when the block goes to another
    sheet, and on its own sheet takes the edge where the block goes, if
    the block moves straight along it and the range stays the right way
    round: B7 cut to B9 leaves SUM(B2:B7) reading SUM(B2:B9). A formula
    ``moving`` with the block to ``target`` names that sheet only where
    it did, and names its old sheet for a reference to it that stays behind.
    """
    body = formula[1:] if formula.startswith("=") else formula
    after = target if moving else owner
    pieces: list[tuple[int, int, str]] = []
    for token in tokenize(body):
        if token.kind != "ref":
            continue
        named, reference = split_sheet(token.text)
        on = named or owner
        corners = [_corner(part) for part in reference.split(":")]
        box = _box(corners)
        prefix = token.text[: len(token.text) - len(reference)]
        if on.casefold() == sheet.casefold() and _inside(box, area):
            moved = ":".join(_shifted(corner, down, across) for corner in corners)
            if target.casefold() != on.casefold() or named:
                prefix = quote_sheet(target) + "!" if (named or target.casefold() != after.casefold()) else ""
            pieces.append((token.at, token.at + len(token.text), prefix + moved))
            continue
        if on.casefold() == target.casefold() and _inside(box, landing):
            pieces.append((token.at, token.at + len(token.text), prefix + "#REF!"))
            continue
        edged = _edged(corners, box, area, down, across, elsewhere=target.casefold() != sheet.casefold()) \
            if on.casefold() == sheet.casefold() else None
        if edged is None and (named or after.casefold() == owner.casefold()):
            continue
        text = reference if edged is None else edged
        if not named and after.casefold() != owner.casefold():
            prefix = quote_sheet(owner) + "!"
        pieces.append((token.at, token.at + len(token.text), prefix + text))
    for start, stop, replacement in reversed(pieces):
        body = body[:start] + replacement + body[stop:]
    return ("=" if formula.startswith("=") else "") + body


def _corner(part: str) -> tuple[str, int | None, str, int | None]:
    column_fixed = part.startswith("$")
    rest = part[1:] if column_fixed else part
    letters = ""
    while rest and rest[0].isalpha():
        letters, rest = letters + rest[0], rest[1:]
    row_fixed = rest.startswith("$")
    digits = rest[1:] if row_fixed else rest
    if not letters:
        # A whole row: its only dollar sign is the row's.
        return "", None, "$" if column_fixed or row_fixed else "", int(digits) if digits else None
    return ("$" if column_fixed else ""), column_number(letters), ("$" if row_fixed else ""), \
        int(digits) if digits else None


def _box(corners: list[tuple[str, int | None, str, int | None]]) -> Area:
    columns = [corner[1] for corner in corners]
    rows = [corner[3] for corner in corners]
    left = min(column for column in columns if column is not None) if any(columns) else 1
    right = max(column for column in columns if column is not None) if any(columns) else MAX_COLUMNS
    top = min(row for row in rows if row is not None) if any(rows) else 1
    bottom = max(row for row in rows if row is not None) if any(rows) else MAX_ROWS
    return Area(top, left, bottom, right)


def _inside(box: Area, area: Area) -> bool:
    return area.top <= box.top and box.bottom <= area.bottom and area.left <= box.left and box.right <= area.right


def _edged(corners: list[tuple[str, int | None, str, int | None]], box: Area, area: Area, down: int, across: int,
           *, elsewhere: bool) -> str | None:
    """A reference that loses a whole edge to a cut block, as moved_references has it; None where it stays."""
    if any(corner[1] is None or corner[3] is None for corner in corners):
        return None  # Whole rows and columns stay.
    top, left = max(box.top, area.top), max(box.left, area.left)
    bottom, right = min(box.bottom, area.bottom), min(box.right, area.right)
    if top > bottom or left > right:
        return None
    wide, tall = (left, right) == (box.left, box.right), (top, bottom) == (box.top, box.bottom)
    side = "top" if wide and top == box.top else "bottom" if wide and bottom == box.bottom else \
        "left" if tall and left == box.left else "right" if tall and right == box.right else ""
    if not side or (not elsewhere and (across if side in ("top", "bottom") else down)):
        return None  # On its own sheet an edge follows only a block moving straight along the range.
    new_top, new_left, new_bottom, new_right = box.top, box.left, box.bottom, box.right
    if side == "top":
        new_top = bottom + 1 if elsewhere else box.top + down
    elif side == "bottom":
        new_bottom = top - 1 if elsewhere else box.bottom + down
    elif side == "left":
        new_left = right + 1 if elsewhere else box.left + across
    else:
        new_right = left - 1 if elsewhere else box.right + across
    if new_top > new_bottom or new_left > new_right or new_top < 1 or new_left < 1 or new_bottom > MAX_ROWS \
            or new_right > MAX_COLUMNS:
        return None
    rows = [corner[3] or 0 for corner in corners]
    columns = [corner[1] or 0 for corner in corners]
    placed_rows = [new_top] if len(corners) == 1 else [new_top, new_bottom] if rows[0] <= rows[-1] else \
        [new_bottom, new_top]
    placed_columns = [new_left] if len(corners) == 1 else [new_left, new_right] if columns[0] <= columns[-1] else \
        [new_right, new_left]
    return ":".join(f"{column_fixed}{column_letter(column)}{row_fixed}{row}"
                    for (column_fixed, _, row_fixed, _), row, column in zip(corners, placed_rows, placed_columns))


def _shifted(corner: tuple[str, int | None, str, int | None], down: int, across: int) -> str:
    column_fixed, column, row_fixed, row = corner
    return (f"{column_fixed}{column_letter(column + across)}" if column is not None else "") + \
        (f"{row_fixed}{row + down}" if row is not None else "")
