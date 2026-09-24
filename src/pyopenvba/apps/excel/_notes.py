"""The notes on a sheet's cells, which VBA calls comments.

``Range.AddComment "text"`` puts a note on a cell, ``Range.Comment`` is
it or Nothing, and ``Worksheet.Comments`` lists a sheet's notes row by
row. A note is also a shape, "Comment 1", of Type 4, counted with the
sheet's other shapes, and its box is placed as Excel places a new one:
144 by 79 pixels, 15 right of the cell and 10 above it, the right edge
of a merged cell's block standing for the cell's; 2 below the sheet's
top at most; to the left of the cell where it would run off the sheet's
last column; and above the cell, its bottom 10 below the cell's top but
no nearer than 12 to the sheet's bottom, where it would run off the last
row. What the object model answers, from AddComment's refusals to what
Text and NoteText do with a start and a length, is what live Excel
answered in tests/fixtures/excel_model/probes.txt; the placement is from
tests/fixtures/notes.json.

A note moves with its cell when rows, columns or cells shift, and goes
with a deleted row or column; Copy takes it along, AutoFill does not,
Clear removes it, ClearContents and ClearFormats leave it, and a merge
keeps only its first cell's. The file keeps notes in two parts, the text
in a comments part and the box in the sheet's VML part (see
_notes_file).
"""

from __future__ import annotations

import uuid
from collections.abc import Callable
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NOTHING,
    VBAInt,
    VBASingle,
    error,
    to_bool,
    to_integer,
    to_text,
)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

#: A new note's box in pixels, 108 by 59.25 points, and how far it sits right of its cell and above it.
WIDTH, HEIGHT = 144, 79
_RIGHT, _ABOVE = 15, 10
#: The nearest a box comes to the sheet's top, and to its bottom when it goes above its cell.
_TOP_MARGIN, _BOTTOM_MARGIN = 2, 12
#: How many shape ids each VML part has to itself: a sheet's shapes are numbered from its block's start.
BLOCK = 1024
#: What Comment.Creator answers, xlCreatorCode; and Shape.Type for a note's box, msoComment.
_CREATOR = 1480803660
_MSO_COMMENT = 4

#: Where a box is, as the file's anchor holds it: the column its left edge is in and how many pixels into it, the
#: row its top edge is in and how far into it, then the same for its right and bottom edges, all from zero.
Anchor = tuple[int, int, int, int, int, int, int, int]


@dataclass
class Note:
    """A note on one cell."""

    text: str
    author: str
    visible: bool = False
    #: The box: its shape's name and id, where it is, and how big in pixels. Its first corner moves with the cells
    #: under it; its size stays as it is when they are resized, the far corner going with it
    #: (tests/fixtures/notes.json).
    name: str = ""
    shape_id: int = 0
    anchor: Anchor = (0, 0, 0, 0, 0, 0, 0, 0)
    width: int = WIDTH
    height: int = HEIGHT
    #: The random id Excel gives each note in the file, {GUID}.
    uid: str = ""
    #: The markup a note read from a file came with, its <comment> and its box, so a save writes an untouched one
    #: back as it arrived; each is "" once what it holds changes.
    comment_xml: str = ""
    shape_xml: str = ""


# --- where a box goes -----------------------------------------------------------------------------------------


def _column_left(sheet: Worksheet, column: int) -> int:
    """The pixels from the sheet's left edge to the column's."""
    return sheet.dims.columns_pixels(1, column - 1)


def _row_top(sheet: Worksheet, row: int) -> int:
    return sheet.dims.rows_pixels(1, row - 1)


def _line_at(position: int, before: Callable[[int], int], last: int) -> tuple[int, int]:
    """The line, counted from zero, the pixel ``position`` falls in, and how many pixels into it; ``before(n)`` is
    the pixels before line n, counted from one."""
    low, high = 1, last
    while low < high:
        middle = (low + high + 1) // 2
        if before(middle) <= position:
            low = middle
        else:
            high = middle - 1
    return low - 1, position - before(low)


def placed(sheet: Worksheet, row: int, column: int) -> Anchor:
    """Where Excel puts a new note's box on the cell (tests/fixtures/notes.json)."""
    from pyopenvba.apps.excel import _merges

    block = _merges.at(sheet, row, column) or Area(row, column, row, column)
    width = sheet.dims.columns_pixels(1, MAX_COLUMNS)
    height = sheet.dims.rows_pixels(1, MAX_ROWS)
    left = _column_left(sheet, block.right + 1) + _RIGHT
    if left + WIDTH > width:
        left = _column_left(sheet, block.left) - _RIGHT - WIDTH
    top = max(_row_top(sheet, row) - _ABOVE, _TOP_MARGIN)
    if top + HEIGHT > height:
        top = min(_row_top(sheet, row) + _ABOVE, height - _BOTTOM_MARGIN) - HEIGHT
    return _anchor(sheet, left, top, left + WIDTH, top + HEIGHT)


def _anchor(sheet: Worksheet, left: int, top: int, right: int, bottom: int) -> Anchor:
    columns: Callable[[int], int] = lambda column: _column_left(sheet, column)  # noqa: E731
    rows: Callable[[int], int] = lambda row: _row_top(sheet, row)  # noqa: E731
    first_column, x = _line_at(left, columns, MAX_COLUMNS)
    first_row, y = _line_at(top, rows, MAX_ROWS)
    last_column, x2 = _line_at(right, columns, MAX_COLUMNS)
    last_row, y2 = _line_at(bottom, rows, MAX_ROWS)
    return first_column, x, first_row, y, last_column, x2, last_row, y2


def box(sheet: Worksheet, note: Note) -> tuple[int, int, int, int]:
    """A note's box as it stands, in pixels from the sheet's corner: its left, top, right and bottom edges."""
    column, x, row, y = note.anchor[:4]
    left, top = _column_left(sheet, column + 1) + x, _row_top(sheet, row + 1) + y
    return left, top, left + note.width, top + note.height


def anchor_of(sheet: Worksheet, note: Note) -> Anchor:
    """A note's anchor as a file writes it, its far corner where the box's size puts it on the sheet as it is."""
    return _anchor(sheet, *box(sheet, note))


def size_of(sheet: Worksheet, anchor: Anchor) -> tuple[int, int]:
    """How wide and tall in pixels the box a file's anchor gives is, on the sheet as it came."""
    column, x, row, y, last_column, x2, last_row, y2 = anchor
    return (_column_left(sheet, last_column + 1) + x2 - _column_left(sheet, column + 1) - x,
            _row_top(sheet, last_row + 1) + y2 - _row_top(sheet, row + 1) - y)


# --- the sheet's notes ------------------------------------------------------------------------------------------


def vml_block(sheet: Worksheet) -> int:
    """The block of ids the sheet's VML part numbers its shapes from, the next free one if it has none yet: each
    part claims a block of its own (tests/fixtures/notes.json)."""
    if not sheet.vml_block:
        sheet.vml_block = 1 + max((one.vml_block for one in sheet.book.sheets_), default=0)
    return sheet.vml_block


def next_shape_id(sheet: Worksheet) -> int:
    """The id the sheet's next note or form control takes: one past the last it gave, never one a deleted shape
    had (tests/fixtures/notes.json)."""
    start = vml_block(sheet) * BLOCK
    sheet.vml_last_id = max(sheet.vml_last_id, start) + 1
    return sheet.vml_last_id


def note_at(sheet: Worksheet, row: int, column: int) -> Note | None:
    return sheet.notes.get((row, column))


def existing(sheet: Worksheet, row: int, column: int) -> Note:
    """The cell's note, which an object standing for it expects; error 1004 once it is deleted."""
    found = note_at(sheet, row, column)
    if found is None:
        raise error(1004, "the note has been deleted")
    return found


def edited(sheet: Worksheet, row: int, column: int, text: str) -> None:
    """The cell's note given new text."""
    note = existing(sheet, row, column)
    if text != note.text:
        sheet.notes[(row, column)] = replace(note, text=text, comment_xml="")
        changed(sheet)


def in_order(sheet: Worksheet) -> list[tuple[tuple[int, int], Note]]:
    """The sheet's notes as Comments lists them: row by row, and along each row (tests/fixtures/excel_model/)."""
    return sorted(sheet.notes.items())


def add(sheet: Worksheet, row: int, column: int, text: str) -> Note:
    """A new note on the cell, written by the application's user."""
    sheet.shape_count += 1
    note = Note(text=text, author=sheet.book.application.user_name, name=f"Comment {sheet.shape_count}",
                shape_id=next_shape_id(sheet), anchor=placed(sheet, row, column),
                uid="{" + str(uuid.uuid4()).upper() + "}")
    sheet.notes[(row, column)] = note
    changed(sheet)
    return note


def changed(sheet: Worksheet) -> None:
    sheet.notes_changed = True
    sheet.book.saved = False


def remove(sheet: Worksheet, row: int, column: int) -> None:
    if sheet.notes.pop((row, column), None) is not None:
        changed(sheet)


def clear(sheet: Worksheet, areas: list[Area]) -> None:
    """Take the notes off the cells of the areas, as Clear, ClearComments and ClearNotes do."""
    for row, column in [key for key in sheet.notes if any(area.contains(*key) for area in areas)]:
        remove(sheet, row, column)


def merged(sheet: Worksheet, area: Area) -> None:
    """A merge keeps the note on the block's first cell and drops the rest (tests/fixtures/excel_model/)."""
    for row, column in [key for key in sheet.notes if area.contains(*key) and key != (area.top, area.left)]:
        remove(sheet, row, column)


# --- moving with the cells ---------------------------------------------------------------------------------------


def _moved_edge(line: int, offset: int, start: int, count: int, delete: bool) -> tuple[int, int]:
    """One edge of a box, a zero-based line and pixels into it, after ``count`` lines from ``start`` (from zero)
    are inserted or deleted: it moves with the lines after it, and one inside what is deleted comes to rest at the
    start of what follows."""
    if not delete:
        return (line + count, offset) if line >= start else (line, offset)
    if line >= start + count:
        return line - count, offset
    if line >= start:
        return start, 0
    return line, offset


def lines_moved(sheet: Worksheet, start: int, count: int, *, delete: bool, rows: bool) -> None:
    """Whole rows or columns from ``start`` (from one) inserted or deleted: each note moves with its cell, goes with
    a deleted one, and its box moves and sizes with the lines under it (tests/fixtures/notes.json)."""
    if not sheet.notes:
        return
    moved: dict[tuple[int, int], Note] = {}
    for (row, column), note in sheet.notes.items():
        line = row if rows else column
        if delete and start <= line < start + count:
            continue
        if line >= start:
            line = line - count if delete else line + count
        key = (line, column) if rows else (row, line)
        column_edge, x, row_edge, y, last_column, x2, last_row, y2 = note.anchor
        if rows:
            row_edge, y = _moved_edge(row_edge, y, start - 1, count, delete)
            last_row, y2 = _moved_edge(last_row, y2, start - 1, count, delete)
        else:
            column_edge, x = _moved_edge(column_edge, x, start - 1, count, delete)
            last_column, x2 = _moved_edge(last_column, x2, start - 1, count, delete)
        anchor = (column_edge, x, row_edge, y, last_column, x2, last_row, y2)
        moved[key] = note if anchor == note.anchor else replace(note, anchor=anchor, shape_xml="")
    if moved != sheet.notes or list(moved) != list(sheet.notes):
        sheet.notes = moved
        changed(sheet)


def cells_moved(sheet: Worksheet, moves: dict[tuple[int, int], tuple[int, int] | None]) -> None:
    """Cells shifted by an insert or a delete that does not take whole rows or columns: each note goes where its
    cell goes, None for a deleted cell (tests/fixtures/excel_model/), and gets the box a new note there would have,
    where its box goes in Excel not having been measured."""
    if not any(key in moves for key in sheet.notes):
        return
    moved: dict[tuple[int, int], Note] = {}
    for key, note in sheet.notes.items():
        if key not in moves:
            moved[key] = note
            continue
        target = moves[key]
        if target is not None:
            moved[target] = replace(note, anchor=placed(sheet, *target), shape_xml="")
    sheet.notes = moved
    changed(sheet)


def cut(source: Worksheet, area: Area, target: Worksheet, down: int, across: int) -> None:
    """Cut moves the notes of ``area`` ``down`` and ``across`` on ``target`` and takes off those it lands on
    (tests/fixtures/excel_model/): on its own sheet a note keeps its shape, on another it gets one there; either
    way its box goes where a new note's would, where Excel puts it not having been measured."""
    moving = {key: note for key, note in source.notes.items() if area.contains(*key)}
    landing = Area(area.top + down, area.left + across, area.bottom + down, area.right + across)
    if not moving and not any(landing.contains(*key) for key in target.notes):
        return
    for key in moving:
        del source.notes[key]
    clear(target, [landing])
    for (row, column), note in moving.items():
        at = (row + down, column + across)
        if target is not source:
            target.shape_count += 1
            note = replace(note, name=f"Comment {target.shape_count}", shape_id=next_shape_id(target))
        target.notes[at] = replace(note, anchor=placed(target, *at), shape_xml="")
    changed(source)
    changed(target)


def copied(target: Worksheet, notes: dict[tuple[int, int], Note], landing: list[Area]) -> None:
    """What a copy, or a paste of comments, does to the notes of the cells it lands on: each takes the note its
    source cell had, ``notes`` by the cell it lands on, as a new note in a box of its own, or none, which takes
    off the note it had (tests/fixtures/excel_model/, notes.json)."""
    clear(target, landing)
    for (row, column), note in notes.items():
        target.shape_count += 1
        target.notes[(row, column)] = replace(
            note, name=f"Comment {target.shape_count}", shape_id=next_shape_id(target),
            anchor=placed(target, row, column), uid="{" + str(uuid.uuid4()).upper() + "}", comment_xml="",
            shape_xml="")
    if notes:
        changed(target)


# --- the object model ---------------------------------------------------------------------------------------------


def _text_argument(value: object) -> str:
    """The text AddComment or NoteText is given: a string; a number is error 1004 (tests/fixtures/excel_model/)."""
    if isinstance(value, str):
        return value
    raise error(1004, "Application-defined or object-defined error")


class Comment(ExcelObject):
    """One cell's note."""

    vba_type_name = "Comment"

    def __init__(self, sheet: Worksheet, row: int, column: int) -> None:
        self.sheet = sheet
        self.row, self.column = row, column

    @property
    def note(self) -> Note:
        return existing(self.sheet, self.row, self.column)

    @method
    def Text(self, Text: object = MISSING, Start: object = MISSING, Overwrite: object = MISSING) -> object:
        """The note's text; given text, it replaces the text from Start on, or goes in at Start without
        overwriting when Overwrite is False; Overwrite True is error 5 (tests/fixtures/excel_model/)."""
        current = self.note.text
        if Text is MISSING:
            return current
        added = to_text(Text)
        if Start is MISSING:
            edited(self.sheet, self.row, self.column, added)
            return added
        start = int(to_integer(Start, "Long"))
        if start < 1 or (Overwrite is not MISSING and to_bool(Overwrite)):
            raise error(5, "Invalid procedure call or argument")
        at = start - 1
        wanted = current[:at] + added + (current[at:] if Overwrite is not MISSING else "")
        edited(self.sheet, self.row, self.column, wanted)
        return wanted

    @member
    def Author(self) -> object:
        return self.note.author

    @member
    def Visible(self) -> object:
        return self.note.visible

    @setter("Visible")
    def _set_visible(self, value: object) -> None:
        shown = to_bool(value)
        note = self.note
        if shown != note.visible:
            self.sheet.notes[(self.row, self.column)] = replace(note, visible=shown, shape_xml="")
            changed(self.sheet)

    @method
    def Delete(self) -> object:
        existing(self.sheet, self.row, self.column)
        remove(self.sheet, self.row, self.column)
        return EMPTY

    def _neighbour(self, step: int) -> object:
        order = [key for key, _ in in_order(self.sheet)]
        at = order.index((self.row, self.column)) + step
        return Comment(self.sheet, *order[at]) if 0 <= at < len(order) else NOTHING

    @method
    def Next(self) -> object:
        return self._neighbour(1)

    @method
    def Previous(self) -> object:
        return self._neighbour(-1)

    @member
    def Parent(self) -> object:
        return Range(self.sheet, [Area(self.row, self.column, self.row, self.column, self.sheet.name)])

    @member
    def Shape(self) -> object:
        existing(self.sheet, self.row, self.column)
        return NoteShape(self.sheet, self.row, self.column)

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(_CREATOR, "Long")


class Comments(VBACollection, ExcelObject):
    """A sheet's notes, row by row."""

    vba_type_name = "Comments"

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def vba_items(self) -> list[object]:
        return [Comment(self.sheet, *key) for key, _ in in_order(self.sheet)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        position = int(to_integer(index, "Long"))
        if position < 1:
            raise error(9, "Subscript out of range")
        if position > len(items):
            raise error(1004, "Application-defined or object-defined error")
        return items[position - 1]

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(_CREATOR, "Long")


class NoteShape(ExcelObject):
    """A note's box, as Shapes and Comment.Shape answer it."""

    vba_type_name = "Shape"

    def __init__(self, sheet: Worksheet, row: int, column: int) -> None:
        self.sheet = sheet
        self.row, self.column = row, column

    @property
    def note(self) -> Note:
        return existing(self.sheet, self.row, self.column)

    def _edges(self) -> tuple[int, int, int, int]:
        return box(self.sheet, self.note)

    @member
    def Name(self) -> object:
        return self.note.name

    @member
    def Type(self) -> object:
        return VBAInt(_MSO_COMMENT, "Long")

    @member
    def Left(self) -> object:
        return VBASingle(self._edges()[0] * 0.75)

    @member
    def Top(self) -> object:
        return VBASingle(self._edges()[1] * 0.75)

    @member
    def Width(self) -> object:
        left, _, right, _ = self._edges()
        return VBASingle((right - left) * 0.75)

    @member
    def Height(self) -> object:
        _, top, _, bottom = self._edges()
        return VBASingle((bottom - top) * 0.75)

    @member
    def Visible(self) -> object:
        """msoTrue or msoFalse, as a Long (tests/fixtures/excel_model/)."""
        return VBAInt(-1 if self.note.visible else 0, "Long")

    @member
    def TextFrame(self) -> object:
        return NoteTextFrame(self.sheet, self.row, self.column)

    @method
    def Delete(self) -> object:
        remove(self.sheet, self.row, self.column)
        return EMPTY

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application


class NoteTextFrame(ExcelObject):
    """The text in a note's box."""

    vba_type_name = "TextFrame"

    def __init__(self, sheet: Worksheet, row: int, column: int) -> None:
        self.sheet = sheet
        self.row, self.column = row, column

    @method
    def Characters(self, Start: object = MISSING, Length: object = MISSING) -> object:
        if Start is not MISSING or Length is not MISSING:
            from pyopenvba.exceptions import VBAUnsupportedError

            raise VBAUnsupportedError("TextFrame.Characters of part of a note's text is not implemented")
        return NoteCharacters(self.sheet, self.row, self.column)


class NoteCharacters(ExcelObject):
    """All of a note's text, as Characters answers it."""

    vba_type_name = "Characters"

    def __init__(self, sheet: Worksheet, row: int, column: int) -> None:
        self.sheet = sheet
        self.row, self.column = row, column

    @member
    def Text(self) -> object:
        return Comment(self.sheet, self.row, self.column).note.text

    @setter("Text")
    def _set_text(self, value: object) -> None:
        edited(self.sheet, self.row, self.column, to_text(value))


# --- what Range does with notes ---------------------------------------------------------------------------------


def add_comment(target: Range, text: object) -> object:
    """Range.AddComment: a note on one cell that has none; error 5 for a range of several cells and 1004 for a cell
    with a note already (tests/fixtures/excel_model/)."""
    if not target.single:
        raise error(5, "Invalid procedure call or argument")
    row, column = target.first.top, target.first.left
    wanted = "" if text is MISSING else _text_argument(text)
    if note_at(target.sheet, row, column) is not None:
        raise error(1004, "Application-defined or object-defined error")
    add(target.sheet, row, column, wanted)
    return Comment(target.sheet, row, column)


def comment_of(target: Range) -> object:
    """Range.Comment: the note of a range of one cell, or Nothing; a range of several is Nothing, whatever its
    cells hold (tests/fixtures/excel_model/)."""
    if not target.single:
        return NOTHING
    row, column = target.first.top, target.first.left
    return NOTHING if note_at(target.sheet, row, column) is None else Comment(target.sheet, row, column)


def note_text(target: Range, text: object, start: object, length: object) -> object:
    """Range.NoteText, the note of the range's first cell: read, all of it or Length characters from Start; written,
    it replaces the text from Start on, or Length characters of it, and "" written whole takes the note off
    (tests/fixtures/excel_model/)."""
    sheet, row, column = target.sheet, target.first.top, target.first.left
    note = note_at(sheet, row, column)
    current = "" if note is None else note.text
    at = 0 if start is MISSING else int(to_integer(start, "Long")) - 1
    if at < 0:
        raise error(1004, "Application-defined or object-defined error")
    if text is MISSING:
        if length is MISSING:
            return current[at:]
        return current[at:at + int(to_integer(length, "Long"))]
    added = _text_argument(text)
    if start is MISSING:
        wanted = added
    elif length is MISSING:
        wanted = current[:at] + added
    else:
        wanted = current[:at] + added + current[at + int(to_integer(length, "Long")):]
    if not wanted and start is MISSING:
        remove(sheet, row, column)
        return ""
    if note is None:
        add(sheet, row, column, wanted)
    else:
        edited(sheet, row, column, wanted)
    return wanted


__all__ = [
    "BLOCK",
    "HEIGHT",
    "WIDTH",
    "Anchor",
    "Comment",
    "Comments",
    "Note",
    "NoteCharacters",
    "NoteShape",
    "NoteTextFrame",
    "add",
    "add_comment",
    "anchor_of",
    "box",
    "cells_moved",
    "changed",
    "clear",
    "comment_of",
    "copied",
    "cut",
    "edited",
    "existing",
    "in_order",
    "lines_moved",
    "merged",
    "next_shape_id",
    "note_at",
    "note_text",
    "placed",
    "remove",
    "size_of",
    "vml_block",
]
