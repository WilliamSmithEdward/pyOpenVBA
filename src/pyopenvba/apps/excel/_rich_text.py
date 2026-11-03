"""A cell's text in runs of fonts, and Range.Characters, as Excel keeps them (scripts/measure_characters.py,
tests/fixtures/characters.json).

A cell's own font is its first character's: a run of its text in another font keeps that font, which a file
writes as the run's rPr, the first run with none. Two runs side by side in one font are one, and a cell whose
characters all share a font has no runs. Writing a value, the same text too, drops them, as do Replace and pasted
values; a copy, a sort and a save keep them; a format applied over the cell, Style or ClearFormats, changes its own
font and so the first run, and leaves the rest as they are.

Range.Characters(Start, Length) reads and changes part of one cell's text:

- Start below 1 is 1, and one past the text's end reads nothing; Length left out, or below 0, reaches to the end,
  0 is nothing at Start. A font set with Length 0 reaches to the end, though, and one set past the end changes
  nothing. Start and Length are taken as Longs are, Null as left out; other text is error 1004.
- Count answers Start where Length is more than 0, -1 with Start left out, and the text's length otherwise.
- Text, Caption, Insert and Delete want a cell of text, or an empty one: a number, a logical, a date or a formula
  is error 1004, though its font changes the whole cell's. What replaces characters takes the first one's font,
  or the cell's own where it replaces none, and an edit that would leave more than 255 characters changes nothing,
  with no error. Insert answers what it inserted, Delete Null.
- A font read wants one cell; an empty part reads the cell's own font.
- Any font set, or text set, leaves the workbook unsaved, even one that changes nothing or reaches past the text's
  end; reads, and an edit too long to make, do not.

A run's font counts toward its row's height as a cell's font does, and AutoFit measures each run in its own font
(scripts/measure_rich_text_fit.py, tests/fixtures/rich_text_fit.json; see _dimensions and _column_fit).
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterator, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING, Final

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba.apps.excel import _styles as S
from pyopenvba.apps.excel._formats import Font
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, VBAInt, error, to_integer, to_number, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Cell, Range, Worksheet

_RUN: Final = re.compile(r"<r>(.*?)</r>", re.DOTALL)
_RUN_FONT: Final = re.compile(r"<rPr>(.*?)</rPr>", re.DOTALL)
_TEXT: Final = re.compile(r"<t\b[^>]*>(.*?)</t>|<t\b[^>]*/>", re.DOTALL)
#: Past this many positions a part's cells are found among the sheet's cells rather than position by position.
_FEW: Final = 4096
#: The longest text an edit of characters leaves; one that would leave more changes nothing.
_LONGEST_EDIT: Final = 255


@dataclass(frozen=True, slots=True)
class Run:
    """A stretch of a cell's text in one font, None for the cell's own."""

    text: str
    font: S.Font | None = None


def runs_of(cell: Cell) -> tuple[Run, ...] | None:
    """A cell's runs while it holds the text they make, and no formula; None for a cell in its own font."""
    runs = cell.runs
    if runs is None:
        return None
    if cell.formula or not isinstance(cell.value, str) or "".join(run.text for run in runs) != cell.value:
        cell.runs = None
        return None
    return runs


def _text_of(cell: Cell | None) -> str | None:
    """A cell's text as Characters reads it: its own, "" for none, or None for a number, a logical, a date, an
    error or a formula."""
    if cell is None or cell.value is EMPTY and not cell.formula:
        return ""
    if cell.formula or not isinstance(cell.value, str):
        return None
    return cell.value


def fonts(sheet: Worksheet, row: int, column: int) -> list[S.Font]:
    """The font of each character of a cell's text."""
    cell = sheet.cells_.get((row, column))
    text = _text_of(cell)
    own = sheet.style_at(row, column).font
    if cell is None or not text:
        return []
    runs = runs_of(cell)
    if runs is None:
        return [own] * len(text)
    return [run.font or own for run in runs for _ in run.text]


def settle(sheet: Worksheet, row: int, column: int, given: list[S.Font]) -> None:
    """A text cell's characters given these fonts: its own font the first one's, and runs for the rest. Runs that
    change leave the sheet to save again and its row to be measured again, as a run's font counts in its row's
    height (see _dimensions)."""
    cell = sheet.cells_.get((row, column))
    if cell is None or not isinstance(cell.value, str) or not cell.value:
        if cell is not None:
            cell.runs = None
        return
    text = cell.value
    style = sheet.style_at(row, column)
    if style.font != given[0]:
        sheet.restyle(row, column, S.applying(style, "font", font=given[0]))
        cell = sheet.cells_[(row, column)]
    pieces: list[tuple[str, S.Font]] = []
    for character, font in zip(text, given, strict=True):
        if pieces and pieces[-1][1] == font:
            pieces[-1] = (pieces[-1][0] + character, font)
        else:
            pieces.append((character, font))
    runs = None if len(pieces) == 1 else tuple(
        Run(piece, None if index == 0 else font) for index, (piece, font) in enumerate(pieces))
    if runs != cell.runs:
        cell.runs = runs
        sheet.touched()
        sheet.dims.fonts_changed(row)


def _cells(target: Range) -> Iterator[tuple[int, int, Cell]]:
    """The cells a range holds, position by position where it is small."""
    sheet = target.sheet
    for area in target.areas:
        bottom, right = min(area.bottom, MAX_ROWS), min(area.right, MAX_COLUMNS)
        if (bottom - area.top + 1) * (right - area.left + 1) <= _FEW:
            for row in range(area.top, bottom + 1):
                for column in range(area.left, right + 1):
                    found = sheet.cells_.get((row, column))
                    if found is not None:
                        yield row, column, found
        else:
            for (row, column), found in list(sheet.cells_.items()):
                if area.contains(row, column):
                    yield row, column, found


def run_fonts(target: Range) -> list[S.Font]:
    """The fonts of the runs a range's cells' text is in apart from their own."""
    return [run.font for _, _, cell in _cells(target) for run in runs_of(cell) or () if run.font is not None]


def refont(target: Range, change: Callable[[S.Font], S.Font]) -> None:
    """A font change over a range's whole cells reaching each run of their text, as Range.Font's does."""
    for row, column, cell in list(_cells(target)):
        runs = runs_of(cell)
        if runs is not None:
            own = target.sheet.style_at(row, column).font
            settle(target.sheet, row, column,
                   [own if run.font is None else change(run.font) for run in runs for _ in run.text])


# --- the file ----------------------------------------------------------------------------


def read_runs(item: str, text_of: Callable[[str], str]) -> tuple[Run, ...] | None:
    """The runs of a string item, an <si> or an <is>, or None for one in a single run in the cell's own font."""
    runs: list[Run] = []
    for body in _RUN.findall(item):
        found = _RUN_FONT.search(body)
        text = "".join(text_of(piece) for piece in _TEXT.findall(body))
        runs.append(Run(text, S.parse_font(found.group(1).replace("<rFont", "<name")) if found else None))
    if len(runs) > 1 or runs and runs[0].font is not None:
        return tuple(runs)
    return None


def runs_xml(runs: tuple[Run, ...], text_xml: Callable[[str], str]) -> str:
    """Runs as an <is> writes them: each in an <r>, those in a font of their own with its <rPr>."""
    out: list[str] = []
    for run in runs:
        font = "" if run.font is None else _run_font_xml(run.font)
        out.append(f'<r>{font}<t xml:space="preserve">{text_xml(run.text)}</t></r>')
    return "".join(out)


def _run_font_xml(font: S.Font) -> str:
    """A run's font as its <rPr>: a font's elements, the name as rFont."""
    inner = S.font_xml(font).removeprefix("<font>").removesuffix("</font>")
    return "<rPr>" + inner.replace("<name ", "<rFont ") + "</rPr>"


# --- Range.Characters -----------------------------------------------------------------------


def characters(target: Range, start: object, length: object) -> Characters:
    """Range.Characters(Start, Length)."""
    return Characters(target, _position(start), _position(length))


def _position(value: object) -> int | None:
    if value is MISSING or value is NULL:
        return None
    if isinstance(value, str):
        try:
            to_number(value)
        except Exception:
            raise error(1004, "Unable to get the Characters property of the Range class") from None
    return int(to_integer(value, "Long"))


class Characters(ExcelObject):
    """Part of a cell's text: its characters from Start, Length of them."""

    vba_type_name = "Characters"

    def __init__(self, target: Range, start: int | None, length: int | None) -> None:
        self.target = target
        self.start = start
        self.length = length

    def _first(self) -> int:
        return max(1, self.start or 1) - 1

    def read_span(self, text: str) -> tuple[int, int]:
        """The characters a read reaches: Length 0 none, below 0 or left out all to the end."""
        first = min(self._first(), len(text))
        if self.length is None or self.length < 0:
            return first, len(text)
        return first, min(len(text), first + self.length)

    def write_span(self, text: str) -> tuple[int, int]:
        """The characters a font set reaches: as a read's, but Length 0 reaching to the end."""
        first = min(self._first(), len(text))
        if self.length is None or self.length <= 0:
            return first, len(text)
        return first, min(len(text), first + self.length)

    def _one_text(self, what: str, kind: str = "get") -> tuple[int, int, str]:
        """The one cell's position and text a member wants; error 1004 for several cells or a cell of no text."""
        if not self.target.single:
            raise error(1004, f"Unable to {kind} the {what} property of the Characters class")
        area = self.target.first
        text = _text_of(self.target.sheet.cells_.get((area.top, area.left)))
        if text is None:
            if kind == "method":
                raise error(1004, f"{what} method of Characters class failed")
            raise error(1004, f"Unable to {kind} the {what} property of the Characters class")
        return area.top, area.left, text

    @member(default=True)
    def Text(self) -> object:
        _, _, text = self._one_text("Text")
        first, last = self.read_span(text)
        return text[first:last]

    @setter("Text")
    def _set_text(self, value: object) -> None:
        self._replace(to_text(value), "Text", "set")

    @member
    def Caption(self) -> object:
        _, _, text = self._one_text("Caption")
        first, last = self.read_span(text)
        return text[first:last]

    @setter("Caption")
    def _set_caption(self, value: object) -> None:
        self._replace(to_text(value), "Caption", "set")

    @member
    def Count(self) -> object:
        """Start where Length is more than 0, as Excel answers, -1 for Start left out; else the text's length."""
        if self.length is not None and self.length > 0:
            return VBAInt(self.start if self.start is not None else -1, "Long")
        _, _, text = self._one_text("Count")
        return VBAInt(len(text), "Long")

    @member
    def Font(self) -> object:
        return CharactersFont(self)

    @method
    def Insert(self, String: object = MISSING) -> object:
        if String is MISSING:
            raise error(449)
        inserted = to_text(String)
        self._replace(inserted, "Insert", "method")
        return inserted

    @method
    def Delete(self) -> object:
        self._replace("", "Delete", "method")
        return NULL

    @member
    def Parent(self) -> object:
        return self.target

    @member
    def Application(self) -> object:
        return self.target.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")

    def _replace(self, new: str, what: str, kind: str) -> None:
        """The characters replaced by ``new``, which takes the first one's font, or the cell's own for none."""
        from pyopenvba.apps.excel import _events

        if not self.target.single:
            raise VBAUnsupportedError(f"Characters.{what} over several cells is not implemented")
        row, column, text = self._one_text(what, kind)
        sheet = self.target.sheet
        given = fonts(sheet, row, column)
        own = sheet.style_at(row, column).font
        first, last = self.read_span(text)
        font = given[first] if first < last else own
        edited = text[:first] + new + text[last:]
        if len(edited) > _LONGEST_EDIT:
            return
        cell = sheet.cell(row, column, create=True)
        assert cell is not None
        cell.value = edited
        cell.runs = None
        sheet.cell_changed(row, column)
        settle(sheet, row, column, given[:first] + [font] * len(new) + given[last:])
        from pyopenvba.apps.excel._model import Range as RangeClass

        _events.after_edit(RangeClass(sheet, [Area(row, column, row, column, sheet.name)]))


class CharactersFont(Font):
    """The font of part of a cell's text: what a read finds there, and a change made to those characters."""

    vba_type_name = "Font"

    def __init__(self, part: Characters) -> None:
        super().__init__(part.target)
        self.part = part

    @member
    def Parent(self) -> object:
        return self.part

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        spec = self.vba_member(name)
        if spec is not None and spec.name not in ("Parent", "Application", "Creator") and not self.target.single:
            raise error(1004, f"Unable to get the {spec.name} property of the Font class")
        return super().vba_get(name, args, named)

    def _fonts(self) -> list[S.Font]:
        area = self.target.first
        sheet = self.target.sheet
        own = sheet.style_at(area.top, area.left).font
        text = _text_of(sheet.cells_.get((area.top, area.left)))
        if not text:
            return [own]
        first, last = self.part.read_span(text)
        return fonts(sheet, area.top, area.left)[first:last] or [own]

    def _write(self, change: Callable[[S.Font], S.Font]) -> None:
        sheet = self.target.sheet
        # Any font set leaves the workbook to save, one that changes nothing too.
        sheet.touched()
        for area in self.target.areas:
            if area.rows * area.columns > 1048576:
                raise VBAUnsupportedError("Characters over more than 1,048,576 cells is not implemented")
            for row in range(area.top, area.bottom + 1):
                for column in range(area.left, area.right + 1):
                    text = _text_of(sheet.cells_.get((row, column)))
                    style = sheet.style_at(row, column)
                    if not text:
                        # A number, a logical, a date, a formula or nothing: the whole cell's font.
                        sheet.restyle(row, column, S.applying(style, "font", font=change(style.font)))
                        continue
                    given = fonts(sheet, row, column)
                    first, last = self.part.write_span(text)
                    if first < last:
                        settle(sheet, row, column, given[:first] + [change(font) for font in given[first:last]]
                               + given[last:])
