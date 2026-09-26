"""What a macro reads and sets on a range's format: Font, Interior, Borders, alignment, protection.

Every answer here is Excel's own, replayed from tests/fixtures/range_format.json:

* A property read over several cells answers the shared value, or Null
  when the cells differ. Colours are compared as colours: when two cells
  differ in colour, their ``Color``, ``ColorIndex``, ``ThemeColor`` and
  ``TintAndShade`` are all Null, except that a mixed ``Interior.Color``
  answers 0.
* A border between two cells is stored on one of them. Reading an edge
  looks at the cell's own side first and then at the neighbour's
  opposite side. Setting an edge writes it on the cell, first copying
  in whatever the cell showed from its neighbours, and clears the
  neighbour's copy; an inside border is written on both cells. A range's
  edge answers for its first cell along that edge.
* ``LineStyle`` and ``Weight`` pair up into the file's thirteen styles.
  A style keeps the current weight when the pair exists and otherwise
  takes its own default; a weight the style cannot take turns the line
  continuous.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable
from dataclasses import replace
from typing import TYPE_CHECKING, Final

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba.apps.excel import _row_formats
from pyopenvba.apps.excel import _styles as S
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.apps.excel._visible import visible_areas
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, setter
from pyopenvba.interpreter._values import MISSING, NULL, VBAInt, error, to_bool, to_integer, to_number, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

# --- the constants Excel answers with ---------------------------------------------------------

UNDERLINES = {"": -4142, "single": 2, "double": -4119, "singleAccounting": 4, "doubleAccounting": 5}
#: True is -1, and 3 is double, as Excel reads them.
_UNDERLINE_OF = {-4142: "", 2: "single", -1: "single", -4119: "double", 3: "double", 4: "singleAccounting",
                 5: "doubleAccounting", 0: ""}
PATTERNS = {"none": -4142, "solid": 1, "darkGray": -4126, "mediumGray": -4125, "lightGray": -4124,
            "gray125": 17, "gray0625": 18, "darkHorizontal": -4128, "darkVertical": -4166, "darkDown": -4121,
            "darkUp": -4162, "darkGrid": 9, "darkTrellis": 10, "lightHorizontal": 11, "lightVertical": 12,
            "lightDown": 13, "lightUp": 14, "lightGrid": 15, "lightTrellis": 16}
_PATTERN_OF = {value: name for name, value in PATTERNS.items()}
HORIZONTAL = {"": 1, "general": 1, "left": -4131, "center": -4108, "right": -4152, "fill": 5, "justify": -4130,
              "centerContinuous": 7, "distributed": -4117}
_HORIZONTAL_OF = {1: "", -4131: "left", -4108: "center", -4152: "right", 5: "fill", -4130: "justify",
                  7: "centerContinuous", -4117: "distributed"}
VERTICAL = {"": -4107, "bottom": -4107, "top": -4160, "center": -4108, "justify": -4130, "distributed": -4117}
_VERTICAL_OF = {-4107: "", -4160: "top", -4108: "center", -4130: "justify", -4117: "distributed"}
_READING = {0: -5002, 1: -5003, 2: -5004}
_READING_OF = {value: key for key, value in _READING.items()}
#: (LineStyle, Weight) for each of the file's border styles.
LINES = {"hair": (1, 1), "thin": (1, 2), "medium": (1, -4138), "thick": (1, 4), "dashed": (-4115, 2),
         "mediumDashed": (-4115, -4138), "dashDot": (4, 2), "mediumDashDot": (4, -4138), "dashDotDot": (5, 2),
         "mediumDashDotDot": (5, -4138), "dotted": (-4118, 2), "double": (-4119, 4), "slantDashDot": (13, -4138)}
_STYLE_OF = {value: name for name, value in LINES.items()}
#: The weight a line style takes when the current weight does not pair with it.
_DEFAULT_WEIGHT = {1: 2, -4115: 2, 4: 2, 5: 2, -4118: 2, -4119: 4, 13: -4138}
_WEIGHTS = (1, 2, -4138, 4)
XL_NONE = -4142
XL_AUTOMATIC = -4105
#: Borders(1) to Borders(4) are the old names for the four edges.
_INDEX = {1: 7, 2: 10, 3: 8, 4: 9, 5: 5, 6: 6, 7: 7, 8: 8, 9: 9, 10: 10, 11: 11, 12: 12}
EDGES = {7: "left", 8: "top", 9: "bottom", 10: "right"}
_OPPOSITE = {"left": ("right", 0, -1), "right": ("left", 0, 1), "top": ("bottom", -1, 0), "bottom": ("top", 1, 0)}


# --- reading and writing a range's styles ------------------------------------------------------


def styles_of(target: Range) -> list[S.Style]:
    """Every distinct format the range's positions show, without walking positions nobody has touched."""
    sheet = target.sheet
    found: dict[S.Style, None] = {}
    for area in target.areas:
        bottom, right = min(area.bottom, MAX_ROWS), min(area.right, MAX_COLUMNS)
        size = (bottom - area.top + 1) * (right - area.left + 1)
        if size <= 4096:
            for row in range(area.top, bottom + 1):
                for column in range(area.left, right + 1):
                    found[sheet.style_at(row, column)] = None
            continue
        found.update(dict.fromkeys(_row_formats.distinct_styles(sheet, area)))
    return list(found)


def uniform(values: Iterable[object]) -> object:
    """The one value every cell shares, or Null."""
    distinct: list[object] = []
    for value in values:
        if not any(value == seen and type(value) is type(seen) for seen in distinct):
            distinct.append(value)
    return distinct[0] if len(distinct) == 1 else NULL


def restyle(target: Range, change: Callable[[S.Style], S.Style]) -> None:
    """Change the format every position of the range shows: whole rows and columns through their own formats.

    On a filtered sheet only the visible cells change, as _visible has it.
    """
    sheet = target.sheet
    parts = [area for area in visible_areas(target) or target.areas
             if not _row_formats.format_area(sheet, area, change)]
    for row, column in _positions(parts):
        sheet.restyle(row, column, change(sheet.style_at(row, column)))


def _within_limit(areas: list[Area]) -> None:
    """Formatting more than 1,048,576 positions one by one reports itself; whole rows and columns go another way."""
    if any(area.rows * area.columns > 1048576 for area in areas):
        raise VBAUnsupportedError("formatting more than 1,048,576 cells at once, short of whole rows or "
                                  "columns, is not implemented")


def _positions(areas: list[Area]) -> list[tuple[int, int]]:
    """Every position of areas that are not whole rows or columns, which a format change visits one by one."""
    _within_limit(areas)
    return [(row, column) for area in areas for row in range(area.top, area.bottom + 1)
            for column in range(area.left, area.right + 1)]


def _long(value: int) -> VBAInt:
    return VBAInt(value, "Long")


def _colors(target: Range) -> S.Colors:
    return target.sheet.book.stylesheet.colors


def _color_answer(colors: S.Colors, color: S.Color | None, automatic: str, what: str) -> object:
    """One colour's Color, ColorIndex, ThemeColor or TintAndShade."""
    if what == "Color":
        return float(S.bgr(colors.rrggbb(color, automatic)))
    if what == "ColorIndex":
        return _long(colors.color_index(color, automatic))
    if what == "ThemeColor":
        if color is None or color.kind != "theme":
            raise error(5)
        return _long(int(color.value) + 1)
    return (color.tint / S.TINT_SCALE) if color is not None else 0.0


def _palette_color(value: object) -> S.Color | None:
    """ColorIndex n as a colour; None for automatic. xlColorIndexNone changes nothing."""
    index = int(to_integer(value, "Long"))
    if index == XL_AUTOMATIC:
        return None
    if not 1 <= index <= 56:
        raise error(9)
    return S.Color("indexed", index + 7)


def _theme_color(value: object) -> S.Color:
    index = int(to_integer(value, "Long"))
    if not 1 <= index <= 12:
        raise error(9)
    return S.Color("theme", index - 1)


def _tint(value: object) -> int:
    tint = float(to_number(value))
    if not -1 <= tint <= 1:
        raise error(5)
    return S.tint_of(tint)


# --- Font -----------------------------------------------------------------------------------------


class Font(ExcelObject):
    """A range's font, read over every cell of it and written to each."""

    vba_type_name = "Font"

    def __init__(self, target: Range) -> None:
        self.target = target

    def guard_set(self, member: str) -> None:
        from pyopenvba.apps.excel._protection import check_format_set

        check_format_set(self.target.sheet, self.vba_type_name, member)

    def _read(self, get: Callable[[S.Font], object]) -> object:
        return uniform(get(style.font) for style in styles_of(self.target))

    def _write(self, change: Callable[[S.Font], S.Font]) -> None:
        restyle(self.target, lambda style: S.applying(style, "font", font=change(style.font)))

    def _color(self, what: str) -> object:
        colors = {style.font.color for style in styles_of(self.target)}
        if len(colors) != 1:
            return NULL
        return _color_answer(_colors(self.target), next(iter(colors)), "000000", what)

    @member
    def Name(self) -> object:
        return self._read(lambda font: font.name)

    @setter("Name")
    def _set_name(self, value: object) -> None:
        name = to_text(value)
        if not name:
            return
        self._write(lambda font: font if font.name == name else replace(
            font, name=name, scheme="", family=S.font_family(name, font.family)))

    @member
    def Size(self) -> object:
        return self._read(lambda font: font.size)

    @setter("Size")
    def _set_size(self, value: object) -> None:
        size = float(to_number(value))
        if not 1 <= size <= 409.5:
            raise error(1004, "Unable to set the Size property of the Font class")
        self._write(lambda font: replace(font, size=size))

    @member
    def Bold(self) -> object:
        return self._read(lambda font: font.bold)

    @setter("Bold")
    def _set_bold(self, value: object) -> None:
        bold = to_bool(value)
        self._write(lambda font: replace(font, bold=bold))

    @member
    def Italic(self) -> object:
        return self._read(lambda font: font.italic)

    @setter("Italic")
    def _set_italic(self, value: object) -> None:
        italic = to_bool(value)
        self._write(lambda font: replace(font, italic=italic))

    @member
    def Strikethrough(self) -> object:
        return self._read(lambda font: font.strike)

    @setter("Strikethrough")
    def _set_strikethrough(self, value: object) -> None:
        strike = to_bool(value)
        self._write(lambda font: replace(font, strike=strike))

    @member
    def Superscript(self) -> object:
        return self._read(lambda font: font.vert_align == "superscript")

    @setter("Superscript")
    def _set_superscript(self, value: object) -> None:
        on = to_bool(value)
        self._write(lambda font: replace(font, vert_align="superscript" if on else (
            "" if font.vert_align == "superscript" else font.vert_align)))

    @member
    def Subscript(self) -> object:
        return self._read(lambda font: font.vert_align == "subscript")

    @setter("Subscript")
    def _set_subscript(self, value: object) -> None:
        on = to_bool(value)
        self._write(lambda font: replace(font, vert_align="subscript" if on else (
            "" if font.vert_align == "subscript" else font.vert_align)))

    @member
    def Underline(self) -> object:
        return self._read(lambda font: _long(UNDERLINES.get(font.underline, 2)))

    @setter("Underline")
    def _set_underline(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        if wanted not in _UNDERLINE_OF:
            raise VBAUnsupportedError(f"Font.Underline = {wanted} is not implemented")
        underline = _UNDERLINE_OF[wanted]
        self._write(lambda font: replace(font, underline=underline))

    @member
    def FontStyle(self) -> object:
        return self._read(lambda font: {(False, False): "Regular", (True, False): "Bold", (False, True): "Italic",
                                        (True, True): "Bold Italic"}[(font.bold, font.italic)])

    @setter("FontStyle")
    def _set_font_style(self, value: object) -> None:
        wanted = to_text(value).strip().lower()
        known = {"regular": (False, False), "bold": (True, False), "italic": (False, True),
                 "bold italic": (True, True)}
        if wanted not in known:
            return
        bold, italic = known[wanted]
        self._write(lambda font: replace(font, bold=bold, italic=italic))

    @member
    def Color(self) -> object:
        return self._color("Color")

    @setter("Color")
    def _set_color(self, value: object) -> None:
        color = S.Color.from_bgr(int(to_integer(value, "Long")))
        self._write(lambda font: replace(font, color=color))

    @member
    def ColorIndex(self) -> object:
        return self._color("ColorIndex")

    @setter("ColorIndex")
    def _set_color_index(self, value: object) -> None:
        index = int(to_integer(value, "Long"))
        if index == XL_NONE:
            return
        color = None if index == 0 else _palette_color(value)
        self._write(lambda font: replace(font, color=color))

    @member
    def ThemeColor(self) -> object:
        return self._color("ThemeColor")

    @setter("ThemeColor")
    def _set_theme_color(self, value: object) -> None:
        color = _theme_color(value)
        self._write(lambda font: replace(font, color=color))

    @member
    def TintAndShade(self) -> object:
        return self._color("TintAndShade")

    @setter("TintAndShade")
    def _set_tint(self, value: object) -> None:
        tint = _tint(value)

        def change(font: S.Font) -> S.Font:
            # An automatic colour takes the tint for as long as the workbook is open;
            # the file has nowhere to keep it, so a save drops it.
            color = font.color if font.color is not None else S.Color("auto")
            return replace(font, color=replace(color, tint=tint))

        self._write(change)

    @member
    def Parent(self) -> object:
        return self.target

    @member
    def Application(self) -> object:
        return self.target.sheet.book.application


# --- Interior -------------------------------------------------------------------------------------


def cell_color(fill: S.Fill) -> S.Color:
    """The colour a macro calls Interior.Color: a solid fill's foreground, any other's background."""
    return fill.foreground if fill.pattern == "solid" else fill.background


def pattern_color(fill: S.Fill) -> S.Color:
    return fill.background if fill.pattern == "solid" else fill.foreground


def with_cell_color(fill: S.Fill, color: S.Color) -> S.Fill:
    if fill.pattern == "none":
        return S.Fill("solid", foreground=color, background=S.FOREGROUND)
    if fill.pattern == "solid":
        return replace(fill, foreground=color)
    return replace(fill, background=color)


def with_pattern(fill: S.Fill, pattern: str) -> S.Fill:
    """A new pattern keeps the cell colour and the pattern colour, whichever side each lives on."""
    if pattern == "none":
        return S.Fill()
    if fill.pattern == "none":
        cell, lines = S.BACKGROUND, S.FOREGROUND
    else:
        cell, lines = cell_color(fill), pattern_color(fill)
    if pattern == "solid":
        return S.Fill(pattern, foreground=cell, background=lines)
    return S.Fill(pattern, foreground=lines, background=cell)


class Interior(ExcelObject):
    """A range's fill."""

    vba_type_name = "Interior"

    def __init__(self, target: Range) -> None:
        self.target = target

    def guard_set(self, member: str) -> None:
        from pyopenvba.apps.excel._protection import check_format_set

        check_format_set(self.target.sheet, self.vba_type_name, member)

    def _fills(self) -> list[S.Fill]:
        fills = list({style.fill: None for style in styles_of(self.target)})
        if any(fill.gradient for fill in fills):
            raise VBAUnsupportedError("a gradient fill's Interior is not implemented")
        return fills

    def _write(self, change: Callable[[S.Fill], S.Fill]) -> None:
        def apply(style: S.Style) -> S.Style:
            if style.fill.gradient:
                raise VBAUnsupportedError("a gradient fill's Interior is not implemented")
            return S.applying(style, "fill", fill=change(style.fill))

        restyle(self.target, apply)

    def _answer(self, pick: Callable[[S.Fill], S.Color], what: str, mixed: object) -> object:
        fills = self._fills()
        colors = {pick(fill) for fill in fills}
        if len(colors) != 1:
            return mixed
        color = next(iter(colors))
        automatic = "FFFFFF" if pick is cell_color else "000000"
        if what in ("ColorIndex", "ThemeColor") and all(fill.pattern == "none" for fill in fills):
            return _long(XL_NONE)
        if what == "ThemeColor":
            return _long(int(color.value) + 1 if color.kind == "theme" else 0)
        return _color_answer(_colors(self.target), color, automatic, what)

    @member
    def Color(self) -> object:
        return self._answer(cell_color, "Color", 0.0)

    @setter("Color")
    def _set_color(self, value: object) -> None:
        color = S.Color.from_bgr(int(to_integer(value, "Long")))
        self._write(lambda fill: with_cell_color(fill, color))

    @member
    def ColorIndex(self) -> object:
        return self._answer(cell_color, "ColorIndex", NULL)

    @setter("ColorIndex")
    def _set_color_index(self, value: object) -> None:
        if int(to_integer(value, "Long")) in (XL_NONE, 0):
            self._write(lambda _: S.Fill())
            return
        color = _palette_color(value) or S.BACKGROUND
        self._write(lambda fill: with_cell_color(fill, color))

    @member
    def ThemeColor(self) -> object:
        return self._answer(cell_color, "ThemeColor", NULL)

    @setter("ThemeColor")
    def _set_theme_color(self, value: object) -> None:
        color = _theme_color(value)
        self._write(lambda fill: with_cell_color(fill, color))

    @member
    def TintAndShade(self) -> object:
        return self._answer(cell_color, "TintAndShade", NULL)

    @setter("TintAndShade")
    def _set_tint(self, value: object) -> None:
        tint = _tint(value)

        def change(fill: S.Fill) -> S.Fill:
            if fill.pattern == "none":
                return fill
            color = cell_color(fill)
            if color in (S.FOREGROUND, S.BACKGROUND) or color.kind == "auto":
                raise VBAUnsupportedError("TintAndShade on an automatic fill colour is not implemented")
            return with_cell_color(fill, replace(color, tint=tint))

        self._write(change)

    @member
    def Pattern(self) -> object:
        found = uniform(fill.pattern for fill in self._fills())
        return found if found is NULL else _long(PATTERNS.get(str(found), 1))

    @setter("Pattern")
    def _set_pattern(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        if wanted == XL_AUTOMATIC:
            wanted = 1  # xlPatternAutomatic is a solid fill
        if wanted not in _PATTERN_OF:
            raise VBAUnsupportedError(f"Interior.Pattern = {wanted} is not implemented")
        pattern = _PATTERN_OF[wanted]
        self._write(lambda fill: with_pattern(fill, pattern))

    @member
    def PatternColor(self) -> object:
        found = self._answer(pattern_color, "Color", 0.0)
        return _long(int(found)) if isinstance(found, float) else found

    @setter("PatternColor")
    def _set_pattern_color(self, value: object) -> None:
        color = S.Color.from_bgr(int(to_integer(value, "Long")))
        self._write(lambda fill: self._with_pattern_color(fill, color))

    @staticmethod
    def _with_pattern_color(fill: S.Fill, color: S.Color) -> S.Fill:
        if fill.pattern == "none":
            # A pattern colour given to a cell with no fill makes it solid.
            return S.Fill("solid", foreground=S.BACKGROUND, background=color)
        if fill.pattern == "solid":
            return replace(fill, background=color)
        return replace(fill, foreground=color)

    @member
    def PatternColorIndex(self) -> object:
        return self._answer(pattern_color, "ColorIndex", NULL)

    @setter("PatternColorIndex")
    def _set_pattern_color_index(self, value: object) -> None:
        if int(to_integer(value, "Long")) == XL_NONE:
            self._write(lambda _: S.Fill())
            return
        color = _palette_color(value) or S.FOREGROUND
        self._write(lambda fill: self._with_pattern_color(fill, color))

    @member
    def PatternThemeColor(self) -> object:
        return self._answer(pattern_color, "ThemeColor", NULL)

    @setter("PatternThemeColor")
    def _set_pattern_theme_color(self, value: object) -> None:
        color = _theme_color(value)
        self._write(lambda fill: self._with_pattern_color(fill, color))

    @member
    def PatternTintAndShade(self) -> object:
        return self._answer(pattern_color, "TintAndShade", NULL)

    @setter("PatternTintAndShade")
    def _set_pattern_tint(self, value: object) -> None:
        tint = _tint(value)

        def change(fill: S.Fill) -> S.Fill:
            color = pattern_color(fill)
            if fill.pattern == "none" or color in (S.FOREGROUND, S.BACKGROUND) or color.kind == "auto":
                raise VBAUnsupportedError("PatternTintAndShade on an automatic pattern colour is not implemented")
            return self._with_pattern_color(fill, replace(color, tint=tint))

        self._write(change)

    @member
    def Parent(self) -> object:
        return self.target

    @member
    def Application(self) -> object:
        return self.target.sheet.book.application


# --- borders ------------------------------------------------------------------------------------------


def own_side(sheet: Worksheet, row: int, column: int, side: str) -> S.Side:
    side_of: S.Side = getattr(sheet.style_at(row, column).border, side)
    return side_of


def effective_side(sheet: Worksheet, row: int, column: int, side: str) -> S.Side:
    """The edge a cell shows: its own side, else its neighbour's side facing it.

    Two cells can both draw the edge between them where their cell styles
    give it -- a macro's border clears the neighbour's -- and then the
    stronger line shows, and between two of one line style the darker
    colour (tests/fixtures/cell_styles/shared_edges.json).
    """
    mine = own_side(sheet, row, column, side)
    opposite, down, across = _OPPOSITE[side]
    row, column = row + down, column + across
    theirs = own_side(sheet, row, column, opposite) if 1 <= row <= MAX_ROWS and 1 <= column <= MAX_COLUMNS \
        else S.NO_SIDE
    if mine.style and theirs.style and mine != theirs:
        return _stronger(sheet.book.stylesheet.colors, mine, theirs)
    return mine if mine.style else theirs if theirs.style else S.NO_SIDE


#: The file's line styles from the weakest to the one that shows over every other on an edge two cells draw.
_LINE_STRENGTH: Final = ("hair", "dashDotDot", "dashDot", "dotted", "dashed", "thin", "mediumDashDotDot",
                         "slantDashDot", "mediumDashDot", "mediumDashed", "medium", "thick", "double")


def _stronger(colors: S.Colors, one: S.Side, other: S.Side) -> S.Side:
    """Which of two sides shows on an edge: the stronger line style, then the darker colour, 2R + 5G + B."""

    def strength(side: S.Side) -> tuple[int, int]:
        rrggbb = colors.rrggbb(side.color, "000000")
        red, green, blue = (int(rrggbb[index:index + 2], 16) for index in (0, 2, 4))
        return _LINE_STRENGTH.index(side.style) if side.style in _LINE_STRENGTH else -1, -(2 * red + 5 * green + blue)

    first, second = strength(one), strength(other)
    if first == second:
        raise VBAUnsupportedError("which of two borders shows on an edge two cells draw in different colours of "
                                  "one darkness is not implemented")
    return one if first > second else other


def _diagonal(border: S.Border, index: int) -> S.Side:
    shown = border.diagonal_down if index == 5 else border.diagonal_up
    return border.diagonal if shown else S.NO_SIDE


def with_diagonal(border: S.Border, index: int, change: Callable[[S.Side], S.Side]) -> S.Border:
    """A border with one diagonal changed.

    Both diagonals share one line: xlDiagonalDown (5) and xlDiagonalUp (6)
    only say which of them show it.
    """
    side = change(_diagonal(border, index))
    up, down = border.diagonal_up, border.diagonal_down
    if index == 6:
        up = bool(side.style)
    else:
        down = bool(side.style)
    shared = side if side.style else (border.diagonal if up or down else S.NO_SIDE)
    return replace(border, diagonal=shared, diagonal_up=up, diagonal_down=down)


def change_side(sheet: Worksheet, row: int, column: int, index: int, change: Callable[[S.Side], S.Side],
                *, clear_neighbour: bool = True) -> None:
    """Set one edge of one cell the way Excel does it."""
    style = sheet.style_at(row, column)
    sides = {name: effective_side(sheet, row, column, name) for name in ("left", "right", "top", "bottom")}
    border = replace(style.border, **sides)
    if index in (5, 6):
        sheet.restyle(row, column, S.applying(style, "border", border=with_diagonal(border, index, change)))
        return
    name = EDGES[index]
    border = replace(border, **{name: change(getattr(border, name))})
    sheet.restyle(row, column, S.applying(style, "border", border=border))
    if not clear_neighbour:
        return
    opposite, down, across = _OPPOSITE[name]
    row, column = row + down, column + across
    if 1 <= row <= MAX_ROWS and 1 <= column <= MAX_COLUMNS and own_side(sheet, row, column, opposite).style:
        theirs = sheet.style_at(row, column)
        cleared = replace(theirs.border, **{opposite: S.NO_SIDE})
        sheet.restyle(row, column, S.applying(theirs, "border", border=cleared))


def inside_current(upper: S.Side, lower: S.Side) -> S.Side:
    """What an inside border between two cells is before a change, from the sides each one stores.

    When only one of them stores it, Excel keeps its line and weight but
    not its colour, which becomes automatic and is written ``auto="1"``.
    """
    if upper == lower:
        return upper
    if not upper.style or not lower.style:
        found = upper if upper.style else lower
        return S.Side(found.style, S.Color("auto"))
    raise VBAUnsupportedError("an inside border whose two cells store different lines is not implemented")


def with_line_style(side: S.Side, line_style: int) -> S.Side:
    if line_style == XL_NONE:
        return S.NO_SIDE
    if line_style not in _DEFAULT_WEIGHT:
        raise error(1004, "Unable to set the LineStyle property of the Border class")
    weight = LINES.get(side.style, (1, 2))[1] if side.style else 2
    if (line_style, weight) not in _STYLE_OF:
        weight = _DEFAULT_WEIGHT[line_style]
    return S.Side(_STYLE_OF[(line_style, weight)], side.color or S.FOREGROUND)


def with_weight(side: S.Side, weight: int) -> S.Side:
    if weight == 3:
        weight = -4138  # Excel reads 3 as xlMedium
    if weight not in _WEIGHTS:
        raise error(1004, "Unable to set the Weight property of the Border class")
    line = LINES.get(side.style, (1, 2))[0] if side.style else 1
    if (line, weight) not in _STYLE_OF:
        line = 1
    return S.Side(_STYLE_OF[(line, weight)], side.color or S.FOREGROUND)


def with_side_color(side: S.Side, color: S.Color) -> S.Side:
    return S.Side(side.style or "thin", color)


def side_answer(colors: S.Colors, side: S.Side, what: str) -> object:
    if what == "LineStyle":
        return _long(LINES.get(side.style, (1, 2))[0] if side.style else XL_NONE)
    if what == "Weight":
        return _long(LINES.get(side.style, (1, 2))[1] if side.style else 2)
    if not side.style:
        return {"Color": 0.0, "ColorIndex": _long(XL_NONE)}.get(what, NULL)
    return _color_answer(colors, side.color, "000000", what)


def _side_change(what: str, value: object) -> Callable[[S.Side], S.Side]:
    """How setting one Border property changes a side."""
    if what == "LineStyle":
        line_style = int(to_integer(value, "Long"))
        return lambda side: with_line_style(side, line_style)
    if what == "Weight":
        weight = int(to_integer(value, "Long"))
        return lambda side: with_weight(side, weight)
    if what == "Color":
        color = S.Color.from_bgr(int(to_integer(value, "Long")))
        return lambda side: with_side_color(side, color)
    if what == "ColorIndex":
        if int(to_integer(value, "Long")) == XL_NONE:
            return lambda side: side
        palette = _palette_color(value) or S.FOREGROUND
        return lambda side: with_side_color(side, palette)
    if what == "ThemeColor":
        theme = _theme_color(value)
        return lambda side: with_side_color(side, theme)
    tint = _tint(value)

    def tinted(side: S.Side) -> S.Side:
        # A tint draws a missing border, but only an RGB or theme colour takes it.
        if not side.style:
            return S.Side("thin", S.FOREGROUND)
        if side.color is None or side.color.kind not in ("rgb", "theme"):
            return side
        return S.Side(side.style, replace(side.color, tint=tint))

    return tinted


def _exactly(side: S.Side) -> Callable[[S.Side], S.Side]:
    return lambda _: side


def _inside_pairs(area: Area, index: int) -> list[tuple[tuple[int, int], tuple[int, int]]]:
    """The cells either side of each inside edge: above and below for xlInsideHorizontal, else left and right."""
    if index == 12:
        return [((row, column), (row + 1, column)) for row in range(area.top, area.bottom)
                for column in range(area.left, area.right + 1)]
    return [((row, column), (row, column + 1)) for row in range(area.top, area.bottom + 1)
            for column in range(area.left, area.right)]


class Border(ExcelObject):
    """One edge, one inside edge or one diagonal of a range."""

    vba_type_name = "Border"

    def __init__(self, target: Range, index: int) -> None:
        self.target = target
        self.index = index

    def guard_set(self, member: str) -> None:
        from pyopenvba.apps.excel._protection import check_format_set

        check_format_set(self.target.sheet, self.vba_type_name, member)

    def side(self) -> S.Side:
        """What the range answers for this border: its first cell along it."""
        sheet, area = self.target.sheet, self.target.first
        if self.index in EDGES:
            name = EDGES[self.index]
            row = area.bottom if name == "bottom" else area.top
            column = area.right if name == "right" else area.left
            return effective_side(sheet, row, column, name)
        if self.index == 11:
            return effective_side(sheet, area.top, area.left, "right") if area.columns > 1 else S.NO_SIDE
        if self.index == 12:
            return effective_side(sheet, area.top, area.left, "bottom") if area.rows > 1 else S.NO_SIDE
        return _diagonal(sheet.style_at(area.top, area.left).border, self.index)

    def answer(self, what: str) -> object:
        return side_answer(_colors(self.target), self.side(), what)

    def assign(self, what: str, value: object) -> None:
        self.apply(_side_change(what, value))

    def apply(self, change: Callable[[S.Side], S.Side]) -> None:
        """Set the border on each area -- each visible area, on a filtered sheet, measured for the bottom edge."""
        sheet = self.target.sheet
        parts = [area for area in visible_areas(self.target) or self.target.areas
                 if not _row_formats.border_area(sheet, area, self.index, change)]
        _within_limit(parts)
        for area in parts:
            if self.index in (5, 6):
                for row in range(area.top, area.bottom + 1):
                    for column in range(area.left, area.right + 1):
                        change_side(sheet, row, column, self.index, change)
            elif self.index in EDGES:
                name = EDGES[self.index]
                if name in ("top", "bottom"):
                    row = area.top if name == "top" else area.bottom
                    for column in range(area.left, area.right + 1):
                        change_side(sheet, row, column, self.index, change)
                else:
                    column = area.left if name == "left" else area.right
                    for row in range(area.top, area.bottom + 1):
                        change_side(sheet, row, column, self.index, change)
            else:
                # An inside border is written on both of the cells it separates.
                for (row, column), (other_row, other_column) in _inside_pairs(area, self.index):
                    first, second = (9, 8) if self.index == 12 else (10, 7)
                    wanted = change(inside_current(own_side(sheet, row, column, EDGES[first]),
                                                   own_side(sheet, other_row, other_column, EDGES[second])))
                    change_side(sheet, row, column, first, _exactly(wanted), clear_neighbour=False)
                    change_side(sheet, other_row, other_column, second, _exactly(wanted), clear_neighbour=False)

    def sides(self) -> list[S.Side]:
        """Every side along this border, which a property of the whole collection answers over."""
        sheet, area = self.target.sheet, self.target.first
        # Along a whole row or column most positions stand for each other: a sample of them answers.
        columns = (range(area.left, area.right + 1) if area.columns <= 4096 else
                   _row_formats.sample_columns(sheet, area.left, area.right, area.top, area.bottom))
        rows = (range(area.top, area.bottom + 1) if area.rows <= 4096 else
                _row_formats.sample_rows(sheet, area.top, area.bottom, area.left, area.right))
        if self.index in EDGES:
            name = EDGES[self.index]
            if name in ("top", "bottom"):
                row = area.top if name == "top" else area.bottom
                return [effective_side(sheet, row, column, name) for column in columns]
            column = area.left if name == "left" else area.right
            return [effective_side(sheet, row, column, name) for row in rows]
        if self.index == 12:
            return [effective_side(sheet, row, column, "bottom") for row in rows if row < area.bottom
                    for column in columns]
        return [effective_side(sheet, row, column, "right") for row in rows for column in columns
                if column < area.right]

    @member
    def LineStyle(self) -> object:
        return self.answer("LineStyle")

    @setter("LineStyle")
    def _set_line_style(self, value: object) -> None:
        self.assign("LineStyle", value)

    @member
    def Weight(self) -> object:
        return self.answer("Weight")

    @setter("Weight")
    def _set_weight(self, value: object) -> None:
        self.assign("Weight", value)

    @member
    def Color(self) -> object:
        return self.answer("Color")

    @setter("Color")
    def _set_color(self, value: object) -> None:
        self.assign("Color", value)

    @member
    def ColorIndex(self) -> object:
        return self.answer("ColorIndex")

    @setter("ColorIndex")
    def _set_color_index(self, value: object) -> None:
        self.assign("ColorIndex", value)

    @member
    def ThemeColor(self) -> object:
        return self.answer("ThemeColor")

    @setter("ThemeColor")
    def _set_theme_color(self, value: object) -> None:
        self.assign("ThemeColor", value)

    @member
    def TintAndShade(self) -> object:
        return self.answer("TintAndShade")

    @setter("TintAndShade")
    def _set_tint(self, value: object) -> None:
        self.assign("TintAndShade", value)

    @member
    def Parent(self) -> object:
        return self.target

    @member
    def Application(self) -> object:
        return self.target.sheet.book.application


class Borders(VBACollection, ExcelObject):
    """A range's borders: Borders(xlEdgeTop) is one, a property set here sets the edges and insides."""

    vba_type_name = "Borders"

    def __init__(self, target: Range) -> None:
        self.target = target

    def guard_set(self, member: str) -> None:
        from pyopenvba.apps.excel._protection import check_format_set

        check_format_set(self.target.sheet, self.vba_type_name, member)

    def vba_items(self) -> list[object]:
        return [Border(self.target, index) for index in (7, 8, 9, 10, 5, 6)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        number = int(to_integer(index, "Long"))
        if number not in _INDEX:
            raise error(1004, "Unable to get the Item property of the Borders class")
        return Border(self.target, _INDEX[number])

    def _indexes(self) -> list[int]:
        """The borders a property of the collection covers: the edges, and the insides a range has."""
        area = self.target.first
        found = [7, 8, 9, 10]
        if area.columns > 1:
            found.append(11)
        if area.rows > 1:
            found.append(12)
        return found

    def _all(self, what: str) -> object:
        """Every cell along every edge and inside edge; mixed is Null, and a mixed Color is 0."""
        colors = _colors(self.target)
        found = uniform(side_answer(colors, side, what)
                        for index in self._indexes() for side in Border(self.target, index).sides())
        return 0.0 if found is NULL and what == "Color" else found

    def _set_all(self, what: str, value: object) -> None:
        for index in self._indexes():
            Border(self.target, index).assign(what, value)

    @member
    def LineStyle(self) -> object:
        return self._all("LineStyle")

    @setter("LineStyle")
    def _set_line_style(self, value: object) -> None:
        self._set_all("LineStyle", value)

    @member
    def Value(self) -> object:
        return self._all("LineStyle")

    @setter("Value")
    def _set_value(self, value: object) -> None:
        self._set_all("LineStyle", value)

    @member
    def Weight(self) -> object:
        return self._all("Weight")

    @setter("Weight")
    def _set_weight(self, value: object) -> None:
        self._set_all("Weight", value)

    @member
    def Color(self) -> object:
        return self._all("Color")

    @setter("Color")
    def _set_color(self, value: object) -> None:
        self._set_all("Color", value)

    @member
    def ColorIndex(self) -> object:
        return self._all("ColorIndex")

    @setter("ColorIndex")
    def _set_color_index(self, value: object) -> None:
        self._set_all("ColorIndex", value)

    @member
    def ThemeColor(self) -> object:
        return self._all("ThemeColor")

    @setter("ThemeColor")
    def _set_theme_color(self, value: object) -> None:
        self._set_all("ThemeColor", value)

    @member
    def TintAndShade(self) -> object:
        return self._all("TintAndShade")

    @setter("TintAndShade")
    def _set_tint(self, value: object) -> None:
        self._set_all("TintAndShade", value)

    @member
    def Parent(self) -> object:
        return self.target

    @member
    def Application(self) -> object:
        return self.target.sheet.book.application


def border_around(target: Range, line_style: object, weight: object, color_index: object, color: object,
                  theme_color: object) -> None:
    """Range.BorderAround: the four edges, each set to the line, weight and colour given."""
    for index in (7, 8, 9, 10):
        border = Border(target, index)
        border.assign("LineStyle", 1 if line_style is MISSING else line_style)
        for what, given in (("Weight", weight), ("ColorIndex", color_index), ("Color", color),
                            ("ThemeColor", theme_color)):
            if given is not MISSING:
                border.assign(what, given)


# --- alignment and protection -------------------------------------------------------------------------


def read_alignment(target: Range, what: str) -> object:
    values = [_alignment_answer(style, what) for style in styles_of(target)]
    return uniform(values)


def _alignment_answer(style: S.Style, what: str) -> object:
    alignment = style.alignment
    if what == "HorizontalAlignment":
        return _long(HORIZONTAL.get(alignment.horizontal, 1))
    if what == "VerticalAlignment":
        return _long(VERTICAL.get(alignment.vertical, -4107))
    if what == "WrapText":
        return alignment.wrap
    if what == "ShrinkToFit":
        return alignment.shrink
    if what == "AddIndent":
        return alignment.justify_last_line
    if what == "IndentLevel":
        return _long(alignment.indent)
    if what == "ReadingOrder":
        return _long(_READING.get(alignment.reading_order, -5002))
    if what == "Orientation":
        rotation = alignment.text_rotation
        special = {0: -4128, 255: -4166, 90: -4171, 180: -4170}
        return _long(special.get(rotation, rotation if rotation <= 90 else 90 - rotation))
    if what == "Locked":
        return style.protection.locked
    return style.protection.hidden


def write_alignment(target: Range, what: str, value: object) -> None:
    if what in ("Locked", "FormulaHidden"):
        on = to_bool(value)
        field_name = "locked" if what == "Locked" else "hidden"
        restyle(target, lambda style: S.applying(style, "protection",
                                                 protection=replace(style.protection, **{field_name: on})))
        return
    change = _alignment_change(what, value)
    restyle(target, lambda style: S.applying(style, "alignment", alignment=change(style.alignment)))


def _alignment_change(what: str, value: object) -> Callable[[S.Alignment], S.Alignment]:
    if what in ("WrapText", "ShrinkToFit", "AddIndent"):
        on = to_bool(value)
        name = {"WrapText": "wrap", "ShrinkToFit": "shrink", "AddIndent": "justify_last_line"}[what]
        return lambda alignment: replace(alignment, **{name: on})
    number = int(to_integer(value, "Long"))
    if what == "HorizontalAlignment":
        if number not in _HORIZONTAL_OF:
            raise error(1004, "Unable to set the HorizontalAlignment property of the Range class")
        horizontal = _HORIZONTAL_OF[number]
        keeps_indent = horizontal in ("left", "right", "distributed")
        return lambda alignment: replace(alignment, horizontal=horizontal,
                                         indent=alignment.indent if keeps_indent else 0)
    if what == "VerticalAlignment":
        if number not in _VERTICAL_OF:
            raise error(1004, "Unable to set the VerticalAlignment property of the Range class")
        vertical = _VERTICAL_OF[number]
        return lambda alignment: replace(alignment, vertical=vertical)
    if what == "IndentLevel":
        if not 0 <= number <= 250:
            raise error(1004, "Unable to set the IndentLevel property of the Range class")

        def indent(alignment: S.Alignment) -> S.Alignment:
            horizontal = alignment.horizontal
            if number and horizontal not in ("left", "right", "distributed"):
                horizontal = "left"
            return replace(alignment, indent=number, horizontal=horizontal)

        return indent
    if what == "ReadingOrder":
        if number not in _READING_OF:
            raise error(1004, "Unable to set the ReadingOrder property of the Range class")
        order = _READING_OF[number]
        return lambda alignment: replace(alignment, reading_order=order)
    # Orientation
    special = {-4128: 0, -4166: 255, -4171: 90, -4170: 180}
    if number in special:
        rotation = special[number]
    elif -90 <= number <= 90:
        rotation = number if number >= 0 else 90 - number
    else:
        raise error(1004, "Unable to set the Orientation property of the Range class")
    return lambda alignment: replace(alignment, text_rotation=rotation)


def clear_formats(target: Range) -> None:
    """Range.ClearFormats: every cell back to the default format, and merged cells unmerged.

    On a filtered sheet only the visible cells, as _visible has it.
    """
    from pyopenvba.apps.excel import _merges
    from pyopenvba.apps.excel._visible import visible

    target = visible(target)
    _merges.unmerge(target)
    sheet = target.sheet
    parts = [area for area in target.areas if not _row_formats.clear_area(sheet, area)]
    default = sheet.book.stylesheet.default
    for row, column in _positions(parts):
        # A position a row or column format reaches keeps a cell in the default format.
        sheet.restyle(row, column, default)
    sheet.touched()
