"""Cell styles by name: Range.Style, Workbook.Styles and the Style object, as live Excel keeps them.

Every rule here was measured (scripts/measure_cell_styles.py, tests/fixtures/cell_styles):

* A new workbook's Styles collection holds Excel's 47 standard built-in
  styles, and its file only Normal until a cell uses another. Styles
  lists them, and any a macro adds, in the order Windows word sort puts
  their names; a name finds its style in any case.
* Giving a cell a style replaces the parts of its format the style
  includes -- IncludeNumber, IncludeFont, IncludeAlignment,
  IncludeBorder, IncludePatterns, IncludeProtection -- with the style's
  and keeps the rest; a border the style includes replaces all four
  sides. A range whose cells have different styles answers Nothing for
  Style; a name no style has, or anything but a name or a Style, is
  error 450.
* A save writes a built-in style only while a cell uses it, and every
  style a macro added whether used or not: after the file's own, the
  built-in ones in Excel's order, then the rest in the order they were
  made. A built-in style's fonts are the theme's body or heading font;
  one that wants the body font as it is gets a copy of it written first,
  and one that wants the Normal font gets a copy of that, never the
  file's first font itself (Stylesheet.write_cell_styles).
* Styles.Add copies Normal, or the first cell of the range it is based
  on, whose parts the range's cells do not share it leaves out.
"""

from __future__ import annotations

from dataclasses import replace
from typing import TYPE_CHECKING, Final

from pyopenvba.apps.excel import _styles as S
from pyopenvba.apps.excel._formats import (
    PATTERNS,
    UNDERLINES,
    XL_NONE,
    cell_color,
    side_answer,
    styles_of,
)
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method
from pyopenvba.interpreter._values import MISSING, NOTHING, NULL, VBAInt, error, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook

# --- Excel's built-in styles -------------------------------------------------------------------------

#: Excel writes a built-in style's font in the theme's body font ({minor}) or heading font ({major}).
_BODY = '<font><sz val="11"/><color theme="1"/><name val="{minor}"/><family val="2"/><scheme val="minor"/></font>'
_NO_FILL = '<fill><patternFill patternType="none"/></fill>'
_NO_BORDER = "<border><left/><right/><top/><bottom/><diagonal/></border>"


def _font(color: str, size: int = 11, *, bold: bool = False, italic: bool = False, heading: bool = False) -> str:
    kind = "major" if heading else "minor"
    return ("<font>" + ("<b/>" if bold else "") + ("<i/>" if italic else "") + f'<sz val="{size}"/><color {color}/>'
            f'<name val="{{{kind}}}"/><family val="2"/><scheme val="{kind}"/></font>')


def _solid(color: str) -> str:
    return f'<fill><patternFill patternType="solid"><fgColor {color}/></patternFill></fill>'


def _tinted(theme: int, tint: str) -> str:
    return (f'<fill><patternFill patternType="solid"><fgColor theme="{theme}" tint="{tint}"/><bgColor indexed="65"/>'
            "</patternFill></fill>")


def _boxed(line: str, color: str) -> str:
    return ("<border>" + "".join(f'<{edge} style="{line}"><color rgb="{color}"/></{edge}>'
                                 for edge in ("left", "right", "top", "bottom")) + "<diagonal/></border>")


def _under(line: str, color: str) -> str:
    return f'<border><left/><right/><top/><bottom style="{line}"><color {color}/></bottom><diagonal/></border>'


#: The parts a style's Include properties name, by the letters the table below gives them.
_PARTS = {"N": "number_format", "F": "font", "A": "alignment", "B": "border", "P": "fill", "R": "protection"}
_TINTS = ("0.79998168889431442", "0.59999389629810485", "0.39997558519241921")

#: Excel's standard built-in styles after Normal, in its own order, which a save writes new ones in: name,
#: builtinId, the parts it includes, its number format's id, its font (None: a copy of the Normal font), its
#: fill and its border, as Excel writes them (tests/fixtures/cell_styles/every_style.xlsx).
STANDARD: Final[tuple[tuple[str, int, str, int, str | None, str, str], ...]] = (
    ("Comma", 3, "N", 43, None, _NO_FILL, _NO_BORDER),
    ("Comma [0]", 6, "N", 41, None, _NO_FILL, _NO_BORDER),
    ("Currency", 4, "N", 44, None, _NO_FILL, _NO_BORDER),
    ("Currency [0]", 7, "N", 42, None, _NO_FILL, _NO_BORDER),
    ("Percent", 5, "N", 9, None, _NO_FILL, _NO_BORDER),
    ("Title", 15, "F", 0, _font('theme="3"', 18, heading=True), _NO_FILL, _NO_BORDER),
    ("Heading 1", 16, "FB", 0, _font('theme="3"', 15, bold=True), _NO_FILL, _under("thick", 'theme="4"')),
    ("Heading 2", 17, "FB", 0, _font('theme="3"', 13, bold=True), _NO_FILL,
     _under("thick", 'theme="4" tint="0.499984740745262"')),
    ("Heading 3", 18, "FB", 0, _font('theme="3"', bold=True), _NO_FILL,
     _under("medium", 'theme="4" tint="0.39997558519241921"')),
    ("Heading 4", 19, "F", 0, _font('theme="3"', bold=True), _NO_FILL, _NO_BORDER),
    ("Good", 26, "FP", 0, _font('rgb="FF006100"'), _solid('rgb="FFC6EFCE"'), _NO_BORDER),
    ("Bad", 27, "FP", 0, _font('rgb="FF9C0006"'), _solid('rgb="FFFFC7CE"'), _NO_BORDER),
    ("Neutral", 28, "FP", 0, _font('rgb="FF9C5700"'), _solid('rgb="FFFFEB9C"'), _NO_BORDER),
    ("Input", 20, "FBP", 0, _font('rgb="FF3F3F76"'), _solid('rgb="FFFFCC99"'), _boxed("thin", "FF7F7F7F")),
    ("Output", 21, "FBP", 0, _font('rgb="FF3F3F3F"', bold=True), _solid('rgb="FFF2F2F2"'),
     _boxed("thin", "FF3F3F3F")),
    ("Calculation", 22, "FBP", 0, _font('rgb="FFFA7D00"', bold=True), _solid('rgb="FFF2F2F2"'),
     _boxed("thin", "FF7F7F7F")),
    ("Linked Cell", 24, "FB", 0, _font('rgb="FFFA7D00"'), _NO_FILL, _under("double", 'rgb="FFFF8001"')),
    ("Check Cell", 23, "FBP", 0, _font('theme="0"', bold=True), _solid('rgb="FFA5A5A5"'),
     _boxed("double", "FF3F3F3F")),
    ("Warning Text", 11, "F", 0, _font('rgb="FFFF0000"'), _NO_FILL, _NO_BORDER),
    ("Note", 10, "BP", 0, None, _solid('rgb="FFFFFFCC"'), _boxed("thin", "FFB2B2B2")),
    ("Explanatory Text", 53, "F", 0, _font('rgb="FF7F7F7F"', italic=True), _NO_FILL, _NO_BORDER),
    ("Total", 25, "FB", 0, _font('theme="1"', bold=True), _NO_FILL,
     '<border><left/><right/><top style="thin"><color theme="4"/></top><bottom style="double"><color theme="4"/>'
     "</bottom><diagonal/></border>"),
    *(entry for accent in range(1, 7) for entry in (
        (f"Accent{accent}", 25 + accent * 4, "FP", 0, _font('theme="0"'), _solid(f'theme="{accent + 3}"'),
         _NO_BORDER),
        *((f"{shade}% - Accent{accent}", 26 + accent * 4 + step, "FP", 0, _BODY, _tinted(accent + 3, tint),
           _NO_BORDER) for step, (shade, tint) in enumerate(zip((20, 40, 60), _TINTS, strict=True)))))
)
#: The standard styles' places in Excel's order, Normal being 0, by name in lower case.
_RANKS: Final = {name.casefold(): rank for rank, (name, *_) in enumerate(STANDARD, start=1)}
#: The theme typefaces Excel writes with family 2, the only ones a built-in style is written in here.
_SWISS: Final = frozenset({"Aptos", "Aptos Narrow", "Aptos Display", "Calibri", "Calibri Light"})


def _body_font(stylesheet: S.Stylesheet) -> S.Font | None:
    """The theme's body font as Excel writes it for a built-in style; None for a typeface none is written in."""
    minor = stylesheet.theme_fonts[1]
    return S.parse_font(_BODY.replace("{minor}", _typeface(minor))) if minor in _SWISS else None


def _typeface(name: str) -> str:
    if name not in _SWISS:
        raise VBAUnsupportedError(f"a built-in cell style in a theme whose fonts are {name!r} is not implemented")
    from pyopenvba._xml import escape

    return escape(name)


def _normal(stylesheet: S.Stylesheet) -> S.CellStyle:
    found = stylesheet.named("Normal")
    return stylesheet.cell_styles[0 if found is None else found]


def _built_in(stylesheet: S.Stylesheet, name: str) -> S.CellStyle | None:
    """A standard built-in style the file does not hold, as Excel defines it under the workbook's theme."""
    rank = _RANKS.get(name.casefold())
    if rank is None:
        return None
    title, builtin, letters, number, font, fill, border = STANDARD[rank - 1]
    if font is None:
        value = _normal(stylesheet).format.font
    else:
        major, minor = stylesheet.theme_fonts
        value = S.parse_font(font.replace("{major}", _typeface(major)).replace("{minor}", _typeface(minor)))
    own = S.Style(number_format=stylesheet.number_formats[number], font=value, fill=S.parse_fill(fill),
                  border=S.parse_border(border), styled=S.PARTS)
    return S.CellStyle(own, frozenset(_PARTS[letter] for letter in letters), name=title, builtin=builtin, rank=rank)


def find(book: Workbook, name: str) -> int | None:
    """The place of the style a name names, in any case: the stylesheet's, or a built-in one it now holds."""
    stylesheet = book.stylesheet
    found = stylesheet.named(name)
    if found is not None:
        return found
    entry = _built_in(stylesheet, name)
    return None if entry is None else stylesheet.add_cell_style(entry)


def _listed(book: Workbook) -> list[str]:
    """Every style's name, as the Styles collection lists them."""
    from pyopenvba.apps.excel._sort import text_key

    stylesheet = book.stylesheet
    names = {entry.name.casefold(): entry.name for entry in stylesheet.cell_styles if entry.name}
    for name, *_ in STANDARD:
        names.setdefault(name.casefold(), name)
    return sorted(names.values(), key=lambda name: text_key(name, False))


# --- giving cells a style --------------------------------------------------------------------------------


def given(style: S.Style, index: int, entry: S.CellStyle) -> S.Style:
    """A format given a cell style: the parts the style includes become the style's own, the rest stay."""
    own = entry.format
    changes = {part: getattr(own, part) for part in entry.includes}
    return replace(style, base=index, applied=style.applied - entry.includes, styled=entry.includes,
                   **changes)  # type: ignore[arg-type]


def style_of(target: Range) -> object:
    """Range.Style: the style every cell of the range has, or Nothing where they differ."""
    bases = {style.base for style in styles_of(target)}
    if len(bases) != 1:
        return NOTHING
    entry = target.sheet.book.stylesheet.cell_styles[bases.pop()]
    if not entry.name:
        raise VBAUnsupportedError("Range.Style of a cell whose cell style has no name is not implemented")
    return StyleObject(target.sheet.book, entry.name)


def set_style(target: Range, value: object) -> None:
    """Give every cell of the range a style, named or as a Style object."""
    from pyopenvba.apps.excel._formats import restyle

    book = target.sheet.book
    if isinstance(value, StyleObject):
        if value.book is not book:
            raise VBAUnsupportedError("giving a range a Style object of another workbook is not implemented")
        name = value.name
    elif isinstance(value, str):
        name = value
    else:
        raise error(450)
    index = find(book, name)
    if index is None:
        raise error(450)
    entry = book.stylesheet.cell_styles[index]
    restyle(target, lambda style: given(style, index, entry))


def prepare_save(book: Workbook) -> None:
    """Write the styles a save has to before the sheets, as Excel's own table holds them ahead of cells' parts."""
    stylesheet = book.stylesheet
    if all(entry.xf is not None for entry in stylesheet.cell_styles):
        return
    in_use: set[int] = set()
    for sheet in book.sheets_:
        in_use.update(cell.style.base for cell in sheet.cells_.values() if cell.style is not None)
        in_use.update(record.style.base for record in sheet.dims.rows.values() if record.style is not None)
        in_use.update(record.style.base for record in sheet.dims.columns.values() if record.style is not None)
    stylesheet.write_cell_styles(in_use, lead_font=_body_font(stylesheet))


# --- Workbook.Styles ------------------------------------------------------------------------------------


class Styles(VBACollection, ExcelObject):
    """A workbook's cell styles, built-in ones and those added, in the order their names sort."""

    vba_type_name = "Styles"

    def __init__(self, book: Workbook) -> None:
        self.book = book

    def vba_items(self) -> list[object]:
        return [StyleObject(self.book, name) for name in _listed(self.book)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            found = find(self.book, index)
            if found is None:
                raise error(9)
            return StyleObject(self.book, self.book.stylesheet.cell_styles[found].name)
        position = int(to_integer(index, "Long"))
        if 1 <= position <= len(items):
            return items[position - 1]
        raise error(9)

    @method
    def Add(self, Name: object = MISSING, BasedOn: object = MISSING) -> object:
        if Name is MISSING:
            raise error(449)
        name = to_text(Name).lstrip(" ")
        if not name:
            raise VBAUnsupportedError("a cell style with no name, which Excel calls Style 1, is not implemented")
        stylesheet = self.book.stylesheet
        key = name.casefold()
        if len(name) > 255 or any(one.casefold() == key for one in _listed(self.book)):
            raise error(1004, "Add method of Styles class failed")
        if BasedOn is MISSING:
            own, includes = _normal(stylesheet).format, S.PARTS
        elif isinstance(BasedOn, Range) and BasedOn.sheet.book is self.book:
            own, includes = _based_on(BasedOn)
        else:
            raise error(1004, "Add method of Styles class failed")
        stylesheet.add_cell_style(S.CellStyle(replace(own, applied=frozenset(), quote_prefix=False), includes,
                                              name=name))
        return StyleObject(self.book, name)

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Application(self) -> object:
        return self.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")


def _based_on(target: Range) -> tuple[S.Style, frozenset[str]]:
    """Styles.Add's BasedOn: its first cell's format, leaving out the parts the range's cells do not share."""
    first = target.first
    own = target.sheet.style_at(first.top, first.left)
    formats = styles_of(target)
    includes = frozenset(part for part in S.PARTS if all(getattr(one, part) == getattr(own, part) for one in formats))
    return own, includes


# --- the Style object -----------------------------------------------------------------------------------


class StyleObject(ExcelObject):
    """One cell style, found by its name each time it is asked."""

    vba_type_name = "Style"

    def __init__(self, book: Workbook, name: str) -> None:
        self.book = book
        self.name = name

    @property
    def cell_style(self) -> S.CellStyle:
        found = find(self.book, self.name)
        if found is None:
            raise error(424)
        return self.book.stylesheet.cell_styles[found]

    def guard_set(self, member: str) -> None:
        if member.casefold() in _SETTABLE:
            raise VBAUnsupportedError(f"changing a cell style's {member} is not implemented")

    @member(default=True, name="_Default")
    def Default(self) -> object:
        return self.cell_style.name

    @member
    def Name(self) -> object:
        return self.cell_style.name

    @member
    def NameLocal(self) -> object:
        return self.cell_style.name

    @member
    def Value(self) -> object:
        return self.cell_style.name

    @member
    def BuiltIn(self) -> object:
        return self.cell_style.builtin is not None

    def _includes(self, part: str) -> object:
        return part in self.cell_style.includes

    @member
    def IncludeNumber(self) -> object:
        return self._includes("number_format")

    @member
    def IncludeFont(self) -> object:
        return self._includes("font")

    @member
    def IncludeAlignment(self) -> object:
        return self._includes("alignment")

    @member
    def IncludeBorder(self) -> object:
        return self._includes("border")

    @member
    def IncludePatterns(self) -> object:
        return self._includes("fill")

    @member
    def IncludeProtection(self) -> object:
        return self._includes("protection")

    @member
    def NumberFormat(self) -> object:
        return self.cell_style.format.number_format

    @member
    def NumberFormatLocal(self) -> object:
        return self.cell_style.format.number_format

    def _alignment(self, what: str) -> object:
        from pyopenvba.apps.excel._formats import alignment_answer

        return alignment_answer(self.cell_style.format, what)

    @member
    def HorizontalAlignment(self) -> object:
        return self._alignment("HorizontalAlignment")

    @member
    def VerticalAlignment(self) -> object:
        return self._alignment("VerticalAlignment")

    @member
    def WrapText(self) -> object:
        return self._alignment("WrapText")

    @member
    def Orientation(self) -> object:
        return self._alignment("Orientation")

    @member
    def IndentLevel(self) -> object:
        indent = self.cell_style.format.alignment.indent
        return VBAInt(indent, "Long") if indent else NULL

    @member
    def ShrinkToFit(self) -> object:
        return self._alignment("ShrinkToFit")

    @member
    def ReadingOrder(self) -> object:
        return self._alignment("ReadingOrder")

    @member
    def AddIndent(self) -> object:
        return self._alignment("AddIndent")

    @member
    def Locked(self) -> object:
        return self._alignment("Locked")

    @member
    def FormulaHidden(self) -> object:
        return self._alignment("FormulaHidden")

    @member
    def MergeCells(self) -> object:
        raise error(1004, "Unable to get the MergeCells property of the Style class")

    @member
    def Font(self) -> object:
        return StyleFont(self)

    @member
    def Interior(self) -> object:
        return StyleInterior(self)

    @member
    def Borders(self, Index: object = MISSING) -> object:
        borders = StyleBorders(self)
        return borders if Index is MISSING else borders.vba_get("Item", [Index])

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Application(self) -> object:
        return self.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")


#: The Style properties a macro may set in Excel, which a style here does not take yet.
_SETTABLE: Final = frozenset(name.casefold() for name in (
    "AddIndent", "FormulaHidden", "HorizontalAlignment", "IncludeAlignment", "IncludeBorder", "IncludeFont",
    "IncludeNumber", "IncludePatterns", "IncludeProtection", "IndentLevel", "Locked", "MergeCells", "NumberFormat",
    "NumberFormatLocal", "Orientation", "ShrinkToFit", "VerticalAlignment", "WrapText", "ReadingOrder"))


class _StylePart(ExcelObject):
    """Font, Interior or a border of a cell style: read here, set not yet."""

    def __init__(self, owner: StyleObject) -> None:
        self.owner = owner

    @property
    def colors(self) -> S.Colors:
        return self.owner.book.stylesheet.colors

    def guard_set(self, member: str) -> None:
        raise VBAUnsupportedError(f"changing a cell style's {self.vba_type_name}.{member} is not implemented")

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.book.application


def _long(value: int) -> VBAInt:
    return VBAInt(value, "Long")


def _color(colors: S.Colors, color: S.Color | None, automatic: str, what: str) -> object:
    """One colour's Color, ColorIndex, ThemeColor or TintAndShade, as a style answers them."""
    if what == "Color":
        return float(S.bgr(colors.rrggbb(color, automatic)))
    if what == "ColorIndex":
        return _long(colors.color_index(color, automatic))
    if what == "ThemeColor":
        if color is None or color.kind != "theme":
            raise error(5)
        return _long(int(color.value) + 1)
    return (color.tint / S.TINT_SCALE) if color is not None else 0.0


class StyleFont(_StylePart):
    """A cell style's font."""

    vba_type_name = "Font"

    @property
    def font(self) -> S.Font:
        return self.owner.cell_style.format.font

    @member
    def Name(self) -> object:
        return self.font.name

    @member
    def Size(self) -> object:
        return self.font.size

    @member
    def Bold(self) -> object:
        return self.font.bold

    @member
    def Italic(self) -> object:
        return self.font.italic

    @member
    def Strikethrough(self) -> object:
        return self.font.strike

    @member
    def Superscript(self) -> object:
        return self.font.vert_align == "superscript"

    @member
    def Subscript(self) -> object:
        return self.font.vert_align == "subscript"

    @member
    def Underline(self) -> object:
        return _long(UNDERLINES.get(self.font.underline, XL_NONE))

    @member
    def FontStyle(self) -> object:
        raise error(1004, "Unable to get the FontStyle property of the Font class")

    @member
    def Color(self) -> object:
        return _color(self.colors, self.font.color, "000000", "Color")

    @member
    def ColorIndex(self) -> object:
        return _color(self.colors, self.font.color, "000000", "ColorIndex")

    @member
    def ThemeColor(self) -> object:
        return _color(self.colors, self.font.color, "000000", "ThemeColor")

    @member
    def TintAndShade(self) -> object:
        return _color(self.colors, self.font.color, "000000", "TintAndShade")

    @member
    def ThemeFont(self) -> object:
        return _long({"major": 1, "minor": 2}.get(self.font.scheme, 0))


class StyleInterior(_StylePart):
    """A cell style's fill, which answers apart from a range's: its pattern colour is white, not black."""

    vba_type_name = "Interior"

    @property
    def fill(self) -> S.Fill:
        return self.owner.cell_style.format.fill

    def _answer(self, what: str, *, pattern: bool = False) -> object:
        fill = self.fill
        if fill.pattern == "none":
            if pattern and what == "Color":
                return _long(0)
            return {"Color": 16777215.0, "TintAndShade": 0.0}.get(what, _long(XL_NONE))
        if pattern:
            # A style's pattern colour is its automatic one, which it answers as white, and PatternColor a Long.
            return {"Color": _long(16777215), "ColorIndex": _long(-4105), "ThemeColor": _long(0),
                    "TintAndShade": 0.0}[what]
        color = cell_color(fill)
        if what == "ThemeColor":
            return _long(int(color.value) + 1 if color.kind == "theme" else 0)
        return _color(self.colors, color, "FFFFFF", what)

    @member
    def Color(self) -> object:
        return self._answer("Color")

    @member
    def ColorIndex(self) -> object:
        return self._answer("ColorIndex")

    @member
    def ThemeColor(self) -> object:
        return self._answer("ThemeColor")

    @member
    def TintAndShade(self) -> object:
        return self._answer("TintAndShade")

    @member
    def Pattern(self) -> object:
        pattern = self.fill.pattern
        return _long(XL_NONE if pattern == "none" else PATTERNS.get(pattern, 1))

    @member
    def PatternColor(self) -> object:
        return self._answer("Color", pattern=True)

    @member
    def PatternColorIndex(self) -> object:
        return self._answer("ColorIndex", pattern=True)

    @member
    def PatternThemeColor(self) -> object:
        return self._answer("ThemeColor", pattern=True)

    @member
    def PatternTintAndShade(self) -> object:
        return self._answer("TintAndShade", pattern=True)


#: A style's Borders by index: xlLeft, xlRight, xlTop and xlBottom are its edges, as 1 to 4 are, 5 and 6 its
#: diagonals; xlEdgeLeft to xlInsideHorizontal, 7 to 12, name nothing it has.
_STYLE_EDGES: Final = {-4131: "left", -4152: "right", -4160: "top", -4107: "bottom", 1: "left", 2: "right",
                       3: "top", 4: "bottom", 5: "down", 6: "up"}


class StyleBorders(VBACollection, _StylePart):
    """A cell style's borders: its edges by xlLeft, xlRight, xlTop and xlBottom; xlEdgeLeft and the like
    answer Null."""

    vba_type_name = "Borders"

    def vba_items(self) -> list[object]:
        return [StyleBorder(self.owner, index) for index in (-4131, -4152, -4160, -4107, 5, 6)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        number = int(to_integer(index, "Long"))
        if number not in _STYLE_EDGES and number not in range(7, 13):
            raise error(1004, "Unable to get the Item property of the Borders class")
        return StyleBorder(self.owner, number)


class StyleBorder(_StylePart):
    """One of a cell style's borders."""

    vba_type_name = "Border"

    def __init__(self, owner: StyleObject, index: int) -> None:
        super().__init__(owner)
        self.index = index

    def _answer(self, what: str) -> object:
        edge = _STYLE_EDGES.get(self.index)
        if edge is None:
            # xlEdgeLeft to xlInsideHorizontal name nothing a style has.
            return {"Color": 0.0, "ColorIndex": _long(XL_NONE)}.get(what, NULL)
        border = self.owner.cell_style.format.border
        if edge in ("down", "up"):
            shown = border.diagonal_down if edge == "down" else border.diagonal_up
            side = border.diagonal if shown else S.NO_SIDE
        else:
            side = getattr(border, edge)
        return side_answer(self.colors, side, what)

    @member
    def LineStyle(self) -> object:
        return self._answer("LineStyle")

    @member
    def Weight(self) -> object:
        return self._answer("Weight")

    @member
    def Color(self) -> object:
        return self._answer("Color")

    @member
    def ColorIndex(self) -> object:
        return self._answer("ColorIndex")

    @member
    def ThemeColor(self) -> object:
        return self._answer("ThemeColor")

    @member
    def TintAndShade(self) -> object:
        return self._answer("TintAndShade")
