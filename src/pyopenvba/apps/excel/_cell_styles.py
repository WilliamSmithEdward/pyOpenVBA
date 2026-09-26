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
  file's first font itself (Stylesheet.prepare).
* Styles.Add copies Normal, or the first cell of the range it is based
  on, whose parts the range's cells do not share it leaves out.

scripts/measure_style_changes.py and tests/fixtures/cell_styles/
style_changes.json add what a macro does to a style:

* A property set on a style changes it, and each format that took the
  changed part from the style when given it follows, while the style
  includes the part; a part included again reaches them then. Normal's
  font is the file's first, so every format in it follows, whatever its
  style. A built-in style a macro touches is written customBuiltin.
* Style.Delete moves its cells to Normal, each part they did not set
  themselves following Normal from then on; a built-in style stays in the
  file, hidden. Styles.Merge brings another workbook's styles in, one
  named as one here taking its place.
* A protected sheet in the workbook refuses both, with error 1004.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import replace
from typing import TYPE_CHECKING, Final

from pyopenvba.apps.excel import _styles as S
from pyopenvba.apps.excel._formats import Border as RangeBorder
from pyopenvba.apps.excel._formats import Borders as RangeBorders
from pyopenvba.apps.excel._formats import Font as RangeFont
from pyopenvba.apps.excel._formats import Interior as RangeInterior
from pyopenvba.apps.excel._formats import (
    XL_NONE,
    alignment_answer,
    alignment_change,
    cell_color,
    pattern_color,
    side_answer,
    side_change,
    styles_of,
    with_diagonal,
)
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NOTHING,
    NULL,
    VBAInt,
    error,
    to_bool,
    to_integer,
    to_text,
)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._dimensions import ColumnRecord, RowRecord
    from pyopenvba.apps.excel._model import Cell, Workbook, Worksheet

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
    if any(entry.builtin == builtin for entry in stylesheet.cell_styles):
        return None  # the file holds it, or held it until a macro deleted it
    if font is None:
        # The copy of the Normal font Excel's table held from the start, whatever a macro made Normal since.
        value = stylesheet.opening_font
    else:
        major, minor = stylesheet.theme_fonts
        value = S.parse_font(font.replace("{major}", _typeface(major)).replace("{minor}", _typeface(minor)))
    own = S.Style(number_format=stylesheet.number_formats[number], font=value, fill=S.parse_fill(fill),
                  border=S.parse_border(border))
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
    names = {entry.name.casefold(): entry.name for entry in stylesheet.cell_styles if entry.name and not entry.deleted}
    held = {entry.builtin for entry in stylesheet.cell_styles}
    for name, builtin, *_ in STANDARD:
        if builtin not in held:
            names.setdefault(name.casefold(), name)
    return sorted(names.values(), key=lambda name: text_key(name, False))


# --- giving cells a style --------------------------------------------------------------------------------


def given(stylesheet: S.Stylesheet, style: S.Style, index: int, entry: S.CellStyle) -> S.Style:
    """A format given a cell style: the parts the style includes become the style's own, the rest stay.

    A part the style leaves alone counts from then on as the cell's own
    where it is not the style's, and a save writes its flag even if the
    style changes to match (tests/fixtures/cell_styles/follow.xlsx).
    """
    own = entry.format
    changes = {part: getattr(own, part) for part in entry.includes}
    kept = {part for part in S.PARTS - entry.includes if _differs(stylesheet, style, index, part)}
    return replace(style, base=index, applied=(style.applied - entry.includes) | kept,
                   **changes)  # type: ignore[arg-type]


def _differs(stylesheet: S.Stylesheet, style: S.Style, index: int, part: str) -> bool:
    """Whether a format's part is another entry of the file's than the one cell style ``index`` has.

    Parts are compared as the entries they point at: a cell in the Normal
    font points at the file's first font, a style that wants the Normal
    font at a copy of it, so the two differ though they read the same.
    """
    mine, theirs = getattr(style, part), getattr(stylesheet.cell_styles[index].format, part)
    if mine != theirs:
        return True
    if part != "font" or mine != _normal(stylesheet).format.font:
        return False
    normal = stylesheet.named("Normal") or 0
    copied = "font" not in style.applied and style.base != normal
    return copied != (index != normal)


def style_of(target: Range) -> object:
    """Range.Style: the style every cell of the range has, or Nothing where they differ."""
    bases = {style.base for style in styles_of(target)}
    if len(bases) != 1:
        return NOTHING
    entry = target.sheet.book.stylesheet.cell_styles[bases.pop()]
    if not entry.name:
        raise VBAUnsupportedError("Range.Style of a cell whose cell style has no name is not implemented")
    return StyleObject(target.sheet.book, entry)


def set_style(target: Range, value: object) -> None:
    """Give every cell of the range a style, named or as a Style object."""
    from pyopenvba.apps.excel._formats import restyle

    book = target.sheet.book
    if isinstance(value, StyleObject):
        if value.book is not book:
            raise VBAUnsupportedError("giving a range a Style object of another workbook is not implemented")
        index: int | None = _place(book.stylesheet, value.cell_style)
    elif isinstance(value, str):
        index = find(book, value)
    else:
        raise error(450)
    if index is None:
        raise error(450)
    place, entry = index, book.stylesheet.cell_styles[index]
    restyle(target, lambda style: given(book.stylesheet, style, place, entry))


def prepare_save(book: Workbook) -> None:
    """Give the stylesheet what the save adds before the sheets are written, in Excel's order: the styles
    cells use, then the cells' new formats, each part where Excel's own table holds it."""
    stylesheet = book.stylesheet
    formats = [(holder.style, holder.xf) for holder in _holders(book) if holder.style is not None]
    stylesheet.prepare(formats, latent(stylesheet))


def finish_save(book: Workbook, moved: dict[int, int]) -> None:
    """After a save, each cell, row and column format points at the place its xf went."""
    for holder in _holders(book):
        if holder.xf >= 0:
            holder.xf = moved.get(holder.xf, -1)


def _holders(book: Workbook) -> list[Cell | RowRecord | ColumnRecord]:
    """Every cell, row and column that may hold a format: each sheet's columns, rows and cells in order."""
    found: list[Cell | RowRecord | ColumnRecord] = []
    for sheet in book.sheets_:
        found += [*(record for _, record in sorted(sheet.dims.columns.items())),
                  *(record for _, record in sorted(sheet.dims.rows.items())),
                  *(cell for _, cell in sorted(sheet.cells_.items()))]
    return found


def latent(stylesheet: S.Stylesheet) -> S.Latent:
    """The fonts, fills and borders of the standard built-in styles, in Excel's own order, the theme's body
    font first: the places a save's new parts take ahead of what a macro made (see Stylesheet.meet)."""
    major, minor = stylesheet.theme_fonts
    themed = major in _SWISS and minor in _SWISS
    lead = _body_font(stylesheet)
    fonts: list[S.Font] = [] if lead is None else [lead]
    fills: list[S.Fill] = []
    borders: list[S.Border] = []
    for _, _, _, _, font, fill, border in STANDARD:
        if font is None:
            value: S.Font | None = stylesheet.opening_font
        else:
            value = S.parse_font(font.replace("{major}", _typeface(major)).replace("{minor}", _typeface(minor))) \
                if themed else None
        if value is not None and value not in fonts:
            fonts.append(value)
        filled, edged = S.parse_fill(fill), S.parse_border(border)
        if filled not in fills:
            fills.append(filled)
        if edged not in borders:
            borders.append(edged)
    return S.Latent(tuple(fonts), tuple(fills), tuple(borders))


# --- Workbook.Styles ------------------------------------------------------------------------------------


class Styles(VBACollection, ExcelObject):
    """A workbook's cell styles, built-in ones and those added, in the order their names sort."""

    vba_type_name = "Styles"

    def __init__(self, book: Workbook) -> None:
        self.book = book

    def vba_items(self) -> list[object]:
        return [self._named(name) for name in _listed(self.book)]

    def _named(self, name: str) -> StyleObject:
        found = find(self.book, name)
        if found is None:
            raise error(9)
        return StyleObject(self.book, self.book.stylesheet.cell_styles[found])

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            return self._named(index)
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
        entry = S.CellStyle(replace(own, applied=frozenset(), quote_prefix=False), includes, name=name)
        stylesheet.add_cell_style(entry)
        stylesheet.touch()
        return StyleObject(self.book, entry)

    @method
    def Merge(self, Workbook: object = MISSING) -> object:
        """Bring another workbook's styles in: one named as one here takes its place, keeping this workbook's
        spelling, and reaches the cells that follow it; the rest come in the order their names sort."""
        from pyopenvba.apps.excel._model import Workbook as Book

        source = Workbook
        if not isinstance(source, Book) or source is self.book:
            raise error(1004, "Merge method of Styles class failed")
        _guard(self.book)
        stylesheet, theirs_sheet = self.book.stylesheet, source.stylesheet
        stylesheet.touch()
        if (theirs_sheet.colors.theme, theirs_sheet.theme_fonts) != (stylesheet.colors.theme, stylesheet.theme_fonts):
            raise VBAUnsupportedError("merging the styles of a workbook with another theme is not implemented")
        for name in _listed(source):
            found = find(source, name)
            assert found is not None
            theirs = theirs_sheet.cell_styles[found]
            place = find(self.book, name)
            if place is None:
                if theirs.builtin is not None:
                    raise VBAUnsupportedError("merging a built-in style this workbook deleted is not implemented")
                stylesheet.add_cell_style(S.CellStyle(replace(theirs.format, applied=frozenset()), theirs.includes,
                                                      name=theirs.name))
                continue
            mine = stylesheet.cell_styles[place]
            if mine.includes == theirs.includes and all(getattr(mine.format, part) == getattr(theirs.format, part)
                                                        for part in S.PARTS):
                continue
            if mine.builtin == 0:
                raise VBAUnsupportedError("merging a Normal style that differs from this workbook's is not "
                                          "implemented")
            moved = frozenset(part for part in S.PARTS if getattr(mine.format, part) != getattr(theirs.format, part))
            own = replace(mine.format, **{part: getattr(theirs.format, part) for part in S.PARTS})
            _redefine(self.book, place, mine, own, theirs.includes, moved & theirs.includes)
        return EMPTY

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


# --- changing a style -----------------------------------------------------------------------------------


def _place(stylesheet: S.Stylesheet, entry: S.CellStyle) -> int:
    """A cell style's place in the stylesheet, which the formats that derive from it name as their base."""
    return next(index for index, one in enumerate(stylesheet.cell_styles) if one is entry)


def _guard(book: Workbook) -> None:
    """Excel changes and deletes a style only through the active workbook, and not while one of its sheets is
    protected, whatever the protection allows (tests/fixtures/cell_styles/style_changes.json, Guards)."""
    from pyopenvba.apps.excel._protection import enforced

    if book.application.active_book is not book:
        raise VBAUnsupportedError("changing a cell style of a workbook that is not the active one is not "
                                  "implemented: Excel changes the active workbook's style of that name")
    if any(enforced(sheet) is not None for sheet in book.sheets_):
        raise error(1004, "A cell style cannot change while a sheet of the workbook is protected")


def _change(owner: StyleObject, change: Callable[[S.Style], S.Style]) -> None:
    """A macro sets a property of a cell style: the style changes, and so does each format that took the
    changed part from it, while the style includes the part."""
    book = owner.book
    _guard(book)
    entry = owner.cell_style
    old, new = entry.format, change(entry.format)
    moved = frozenset(part for part in S.PARTS if getattr(old, part) != getattr(new, part))
    _redefine(book, _place(book.stylesheet, entry), entry, new, entry.includes, moved & entry.includes)


def _include(owner: StyleObject, part: str, value: object) -> None:
    """IncludeNumber and the rest: a part the style includes again reaches the formats that took it from the
    style; one it leaves out stays with them as it is."""
    book = owner.book
    _guard(book)
    entry = owner.cell_style
    on = to_bool(value)
    includes = entry.includes | {part} if on else entry.includes - {part}
    _redefine(book, _place(book.stylesheet, entry), entry, entry.format, includes,
              frozenset({part}) if on else frozenset())


def _redefine(book: Workbook, place: int, entry: S.CellStyle, own: S.Style, includes: frozenset[str],
              reach: frozenset[str]) -> None:
    """A cell style defined anew: its format, the parts it includes, and the parts that reach the formats that
    took them from it (tests/fixtures/cell_styles/style_changes.json).

    A built-in style a macro touches is customBuiltin from then on, even
    set to what it was. Normal's font is the file's first, which every
    format in that font points at, so all of them follow it, whatever
    their style.
    """
    stylesheet = book.stylesheet
    old = entry.format
    entry.format, entry.includes = replace(own, base=place, applied=frozenset()), includes
    if entry.builtin is not None:
        entry.custom = True
    stylesheet.restyle_cell_style(entry)
    new = entry.format
    normal = _normal_place(stylesheet)
    if place == normal and old.font != new.font:
        stylesheet.renormal_font(new.font)
        _follow(book, lambda style: _in_first_font(style, normal, old.font, new.font))
        reach -= {"font"}
    if not reach:
        return

    def follow(style: S.Style) -> S.Style | None:
        taken: frozenset[str] = reach - style.applied if style.base == place else frozenset()
        return replace(style, **{part: getattr(new, part) for part in taken}) if taken else None  # type: ignore[arg-type]

    _follow(book, follow)


def _normal_place(stylesheet: S.Stylesheet) -> int:
    found = stylesheet.named("Normal")
    return 0 if found is None else found


def _in_first_font(style: S.Style, normal: int, old: S.Font, new: S.Font) -> S.Style | None:
    """A format in the file's first font, the Normal font, which changes in place: one whose font is its
    style's own entry keeps it."""
    if style.font != old or ("font" not in style.applied and style.base != normal):
        return None
    return replace(style, font=new)


def _follow(book: Workbook, change: Callable[[S.Style], S.Style | None]) -> None:
    """Give a change to the table's cell xfs, in place, and to every cell, row and column format of the book,
    each keeping its own xf, so that two that come to say the same stay apart as Excel keeps them."""
    stylesheet = book.stylesheet
    moved: list[tuple[Worksheet, int, int, Cell | RowRecord | ColumnRecord, int, S.Style]] = []
    for sheet in book.sheets_:
        dims = sheet.dims
        holders: list[tuple[int, int, Cell | RowRecord | ColumnRecord]] = [
            *((row, column, cell) for (row, column), cell in sheet.cells_.items()),
            *((row, 0, record) for row, record in dims.rows.items()),
            *((0, column, record) for column, record in dims.columns.items())]
        for row, column, holder in holders:
            if holder.style is None:
                continue
            new = change(holder.style)
            if new is not None and new != holder.style:
                moved.append((sheet, row, column, holder, stylesheet.entry_of(holder.style, holder.xf), new))
    stylesheet.restyle_xfs(change)
    for sheet, row, column, holder, entry, new in moved:
        fonts = holder.style is not None and holder.style.font != new.font
        holder.style, holder.xf = new, entry
        if fonts and row:
            sheet.dims.fonts_changed(row)
        elif fonts:
            sheet.dims.column_fonts_changed()
        sheet.touched()


def _delete(owner: StyleObject) -> None:
    """Style.Delete: the formats of its cells go to Normal, each part they did not set themselves following
    Normal from then on; a built-in style stays in the file, hidden, and Normal cannot go."""
    book = owner.book
    _guard(book)
    entry = owner.cell_style
    stylesheet = book.stylesheet
    place, normal = _place(stylesheet, entry), _normal_place(stylesheet)
    if place == normal:
        raise error(1004, "Delete method of Style class failed")
    entry.deleted = True
    stylesheet.touch()
    own = stylesheet.cell_styles[normal].format

    def to_normal(style: S.Style) -> S.Style | None:
        if style.base != place:
            return None
        kept = S.PARTS - style.applied
        return replace(style, base=normal, **{part: getattr(own, part) for part in kept})  # type: ignore[arg-type]

    _follow(book, to_normal)


# --- the Style object -----------------------------------------------------------------------------------


class StyleObject(ExcelObject):
    """One cell style, held as itself: once deleted it is gone, whatever takes its name later."""

    vba_type_name = "Style"

    def __init__(self, book: Workbook, entry: S.CellStyle) -> None:
        self.book = book
        self.entry = entry

    @property
    def cell_style(self) -> S.CellStyle:
        if self.entry.deleted:
            raise error(424)
        return self.entry

    @member(default=True, name="_Default")
    def Default(self) -> object:
        return self.cell_style.name

    @member
    def Name(self) -> object:
        return self.cell_style.name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        raise error(450)

    @member
    def NameLocal(self) -> object:
        return self.cell_style.name

    @setter("NameLocal")
    def _set_name_local(self, value: object) -> None:
        raise error(1004, "Unable to set the NameLocal property of the Style class")

    @member
    def Value(self) -> object:
        return self.cell_style.name

    @setter("Value")
    def _set_value(self, value: object) -> None:
        raise error(450)

    @member
    def BuiltIn(self) -> object:
        return self.cell_style.builtin is not None

    @setter("BuiltIn")
    def _set_built_in(self, value: object) -> None:
        raise error(450)

    @method
    def Delete(self) -> object:
        _delete(self)
        return EMPTY

    def _includes(self, part: str) -> object:
        return part in self.cell_style.includes

    @member
    def IncludeNumber(self) -> object:
        return self._includes("number_format")

    @setter("IncludeNumber")
    def _set_include_number(self, value: object) -> None:
        _include(self, "number_format", value)

    @member
    def IncludeFont(self) -> object:
        return self._includes("font")

    @setter("IncludeFont")
    def _set_include_font(self, value: object) -> None:
        _include(self, "font", value)

    @member
    def IncludeAlignment(self) -> object:
        return self._includes("alignment")

    @setter("IncludeAlignment")
    def _set_include_alignment(self, value: object) -> None:
        _include(self, "alignment", value)

    @member
    def IncludeBorder(self) -> object:
        return self._includes("border")

    @setter("IncludeBorder")
    def _set_include_border(self, value: object) -> None:
        _include(self, "border", value)

    @member
    def IncludePatterns(self) -> object:
        return self._includes("fill")

    @setter("IncludePatterns")
    def _set_include_patterns(self, value: object) -> None:
        _include(self, "fill", value)

    @member
    def IncludeProtection(self) -> object:
        return self._includes("protection")

    @setter("IncludeProtection")
    def _set_include_protection(self, value: object) -> None:
        _include(self, "protection", value)

    @member
    def NumberFormat(self) -> object:
        return self.cell_style.format.number_format

    @setter("NumberFormat")
    def _set_number_format(self, value: object) -> None:
        from pyopenvba.apps.excel._number_format import normalized

        code = normalized(to_text(value))
        _change(self, lambda own: replace(own, number_format=code))

    @member
    def NumberFormatLocal(self) -> object:
        return self.cell_style.format.number_format

    @setter("NumberFormatLocal")
    def _set_number_format_local(self, value: object) -> None:
        self._set_number_format(value)

    def _alignment(self, what: str) -> object:
        return alignment_answer(self.cell_style.format, what)

    def _align(self, what: str, value: object) -> None:
        if what in ("Locked", "FormulaHidden"):
            on = to_bool(value)
            field_name = "locked" if what == "Locked" else "hidden"
            _change(self, lambda own: replace(own, protection=replace(own.protection, **{field_name: on})))
            return
        change = alignment_change(what, value)
        _change(self, lambda own: replace(own, alignment=change(own.alignment)))

    @member
    def HorizontalAlignment(self) -> object:
        return self._alignment("HorizontalAlignment")

    @setter("HorizontalAlignment")
    def _set_horizontal_alignment(self, value: object) -> None:
        self._align("HorizontalAlignment", value)

    @member
    def VerticalAlignment(self) -> object:
        return self._alignment("VerticalAlignment")

    @setter("VerticalAlignment")
    def _set_vertical_alignment(self, value: object) -> None:
        self._align("VerticalAlignment", value)

    @member
    def WrapText(self) -> object:
        return self._alignment("WrapText")

    @setter("WrapText")
    def _set_wrap_text(self, value: object) -> None:
        self._align("WrapText", value)

    @member
    def Orientation(self) -> object:
        return self._alignment("Orientation")

    @setter("Orientation")
    def _set_orientation(self, value: object) -> None:
        self._align("Orientation", value)

    @member
    def IndentLevel(self) -> object:
        """Null unless the style's text can be indented: aligned left, right or distributed, and not turned."""
        alignment = self.cell_style.format.alignment
        if alignment.horizontal not in ("left", "right", "distributed") or alignment.text_rotation:
            return NULL
        return VBAInt(alignment.indent, "Long")

    @setter("IndentLevel")
    def _set_indent_level(self, value: object) -> None:
        self._align("IndentLevel", value)

    @member
    def ShrinkToFit(self) -> object:
        return self._alignment("ShrinkToFit")

    @setter("ShrinkToFit")
    def _set_shrink_to_fit(self, value: object) -> None:
        self._align("ShrinkToFit", value)

    @member
    def ReadingOrder(self) -> object:
        return self._alignment("ReadingOrder")

    @setter("ReadingOrder")
    def _set_reading_order(self, value: object) -> None:
        self._align("ReadingOrder", value)

    @member
    def AddIndent(self) -> object:
        return self._alignment("AddIndent")

    @setter("AddIndent")
    def _set_add_indent(self, value: object) -> None:
        self._align("AddIndent", value)

    @member
    def Locked(self) -> object:
        return self._alignment("Locked")

    @setter("Locked")
    def _set_locked(self, value: object) -> None:
        self._align("Locked", value)

    @member
    def FormulaHidden(self) -> object:
        return self._alignment("FormulaHidden")

    @setter("FormulaHidden")
    def _set_formula_hidden(self, value: object) -> None:
        self._align("FormulaHidden", value)

    @member
    def MergeCells(self) -> object:
        raise error(1004, "Unable to get the MergeCells property of the Style class")

    @setter("MergeCells")
    def _set_merge_cells(self, value: object) -> None:
        raise error(1004, "Unable to set the MergeCells property of the Style class")

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


#: The properties of Normal's font a macro may change here: the rest move the sheets' column widths and row
#: heights, which rest on the Normal font and which the model measures for Aptos Narrow and Calibri 11 alone.
_NORMAL_FONT: Final = frozenset({"color", "colorindex", "themecolor", "tintandshade", "strikethrough"})


class StyleFont(RangeFont):
    """A cell style's font: set as a range's is, it changes the style and the formats that follow it."""

    vba_type_name = "Font"

    def __init__(self, owner: StyleObject) -> None:
        self.owner = owner

    @property
    def font(self) -> S.Font:
        """The style's font as its properties read it: neither super- nor subscript while the style leaves its
        font out (tests/fixtures/cell_styles/style_changes.json, Quirks)."""
        entry = self.owner.cell_style
        font = entry.format.font
        return font if "font" in entry.includes else replace(font, vert_align="")

    def guard_set(self, member: str) -> None:
        entry = self.owner.cell_style
        if entry.builtin == 0 and member.casefold() not in _NORMAL_FONT:
            raise VBAUnsupportedError(f"changing the Normal style's Font.{member}, which the sheets' column widths "
                                      "and row heights rest on, is not implemented")

    def _read(self, get: Callable[[S.Font], object]) -> object:
        return get(self.font)

    def _write(self, change: Callable[[S.Font], S.Font]) -> None:
        _change(self.owner, lambda own: replace(own, font=change(own.font)))

    def _color(self, what: str) -> object:
        return _color(self.owner.book.stylesheet.colors, self.font.color, "000000", what)

    @member
    def FontStyle(self) -> object:
        raise error(1004, "Unable to get the FontStyle property of the Font class")

    @setter("FontStyle")
    def _set_font_style(self, value: object) -> None:
        raise error(1004, "Unable to set the FontStyle property of the Font class")

    @member
    def ThemeFont(self) -> object:
        return _long({"major": 1, "minor": 2}.get(self.font.scheme, 0))

    @setter("ThemeFont")
    def _set_theme_font(self, value: object) -> None:
        which = int(to_integer(value, "Long"))
        if which not in (1, 2):
            raise VBAUnsupportedError(f"Font.ThemeFont = {which} on a cell style is not implemented")
        major, minor = self.owner.book.stylesheet.theme_fonts
        name, scheme = (major, "major") if which == 1 else (minor, "minor")
        self._write(lambda font: replace(font, name=name, scheme=scheme, family=S.font_family(name, font.family)))

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.book.application


class StyleInterior(RangeInterior):
    """A cell style's fill, which answers and changes apart from a range's: its automatic pattern colour stays
    what it is, white or black, where a cell's turns black."""

    vba_type_name = "Interior"

    def __init__(self, owner: StyleObject) -> None:
        self.owner = owner

    @property
    def fill(self) -> S.Fill:
        fill = self.owner.cell_style.format.fill
        if fill.gradient:
            raise VBAUnsupportedError("a gradient fill's Interior is not implemented")
        return fill

    def guard_set(self, member: str) -> None:
        return

    def _fills(self) -> list[S.Fill]:
        return [self.fill]

    def _write(self, change: Callable[[S.Fill], S.Fill]) -> None:
        self._fills()  # a gradient fill reports itself before anything changes
        _change(self.owner, lambda own: replace(own, fill=change(own.fill)))

    def _colored(self, fill: S.Fill, color: S.Color) -> S.Fill:
        if fill.pattern == "none":
            return S.Fill("solid", foreground=color, background=S.FOREGROUND)
        if fill.pattern == "solid":
            return replace(fill, foreground=color)
        return replace(fill, background=color)

    def _patterned(self, fill: S.Fill, pattern: str) -> S.Fill:
        if pattern == "none":
            return S.Fill()
        cell, lines = (S.BACKGROUND, S.FOREGROUND) if fill.pattern == "none" else (cell_color(fill), pattern_color(fill))
        if pattern == "solid":
            return S.Fill(pattern, foreground=cell, background=lines)
        return S.Fill(pattern, foreground=lines, background=cell)

    def _answer(self, pick: Callable[[S.Fill], S.Color], what: str, mixed: object) -> object:
        fill = self.fill
        pattern = pick is pattern_color
        if fill.pattern == "none":
            if pattern and what == "Color":
                return 0.0
            return {"Color": 16777215.0, "TintAndShade": 0.0}.get(what, _long(XL_NONE))
        color = pick(fill)
        if pattern and color in (S.FOREGROUND, S.BACKGROUND):
            # A style's automatic pattern colour answers as the system colour it is: white or black.
            return {"Color": 16777215.0 if color == S.BACKGROUND else 0.0, "ColorIndex": _long(S.XL_AUTOMATIC),
                    "ThemeColor": _long(0), "TintAndShade": 0.0}[what]
        if what == "ThemeColor":
            return _long(int(color.value) + 1 if color.kind == "theme" else 0)
        return _color(self.owner.book.stylesheet.colors, color, "000000" if pattern else "FFFFFF", what)

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.book.application


#: A style's Borders by index: xlLeft, xlRight, xlTop and xlBottom are its edges, as 1 to 4 are, 5 and 6 its
#: diagonals; xlEdgeLeft to xlInsideHorizontal, 7 to 12, name nothing it has.
_STYLE_EDGES: Final = {-4131: "left", -4152: "right", -4160: "top", -4107: "bottom", 1: "left", 2: "right",
                       3: "top", 4: "bottom", 5: "down", 6: "up"}
#: What Borders(7) to Borders(12) of a style read, in turn: its left, right, top and bottom edges and its
#: diagonals down and up (tests/fixtures/cell_styles/style_changes.json, Reach).
_SLOTS: Final = ("left", "right", "top", "bottom", "down", "up")


def _side(border: S.Border, edge: str) -> S.Side:
    if edge in ("down", "up"):
        shown = border.diagonal_down if edge == "down" else border.diagonal_up
        return border.diagonal if shown else S.NO_SIDE
    side: S.Side = getattr(border, edge)
    return side


def _with_edge(border: S.Border, edge: str, change: Callable[[S.Side], S.Side]) -> S.Border:
    if edge in ("down", "up"):
        return with_diagonal(border, 5 if edge == "down" else 6, change)
    return replace(border, **{edge: change(getattr(border, edge))})


class StyleBorders(RangeBorders):
    """A cell style's borders: its edges by xlLeft, xlRight, xlTop and xlBottom, and its diagonals; a property
    set here sets all six."""

    vba_type_name = "Borders"

    def __init__(self, owner: StyleObject) -> None:
        self.owner = owner

    def guard_set(self, member: str) -> None:
        return

    def vba_items(self) -> list[object]:
        return [StyleBorder(self.owner, index) for index in (-4131, -4152, -4160, -4107, 5, 6)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        number = int(to_integer(index, "Long"))
        if number not in _STYLE_EDGES and number not in range(7, 13):
            raise error(1004, "Unable to get the Item property of the Borders class")
        return StyleBorder(self.owner, number)

    def _all(self, what: str) -> object:
        raise VBAUnsupportedError(f"reading a cell style's Borders.{what} is not implemented")

    def _set_all(self, what: str, value: object) -> None:
        change = side_change(what, value)

        def every(border: S.Border) -> S.Border:
            for edge in _SLOTS:
                border = _with_edge(border, edge, change)
            return border

        _change(self.owner, lambda own: replace(own, border=every(own.border)))

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.book.application


class StyleBorder(RangeBorder):
    """One of a cell style's borders."""

    vba_type_name = "Border"

    def __init__(self, owner: StyleObject, index: int) -> None:
        self.owner = owner
        self.index = index

    def guard_set(self, member: str) -> None:
        return

    def answer(self, what: str) -> object:
        border = self.owner.cell_style.format.border
        edge = _STYLE_EDGES.get(self.index)
        if edge is None:
            # xlEdgeLeft to xlInsideHorizontal read a style's edges in turn, and only as there being a line in
            # the automatic colour, which is what a macro draws when it gives none.
            side = _side(border, _SLOTS[self.index - 7])
            if side.style and side.color in (None, S.FOREGROUND):
                return {"LineStyle": _long(XL_NONE), "Weight": _long(2), "Color": 0.0,
                        "ColorIndex": _long(XL_NONE)}.get(what, NULL)
            return {"Color": 0.0, "ColorIndex": _long(XL_NONE)}.get(what, NULL)
        return side_answer(self.owner.book.stylesheet.colors, _side(border, edge), what)

    def assign(self, what: str, value: object) -> None:
        change = side_change(what, value)
        edge = _STYLE_EDGES.get(self.index)
        if edge is None:
            return  # xlEdgeLeft and the rest name nothing a style has, and setting one changes nothing
        _change(self.owner, lambda own: replace(own, border=_with_edge(own.border, edge, change)))

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.book.application
