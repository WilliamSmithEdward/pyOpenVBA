"""Range.Sort and the Sort object a recorded macro uses, as Excel sorts.

Measured in live Excel (scripts/measure_sort_order.py, scripts/measure_sort.py):

- Ascending, numbers come first, then text, then FALSE and TRUE, then
  errors, all of them equal; descending, errors, TRUE and FALSE, text,
  numbers. Blank cells go last either way. The sort is stable: rows that
  compare equal keep their order.
- Text sorts as Windows sorts words, which ``text_key`` reproduces for
  ASCII, checked on 400 random strings: spaces and marks before digits,
  digits before letters, case ignored unless MatchCase asks for lower
  case first, and hyphens and apostrophes ignored but for breaking ties.
  Text with other characters reports itself unsupported.
- Each row of the range moves whole, formats and all, and its formulas
  shift with it as a copy's would; nothing outside the range changes.
  xlSortTextAsNumbers compares text that would type as a number as that
  number. A single cell sorts its current region.
- A header row is kept out of the sort with xlYes; an omitted Header is
  xlNo. xlGuess takes the first row for a header where, in any column,
  its cell and the one under it both hold something and differ in kind,
  in format, or in the first being in capitals and the second not.
- The Sort object sorts its range by its fields in order, keeping them
  after Apply; it keeps a header only when told xlYes.
- On a filtered sheet (scripts/measure_autofilter_edits.py) the visible
  rows sort among themselves into the rows they fill; the hidden ones
  stay, and a range with no visible row does not change.
- A sort field on a cell's colour or its font's (scripts/measure_sort_colour.py)
  puts the rows of that colour on top, or at the bottom with xlDescending,
  the rest keeping their order. A cell shows the colour with a solid fill
  of it, a conditional format's or a theme colour's included; a patterned
  fill, white, and no fill at all do not, but a field whose SortOnValue
  was never set sorts on no fill. A font shows it only in that colour
  set outright: the automatic colour, and a field left at it, match
  nothing. SortOnValue is an Interior or a Font holding the colour, error
  1004 on a field of values, and the file keeps the colour as a dxf.
- With no range or no field, Apply sorts nothing and raises no error.
  The file keeps the rows a sort sorts, under the header given xlYes.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, replace
from functools import cmp_to_key
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba._opc_xml import bounded_child, opc_part
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.apps.excel._visible import filtering, hidden_lines, unmeasured
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import shift_text
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAInt, error, to_bool, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Cell, Workbook, Worksheet
    from pyopenvba.apps.excel._styles import Style

ASCENDING, DESCENDING = 1, 2
GUESS, YES, NO = 0, 1, 2
TOP_TO_BOTTOM, LEFT_TO_RIGHT = 1, 2
TEXT_AS_NUMBERS = 1
#: XlSortOn: what a sort field orders by.
ON_VALUES, ON_CELL_COLOR, ON_FONT_COLOR, ON_ICON = 0, 1, 2, 3
#: Printable ASCII other than letters, hyphens and apostrophes, lowest first, as Excel's sort orders them.
_SYMBOLS = " !\"#$%&()*,./:;?@[\\]^_`{|}~+<=>0123456789"
#: Characters word sort passes over but for breaking ties, apostrophe before hyphen.
_IGNORED = {"'": 0, "-": 1}
_NUMBER, _TEXT, _LOGICAL, _ERROR, _BLANK = range(5)


@dataclass(frozen=True)
class Key:
    """One sort key: the row or column of the range it reads, and how. A key ``on`` a colour orders the rows that
    show ``color``, a VBA colour, before the rest; None is no fill, or the automatic font colour."""

    index: int
    descending: bool = False
    numbers: bool = False
    on: int = ON_VALUES
    color: int | None = None


def text_key(text: str, match_case: bool) -> tuple[tuple[int, ...], tuple[int, ...], tuple[tuple[int, int], ...]]:
    """Text as Windows word sort compares it: weights, then case, then the hyphens and apostrophes it passed over."""
    weights: list[int] = []
    cases: list[int] = []
    passed: list[tuple[int, int]] = []
    for index, char in enumerate(text):
        if ord(char) > 126 or (ord(char) < 32):
            raise VBAUnsupportedError("sorting text with characters beyond printable ASCII is not implemented")
        if char in _IGNORED:
            # A later one sorts earlier, and at the same place an apostrophe before a hyphen.
            passed.append((-index, _IGNORED[char]))
            continue
        weights.append(100 + ord(char.lower()) if char.isalpha() else _SYMBOLS.index(char))
        cases.append(1 if match_case and char.isupper() else 0)
    return tuple(weights), tuple(cases), tuple(passed)


def _classified(value: object, key: Key, match_case: bool) -> tuple[int, object]:
    """A value's class, and what orders it within its class."""
    from pyopenvba.apps.excel._typing import typed_text

    if value is EMPTY or value is None:
        return _BLANK, 0
    if isinstance(value, bool):
        return _LOGICAL, int(value)
    if isinstance(value, (int, float)):
        return _NUMBER, float(value)
    if isinstance(value, ExcelError):
        return _ERROR, 0
    text = str(value)
    if key.numbers:
        found = typed_text(text).value
        if isinstance(found, float):
            return _NUMBER, found
    return _TEXT, text_key(text, match_case)


def _compare(first: tuple[int, object], second: tuple[int, object], descending: bool) -> int:
    (one, left), (other, right) = first, second
    if one == _BLANK or other == _BLANK:
        return (one == _BLANK) - (other == _BLANK)
    if one != other:
        return (one > other) - (one < other) if not descending else (one < other) - (one > other)
    if left == right or one == _ERROR:
        return 0
    less = left < right  # type: ignore[operator]
    return (-1 if less else 1) * (-1 if descending else 1)


def _shown_style(sheet: Worksheet, row: int, column: int) -> Style:
    """A cell's format as its conditional formats show it."""
    from pyopenvba.apps.excel._conditional_eval import display_style
    from pyopenvba.apps.excel._conditional_stats import conditional_read

    if sheet.conditional_formats:
        with conditional_read():
            return display_style(sheet, row, column)
    return sheet.style_at(row, column)


def shows_automatic_font(sheet: Worksheet, row: int, column: int) -> bool:
    """Whether a cell's font shows the automatic colour, which a filter on it finds and a colour sort passes over
    (tests/fixtures/colour_filters/)."""
    color = _shown_style(sheet, row, column).font.color
    return color is None or color.kind == "auto" or color == sheet.book.stylesheet.default.font.color


def shows_color(sheet: Worksheet, row: int, column: int, key: Key) -> bool:
    """Whether a cell shows the colour a colour key sorts on, as its conditional formats show it."""
    from pyopenvba.apps.excel._styles import Color

    style = _shown_style(sheet, row, column)
    wanted = None if key.color is None else str(Color.from_bgr(key.color).value)[-6:]
    if key.on == ON_CELL_COLOR:
        fill = style.fill
        if fill.gradient:
            raise VBAUnsupportedError("sorting on the colour of a gradient fill is not implemented")
        if wanted is None:
            return fill.pattern == "none"
        shown = sheet.book.stylesheet.colors.rrggbb(fill.foreground, "FFFFFF")
        return fill.pattern == "solid" and shown.upper() == wanted
    color = style.font.color
    if color is None or color.kind == "auto" or color == sheet.book.stylesheet.default.font.color:
        # The automatic colour matches nothing, a field left at it included.
        return False
    if color.kind != "rgb":
        raise VBAUnsupportedError("sorting on a font colour from the theme or the palette is not implemented")
    return wanted is not None and str(color.value)[-6:].upper() == wanted


@dataclass
class SortPlan:
    sheet: Worksheet
    block: Area
    moving: dict[tuple[int, int], Cell]
    destination: dict[int, int]
    across: bool
    formulas: dict[tuple[int, int], str]


def plan_sort_area(sheet: Worksheet, area: Area, keys: list[Key], *, header: bool, match_case: bool,
                   across: bool, charge_value: Callable[[object], None] | None = None) -> SortPlan | None:
    """Sort the rows of ``area`` -- its columns, ``across`` -- by ``keys``, the first row or column kept if ``header``."""
    from pyopenvba.apps.excel import _merges

    if any(_merges.intersects(area, one) for one in sheet.merged_areas):
        raise VBAUnsupportedError("sorting merged cells is not implemented")
    if any(area.contains(*key) for key in sheet.notes):
        raise VBAUnsupportedError("sorting cells with notes is not implemented")
    used = sheet.used_bounds()
    if used is None:
        return None
    if across:
        first, last = area.left + (1 if header else 0), min(area.right, used[3])
    else:
        first, last = area.top + (1 if header else 0), min(area.bottom, used[2])
    block = (Area(area.top, first, area.bottom, last) if across else Area(first, area.left, last, area.right)) if first <= last else None
    from pyopenvba.apps.excel._arrays import met, PART_OF_AN_ARRAY

    if block is not None:
        arrays, partial_arrays = met(sheet, [block])
        if arrays or partial_arrays:
            raise error(1004, PART_OF_AN_ARRAY)
        if sheet.data_tables:
            from pyopenvba.apps.excel._data_tables import refuse

            refuse(sheet, [block], "Sorting")
    if block is None or first >= last:
        return None
    calculator = sheet.book.calculator

    def value(line: int, key: Key) -> object:
        row, column = (area.top + key.index, line) if across else (line, area.left + key.index)
        cell = sheet.cells_.get((row, column))
        if cell is None:
            return EMPTY
        return calculator.value_of(sheet.name, row, column) if cell.formula else cell.value

    lines = list(range(first, last + 1))
    if any(not spill.blocked and _merges.intersects(block, spill.area) for spill in sheet.spills.values()):
        raise error(1004, "You can't change part of an array")
    if filtering(sheet):
        # A filtered sheet sorts its visible rows among themselves, into the rows they fill, and none
        # when none shows: measured for Range.Sort and the Sort object.
        if across:
            unmeasured(Range(sheet, [block]), "Sorting left to right")
        else:
            unmeasured(Range(sheet, [block]), "Sorting over hidden columns", lines="columns")
            hidden, _ = hidden_lines(sheet)
            lines = [line for line in lines if line not in hidden]
    classified: dict[int, list[tuple[int, object]]] = {}
    for line in lines:
        row_keys: list[tuple[int, object]] = []
        for key in keys:
            if key.on != ON_VALUES:
                # The colour's rows first: 0 before 1, and after it when descending.
                cell = (area.top + key.index, line) if across else (line, area.left + key.index)
                row_keys.append((_NUMBER, 0.0 if shows_color(sheet, *cell, key) else 1.0))
                continue
            raw = value(line, key)
            if charge_value is not None:
                charge_value(raw)
            row_keys.append(_classified(raw, key, match_case))
        classified[line] = row_keys

    def compare(one: int, other: int) -> int:
        for position, key in enumerate(keys):
            result = _compare(classified[one][position], classified[other][position], key.descending)
            if result:
                return result
        return 0

    order = sorted(lines, key=cmp_to_key(compare))
    destination = dict(zip(order, lines, strict=True))
    moving = {position: cell for position, cell in sheet.cells_.items()
              if block.contains(*position) and position[1 if across else 0] in destination}
    formulas: dict[tuple[int, int], str] = {}
    for (row, column), cell in moving.items():
        shift = (0, destination[column] - column) if across else (destination[row] - row, 0)
        if cell.formula and shift != (0, 0):
            formulas[row, column] = shift_text(cell.formula, *shift)
    return SortPlan(sheet, block, moving, destination, across, formulas)


def apply_sort(plan: SortPlan, *, preserve_cached: bool = False) -> None:
    """Publish only after every comparison and moved formula has been prepared."""
    sheet, block, destination, across = plan.sheet, plan.block, plan.destination, plan.across
    for position in plan.moving:
        del sheet.cells_[position]
    for (row, column), cell in plan.moving.items():
        new = (row, destination[column]) if across else (destination[row], column)
        if (row, column) in plan.formulas:
            cell.formula = plan.formulas[row, column]
            cell.stale = True
            if not preserve_cached:
                cell.value = EMPTY
        sheet.cells_[new] = cell
    if sheet.hyperlinks:
        from pyopenvba.apps.excel._hyperlinks import sorted_lines

        # A link one line high moves with its line; one over several stays (tests/fixtures/hyperlinks/).
        sorted_lines(sheet, block, destination, across=across)
    sheet.touched()
    sheet.book.calculator.rebuild()


def sort_area(sheet: Worksheet, area: Area, keys: list[Key], *, header: bool, match_case: bool,
              across: bool) -> None:
    plan = plan_sort_area(sheet, area, keys, header=header, match_case=match_case, across=across)
    if plan is not None:
        apply_sort(plan)


def guessed_header(sheet: Worksheet, area: Area, across: bool) -> bool:
    """Whether xlGuess takes the first row, or column, for a header."""
    from pyopenvba.apps.excel._region import holds_content

    if (area.columns if across else area.rows) < 2:
        return False
    calculator = sheet.book.calculator
    for offset in range(area.rows if across else area.columns):
        first = (area.top + offset, area.left) if across else (area.top, area.left + offset)
        second = (first[0], first[1] + 1) if across else (first[0] + 1, first[1])
        if not holds_content(sheet, *first) or not holds_content(sheet, *second):
            continue
        values: list[object] = []
        for row, column in (first, second):
            cell = sheet.cells_[(row, column)]
            values.append(calculator.value_of(sheet.name, row, column) if cell.formula else cell.value)
        if _kind(values[0]) != _kind(values[1]):
            return True
        if isinstance(values[0], str) and isinstance(values[1], str) and _capitals(values[0]) \
                and not _capitals(values[1]):
            return True
        if _look(sheet.style_at(*first)) != _look(sheet.style_at(*second)):
            return True
    return False


def _kind(value: object) -> int:
    if isinstance(value, bool):
        return _LOGICAL
    if isinstance(value, (int, float)):
        return _NUMBER
    if isinstance(value, ExcelError):
        return _ERROR
    return _TEXT


def _capitals(text: str) -> bool:
    return any(char.isalpha() for char in text) and text == text.upper()


def _look(style: object) -> tuple[object, ...]:
    """What a format shows, apart from how it was applied."""
    return tuple(getattr(style, part) for part in ("number_format", "font", "fill", "border", "alignment",
                                                    "protection"))


# --- Range.Sort --------------------------------------------------------------------------------------


def range_sort(target: Range, keys: list[tuple[object, object, object]], header: object, order_custom: object,
               match_case: object, orientation: object) -> object:
    """Range.Sort: up to three (key, order, data option) triples, the first required."""
    from pyopenvba.apps.excel._region import current_region

    if len(target.areas) != 1:
        raise VBAUnsupportedError("sorting several areas is not implemented")
    if order_custom is not MISSING and int(to_integer(order_custom, "Long")) != 1:
        raise VBAUnsupportedError("sorting by a custom list is not implemented")
    area = target.first
    if target.single:
        area = current_region(target.sheet, area.top, area.left)
    across = orientation is not MISSING and int(to_integer(orientation, "Long")) == LEFT_TO_RIGHT
    if keys[0][0] is MISSING:
        raise error(1004, "Sort needs a key")
    resolved: list[Key] = []
    for key, order, option in keys:
        if key is MISSING:
            continue
        resolved.append(_key(key, area, across, order, option))
    kept = NO if header is MISSING else int(to_integer(header, "Long"))
    has_header = kept == YES or (kept == GUESS and guessed_header(target.sheet, area, across))
    sort_area(target.sheet, area, resolved, header=has_header,
              match_case=match_case is not MISSING and to_bool(match_case), across=across)
    return True


def _key(key: object, area: Area, across: bool, order: object, option: object) -> Key:
    if not isinstance(key, Range):
        raise error(1004, "A sort key has to be a range")
    index = key.first.top - area.top if across else key.first.left - area.left
    if not 0 <= index < (area.rows if across else area.columns):
        raise error(1004, "The sort key is outside the range sorted")
    descending = order is not MISSING and int(to_integer(order, "Long")) == DESCENDING
    numbers = option is not MISSING and int(to_integer(option, "Long")) == TEXT_AS_NUMBERS
    return Key(index, descending, numbers)


# --- the Sort object ---------------------------------------------------------------------------------


class SortRange(Range):
    vba_type_name = "Range"

    def __init__(self, owner: ExcelObject, selected: Range) -> None:
        super().__init__(selected.sheet, selected.areas, whole=selected.whole, lines=selected.lines)
        self.owner = owner

    def _check_alive(self) -> None:
        super()._check_alive()
        self.owner._check_alive()

    @member
    def Parent(self) -> object:
        return self.owner


class SortState:
    """A sheet's sort settings, which Worksheet.Sort shows."""

    def __init__(self) -> None:
        self.fields: list[SortField] = []
        self.area: Range | None = None
        self.header = GUESS
        self.match_case = False
        self.orientation = TOP_TO_BOTTOM
        self.method = 1
        self.original: tuple[object, ...] = (None, GUESS, False, TOP_TO_BOTTOM, 1, ())


def settings_signature(state: SortState) -> tuple[object, ...]:
    area = (id(state.area.sheet), tuple(state.area.areas)) if state.area is not None else None
    return (area, state.header, state.match_case, state.orientation, state.method,
            tuple((id(field.key.sheet), tuple(field.key.areas), field.order, field.option, field.sort_on, field.color)
                  for field in state.fields))


def changed(sheet: Worksheet) -> bool:
    return sheet.sort_state is not None and settings_signature(sheet.sort_state) != sheet.sort_state.original


def with_sort_state(sheet: Worksheet, xml: str) -> str:
    """Persist measured live worksheet configuration, preserving untouched XML."""
    state = sheet.sort_state
    if state is None or not changed(sheet):
        return xml
    namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    markup = b''
    if state.area is not None and state.fields:
        if len(state.fields) > 64:
            raise VBAUnsupportedError('Persisting more than 64 worksheet sort fields is not implemented')
        if state.area.sheet is not sheet or any(field.key.sheet is not sheet for field in state.fields):
            raise VBAUnsupportedError('Persisting cross-sheet worksheet sort settings is not implemented')
        if state.header not in (GUESS, YES, NO) or any(field.order not in (ASCENDING, DESCENDING) for field in state.fields):
            raise VBAUnsupportedError('Persisting unmeasured worksheet sort enums is not implemented')
        if state.orientation != TOP_TO_BOTTOM or state.method != 1:
            raise VBAUnsupportedError('Persisting non-default worksheet sort methods/orientations is not implemented')
        if any(field.option not in (0, 1) for field in state.fields):
            raise VBAUnsupportedError('Persisting custom worksheet sort data options is not implemented')
        if len(state.area.areas) != 1 or any(len(field.key.areas) != 1 for field in state.fields):
            raise VBAUnsupportedError('Persisting multi-area worksheet sort settings is not implemented')
        from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes

        ns = '{' + namespace + '}'
        block = state.area.first
        if state.header == YES and block.rows > 1:
            # The file keeps the rows sorted, under the header (tests/fixtures/sort_colour/).
            block = Area(block.top + 1, block.left, block.bottom, block.right, block.sheet)
        root = XmlElement(ns + 'sortState', {'xmlns': namespace, 'ref': block.address(absolute=False)})
        if state.match_case:
            root.set('caseSensitive', '1')
        for field in state.fields:
            root.append(XmlElement(ns + 'sortCondition', condition_attributes(sheet.book, field)))
        markup = xml_bytes(root).split(b'?>', 1)[1].lstrip()
    part = opc_part(xml.encode('utf8'), namespace, 'worksheet')
    matches = [(start, end) for tag, _attrs, start, end in part.children if tag == namespace + '}sortState']
    if len(matches) > 1:
        raise VBAUnsupportedError('Repeated worksheet sort metadata is not implemented')
    if matches:
        start, end = matches[0]
        return (part.data[:start] + markup + part.data[end:]).decode('utf8')
    before = {'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting',
              'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter',
              'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags',
              'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls',
              'webPublishItems', 'tableParts', 'extLst'}
    at = next((start for tag, _attrs, start, _end in part.children
               if tag.removeprefix(namespace + '}') in before), part.closing_start)
    return part.insert(at, markup).decode('utf8') if markup else xml


_MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
#: What a sortCondition's sortBy names: the SortOn of its field.
SORTED_BY = {'value': ON_VALUES, 'cellColor': ON_CELL_COLOR, 'fontColor': ON_FONT_COLOR}
_SORT_BY = {ON_CELL_COLOR: 'cellColor', ON_FONT_COLOR: 'fontColor'}


def condition_attributes(book: Workbook, field: SortField) -> dict[str, str]:
    """A field's sortCondition, its attributes in the schema's order as Excel writes them: a colour field names
    the dxf holding its colour."""
    attributes: dict[str, str] = {}
    if field.order == DESCENDING:
        attributes['descending'] = '1'
    if field.sort_on != ON_VALUES:
        attributes['sortBy'] = _SORT_BY[field.sort_on]
    attributes['ref'] = field.key.first.address(absolute=False)
    if field.sort_on != ON_VALUES:
        attributes['dxfId'] = str(book.stylesheet.ensure_dxf(_color_dxf(book, field)))
    return attributes


def _color_dxf(book: Workbook, field: SortField) -> XmlElement:
    """The dxf a file keeps a colour field's colour in (tests/fixtures/sort_colour/): a solid fill of it, or no
    fill; or the workbook's font in it, or in the font's own colour for the automatic one."""
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._styles import Color

    dxf = XmlElement(_MAIN + 'dxf')
    rgb = None if field.color is None else str(Color.from_bgr(field.color).value)
    if field.sort_on == ON_CELL_COLOR:
        pattern = _child(_child(dxf, 'fill'), 'patternFill', {'patternType': 'none' if rgb is None else 'solid'})
        _child(pattern, 'fgColor', {'indexed': '64'} if rgb is None else {'rgb': rgb})
        _child(pattern, 'bgColor', {'indexed': '65'})
        return dxf
    font = book.stylesheet.default.font
    element = _child(dxf, 'font')
    for tag, value in (('b', '0'), ('i', '0'), ('strike', '0'), ('condense', '0'), ('extend', '0'), ('outline', '0'),
                       ('shadow', '0'), ('u', 'none'), ('vertAlign', 'baseline'), ('sz', f'{font.size:g}')):
        _child(element, tag, {'val': value})
    if rgb is not None:
        _child(element, 'color', {'rgb': rgb})
    elif font.color is not None and font.color.kind == 'theme' and not font.color.tint:
        _child(element, 'color', {'theme': str(font.color.value)})
    else:
        raise VBAUnsupportedError("saving a sort on the automatic font colour of a workbook whose font has no "
                                  "theme colour is not implemented")
    _child(element, 'name', {'val': font.name})
    for tag, value in (('family', font.family), ('scheme', font.scheme)):
        if value:
            _child(element, tag, {'val': value})
    return dxf


def _child(parent: XmlElement, tag: str, attributes: dict[str, str] | None = None) -> XmlElement:
    """A new element of the stylesheet's namespace, put last in ``parent``."""
    from pyopenvba.apps.excel._charts_file import XmlElement

    made = XmlElement(_MAIN + tag, attributes or {})
    parent.append(made)
    return made


def condition_color(book: Workbook, sort_by: str, dxf_id: str) -> int | None:
    """The colour a colour condition's dxf holds, None for no fill or for the font's automatic colour."""
    from pyopenvba.apps.excel._styles import bgr

    try:
        dxf = book.stylesheet.dxf(int(dxf_id))
    except (ValueError, IndexError):
        raise VBAUnsupportedError('A colour sort condition naming no dxf is not implemented') from None
    if sort_by == 'cellColor':
        pattern = dxf.find(f'{_MAIN}fill/{_MAIN}patternFill')
        if pattern is not None and pattern.get('patternType') == 'none':
            return None
        color = None if pattern is None else pattern.find(_MAIN + 'fgColor')
        if pattern is None or pattern.get('patternType') != 'solid' or color is None or not color.get('rgb'):
            raise VBAUnsupportedError('A colour sort condition with an unmeasured fill is not implemented')
        return bgr(str(color.get('rgb'))[-6:])
    color = dxf.find(f'{_MAIN}font/{_MAIN}color')
    if color is None or color.get('theme') is not None:
        return None
    if not color.get('rgb'):
        raise VBAUnsupportedError('A font colour sort condition with an unmeasured colour is not implemented')
    return bgr(str(color.get('rgb'))[-6:])


def preflight_save(book: Workbook) -> None:
    """Refuse new sort serialization boundaries before any package is edited."""
    from pyopenvba.apps.excel._table_sort import changed as table_changed
    from pyopenvba.apps.excel._table_sort import write_state
    from pyopenvba.apps.excel._tables import base_table_xml

    package = book.package
    for sheet in book.sheets_:
        if changed(sheet):
            xml = (package.read(sheet.part_name).decode('utf8')
                   if package is not None and sheet.part_name and package.has(sheet.part_name)
                   else '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>')
            with_sort_state(sheet, xml)
        for table in sheet.tables:
            if table_changed(table):
                write_state(table, table.xml or base_table_xml(table))


def _loaded_state(sheet: Worksheet) -> SortState:
    """Hydrate only bounded ordinary value-sort metadata with exact references."""
    state = SortState()
    package = sheet.book.package
    if package is None or not sheet.part_name or not package.has(sheet.part_name):
        return state
    namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    try:
        group = bounded_child(package.entry(sheet.part_name).read(), namespace, 'worksheet', 'sortState',
                              max_bytes=4_000_000, max_elements=65)
    except ValueError:
        raise VBAUnsupportedError('Reading unsupported worksheet sort metadata is not implemented') from None
    if group is None:
        return state
    ns = namespace + '}'
    allowed = {'ref', 'caseSensitive', 'columnSort', 'sortMethod'}
    if (any(key not in allowed and key != 'xmlns' and not key.startswith('xmlns:') for key in group.attributes)
            or group.get('caseSensitive', '0') not in ('0', '1', 'false', 'true')
            or group.get('columnSort', '0') not in ('0', 'false')
            or group.get('sortMethod', 'pinYin') != 'pinYin'
            or len(group.children) > 64 or group.has_text or group.has_opaque):
        raise VBAUnsupportedError('Custom worksheet sort settings are not implemented')
    try:
        area = parse_area(group.get('ref', ''))
        if area.sheet:
            raise VBAUnsupportedError('Qualified worksheet sort ranges are not implemented')
        state.area = Range(sheet, [area])
        state.header = GUESS
        state.match_case = group.get('caseSensitive') in ('1', 'true')
        for condition in group.children:
            sort_by = condition.get('sortBy', 'value')
            if (condition.name != ns + 'sortCondition' or condition.children or condition.has_text or condition.has_opaque
                    or any(key not in ('ref', 'descending', 'sortBy', 'dxfId') and key != 'xmlns'
                           and not key.startswith('xmlns:') for key in condition.attributes)
                    or condition.get('descending', '0') not in ('0', '1', 'false', 'true')
                    or sort_by not in SORTED_BY
                    or (sort_by == 'value') == ('dxfId' in condition.attributes)):
                raise VBAUnsupportedError('Custom worksheet sort conditions are not implemented')
            key = parse_area(condition.get('ref', ''))
            if key.sheet:
                raise VBAUnsupportedError('Qualified worksheet sort keys are not implemented')
            order = DESCENDING if condition.get('descending') in ('1', 'true') else ASCENDING
            field = SortField(Range(sheet, [key]), order, 0, SORTED_BY[sort_by])
            if sort_by != 'value':
                field.color = condition_color(sheet.book, sort_by, str(condition.get('dxfId')))
            state.fields.append(field)
    except ValueError:
        raise VBAUnsupportedError('Invalid worksheet sort references are not implemented') from None
    state.original = settings_signature(state)
    return state


class SortObject(ExcelObject):
    vba_type_name = "Sort"

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet
        if sheet.sort_state is None:
            sheet.sort_state = _loaded_state(sheet)
        self.state: SortState = sheet.sort_state

    def vba_no_member(self, name: str) -> Exception:
        return error(438) if name.casefold() == 'name' else super().vba_no_member(name)

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, 'Long')

    @member
    def SortFields(self) -> object:
        return SortFields(self)

    @method
    def SetRange(self, Rng: object = MISSING) -> object:
        if not isinstance(Rng, Range):
            raise error(1004, "SetRange needs a range")
        self.state.area = Rng
        return EMPTY

    @member
    def Rng(self) -> object:
        if self.state.area is None:
            raise error(1004)
        return SortRange(self, self.state.area)

    @member
    def Header(self) -> object:
        return VBAInt(self.state.header, "Long")

    @setter("Header")
    def _set_header(self, value: object) -> None:
        self.state.header = int(to_integer(value, "Long"))

    @member
    def MatchCase(self) -> object:
        return self.state.match_case

    @setter("MatchCase")
    def _set_match_case(self, value: object) -> None:
        self.state.match_case = to_bool(value)

    @member
    def Orientation(self) -> object:
        return VBAInt(self.state.orientation, "Long")

    @setter("Orientation")
    def _set_orientation(self, value: object) -> None:
        self.state.orientation = int(to_integer(value, "Long"))

    @member
    def SortMethod(self) -> object:
        return VBAInt(self.state.method, "Long")

    @setter("SortMethod")
    def _set_sort_method(self, value: object) -> None:
        self.state.method = int(to_integer(value, "Long"))

    @method
    def Apply(self) -> object:
        state = self.state
        if state.area is None or not state.fields:
            # With no range or no field Excel sorts nothing, and says nothing (tests/fixtures/sort_colour/).
            return EMPTY
        target = state.area
        if len(target.areas) != 1:
            raise VBAUnsupportedError("sorting several areas is not implemented")
        area = target.first
        across = state.orientation == LEFT_TO_RIGHT
        keys = [replace(_key(field.key, area, across, VBAInt(field.order, "Long"), VBAInt(field.option, "Long")),
                        on=field.sort_on, color=field.color) for field in state.fields]
        sort_area(target.sheet, area, keys, header=state.header == YES, match_case=state.match_case, across=across)
        from pyopenvba.apps.excel._events import recalculated

        recalculated(target.sheet.book)
        return EMPTY


class SortFields(VBACollection, ExcelObject):
    vba_type_name = "SortFields"

    def __init__(self, owner: SortObject) -> None:
        self.owner = owner

    def vba_items(self) -> list[object]:
        return list(self.owner.state.fields)

    @method
    def Clear(self) -> object:
        self.owner.state.fields.clear()
        self.owner.state.area = None
        return EMPTY

    @method
    def Add(self, Key: object = MISSING, SortOn: object = MISSING, Order: object = MISSING,
            CustomOrder: object = MISSING, DataOption: object = MISSING) -> object:
        if not isinstance(Key, Range):
            raise error(1004, "A sort field needs a range for its key")
        on = ON_VALUES if SortOn is MISSING else sort_on(SortOn)
        if CustomOrder is not MISSING:
            raise VBAUnsupportedError("sorting by a custom order is not implemented")
        field = SortField(Key, ASCENDING if Order is MISSING else int(to_integer(Order, "Long")),
                          0 if DataOption is MISSING else int(to_integer(DataOption, "Long")), on)
        self.owner.state.fields.append(field)
        return field

    @method
    def Add2(self, Key: object = MISSING, SortOn: object = MISSING, Order: object = MISSING,
             CustomOrder: object = MISSING, DataOption: object = MISSING, SubField: object = MISSING) -> object:
        if SubField is not MISSING:
            raise VBAUnsupportedError("sorting by a data type's field is not implemented")
        return self.Add(Key, SortOn, Order, CustomOrder, DataOption)


class SortField(ExcelObject):
    vba_type_name = "SortField"

    def __init__(self, key: Range, order: int, option: int, sort_on: int = ON_VALUES) -> None:
        self.key = key
        self.order = order
        self.option = option
        self.sort_on = sort_on
        #: The VBA colour a colour field sorts on; None, as it starts, is no fill or the automatic font colour.
        self.color: int | None = None

    def vba_no_member(self, name: str) -> Exception:
        return error(438) if name.casefold() == 'name' else super().vba_no_member(name)

    @member
    def Key(self) -> object:
        return SortRange(self, self.key)

    @member
    def SortOn(self) -> object:
        return VBAInt(self.sort_on, "Long")

    @setter("SortOn")
    def _set_sort_on(self, value: object) -> None:
        self.sort_on = sort_on(value)
        self.color = None

    @member
    def SortOnValue(self) -> object:
        """What a colour field sorts on, as an Interior or a Font holding the colour; a field of values has none,
        error 1004 (tests/fixtures/sort_colour/)."""
        if self.sort_on == ON_CELL_COLOR:
            return SortInterior(self)
        if self.sort_on == ON_FONT_COLOR:
            return SortFont(self)
        raise error(1004, "A sort on values has no SortOnValue")

    @member
    def Order(self) -> object:
        return VBAInt(self.order, "Long")

    @setter("Order")
    def _set_order(self, value: object) -> None:
        self.order = int(to_integer(value, "Long"))

    @member
    def DataOption(self) -> object:
        return VBAInt(self.option, "Long")

    @setter("DataOption")
    def _set_data_option(self, value: object) -> None:
        self.option = int(to_integer(value, "Long"))


def sort_on(value: object) -> int:
    """The SortOn a macro gives: values, a cell's colour or its font's. Icons are not modelled."""
    found = int(to_integer(value, "Long"))
    if found not in (ON_VALUES, ON_CELL_COLOR, ON_FONT_COLOR):
        raise VBAUnsupportedError("sorting on a conditional format's icons is not implemented" if found == ON_ICON
                                  else f"sorting with SortOn {found} is not implemented")
    return found


class SortInterior(ExcelObject):
    """A cell-colour field's SortOnValue: the colour the field sorts on, which reads as white, for no fill, until a
    macro sets it."""

    vba_type_name = "Interior"
    #: What Color reads before a colour is set.
    unset = 0xFFFFFF

    def __init__(self, field: SortField) -> None:
        self.field = field

    @member
    def Color(self) -> object:
        return float(self.unset if self.field.color is None else self.field.color)

    @setter("Color")
    def _set_color(self, value: object) -> None:
        self.field.color = int(to_integer(value, "Long"))


class SortFont(SortInterior):
    """A font-colour field's SortOnValue, which reads as black, for the automatic colour, until set."""

    vba_type_name = "Font"
    unset = 0
