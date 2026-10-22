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
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
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
    from pyopenvba.apps.excel._model import Cell, Workbook, Worksheet

ASCENDING, DESCENDING = 1, 2
GUESS, YES, NO = 0, 1, 2
TOP_TO_BOTTOM, LEFT_TO_RIGHT = 1, 2
TEXT_AS_NUMBERS = 1
#: Printable ASCII other than letters, hyphens and apostrophes, lowest first, as Excel's sort orders them.
_SYMBOLS = " !\"#$%&()*,./:;?@[\\]^_`{|}~+<=>0123456789"
#: Characters word sort passes over but for breaking ties, apostrophe before hyphen.
_IGNORED = {"'": 0, "-": 1}
_NUMBER, _TEXT, _LOGICAL, _ERROR, _BLANK = range(5)


@dataclass(frozen=True)
class Key:
    """One sort key: the row or column of the range it reads, and how."""

    index: int
    descending: bool = False
    numbers: bool = False


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
            tuple((id(field.key.sheet), tuple(field.key.areas), field.order, field.option) for field in state.fields))


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
        root = XmlElement(ns + 'sortState', {'xmlns': namespace, 'ref': state.area.first.address(absolute=False)})
        if state.match_case:
            root.set('caseSensitive', '1')
        for field in state.fields:
            attrs = {'ref': field.key.first.address(absolute=False)}
            if field.order == DESCENDING:
                attrs['descending'] = '1'
            root.append(XmlElement(ns + 'sortCondition', attrs))
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
            if (condition.name != ns + 'sortCondition' or condition.children or condition.has_text or condition.has_opaque
                    or any(key not in ('ref', 'descending', 'sortBy') and key != 'xmlns'
                           and not key.startswith('xmlns:') for key in condition.attributes)
                    or condition.get('descending', '0') not in ('0', '1', 'false', 'true')
                    or condition.get('sortBy', 'value') != 'value'):
                raise VBAUnsupportedError('Custom worksheet sort conditions are not implemented')
            key = parse_area(condition.get('ref', ''))
            if key.sheet:
                raise VBAUnsupportedError('Qualified worksheet sort keys are not implemented')
            state.fields.append(SortField(Range(sheet, [key]),
                                          DESCENDING if condition.get('descending') in ('1', 'true') else ASCENDING, 0))
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
            raise error(1004, "The sort has no range or no key")
        target = state.area
        if len(target.areas) != 1:
            raise VBAUnsupportedError("sorting several areas is not implemented")
        area = target.first
        across = state.orientation == LEFT_TO_RIGHT
        keys = [_key(field.key, area, across, VBAInt(field.order, "Long"), VBAInt(field.option, "Long"))
                for field in state.fields]
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
        if SortOn is not MISSING and int(to_integer(SortOn, "Long")) != 0:
            raise VBAUnsupportedError("sorting on cell colour, font colour or icons is not implemented")
        if CustomOrder is not MISSING:
            raise VBAUnsupportedError("sorting by a custom order is not implemented")
        field = SortField(Key, ASCENDING if Order is MISSING else int(to_integer(Order, "Long")),
                          0 if DataOption is MISSING else int(to_integer(DataOption, "Long")))
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

    def __init__(self, key: Range, order: int, option: int) -> None:
        self.key = key
        self.order = order
        self.option = option

    def vba_no_member(self, name: str) -> Exception:
        return error(438) if name.casefold() == 'name' else super().vba_no_member(name)

    @member
    def Key(self) -> object:
        return SortRange(self, self.key)

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
