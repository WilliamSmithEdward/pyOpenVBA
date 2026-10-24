"""A pivot table's layout, compact, outline or tabular, and its fields' subtotals, as Excel keeps and shows them.

PivotTable.RowAxisLayout lays every field out in one form, and a field's
LayoutForm and LayoutCompactRow lay out that field alone; the table keeps
a field's form as the field's compact and outline flags, and the Values
field's as the table's compactData and outlineData. A field in compact
form shows the fields inside it in its own column; in outline form in
the next column, each of its items on a line of its own above theirs;
in tabular form in the next column, each item on the line of the first
of theirs and its subtotals below them. Subtotals names the functions a
field's subtotals use, Automatic, the data field's own, or any of eleven
others; a lone function shows on an item's own line where the field has
its subtotals at the top, and every other below its lines. A field can
also leave a blank line after each item's lines, and repeat its items'
labels down every line they hold.

Measured with scripts/measure_pivot_layouts.py (tests/fixtures/pivot_layouts/):
a field's compact flag shows the table's first label column as Row Labels
and its first column header as Column Labels, where any field on the rows
or columns is compact; the table's compactData aligns every label cell
left, indented by its depth in its column; the file keeps CompactRowIndent
less one. Subtotals of a page field or of the Values field are error
1004, and so is an index outside 1 to 12 or an array of another length.
"""
from __future__ import annotations

from collections.abc import Iterable, Sequence
from copy import deepcopy
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._pivot_create import encoded, node
from pyopenvba.apps.excel._pivots import NS, decode_text
from pyopenvba.exceptions import VBAError, VBAUnsupportedError
from pyopenvba.interpreter._values import MISSING, VBAArray, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivots import PivotField, PivotTable

VALUES = -2
X14 = '{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}'
#: Excel 2010's extension of a pivot field, where it keeps whether the field repeats its labels.
FIELD_EXTENSION = '{2946ED86-A175-432a-8AC1-64E0C546D7DE}'
#: Excel 2010's extension of a pivot table, where it keeps whether its fields repeat their labels.
TABLE_EXTENSION = '{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}'
#: The functions Subtotals(2) to Subtotals(12) name, in that order, as a subtotal line's type in the file: the
#: attribute marking each on its field, the function its data field's values go through, and the name its label
#: gives it (tests/fixtures/pivot_layouts/ all_functions).
SUBTOTALS: dict[str, tuple[str, str, str]] = {
    'sum': ('sumSubtotal', 'sum', 'Sum'), 'countA': ('countASubtotal', 'count', 'Count'),
    'avg': ('avgSubtotal', 'average', 'Average'), 'max': ('maxSubtotal', 'max', 'Max'),
    'min': ('minSubtotal', 'min', 'Min'), 'product': ('productSubtotal', 'product', 'Product'),
    'count': ('countSubtotal', 'countNums', 'Count'), 'stdDev': ('stdDevSubtotal', 'stdDev', 'StdDev'),
    'stdDevP': ('stdDevPSubtotal', 'stdDevp', 'StdDevp'), 'var': ('varSubtotal', 'var', 'Var'),
    'varP': ('varPSubtotal', 'varp', 'Varp')}
KINDS = tuple(SUBTOTALS)
#: A pivot table's attributes in the order the schema has them, as Excel writes them.
TABLE_ORDER = ('name cacheId dataOnRows dataPosition autoFormatId applyNumberFormats applyBorderFormats '
               'applyFontFormats applyPatternFormats applyAlignmentFormats applyWidthHeightFormats dataCaption '
               'grandTotalCaption errorCaption showError missingCaption showMissing pageStyle pivotTableStyle '
               'vacatedStyle tag updatedVersion minRefreshableVersion asteriskTotals showItems editData '
               'disableFieldList showCalcMbrs visualTotals showMultipleLabel showDataDropDown showDrill printDrill '
               'showMemberPropertyTips showDataTips enableWizard enableDrill enableFieldProperties '
               'preserveFormatting useAutoFormatting pageWrap pageOverThenDown subtotalHiddenItems rowGrandTotals '
               'colGrandTotals fieldPrintTitles itemPrintTitles mergeItem showDropZones createdVersion indent '
               'showEmptyRow showEmptyCol showHeaders compact outline outlineData compactData published '
               'gridDropZones immersive multipleFieldFilters chartFormat rowHeaderCaption colHeaderCaption '
               'fieldListSortAscending mdxSubqueries customListSort').split()
#: Excel 2010's pivot table options in the order the schema has them.
OPTIONS_ORDER = ('fillDownLabelsDefault visualTotalsForSets calculatedMembersInFilters altText altTextSummary '
                 'enableEdit autoApply allocationMethod weightExpression hideValuesRow').split()
#: RowAxisLayout's and LayoutRowDefault's xlCompactRow, xlTabularRow and xlOutlineRow.
COMPACT_ROW, TABULAR_ROW, OUTLINE_ROW = 0, 1, 2


def on(element: XmlElement, name: str, default: bool) -> bool:
    return element.get(name, '1' if default else '0') in ('1', 'true')


@dataclass(frozen=True)
class FieldLayout:
    """How a field on an axis lays out the fields inside it, and its subtotals."""

    compact: bool = True
    outline: bool = True
    #: Whether a lone subtotal shows on each item's own line, LayoutSubtotalLocation xlAtTop.
    top: bool = True
    blank: bool = False
    repeat: bool = False
    #: The subtotals' types, in the order their lines show: 'default' for Automatic, or the functions chosen.
    functions: tuple[str, ...] = ('default',)
    caption: str = ''

    @property
    def mode(self) -> str:
        return 'tabular' if not self.outline else 'compact' if self.compact else 'outline'


def functions_of(field: XmlElement) -> tuple[str, ...]:
    """A field's subtotal types: the functions it names, else Automatic unless it has none (tests/fixtures/
    pivot_layouts/ count_only, where Excel writes no defaultSubtotal beside the one it names)."""
    chosen = tuple(kind for kind in KINDS if on(field, SUBTOTALS[kind][0], False))
    if chosen:
        return chosen
    return ('default',) if on(field, 'defaultSubtotal', True) else ()


def _options(field: XmlElement, create: bool = False) -> XmlElement | None:
    """A pivot field's Excel 2010 options, made where asked and missing."""
    extensions = field.find(NS + 'extLst')
    for extension in extensions if extensions is not None else ():
        if extension.get('uri') == FIELD_EXTENSION:
            return extension.find(X14 + 'pivotField')
    if not create:
        return None
    if extensions is None:
        extensions = node('extLst')
        field.append(extensions)
    extension = node('ext', {'uri': FIELD_EXTENSION, 'xmlns:x14': X14[1:-1]})
    options = XmlElement(X14 + 'pivotField')
    extension.append(options)
    extensions.append(extension)
    return options


def repeats(field: XmlElement) -> bool:
    options = _options(field)
    return options is not None and on(options, 'fillDownLabels', False)


def layout_of(root: XmlElement, index: int) -> FieldLayout:
    """The layout of a table's field, or of its Values field, which has no subtotals."""
    if index == VALUES:
        return FieldLayout(compact=on(root, 'compactData', True), outline=on(root, 'outlineData', False), functions=())
    field = root.findall(f'{NS}pivotFields/{NS}pivotField')[index]
    return FieldLayout(compact=on(field, 'compact', True), outline=on(field, 'outline', True),
                       top=on(field, 'subtotalTop', True), blank=on(field, 'insertBlankRow', False),
                       repeat=repeats(field), functions=functions_of(field),
                       caption=decode_text(field.get('subtotalCaption', '')))


def label_columns(root: XmlElement, rows: list[int]) -> tuple[list[int], list[int]]:
    """Each row field's label column and its depth there: a field in compact form keeps the fields inside it in
    its own column, indented a step further each (tests/fixtures/pivot_layouts/ three_middle_tabular)."""
    columns: list[int] = []
    depths: list[int] = []
    for position in range(len(rows)):
        if position and layout_of(root, rows[position - 1]).mode == 'compact':
            columns.append(columns[-1])
            depths.append(depths[-1] + 1)
        else:
            columns.append(columns[-1] + 1 if columns else 0)
            depths.append(0)
    return columns, depths


def compact_captions(root: XmlElement, axes: Iterable[int]) -> bool:
    """Whether the table shows Row Labels and Column Labels: where a field on its rows or columns is compact,
    whatever the table's own flags (tests/fixtures/pivot_layouts/ column_compact_header)."""
    return any(layout_of(root, index).compact for index in axes if index != VALUES)


def compact_indent(root: XmlElement) -> int:
    """CompactRowIndent: the file keeps it less one, 0 as 127 (tests/fixtures/pivot_layouts/ indent,
    indent_zero)."""
    return (int(root.get('indent', '1')) + 1) % 128


def ordered(element: XmlElement, order: Sequence[str], *, others_first: bool = True) -> None:
    """Put an element's attributes in the order the schema has them, any others, its namespace declarations among
    them, ahead as they stand, or after where Excel writes them last."""
    rank = {name: place for place, name in enumerate(order)}
    items = sorted(element.attrib.items(), key=lambda item: rank.get(item[0], -1 if others_first else len(order)))
    element.attrib.clear()
    element.attrib.update(items)


def set_flag(element: XmlElement, name: str, value: bool, default: bool, order: Sequence[str] | None = None) -> None:
    """Write a flag as Excel does: left out at its default, else 1 or 0."""
    if value == default:
        element.attrib.pop(name, None)
    else:
        element.set(name, '1' if value else '0')
    if order is not None:
        ordered(element, order)


def set_text(element: XmlElement, name: str, value: str, order: Sequence[str]) -> None:
    element.set(name, encoded(value))
    ordered(element, order)


def _choice(value: object, allowed: tuple[int, ...], code: int) -> int:
    try:
        chosen = to_integer(value)
    except (VBAError, ValueError, TypeError, OverflowError):
        raise error(code) from None
    if chosen not in allowed:
        raise error(code)
    return chosen


def _changed(table: PivotTable, root: XmlElement) -> None:
    from pyopenvba.apps.excel._pivot_filters import rerender

    rerender(table, root)
    table.sheet.book.saved = False


def _field_order() -> Sequence[str]:
    from pyopenvba.apps.excel._pivot_filters import FIELD_ORDER

    return FIELD_ORDER


# -- the table


def row_axis_layout(table: PivotTable, value: object) -> None:
    """RowAxisLayout: every field, the Values field and the table in one form (tests/fixtures/pivot_layouts/
    compact, outline, tabular); an unknown form is error 5."""
    chosen = _choice(value, (COMPACT_ROW, TABULAR_ROW, OUTLINE_ROW), 5)
    table = table.entry
    root = deepcopy(table.root)
    compact, outline = chosen == COMPACT_ROW, chosen != TABULAR_ROW
    set_flag(root, 'compact', compact, True)
    set_flag(root, 'compactData', compact, True)
    set_flag(root, 'outline', outline, False)
    set_flag(root, 'outlineData', outline, False)
    ordered(root, TABLE_ORDER)
    for field in root.findall(f'{NS}pivotFields/{NS}pivotField'):
        set_flag(field, 'compact', compact, True)
        set_flag(field, 'outline', outline, True)
        ordered(field, _field_order())
    _changed(table, root)


def layout_row_default(table: PivotTable) -> int:
    root = table.entry.root
    if on(root, 'compact', True):
        return COMPACT_ROW
    return OUTLINE_ROW if on(root, 'outline', False) else TABULAR_ROW


def set_layout_row_default(table: PivotTable, value: object) -> None:
    """LayoutRowDefault: the table's own compact and outline flags, which no field takes up as it comes to an
    axis (tests/fixtures/pivot_layouts/ default_then_add)."""
    chosen = _choice(value, (COMPACT_ROW, TABULAR_ROW, OUTLINE_ROW), 5)
    table = table.entry
    set_flag(table.root, 'compact', chosen == COMPACT_ROW, True)
    set_flag(table.root, 'outline', chosen != TABULAR_ROW, False, TABLE_ORDER)
    _changed(table, deepcopy(table.root))


def repeat_all_labels(table: PivotTable, value: object) -> None:
    """RepeatAllLabels: xlRepeatLabels (2) marks every field and the table's default, xlDoNotRepeatLabels (1)
    takes the marks away; anything else is error 5 (tests/fixtures/pivot_layouts/ repeat_tabular, repeat_off)."""
    repeat = _choice(value, (1, 2), 5) == 2
    table = table.entry
    root = deepcopy(table.root)
    for field in root.findall(f'{NS}pivotFields/{NS}pivotField'):
        _set_repeat(field, repeat)
    _set_table_default(root, repeat)
    _changed(table, root)


def _set_repeat(field: XmlElement, repeat: bool) -> None:
    if repeat:
        options = _options(field, create=True)
        assert options is not None
        options.set('fillDownLabels', '1')
        return
    extensions = field.find(NS + 'extLst')
    if extensions is None:
        return
    for extension in list(extensions):
        if extension.get('uri') == FIELD_EXTENSION:
            extensions.remove(extension)
    if not len(extensions):
        field.remove(extensions)


def _set_table_default(root: XmlElement, repeat: bool) -> None:
    from pyopenvba.apps.excel._pivot_options import table_options

    options = table_options(root)
    if repeat:
        options.set('fillDownLabelsDefault', '1')
    else:
        options.attrib.pop('fillDownLabelsDefault', None)
    ordered(options, OPTIONS_ORDER, others_first=False)


def set_table_flag(table: PivotTable, name: str, value: object) -> None:
    """RowGrand (rowGrandTotals) or ColumnGrand (colGrandTotals)."""
    table = table.entry
    root = deepcopy(table.root)
    set_flag(root, name, to_bool(value), True, TABLE_ORDER)
    _changed(table, root)


def caption(table: PivotTable, name: str, default: str) -> str:
    return decode_text(table.entry.root.get(name, default))


def set_caption(table: PivotTable, name: str, value: object) -> None:
    """A caption the table shows: CompactLayoutRowHeader, CompactLayoutColumnHeader or GrandTotalName."""
    table = table.entry
    root = deepcopy(table.root)
    set_text(root, name, to_text(value), TABLE_ORDER)
    _changed(table, root)


def display_field_captions(table: PivotTable) -> bool:
    return on(table.entry.root, 'showHeaders', True)


def set_display_field_captions(table: PivotTable, value: object) -> None:
    table = table.entry
    root = deepcopy(table.root)
    set_flag(root, 'showHeaders', to_bool(value), True, TABLE_ORDER)
    _changed(table, root)


def set_compact_indent(table: PivotTable, value: object) -> None:
    """CompactRowIndent: 0 to 127, else error 5 (tests/fixtures/pivot_layouts/ indent_errors)."""
    chosen = _choice(value, tuple(range(128)), 5)
    table = table.entry
    root = deepcopy(table.root)
    # Left out at the schema's 1, as Excel leaves it (tests/fixtures/pivot_layouts/ indent_three).
    if (chosen - 1) % 128 == 1:
        root.attrib.pop('indent', None)
    else:
        root.set('indent', str((chosen - 1) % 128))
    ordered(root, TABLE_ORDER)
    _changed(table, root)


def merge_labels(table: PivotTable) -> bool:
    return on(table.entry.root, 'mergeItem', False)


def set_merge_labels(table: PivotTable, value: object) -> None:
    if to_bool(value):
        raise VBAUnsupportedError('PivotTable.MergeLabels, which merges and centres the label cells, '
                                  'is not implemented')
    table.entry.root.attrib.pop('mergeItem', None)


# -- a field


def _definition(field: PivotField) -> XmlElement | None:
    """The pivotField a field's layout lives on, or None for the Values field and a data field, whose layout the
    table keeps."""
    return None if field.data or field.index == VALUES else field.root


def _source(field: PivotField) -> tuple[PivotTable, XmlElement, XmlElement | None]:
    """The table, a copy of its definition to change, and the copy's pivotField for the field."""
    table = field.table.entry
    root = deepcopy(table.root)
    definition = _definition(field)
    if definition is None:
        return table, root, None
    return table, root, root.findall(f'{NS}pivotFields/{NS}pivotField')[field.index]


def layout_form(field: PivotField) -> int:
    """LayoutForm: xlTabular (0) or xlOutline (1); the Values field's and a data field's are the table's
    outlineData (tests/fixtures/pivot_layouts/ values_field_layout)."""
    definition = _definition(field)
    if definition is None:
        return int(on(field.table.entry.root, 'outlineData', False))
    return int(on(definition, 'outline', True))


def set_layout_form(field: PivotField, value: object) -> None:
    outline = _choice(value, (0, 1), 1004) == 1
    table, root, definition = _source(field)
    if definition is None:
        set_flag(root, 'outlineData', outline, False, TABLE_ORDER)
    else:
        set_flag(definition, 'outline', outline, True, _field_order())
    _changed(table, root)


def compact_row(field: PivotField) -> bool:
    definition = _definition(field)
    if definition is None:
        return on(field.table.entry.root, 'compactData', True)
    return on(definition, 'compact', True)


def set_compact_row(field: PivotField, value: object) -> None:
    compact = to_bool(value)
    table, root, definition = _source(field)
    if definition is None:
        set_flag(root, 'compactData', compact, True, TABLE_ORDER)
    else:
        set_flag(definition, 'compact', compact, True, _field_order())
    _changed(table, root)


def _own(field: PivotField) -> XmlElement:
    """A field's pivotField where the member needs one: the Values field and a data field are error 1004."""
    definition = _definition(field)
    if definition is None:
        raise error(1004)
    return definition


def subtotal_location(field: PivotField) -> int:
    """LayoutSubtotalLocation: xlAtTop (1) or xlAtBottom (2), which a tabular field keeps and ignores."""
    return 1 if on(_own(field), 'subtotalTop', True) else 2


def set_subtotal_location(field: PivotField, value: object) -> None:
    _own(field)
    top = _choice(value, (1, 2), 1004) == 1
    table, root, definition = _source(field)
    assert definition is not None
    set_flag(definition, 'subtotalTop', top, True, _field_order())
    _changed(table, root)


def flag_of(field: PivotField, name: str) -> bool:
    return on(_own(field), name, False)


def set_blank_line(field: PivotField, value: object) -> None:
    """LayoutBlankLine: a blank line after each item's lines, where a field is inside it (tests/fixtures/
    pivot_layouts/ blank, blank_single)."""
    _own(field)
    table, root, definition = _source(field)
    assert definition is not None
    set_flag(definition, 'insertBlankRow', to_bool(value), False, _field_order())
    _changed(table, root)


def set_page_break(field: PivotField, value: object) -> None:
    """LayoutPageBreak: a mark on the field that shows in no cell and puts no break in the sheet (tests/fixtures/
    pivot_layouts/ page_break)."""
    set_flag(_own(field), 'insertPageBreak', to_bool(value), False, _field_order())
    field.table.entry.changed = True
    field.table.entry.sheet.book.saved = False


def repeat_labels(field: PivotField) -> bool:
    return repeats(_own(field))


def set_repeat_labels(field: PivotField, value: object) -> None:
    _own(field)
    table, root, definition = _source(field)
    assert definition is not None
    _set_repeat(definition, to_bool(value))
    _changed(table, root)


def subtotal_name(field: PivotField) -> str:
    return decode_text(_own(field).get('subtotalCaption', ''))


def set_subtotal_name(field: PivotField, value: object) -> None:
    """SubtotalName: the label every subtotal line of the field shows in place of the item's name and Total
    (tests/fixtures/pivot_layouts/ subtotal_name_bottom)."""
    _own(field)
    table, root, definition = _source(field)
    assert definition is not None
    text = to_text(value)
    if text:
        set_text(definition, 'subtotalCaption', text, _field_order())
    else:
        definition.attrib.pop('subtotalCaption', None)
    _changed(table, root)


def _index(value: object) -> int:
    try:
        index = to_integer(value)
    except (VBAError, ValueError, TypeError, OverflowError):
        raise error(1004) from None
    if not 1 <= index <= 12:
        raise error(1004)
    return index


def _flags(definition: XmlElement) -> list[bool]:
    """Subtotals as twelve flags: Automatic, then the eleven functions."""
    chosen = functions_of(definition)
    return [chosen == ('default',)] + [kind in chosen for kind in KINDS]


def subtotals(field: PivotField, index: object) -> object:
    """Subtotals(Index): one flag; Subtotals: all twelve, from 1 (tests/fixtures/pivot_layouts/
    subtotal_bounds)."""
    flags = _flags(_own(field))
    if index is not MISSING:
        return flags[_index(index) - 1]
    return VBAArray([(1, 12)], items=list(flags))


def set_subtotals(field: PivotField, index: object, value: object) -> None:
    """Subtotals(Index) = flag, or Subtotals = twelve flags. Automatic on clears the functions, and a function
    on turns Automatic off (tests/fixtures/pivot_layouts/ count_only, automatic_and_sum)."""
    flags = _flags(_own(field))
    if index is not MISSING:
        position = _index(index) - 1
        flags[position] = to_bool(value)
        if flags[position]:
            if position:
                flags[0] = False
            else:
                flags[1:] = [False] * 11
    else:
        if not isinstance(value, VBAArray) or value.dimensions != 1 or value.size != 12:
            raise error(1004)
        flags = [to_bool(one) for one in value.items]
        if flags[0]:
            flags[1:] = [False] * 11
    from pyopenvba.apps.excel._pivot_filters import ensure_items

    table, root, definition = _source(field)
    assert definition is not None
    for kind, chosen in zip(KINDS, flags[1:], strict=True):
        set_flag(definition, SUBTOTALS[kind][0], chosen, False)
    set_flag(definition, 'defaultSubtotal', any(flags), True)
    ordered(definition, _field_order())
    # A field lists its items as its subtotals change, one on no axis too (tests/fixtures/pivot_layouts/
    # hidden_field_subtotals).
    ensure_items(table, definition, field.index)
    _changed(table, root)
