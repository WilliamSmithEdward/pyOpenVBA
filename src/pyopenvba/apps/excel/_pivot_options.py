"""A pivot table's options, as Excel keeps and shows them.

TableStyle2 names the style Excel draws the table in, which it keeps
apart from the cells' formats, with five flags for the parts it draws;
NullString shows in each value cell no record reaches and ErrorString in
each that works out an error, where the table displays them; the page
fields lie in PageFieldOrder, wrapped after PageFieldWrapCount of them;
and fourteen flags change no cell at all. AddFields puts fields on the
rows, columns and page area, ClearTable takes every field away, Location
moves the table, and ManualUpdate holds its lines back until it is set
off again.

Measured with scripts/measure_pivot_options.py (tests/fixtures/pivot_options/):
a null string reads into its cells as a whole number where, trimmed of
spaces, it is a minus and digits or either alone, "-" and a space being 0,
and as text otherwise; DisplayNullString off shows 0. ManualUpdate on drops
the table's lines at once, though its cells stand until it is laid out
again; laid out while on, it shows its captions alone. An unknown style
is no style. Reading Location is error 445; SubtotalHiddenPageItems and
DisplayImmediateItems off are error 5 and DisplayEmptyRow and
DisplayEmptyColumn error 1004, as on any table not from a cube.
"""
from __future__ import annotations

import re
from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._pivot_create import encoded, node
from pyopenvba.apps.excel._pivot_fit import refit
from pyopenvba.apps.excel._pivot_forms import TABLE_ORDER, X14, on, ordered, set_flag, set_text
from pyopenvba.apps.excel._pivots import NS, STYLE_FLAGS, TABLE_FLAGS as FLAGS, decode_text
from pyopenvba.exceptions import VBAError, VBAUnsupportedError
from pyopenvba.interpreter._values import MISSING, VBAArray, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.apps.excel._pivots import PivotTable

#: Excel's own styles a pivot table takes: its pivot styles and its table styles (a probe in Excel).
BUILT_IN_STYLES = [*(f'PivotStyle{weight}{number}' for weight in ('Light', 'Medium', 'Dark') for number in range(1, 29)),
                   *(f'TableStyle{weight}{number}' for weight, count in (('Light', 21), ('Medium', 28), ('Dark', 11))
                     for number in range(1, count + 1))]
#: What TableStyle2 reads for a table in no style: a NUL (a probe in Excel).
NO_STYLE = '\x00'
#: The table's Excel 2010 options, where AlternativeText and Summary live.
OPTIONS_ORDER = ('fillDownLabelsDefault visualTotalsForSets calculatedMembersInFilters altText altTextSummary '
                 'enableEdit autoApply allocationMethod weightExpression hideValuesRow').split()
TABLE_EXTENSION = '{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}'
X14_MAIN = X14[1:-1]
XM_MAIN = 'http://schemas.microsoft.com/office/excel/2006/main'


def _changed(table: PivotTable) -> None:
    table.changed = True
    table.sheet.book.saved = False


def _laid_out(table: PivotTable, root: XmlElement) -> None:
    from pyopenvba.apps.excel._pivot_filters import rerender

    rerender(table, root)
    table.sheet.book.saved = False


# -- flags that show in no cell


def flag(table: PivotTable, name: str) -> bool:
    attribute, default, opposite = FLAGS[name]
    return on(table.entry.root, attribute, default) != opposite


def set_flag_member(table: PivotTable, name: str, value: object) -> None:
    attribute, default, opposite = FLAGS[name]
    table = table.entry
    set_flag(table.root, attribute, to_bool(value) != opposite, default, TABLE_ORDER)
    _changed(table)


def tag(table: PivotTable) -> str:
    return decode_text(table.entry.root.get('tag', ''))


def set_tag(table: PivotTable, value: object) -> None:
    table = table.entry
    text = to_text(value)
    if text:
        set_text(table.root, 'tag', text, TABLE_ORDER)
    else:
        table.root.attrib.pop('tag', None)
    _changed(table)


def table_options(root: XmlElement) -> XmlElement:
    """The table's Excel 2010 options, made where missing, spelled as Excel spells them."""
    extensions = root.find(NS + 'extLst')
    if extensions is None:
        extensions = node('extLst')
        root.append(extensions)
    extension = next((entry for entry in extensions if entry.get('uri') == TABLE_EXTENSION), None)
    if extension is None:
        extension = node('ext', {'uri': TABLE_EXTENSION, 'xmlns:x14': X14_MAIN})
        extensions.insert(0, extension)
    options = extension.find(X14 + 'pivotTableDefinition')
    if options is None:
        options = XmlElement(X14 + 'pivotTableDefinition', {'xmlns:xm': XM_MAIN})
        extension.append(options)
    return options


def _found_options(root: XmlElement) -> XmlElement | None:
    for extension in root.findall(f'{NS}extLst/{NS}ext'):
        if extension.get('uri') == TABLE_EXTENSION:
            return extension.find(X14 + 'pivotTableDefinition')
    return None


def _prune(root: XmlElement) -> None:
    """Take away the table's Excel 2010 extension where its options hold nothing, as Excel does (tests/fixtures/
    pivot_options/ values_row)."""
    extensions = root.find(NS + 'extLst')
    if extensions is None:
        return
    for extension in list(extensions):
        options = extension.find(X14 + 'pivotTableDefinition') if extension.get('uri') == TABLE_EXTENSION else None
        if options is not None and not any(not key.startswith('xmlns') for key in options.attrib):
            extensions.remove(extension)


def alternative(table: PivotTable, attribute: str) -> str:
    options = _found_options(table.entry.root)
    return decode_text(options.get(attribute, '')) if options is not None else ''


def set_alternative(table: PivotTable, attribute: str, value: object) -> None:
    """AlternativeText (altText) or Summary (altTextSummary), kept in the table's Excel 2010 options."""
    table = table.entry
    options = table_options(table.root)
    text = to_text(value)
    if text:
        options.set(attribute, encoded(text))
    else:
        options.attrib.pop(attribute, None)
    ordered(options, OPTIONS_ORDER, others_first=False)
    _prune(table.root)
    _changed(table)


def shows_values_row(table: PivotTable) -> bool:
    return values_row_shown(table.entry.root)


def values_row_shown(root: XmlElement) -> bool:
    options = _found_options(root)
    return not (options is not None and on(options, 'hideValuesRow', False))


def set_shows_values_row(table: PivotTable, value: object) -> None:
    """ShowValuesRow: the Values row over the data fields' captions where the Values field is the only field on
    the columns, the table's hideValuesRow taken away (tests/fixtures/pivot_options/ values_row)."""
    table = table.entry
    root = deepcopy(table.root)
    options = table_options(root)
    if to_bool(value):
        options.attrib.pop('hideValuesRow', None)
    else:
        options.set('hideValuesRow', '1')
    ordered(options, OPTIONS_ORDER, others_first=False)
    _prune(root)
    _laid_out(table, root)


def set_kept_on(table: PivotTable, value: object, attribute: str) -> None:
    """DisplayImmediateItems and SubtotalHiddenPageItems: on as they are, any other value error 5."""
    if to_bool(value) != on(table.entry.root, attribute, attribute == 'showItems'):
        raise error(5)


# -- the style


def table_style(table: PivotTable) -> str:
    info = table.entry.root.find(NS + 'pivotTableStyleInfo')
    name = decode_text(info.get('name', '')) if info is not None else ''
    return name or NO_STYLE


def _styles(table: PivotTable) -> list[str]:
    """The pivot styles the workbook knows: Excel's own and its file's that pivot tables may take."""
    xml = table.sheet.book.stylesheet.xml
    found = re.search(r'<(?:[\w.-]+:)?tableStyles\b[^>]*>(.*?)</(?:[\w.-]+:)?tableStyles>', xml, re.DOTALL)
    own = [] if found is None else [decode_text(one.group(1)) for one in re.finditer(
        r'<(?:[\w.-]+:)?tableStyle\b(?=[^>]*\bname="([^"]*)")(?![^>]*\bpivot="(?:0|false)")[^>]*>', found.group(1))]
    return [*BUILT_IN_STYLES, *own]


def set_table_style(table: PivotTable, value: object) -> None:
    """TableStyle2: a pivot or table style's name, in any case, or a TableStyle (tests/fixtures/pivot_options/
    style, style_lower, style_table). Excel leaves a table given a name no style has, or none, in no style, which it
    draws with borders of the cells themselves; that reports itself unsupported here."""
    from pyopenvba.apps.excel._model import ExcelObject

    table = table.entry
    text = to_text(value.vba_get('Name') if isinstance(value, ExcelObject) else value)
    name = next((style for style in _styles(table) if style.casefold() == text.casefold()), '')
    if not name:
        raise VBAUnsupportedError('A pivot table in no style, which Excel draws with borders of its cells, is not '
                                  'implemented')
    info = table.root.find(NS + 'pivotTableStyleInfo')
    if info is None:
        info = node('pivotTableStyleInfo', {'showRowHeaders': '1', 'showColHeaders': '1', 'showRowStripes': '0',
                                            'showColStripes': '0', 'showLastColumn': '1'})
        extensions = table.root.find(NS + 'extLst')
        table.root.insert(list(table.root).index(extensions) if extensions is not None else len(table.root), info)
    rest = {key: info.get(key, '') for key in info.attrib if key != 'name'}
    info.attrib.clear()
    info.set('name', encoded(name))
    info.attrib.update(rest)
    if table.manual:
        # A table ManualUpdate holds back lays itself out again as its captions alone (tests/fixtures/
        # pivot_options/ manual_style).
        _laid_out(table, deepcopy(table.root))
    else:
        # Its columns fit its cells in the new style's fonts (tests/fixtures/pivot_options/ style_lower).
        refit(table)
    _changed(table)


def style_flag(table: PivotTable, name: str) -> bool:
    info = table.entry.root.find(NS + 'pivotTableStyleInfo')
    return info is not None and on(info, STYLE_FLAGS[name], False)


def set_style_flag(table: PivotTable, name: str, value: object) -> None:
    table = table.entry
    info = table.root.find(NS + 'pivotTableStyleInfo')
    if info is None:
        info = node('pivotTableStyleInfo', {'showRowHeaders': '1', 'showColHeaders': '1', 'showRowStripes': '0',
                                            'showColStripes': '0', 'showLastColumn': '1'})
        extensions = table.root.find(NS + 'extLst')
        table.root.insert(list(table.root).index(extensions) if extensions is not None else len(table.root), info)
    info.set(STYLE_FLAGS[name], '1' if to_bool(value) else '0')
    # The flags turn the style's bold on or off, which the columns take as the table is next laid out, not now
    # (tests/fixtures/pivot_column_fit/ flags_off).
    _changed(table)


# -- what empty and error cells show


def caption(table: PivotTable, attribute: str) -> str:
    return decode_text(table.entry.root.get(attribute, ''))


def set_caption(table: PivotTable, attribute: str, value: object) -> None:
    """NullString (missingCaption) or ErrorString (errorCaption)."""
    table = table.entry
    root = deepcopy(table.root)
    text = to_text(value)
    if text:
        set_text(root, attribute, text, TABLE_ORDER)
    else:
        root.attrib.pop(attribute, None)
    _laid_out(table, root)


def set_display(table: PivotTable, attribute: str, value: object, default: bool) -> None:
    """DisplayNullString (showMissing) or DisplayErrorString (showError)."""
    table = table.entry
    root = deepcopy(table.root)
    set_flag(root, attribute, to_bool(value), default, TABLE_ORDER)
    _laid_out(table, root)


def shown_text(text: str) -> object:
    """What a cell holds for a null or error string: a whole number where, trimmed of spaces, the text is a minus
    and digits, either alone being 0, else the text (tests/fixtures/pivot_options/ null_letter, and a probe of 46
    texts in Excel)."""
    found = re.fullmatch(r' *(-?)(\d*) *', text)
    if found is None:
        return text
    number = float(found.group(2) or '0')
    return -number if found.group(1) and number else number


# -- the page fields


def page_order(table: PivotTable) -> int:
    """PageFieldOrder: xlDownThenOver (1) or xlOverThenDown (2)."""
    return 2 if on(table.entry.root, 'pageOverThenDown', False) else 1


def set_page_order(table: PivotTable, value: object) -> None:
    chosen = _whole(value, 1004)
    if chosen not in (1, 2):
        raise error(1004)
    table = table.entry
    root = deepcopy(table.root)
    set_flag(root, 'pageOverThenDown', chosen == 2, False, TABLE_ORDER)
    _laid_out(table, root)


def page_wrap(table: PivotTable) -> int:
    return int(table.entry.root.get('pageWrap', '0'))


def set_page_wrap(table: PivotTable, value: object) -> None:
    """PageFieldWrapCount: 0 to 255, else error 1004 (tests/fixtures/pivot_options/ page_errors)."""
    chosen = _whole(value, 1004)
    if not 0 <= chosen <= 255:
        raise error(1004)
    table = table.entry
    root = deepcopy(table.root)
    if chosen:
        root.set('pageWrap', str(chosen))
    else:
        root.attrib.pop('pageWrap', None)
    ordered(root, TABLE_ORDER)
    _laid_out(table, root)


def _whole(value: object, code: int) -> int:
    try:
        return to_integer(value)
    except (VBAError, ValueError, TypeError, OverflowError):
        raise error(code) from None


# -- laying the table out again


def _named(table: PivotTable, value: object) -> list[int]:
    """The fields AddFields names, by name or as an array of names; one no field has, or named twice, is error
    1004 (tests/fixtures/pivot_options/ add_fields_bad, add_fields_twice)."""
    if value is MISSING:
        return []
    names = [to_text(one) for one in value.items] if isinstance(value, VBAArray) else [to_text(value)]
    sources = table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
    data_caption = decode_text(table.root.get('dataCaption', 'Values')).casefold()
    found: list[int] = []
    for name in names:
        if name.casefold() == data_caption:
            index = -2
        else:
            index = next((place for place, field in enumerate(sources)
                          if decode_text(field.get('name', '')).casefold() == name.casefold()), None)
            if index is None:
                raise error(1004)
        if index in found:
            raise error(1004)
        found.append(index)
    return found


def add_fields(table: PivotTable, rows: object, columns: object, pages: object, add_to_table: object) -> bool:
    """AddFields: the fields named on the rows, columns and page area, in the order named, each taken from
    wherever it was; the axes' other fields leave them unless AddToTable is True (tests/fixtures/pivot_options/
    add_fields, add_fields_to_table, add_fields_none)."""
    table = table.entry
    wanted = (_named(table, rows), _named(table, columns), _named(table, pages))
    keep = add_to_table is not MISSING and to_bool(add_to_table)
    if len({index for axis in wanted for index in axis}) != sum(len(axis) for axis in wanted):
        raise error(1004)
    root = deepcopy(table.root)
    named = {index for axis in wanted for index in axis}
    groups = (('rowFields', 'field', 'x'), ('colFields', 'field', 'x'), ('pageFields', 'pageField', 'fld'))
    axes: list[list[XmlElement]] = []
    for (name, _, attribute), axis in zip(groups, wanted, strict=True):
        container = root.find(NS + name)
        old = list(container) if container is not None else []
        kept = [entry for entry in old if keep and int(entry.get(attribute, '0')) not in named]
        if name == 'pageFields':
            # Each page field goes ahead of the ones before it, as it does by Orientation (tests/fixtures/
            # pivot_options/ add_fields_pages).
            axes.append([node('pageField', {'fld': str(index), 'hier': '-1'}) for index in reversed(axis)] + kept)
        else:
            axes.append(kept + [node('field', {'x': str(index)}) for index in axis])
        if container is not None:
            root.remove(container)
    fields = root.findall(f'{NS}pivotFields/{NS}pivotField')
    on_axis: dict[int, str] = {}
    for (name, _, attribute), entries in zip(groups, axes, strict=True):
        for entry in entries:
            on_axis[int(entry.get(attribute, '0'))] = {'rowFields': 'axisRow', 'colFields': 'axisCol',
                                                       'pageFields': 'axisPage'}[name]
    for index, field in enumerate(fields):
        if index in on_axis:
            field.set('axis', on_axis[index])
        else:
            field.attrib.pop('axis', None)
    location = root.find(NS + 'location')
    place = list(root).index(location) + 2 if location is not None else 1
    for (name, _, _), entries in zip(groups, axes, strict=True):
        if entries:
            container = node(name, {'count': str(len(entries))})
            container.extend(entries)
            root.insert(place, container)
            place += 1
    _laid_out(table, root)
    return True


def clear_table(table: PivotTable) -> None:
    """ClearTable: every field off the table, its data fields and filters too, leaving the empty table (tests/
    fixtures/pivot_options/ clear_table, clear_filters). Every table of its cache is laid out again with it, as a
    refresh lays them out but from the cache as it is, so that one whose style flags wait fits its columns to them
    (tests/fixtures/pivot_options/ style; tests/fixtures/calculate_order.json clear_table, clear_table_manual)."""
    from pyopenvba.apps.excel._pivot_layout import apply_render, plan_render
    from pyopenvba.apps.excel._pivot_refresh import announce, cache_tables

    table = table.entry
    root = deepcopy(table.root)
    for name in ('rowFields', 'colFields', 'pageFields', 'dataFields', 'filters'):
        for container in root.findall(NS + name):
            root.remove(container)
    for field in root.findall(f'{NS}pivotFields/{NS}pivotField'):
        for key in ('axis', 'dataField', 'measureFilter'):
            field.attrib.pop(key, None)
    plans = [plan_render(other, root, []) if other is table else plan_render(other, deepcopy(other.root))
             for other in cache_tables(table.cache)]
    for plan in plans:
        apply_render(plan)
    table.sheet.book.saved = False
    announce(table.sheet.book, plans)


def set_location(table: PivotTable, value: object) -> None:
    """Location: the table's top left cell, by an address in A1 style, its sheet's name before it or not; the
    table moves there, its page fields above it (tests/fixtures/pivot_options/ location_page); anything else is
    error 1004."""
    table = table.entry
    text = to_text(value).strip()
    if '!' in text:
        sheet, _, text = text.rpartition('!')
        if sheet.strip("'").casefold() != table.sheet.name.casefold():
            raise error(1004)
    try:
        area = parse_area(text.replace('$', ''))
    except ValueError:
        raise error(1004) from None
    root = deepcopy(table.root)
    location = root.find(NS + 'location')
    assert location is not None
    old = parse_area(location.get('ref', 'A1'))
    moved = Area(area.top, area.left, area.top + old.rows - 1, area.left + old.columns - 1, table.sheet.name)
    location.set('ref', moved.address(absolute=False))
    _laid_out(table, root)


#: A pivot cache's attributes in the order the schema has them.
CACHE_ORDER = ('id invalid saveData refreshOnLoad optimizeMemory enableRefresh refreshedBy refreshedDate '
               'refreshedDateIso backgroundQuery missingItemsLimit createdVersion refreshedVersion '
               'minRefreshableVersion recordCount upgradeOnRefresh tupleCache supportSubquery '
               'supportAdvancedDrill').split()


def set_save_data(table: PivotTable, value: object) -> None:
    """SaveData: off, the cache is saved without its records, until another table is made from it (a probe in
    Excel); error 1004 on a table that edits its values. A cache a file holds already is not saved without them
    here."""
    table = table.entry
    if to_bool(value):
        table.cache.root.attrib.pop('saveData', None)
    elif table.root.get('editData', '0') in ('1', 'true'):
        raise error(1004)
    elif table.cache.part:
        raise VBAUnsupportedError('Saving a pivot cache a file holds without its records is not implemented')
    else:
        table.cache.root.set('saveData', '0')
        ordered(table.cache.root, CACHE_ORDER)
        # The table it was set through marks its totals, which it keeps once the cache saves its records again
        # (tests/fixtures/pivot_options/ save_data).
        table.root.set('asteriskTotals', '1')
        ordered(table.root, TABLE_ORDER)
        table.changed = True
    table.cache.changed = True
    table.sheet.book.saved = False


# -- holding the lines back


def set_manual(table: PivotTable, value: object) -> None:
    """ManualUpdate: on, the table drops its lines and lays itself out as its captions alone until it is set off,
    which lays it out in full (tests/fixtures/pivot_options/ manual_pending, manual_done)."""
    table = table.entry
    wanted = to_bool(value)
    was = table.manual
    table.manual = wanted
    if wanted and not was:
        _drop_lines(table.root)
        _changed(table)
    elif was and not wanted:
        _laid_out(table, deepcopy(table.root))


def settle_manual(book: Workbook) -> None:
    """Before a save, lay out each table ManualUpdate holds back as its captions alone, as Excel saves it (tests/
    fixtures/pivot_options/ manual_unchanged), raising no events."""
    from pyopenvba.apps.excel._pivot_layout import apply_render, plan_render

    for sheet in book.sheets_:
        for table in sheet.pivot_tables:
            if table.entry is table and table.manual:
                apply_render(plan_render(table, deepcopy(table.root)))


def _drop_lines(root: XmlElement) -> None:
    for name in ('rowItems', 'colItems'):
        for container in root.findall(NS + name):
            root.remove(container)
