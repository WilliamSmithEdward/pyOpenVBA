"""Names in a pivot table, as Excel takes and shows them: a data field's, the Values field's, a field's and an item's.

Measured with scripts/measure_pivot_names.py (tests/fixtures/pivot_names/):

- A data field's Name or Caption is the dataField's name, error 1004 where a field of the cache, a field's caption
  or another data field has it in any case; a name with a space more is another name, and an empty one is a name too.
- The Values field's is the table's dataCaption.
- A field's Caption is the name the table shows it by and formulas name it by, its cache's name kept: error 5 where
  another field or a data field has it, and an empty one takes the field back to its cache's name. Name does the same
  on a field on an axis; on a field on none it lists the field's items, takes the name and then is error 1004, and on
  a calculated field it is error 1004.
- An item's Name, Caption and Value are its caption, kept as its n, which Excel writes first: error 1004 where
  another item of the field has it in any case, and an empty one is a caption too. The item shows, sorts and is found
  by its caption from then on, its SourceName still the value.
"""
from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._pivot_create import encoded
from pyopenvba.apps.excel._pivots import NS, decode_text
from pyopenvba.interpreter._values import error, string_order, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivots import PivotField, PivotItem, PivotTable


def _same(left: str, right: str) -> bool:
    return string_order(left.casefold(), right.casefold()) == 0


def _changed(table: PivotTable) -> None:
    from pyopenvba.apps.excel._pivot_filters import rerender

    rerender(table, deepcopy(table.root))
    table.sheet.book.saved = False


def _data_names(table: PivotTable, but: XmlElement | None = None) -> list[str]:
    return [decode_text(entry.get('name', '')) for entry in table.root.findall(f'{NS}dataFields/{NS}dataField')
            if entry is not but]


def rename_data(field: PivotField, value: object) -> None:
    """Name or Caption of a data field: error 1004 where a field shows by it or another data field has it; a field's
    cache name it no longer shows by is free (tests/fixtures/pivot_names/ data_caption, data_source_name,
    data_duplicate, data_empty, data_field_caption, data_longest)."""
    from pyopenvba.apps.excel._pivot_calculated import captions

    table = field.table.entry
    text = to_text(value)
    if any(_same(one, text) for one in [*captions(table), *_data_names(table, field.root)]):
        raise error(1004)
    field.root.set('name', encoded(text))
    _changed(table)


def rename_values(field: PivotField, value: object) -> None:
    """Name or Caption of the Values field: the table's dataCaption, error 5 where a field shows by it or a data
    field has it (tests/fixtures/pivot_names/ values_caption, values_outline, values_field_name)."""
    from pyopenvba.apps.excel._pivot_calculated import captions

    table = field.table.entry
    text = to_text(value)
    if any(_same(one, text) for one in [*captions(table), *_data_names(table)]):
        raise error(5)
    table.root.set('dataCaption', encoded(text))
    _changed(table)


def rename_field(field: PivotField, value: object, *, caption: bool) -> None:
    """Caption or Name of a field (tests/fixtures/pivot_names/ row_outline, hidden_name, hidden_caption,
    field_duplicate, field_page; tests/fixtures/pivot_calculated_fields/ renamed, source_name)."""
    from pyopenvba.apps.excel._pivot_calculated import calculated_view, captions
    from pyopenvba.apps.excel._pivot_filters import ensure_items, field_attributes

    table = field.table.entry
    on_axis = field.root.get('axis') is not None
    text = to_text(value)
    if not caption and (calculated_view(field) or text == ''):
        # (tests/fixtures/pivot_names/ field_empty_name)
        raise error(1004)
    others = [*(one for place, one in enumerate(captions(table)) if place != field.index), *_data_names(table)]
    if any(_same(one, text) for one in others):
        raise error(5)
    quirk = not caption and not on_axis
    if quirk:
        ensure_items(table, field.root, field.index)
    if text == '' and caption:
        field.root.attrib.pop('name', None)
    else:
        field.root.set('name', encoded(text))
    ordered = field_attributes(dict(field.root.attrib))
    field.root.attrib.clear()
    field.root.attrib.update(ordered)
    table.changed = True
    table.sheet.book.saved = False
    if on_axis:
        # Outline and tabular forms show a field's name over its items, and a page field beside its choice.
        _changed(table)
    if quirk:
        raise error(1004)


def item_entry(field: PivotField, place: int) -> XmlElement | None:
    """The field's own entry for the item whose value is at ``place`` in the cache, where it lists one."""
    from pyopenvba.apps.excel._pivot_filters import listed

    return next((item for item in listed(field.root) if item.get('x') == str(place)), None)


def caption_of(field: PivotField, place: int) -> str | None:
    """An item's caption, where it was given one."""
    entry = item_entry(field, place)
    caption = entry.get('n') if entry is not None else None
    return decode_text(caption) if caption is not None else None


def rename_item(item: PivotItem, value: object) -> None:
    """Name, Caption or Value of an item (tests/fixtures/pivot_names/ item_caption, item_duplicate, item_reset,
    item_sorted, item_then_sorted)."""
    from pyopenvba.apps.excel._pivot_filters import ensure_items, restate
    from pyopenvba.apps.excel._pivot_sorts import resort_by_name

    field = item.field
    table = field.table.entry
    text = to_text(value)
    ensure_items(table, field.root, field.index)
    others = [one for one in field.PivotItems().vba_items() if getattr(one, 'x', -1) != item.x]
    if any(_same(str(one.vba_get('Name')), text) for one in others):
        raise error(1004)
    entry = item_entry(field, item.x)
    if entry is None:
        raise error(1004)
    entry.set('n', encoded(text))
    restate(entry, hide=entry.get('h') in ('1', 'true'))
    # A field sorted by its items' names sorts by their captions (item_sorted, item_then_sorted).
    resort_by_name(table, field.root, field.index)
    table.changed = True
    _changed(table)
