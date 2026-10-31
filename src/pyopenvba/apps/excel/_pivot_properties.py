"""Small properties of PivotTable, PivotCache, PivotField and PivotItem over a worksheet range.

Measured in live Excel (scripts/measure_pivot_properties.py, tests/fixtures/pivot_properties.json):

- A flag a save writes sits on the table's, the cache's or the field's element, left out at its default, which a set
  back to it restores: SmallGrid is showDropZones turned round, PrintDrillIndicators printDrill,
  ShowPageMultipleItemLabel showMultipleLabel, DisplayMemberPropertyTooltips showMemberPropertyTips,
  PageFieldStyle and VacatedStyle the text of pageStyle and vacatedStyle, MissingItemsLimit missingItemsLimit
  (below 0 the default -1); a field's DragTo* its dragTo* and dragOff, IncludeNewItemsInFilter
  includeNewItemsInFilter, EnableItemSelection showDropDowns, and ShowAllItems showAll, which Excel writes as 0
  for a new field. Text that is no Boolean, and Null, are error 13; a data field's flags are error 1004.
- What an OLAP or external source alone answers is error 1004, read or set; so are the sets of the cache's and
  table's options such a source has. Hidden and the AutoShow settings are error 450 to set, AutoRefresh 445, and
  SourceCaption is error 5. VisualTotals keeps True whatever is set.
- A pivot item counts the source's records that hold it; its standard name is its name, and an item of no group
  has no parent, #N/A.
"""
from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

from pyopenvba.apps.excel import _pivots as pivots
from pyopenvba.apps.excel._pivot_create import encoded
from pyopenvba.apps.excel._pivot_filters import FIELD_ORDER, PAGE, orientation
from pyopenvba.apps.excel._pivot_forms import TABLE_ORDER, on, ordered, set_flag
from pyopenvba.apps.excel._pivot_options import CACHE_ORDER
from pyopenvba.apps.excel._pivots import NS, decode_text
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import MemberSpec
from pyopenvba.interpreter._values import NULL, VBAErrorValue, VBAInt, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement

#: The most MissingItemsLimit keeps, xlMissingItemsMax2.
MOST_MISSING = 1048576


def _cache_of(owner: object) -> pivots.PivotCache:
    if isinstance(owner, pivots.PivotCache):
        return owner
    if isinstance(owner, pivots.PivotTable):
        return owner.cache
    if isinstance(owner, pivots.PivotField):
        return owner.table.cache
    assert isinstance(owner, pivots.PivotItem)
    return owner.field.table.cache


def _range_source(owner: object) -> None:
    """These answers are a worksheet range's; another source's are not measured."""
    source = _cache_of(owner).root.find(NS + 'cacheSource')
    if source is None or source.get('type', 'worksheet') != 'worksheet':
        raise VBAUnsupportedError("a pivot property of a source other than a worksheet range is not implemented")


def _flag(value: object) -> bool:
    if value is NULL:
        raise error(13)
    return to_bool(value)


def _changed(owner: pivots.PivotTable | pivots.PivotCache) -> None:
    owner.changed = True
    (owner.book if isinstance(owner, pivots.PivotCache) else owner.sheet.book).saved = False


def _source_only(name: str) -> MemberSpec:
    """What only an OLAP or external source answers: error 1004, read or set."""

    def get(owner: object) -> object:
        _range_source(owner)
        raise error(1004)

    def put(owner: object, value: object) -> None:
        _range_source(owner)
        raise error(1004)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _fixed(name: str, value: object, refusal: int | None) -> MemberSpec:
    """A property a range's table answers alike: a set is error ``refusal``, or passed over where it is None."""

    def get(owner: object) -> object:
        _range_source(owner)
        return value

    def put(owner: object, given: object) -> None:
        _range_source(owner)
        if refusal is not None:
            raise error(refusal)
        _flag(given)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _table_flag(name: str, attribute: str, default: bool, *, turned: bool = False,
                refuse_true: bool = False) -> MemberSpec:
    def get(table: pivots.PivotTable) -> object:
        _range_source(table)
        return on(table.entry.root, attribute, default) != turned

    def put(table: pivots.PivotTable, value: object) -> None:
        _range_source(table)
        wanted = _flag(value)
        if refuse_true and wanted:
            # Set back to True, Excel's automation stopped answering (scripts/measure_pivot_properties.py).
            raise VBAUnsupportedError(f"setting PivotTable.{name} to True is not implemented")
        table = table.entry
        set_flag(table.root, attribute, wanted != turned, default, TABLE_ORDER)
        _changed(table)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _table_text(name: str, attribute: str) -> MemberSpec:
    def get(table: pivots.PivotTable) -> object:
        _range_source(table)
        return decode_text(table.entry.root.get(attribute, ''))

    def put(table: pivots.PivotTable, value: object) -> None:
        _range_source(table)
        text = to_text(value)
        table = table.entry
        if text:
            table.root.set(attribute, encoded(text))
            ordered(table.root, TABLE_ORDER)
        else:
            table.root.attrib.pop(attribute, None)
        _changed(table)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _own(field: pivots.PivotField) -> XmlElement:
    """A field's own pivotField; a data field's and the Values field's are error 1004."""
    if field.data or field.index < 0:
        raise error(1004)
    return field.root


def _field_flag(name: str, attribute: str, default: bool) -> MemberSpec:
    def get(field: pivots.PivotField) -> object:
        _range_source(field)
        return on(_own(field), attribute, default)

    def put(field: pivots.PivotField, value: object) -> None:
        _range_source(field)
        element = _own(field)
        set_flag(element, attribute, _flag(value), default, FIELD_ORDER)
        _changed(field.table.entry)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _show_all() -> MemberSpec:
    def get(field: pivots.PivotField) -> object:
        _range_source(field)
        if field.data:
            return False
        return on(_own(field), 'showAll', True)

    def put(field: pivots.PivotField, value: object) -> None:
        _range_source(field)
        if field.data:
            raise VBAUnsupportedError("setting a data field's ShowAllItems is not implemented")
        element = _own(field)
        wanted = _flag(value)
        if wanted:
            axis = orientation(element)
            if axis == 0:
                raise error(1004)
            if axis != PAGE:
                raise VBAUnsupportedError("showing a row or column field's items with no data is not implemented")
            element.attrib.pop('showAll', None)
        else:
            element.set('showAll', '0')
            ordered(element, FIELD_ORDER)
        _changed(field.table.entry)

    return MemberSpec(name='ShowAllItems', kind='property', getter=get, setter=put)


def _auto_show(name: str, value: object) -> MemberSpec:
    """A setting of AutoShow, which Excel sets only through AutoShow: error 450 to set."""

    def get(field: pivots.PivotField) -> object:
        _range_source(field)
        if _own(field).get('autoShow') in ('1', 'true'):
            raise VBAUnsupportedError("reading a field's AutoShow settings is not implemented")
        return value

    def put(field: pivots.PivotField, given: object) -> None:
        raise error(450)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _field_fixed(name: str, value: object) -> MemberSpec:
    """A property a range's field answers alike, a data field error 1004; a set is error 1004."""

    def get(field: pivots.PivotField) -> object:
        _range_source(field)
        _own(field)
        return value

    def put(field: pivots.PivotField, given: object) -> None:
        _range_source(field)
        raise error(1004)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _missing_items() -> MemberSpec:
    def get(cache: pivots.PivotCache) -> object:
        _range_source(cache)
        return VBAInt(int(cache.root.get('missingItemsLimit', '-1')), 'Long')

    def put(cache: pivots.PivotCache, value: object) -> None:
        _range_source(cache)
        limit = int(to_integer(value, 'Long'))
        if limit > MOST_MISSING:
            raise VBAUnsupportedError("a MissingItemsLimit past xlMissingItemsMax2 is not implemented")
        if limit < 0:
            cache.root.attrib.pop('missingItemsLimit', None)
        else:
            cache.root.set('missingItemsLimit', str(limit))
            ordered(cache.root, CACHE_ORDER)
        _changed(cache)

    return MemberSpec(name='MissingItemsLimit', kind='property', getter=get, setter=put)


def _error_getter(name: str, number: int) -> MemberSpec:
    def get(owner: object) -> object:
        raise error(number)

    def put(owner: object, value: object) -> None:
        raise error(number)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _shared(item: pivots.PivotItem) -> XmlElement:
    fields = item.field.table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
    if item.field.data or not 0 <= item.field.index < len(fields) or item.x < 0:
        raise VBAUnsupportedError("this pivot item's source values are not implemented")
    field = fields[item.field.index]
    if field.find(NS + 'fieldGroup') is not None:
        raise VBAUnsupportedError("a grouped field's item records are not implemented")
    shared = field.find(NS + 'sharedItems')
    if shared is None or item.x >= len(shared):
        raise VBAUnsupportedError("this pivot item's source values are not implemented")
    return list(shared)[item.x]


def _record_count(item: pivots.PivotItem) -> object:
    from pyopenvba.apps.excel._pivot_aggregate import records

    _range_source(item)
    wanted = _shared(item)
    return VBAInt(sum(1 for row in records(item.field.table.cache) if row[item.field.index] is wanted), 'Long')


def _standard_name(item: pivots.PivotItem) -> object:
    _range_source(item)
    wanted = _shared(item)
    if wanted.tag != NS + 's':
        raise VBAUnsupportedError("the standard name of a pivot item that is not text is not implemented")
    return decode_text(wanted.get('v', ''))


def _parent_show_detail(item: pivots.PivotItem) -> object:
    _range_source(item)
    _shared(item)
    return VBAErrorValue(2042)


def _getter(name: str, get: Callable[[pivots.PivotItem], object]) -> MemberSpec:
    def put(item: pivots.PivotItem, value: object) -> None:
        raise VBAUnsupportedError(f"setting PivotItem.{name} is not implemented")

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


for _spec in (_missing_items(), _fixed('BackgroundQuery', False, 1004), _fixed('SavePassword', False, 1004),
              _fixed('UpgradeOnRefresh', False, 1004),
              *(_source_only(name) for name in (
                  'RefreshPeriod', 'CommandText', 'CommandType', 'Connection', 'IsConnected', 'LocalConnection',
                  'MaintainConnection', 'QueryType', 'RobustConnect', 'SourceConnectionFile', 'SourceDataFile',
                  'UseLocalConnection', 'WorkbookConnection', 'ADOConnection', 'Recordset'))):
    pivots.PivotCache.vba_add_member(_spec)


def _optimize_cache() -> MemberSpec:
    def get(cache: pivots.PivotCache) -> object:
        _range_source(cache)
        return on(cache.root, 'optimizeMemory', False)

    def put(cache: pivots.PivotCache, value: object) -> None:
        _range_source(cache)
        raise error(1004)

    return MemberSpec(name='OptimizeCache', kind='property', getter=get, setter=put)


pivots.PivotCache.vba_add_member(_optimize_cache())
for _spec in (_table_flag('SmallGrid', 'showDropZones', True, turned=True),
              _table_flag('PrintDrillIndicators', 'printDrill', False),
              _table_flag('ShowPageMultipleItemLabel', 'showMultipleLabel', True),
              _table_flag('DisplayMemberPropertyTooltips', 'showMemberPropertyTips', True, refuse_true=True),
              _table_text('PageFieldStyle', 'pageStyle'), _table_text('VacatedStyle', 'vacatedStyle'),
              _fixed('EnableWriteback', False, 1004), _fixed('TotalsAnnotation', False, 1004),
              _fixed('VisualTotals', True, None), _fixed('VisualTotalsForSets', False, 1004),
              _fixed('CalculatedMembersInFilters', False, 1004), _fixed('Hidden', False, 450),
              _fixed('AutoRefresh', False, 445),
              *(_source_only(name) for name in (
                  'AllocationMethod', 'AllocationValue', 'Allocation', 'AllocationWeightExpression',
                  'ViewCalculatedMembers', 'MDX', 'ChangeList', 'CubeFields', 'CalculatedMembers', 'PivotChart',
                  'SelectionMode', 'PivotSelection', 'PivotSelectionStandard', 'InnerDetail'))):
    pivots.PivotTable.vba_add_member(_spec)
for _spec in (_field_flag('DragToColumn', 'dragToCol', True), _field_flag('DragToRow', 'dragToRow', True),
              _field_flag('DragToPage', 'dragToPage', True), _field_flag('DragToData', 'dragToData', True),
              _field_flag('DragToHide', 'dragOff', True),
              _field_flag('IncludeNewItemsInFilter', 'includeNewItemsInFilter', False),
              _field_flag('EnableItemSelection', 'showDropDowns', True), _show_all(),
              _field_fixed('DatabaseSort', False), _field_fixed('ServerBased', False),
              _auto_show('AutoShowCount', VBAInt(10, 'Long')), _auto_show('AutoShowField', ''),
              _auto_show('AutoShowRange', VBAInt(-4160, 'Long')), _auto_show('AutoShowType', VBAInt(-4135, 'Long')),
              _error_getter('SourceCaption', 5),
              *(_source_only(name) for name in (
                  'CurrentPageList', 'CurrentPageName', 'DisplayAsCaption', 'DisplayAsTooltip', 'DisplayInReport',
                  'DrilledDown', 'Hidden', 'HiddenItemsList', 'IsMemberProperty', 'PropertyOrder',
                  'UseMemberPropertyAsCaption', 'VisibleItemsList', 'CubeField', 'PropertyParentField',
                  'MemberPropertyCaption'))):
    pivots.PivotField.vba_add_member(_spec)
for _spec in (_source_only('DrilledDown'), _getter('RecordCount', _record_count),
              _getter('SourceNameStandard', _standard_name), _getter('ParentShowDetail', _parent_show_detail)):
    pivots.PivotItem.vba_add_member(_spec)
