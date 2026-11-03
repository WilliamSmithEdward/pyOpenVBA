"""The error a collection's Item raises for each kind of index, late bound (tests/fixtures/collection_items.json). No rule
gives these: Sheets answers 13 for Null and Names 1004, and an out-of-range index is 9, 1004 or something else by class.
A kind is listed where every way in that fails agrees on the error; otherwise the class's own lookup answers. A chart's
Points and series are left out: o.Item(x) on an Object answers otherwise than sr.Points(x), which their own fixtures hold."""
from __future__ import annotations

from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL

ITEM_ERRORS: dict[str, dict[str, int]] = {
    'Areas': {'missing': 450, 'past': 1004, 'zero': 1004, 'negative': 1004, 'lacking': 13, 'blank': 13,
        'empty': 1004, 'null': 13, 'true': 1004, 'object': 13},
    'CalculatedFields': {'missing': 450, 'past': 9, 'zero': 1004, 'negative': 1004, 'fraction': 9, 'text': 9,
        'lacking': 9, 'blank': 9, 'empty': 1004, 'null': 13, 'true': 1004, 'object': 13},
    'CalculatedItems': {'missing': 450, 'past': 9, 'zero': 1004, 'negative': 1004, 'fraction': 9, 'text': 9,
        'lacking': 9, 'blank': 9, 'empty': 1004, 'null': 13, 'true': 1004, 'object': 13},
    'ChartGroups': {'missing': 450, 'past': 1004, 'zero': 1004, 'negative': 1004, 'fraction': 1004, 'lacking': 13,
        'blank': 13, 'empty': 1004, 'null': 13, 'true': 1004, 'object': 13},
    'ChartObjects': {'missing': 450, 'past': 1004, 'zero': 1004, 'negative': 1004, 'fraction': 1004, 'text': 1004,
        'empty': 1004, 'true': 1004},
    'ColorScaleCriteria': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'lacking': 9, 'blank': 9, 'empty': 9,
        'null': 9, 'true': 9, 'object': 9},
    'ColorStops': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'lacking': 9, 'blank': 9, 'empty': 9,
        'null': 9, 'true': 9, 'object': 9},
    'Comments': {'missing': 450, 'past': 1004, 'zero': 9, 'negative': 9, 'fraction': 1004, 'text': 1004,
        'lacking': 13, 'blank': 13, 'empty': 9, 'null': 13, 'true': 9, 'object': 13},
    'CommentsThreaded': {'missing': 450, 'past': 1004, 'zero': 1004, 'negative': 1004, 'fraction': 1004,
        'text': 1004, 'lacking': 13, 'blank': 13, 'empty': 1004, 'null': 13, 'true': 1004, 'object': 13},
    'Connections': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'fraction': 9, 'text': 9, 'lacking': 9,
        'blank': 9, 'empty': 9, 'null': 9, 'true': 9, 'object': 9},
    'CustomProperties': {'missing': 450, 'past': 9, 'zero': 1004, 'negative': 1004, 'fraction': 9, 'text': 9,
        'lacking': 13, 'blank': 13, 'empty': 1004, 'null': 13, 'true': 1004, 'object': 13},
    'DataLabels': {'missing': 450, 'past': -2147024809, 'negative': -2147024809, 'lacking': 13, 'blank': 13,
        'null': 13, 'true': -2147024809, 'object': 13},
    'DocumentProperties': {'missing': 450, 'past': 5, 'zero': 5, 'negative': 5, 'fraction': 5, 'text': 5,
        'lacking': 5, 'blank': 5, 'empty': 5, 'null': 5, 'true': 5, 'object': 5},
    'Errors': {'missing': 450, 'zero': 1004, 'negative': 1004, 'lacking': 1004, 'blank': 1004, 'empty': 1004,
        'null': 1004, 'true': 1004, 'object': 1004},
    'Filters': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'fraction': 9, 'lacking': 13, 'blank': 13,
        'empty': 9, 'null': 13, 'true': 9, 'object': 13},
    'FormatConditions': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'fraction': 9, 'text': 9, 'lacking': 9,
        'blank': 9, 'empty': 9, 'null': 9, 'true': 9, 'object': 9},
    'GroupShapes': {'missing': 450, 'past': 1004, 'zero': 1004, 'negative': 1004, 'text': 1004, 'lacking': 1004,
        'blank': 1004, 'empty': 1004, 'null': 1004, 'true': 1004, 'object': 1004},
    'HPageBreaks': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'fraction': 9, 'text': 9, 'lacking': 13,
        'blank': 13, 'empty': 9, 'null': 13, 'true': 9, 'object': 13},
    'Hyperlinks': {'missing': 450},
    'IconCriteria': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'lacking': 9, 'blank': 9, 'empty': 9,
        'null': 9, 'true': 9, 'object': 9},
    'IconSet': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'lacking': 9, 'blank': 9, 'empty': 9, 'null': 9,
        'true': 9, 'object': 9},
    'IconSets': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'lacking': 9, 'blank': 9, 'empty': 9,
        'null': 9, 'true': 9, 'object': 9},
    'LegendEntries': {'missing': 450, 'past': -2147467259, 'zero': -2147467259, 'negative': -2147467259,
        'lacking': -2147467259, 'blank': -2147467259, 'empty': -2147467259, 'null': -2147467259, 'true': -2147467259, 'object': -2147467259},
    'ListColumns': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'text': 9, 'lacking': 9, 'blank': 9,
        'empty': 9, 'null': 9, 'true': 9, 'object': 9},
    'ListObjects': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'fraction': 9, 'text': 9, 'lacking': 9,
        'blank': 9, 'empty': 9, 'null': 9, 'true': 9, 'object': 9},
    'ListRows': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'fraction': 9, 'lacking': 9, 'blank': 9,
        'empty': 9, 'null': 9, 'true': 9, 'object': 9},
    'Names': {'missing': 449, 'past': 9, 'zero': 9, 'negative': 9, 'fraction': 9, 'text': 1004, 'lacking': 1004,
        'blank': 1004, 'empty': 9, 'null': 1004, 'true': 9, 'object': 1004},
    'Panes': {'missing': 450, 'zero': 1004, 'negative': 1004, 'lacking': 13, 'blank': 13, 'empty': 1004, 'null': 13,
        'true': 1004, 'object': 13},
    'PivotCaches': {'missing': 450, 'past': 9, 'zero': 1004, 'negative': 1004, 'fraction': 9, 'lacking': 13,
        'blank': 13, 'empty': 1004, 'null': 13, 'true': 1004, 'object': 13},
    'PivotFields': {'missing': 449, 'past': 1004, 'negative': 1004, 'text': 1004, 'lacking': 1004, 'blank': 1004,
        'object': 1004},
    'PivotFilters': {'missing': 450, 'past': 9, 'zero': 5, 'negative': 5, 'fraction': 9, 'text': 9, 'lacking': 5,
        'blank': 5, 'empty': 5, 'null': 5, 'true': 5, 'object': 5},
    'PivotItemList': {'missing': 450, 'past': 1004, 'zero': 1004, 'negative': 1004, 'fraction': 1004, 'text': 438,
        'lacking': 438, 'blank': 438, 'empty': 1004, 'null': 13, 'true': 1004, 'object': 13},
    'PivotItems': {'missing': 449, 'past': 1004, 'zero': 1004, 'negative': 1004, 'text': 1004, 'lacking': 1004,
        'blank': 1004, 'empty': 1004, 'null': 1004, 'object': 1004},
    'PivotTables': {'missing': 449, 'past': 1004, 'zero': 1004, 'negative': 1004, 'fraction': 1004, 'text': 1004,
        'lacking': 1004, 'blank': 1004, 'empty': 1004, 'null': 1004, 'true': 1004, 'object': 1004},
    'Queries': {'missing': 450, 'past': -2147352565, 'zero': -2147352565, 'negative': -2147352565,
        'fraction': -2147352565, 'text': -2147319765, 'lacking': -2147319765, 'blank': -2147319765, 'empty': -2147352565, 'null': 13, 'true': -2147352565, 'object': 13},
    'Scenarios': {'missing': 449, 'past': 1004, 'zero': 1004, 'negative': 1004, 'fraction': 1004, 'text': 1004,
        'lacking': 1004, 'blank': 1004, 'empty': 1004, 'true': 1004, 'object': 1004},
    'ShapeRange': {'missing': 450, 'past': -2147352565, 'zero': -2147352565, 'negative': -2147352565,
        'fraction': -2147352565, 'text': -2147352565, 'lacking': -2147352565, 'blank': -2147352565, 'empty': -2147352565, 'null': -2147352565, 'true': -2147352565, 'object': -2147352565},
    'Shapes': {'missing': 450, 'past': -2147024809, 'zero': -2147024809, 'negative': -2147024809,
        'fraction': -2147024809, 'text': -2147024809, 'lacking': -2147024809, 'empty': -2147024809, 'true': -2147024809},
    'SheetViews': {'missing': 450, 'past': 1004, 'zero': 1004, 'negative': 1004, 'text': 9, 'lacking': 9, 'blank': 9,
        'empty': 1004, 'null': 13, 'true': 1004, 'object': 13},
    'Sheets': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'fraction': 9, 'text': 9, 'lacking': 9,
        'blank': 9, 'empty': 9, 'null': 13, 'true': 9, 'object': 13},
    'Slicers': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'fraction': 5, 'text': 5, 'lacking': 5,
        'blank': 5, 'empty': 5, 'null': 5, 'true': 5, 'object': 5},
    'SortFields': {'missing': 450, 'past': 9, 'zero': 1004, 'negative': 1004, 'fraction': 9, 'text': 9,
        'lacking': 13, 'blank': 13, 'empty': 1004, 'null': 13, 'true': 1004, 'object': 13},
    'SparklineGroup': {'missing': 450, 'past': 1004, 'zero': 1004, 'negative': 1004, 'fraction': 1004,
        'lacking': 1004, 'blank': 1004, 'empty': 1004, 'null': 1004, 'true': 1004, 'object': 1004},
    'SparklineGroups': {'missing': 450, 'past': 1004, 'zero': 1004, 'negative': 1004, 'fraction': 1004, 'text': 1004,
        'lacking': 13, 'blank': 13, 'empty': 1004, 'null': 13, 'true': 1004, 'object': 13},
    'Styles': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'text': 9, 'lacking': 9, 'blank': 9, 'empty': 9,
        'null': 13, 'true': 9, 'object': 13},
    'TableStyles': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'lacking': 9, 'blank': 9, 'empty': 9,
        'null': 9, 'true': 9, 'object': 9},
    'Trendlines': {'missing': -2147467259, 'past': -2147467259, 'zero': -2147467259, 'negative': -2147467259,
        'fraction': -2147467259, 'lacking': -2147467259, 'blank': -2147467259, 'empty': -2147467259, 'null': -2147467259, 'true': -2147467259, 'object': -2147467259},
    'VPageBreaks': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'fraction': 9, 'text': 9, 'lacking': 13,
        'blank': 13, 'empty': 9, 'null': 13, 'true': 9, 'object': 13},
    'Windows': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'fraction': 9, 'text': 9, 'lacking': 9,
        'blank': 9, 'empty': 9, 'null': 13, 'true': 9, 'object': 13},
    'Workbooks': {'missing': 450, 'past': 9, 'zero': 9, 'negative': 9, 'text': 9, 'lacking': 9, 'blank': 9,
        'empty': 9, 'null': 13, 'true': 9, 'object': 13},
}


def index_kind(index: object, count: int) -> str:
    """Which row of the table an index that failed falls under."""
    if index is MISSING:
        return 'missing'
    if index is NULL:
        return 'null'
    if index is EMPTY:
        return 'empty'
    if isinstance(index, VBAObject):
        return 'object'
    if isinstance(index, bool):
        return 'true' if index else 'zero'
    if isinstance(index, str):
        return 'blank' if not index else 'text' if index.strip().lstrip('-').isdigit() else 'lacking'
    if isinstance(index, (int, float)):
        if index != int(index):
            return 'fraction'
        return 'zero' if index == 0 else 'negative' if index < 0 else 'past' if index > count else ''
    return ''


def item_error(owner: str, index: object, count: int) -> int | None:
    """Excel's error for an index the collection could not use, or None to leave the class's own error standing."""
    return ITEM_ERRORS.get(owner, {}).get(index_kind(index, count))


#: Collections that take any number as a position: its whole part with no sign, and 0, Empty and True as the first.
LOOSE_INDEX = {'PivotItems': False, 'PivotTables': False, 'Scenarios': False, 'PivotFields': True}


def loose_index(owner: str, index: object) -> object:
    """The index a loosely indexed collection reads this as, or the index itself on any other."""
    if owner not in LOOSE_INDEX:
        return index
    if isinstance(index, bool):
        return 1 if index else index
    if isinstance(index, float):
        return int(index)
    if LOOSE_INDEX[owner] and (index is EMPTY or index is NULL or index == 0):
        return 1
    return index
