"""Held-out rounds over PivotField's PARTIAL rows whose code reaches no reported gap and whose interface has none: AddPageItem,
AllItemsVisible, CalculatedItems, ChildField, ChildItems, DataType, DrillTo, GroupLevel, HiddenItems, IsCalculated,
LayoutPageBreak, ParentField, ParentItems, PivotItems, ShowingInAxis, SourceCaption, SourceName, TotalLevels and VisibleItems.
Each is read, set and read back where it sets, on a row field, the first field and the data field.

No row or column goes in on the table's sheet: the model reports an edit that would move a pivot table
(tests/test_excel_pivot_shift_refusal.py).

The book is the table rounds' (scripts/measure_table_pivot_rounds.py): a new book whose second sheet holds a pivot table of the
first sheet's keys and values, the same saved and reopened, or the pivot table Excel saved in tests/fixtures/dxf_table/made.xlsx.

    python scripts/measure_pivot_field_rounds.py predict <seed>
    python scripts/measure_pivot_field_rounds.py measure <seed>

tests/test_excel_pivot_field_rounds.py replays the rounds.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_table_pivot_rounds as base  # noqa: E402

OUT = base.FIXTURES / 'pivot_field_rounds'
REACHES = {'PivotField': base.TABLES.reaches['PivotField']}
COLLECTIONS = ('CalculatedItems', 'HiddenItems', 'PivotItems', 'VisibleItems')
READS = ('AllItemsVisible', 'DataType', 'GroupLevel', 'IsCalculated', 'LayoutPageBreak', 'ShowingInAxis', 'SourceCaption',
         'SourceName', 'TotalLevels', 'ChildField', 'ParentField', 'ChildItems', 'ParentItems')
MEMBERS = {'PivotField': [
    *[f'x = Key(o.{name}, book)' for name in READS],
    *[f'x = Key(o.{name}, book) & o.{name}.Count' for name in COLLECTIONS],
    'x = Key(o.PivotItems(1), book)', 'x = Key(o.VisibleItems(1), book)', 'x = Key(o.HiddenItems.Count, book)',
    'x = Key(o.ChildItems(1), book)', 'x = Key(o.ParentItems(1), book)',
    'o.LayoutPageBreak = True: x = Key(o.LayoutPageBreak, book)', 'o.LayoutPageBreak = "no": x = Key(o.LayoutPageBreak, book)',
    'o.AddPageItem "a": x = Key(o.Orientation, book)', 'o.Orientation = xlPageField: o.AddPageItem "a": x = Key(o.Orientation, book)',
    'o.Orientation = xlPageField: o.AddPageItem "nosuch": x = Key(o.Orientation, book)',
    'o.Orientation = xlPageField: o.AddPageItem "a", True: x = Key(o.Orientation, book)',
    'o.DrillTo "v": x = Key(o.Orientation, book)', 'o.DrillTo "nosuch": x = Key(o.Orientation, book)', 'o.DrillTo o.Name: x = Key(o.Orientation, book)',
    'x = Key(o.PivotItems.Count = o.VisibleItems.Count + o.HiddenItems.Count, book)',
]}
# Fewer statements than the table rounds ask, so a round samples 100 of them.
LISTS = base.Plan(REACHES, MEMBERS, (), {}, base.setup, base.books, base.DIMS, OUT, 100)


if __name__ == '__main__':
    base.main(LISTS)
