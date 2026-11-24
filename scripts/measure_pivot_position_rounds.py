"""Held-out rounds over PivotTable's rows that answer cells: TableRange1, TableRange2, DataBodyRange, RowRange, ColumnRange,
DataLabelRange, PageRange and PageRangeCells. The edits a case makes are on the table's own sheet as well as its source's:
rows and columns that go in or out above, before, beside and below the table, and lines through it, which Excel refuses with
1004 (scripts/measure_pivot_shifts.py). No cell goes in or out inside the table, which the model reports. Each edit comes
before the table is reached, or after with the table held.

The book is the table rounds' (scripts/measure_table_pivot_rounds.py): a new book whose second sheet holds a pivot table of the
first sheet's keys and values, the same saved and reopened, or the pivot table Excel saved in tests/fixtures/dxf_table/made.xlsx.

    python scripts/measure_pivot_position_rounds.py predict <seed>
    python scripts/measure_pivot_position_rounds.py measure <seed>

tests/test_excel_pivot_position_rounds.py replays the rounds.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_table_pivot_rounds as base  # noqa: E402

OUT = base.FIXTURES / 'pivot_position_rounds'
REACHES = {'PivotTable': base.TABLES.reaches['PivotTable']}
RANGES = ('TableRange1', 'TableRange2', 'DataBodyRange', 'RowRange', 'ColumnRange', 'DataLabelRange', 'PageRange', 'PageRangeCells')
MEMBERS = {'PivotTable': [
    *[f'x = Key(o.{name}, book)' for name in RANGES],
    *[f'x = Key(o.{name}, book) & " " & o.{name}.Count' for name in RANGES],
    f'o.PivotFields(pfield).Orientation = xlPageField: x = Key(o.PageRange, book) & " " & Key(o.TableRange2, book)',
    f'o.PivotFields(pfield).Orientation = xlPageField: x = Key(o.PageRangeCells, book) & " " & Key(o.TableRange1, book)',
    f'o.PivotFields(pfield).Orientation = xlColumnField: x = Key(o.ColumnRange, book) & " " & Key(o.RowRange, book)',
    f'o.PivotFields(pfield).Orientation = xlColumnField: x = Key(o.DataLabelRange, book) & " " & Key(o.DataBodyRange, book)',
    'x = Key(o.TableRange2.Cells(1, 1), book) & " " & Key(o.TableRange1.Cells(1, 1), book)',
]}
#: Above, before, beside and below the table on its own sheet, through it, and on its source's sheet.
EDITS = ('', 'ws2.Rows(1).Insert', 'ws2.Rows(1).Delete', 'ws2.Columns(1).Insert', 'ws2.Rows("1:2").Insert', 'ws2.Columns("Z").Insert',
         'ws2.Rows(40).Delete', 'ws2.Columns("Z").Delete', 'ws2.Rows(4).Insert', 'ws.Rows(1).Insert')
# Fewer statements than the table rounds ask, so a round samples 80 of them.
LISTS = base.Plan(REACHES, MEMBERS, (), {}, base.setup, base.books, base.DIMS, OUT, 80, EDITS)


if __name__ == '__main__':
    base.main(LISTS)
