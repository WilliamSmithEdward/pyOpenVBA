"""Held-out rounds over PivotTable's PARTIAL rows whose code reaches no reported gap and whose interface has none: its axes and
ranges, its field collections, its display, enable and table style switches, its text properties, and the members a cube's
table answers (AllocateChanges, ChangeConnection, CommitChanges, ConvertToFormulas, DiscardChanges, RefreshDataSourceValues),
which a range's table refuses. Each is read, set and read back where it sets.

RefreshDate is read for its type only: Excel answers the moment the table was refreshed, which no prediction can know.

The book is the table rounds' (scripts/measure_table_pivot_rounds.py): a new book whose second sheet holds a pivot table of the
first sheet's keys and values, the same saved and reopened, or the pivot table Excel saved in tests/fixtures/dxf_table/made.xlsx.

    python scripts/measure_pivot_table_rounds.py predict <seed>
    python scripts/measure_pivot_table_rounds.py measure <seed>

tests/test_excel_pivot_table_rounds.py replays the rounds.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_table_pivot_rounds as base  # noqa: E402

OUT = base.FIXTURES / 'pivot_table_rounds'
REACHES = {'PivotTable': base.TABLES.reaches['PivotTable']}
FLAGS = ('AllowMultipleFilters', 'DisplayContextTooltips', 'DisplayEmptyColumn', 'DisplayEmptyRow', 'DisplayImmediateItems',
         'EnableDataValueEditing', 'EnableDrilldown', 'EnableFieldDialog', 'EnableFieldList', 'EnableWizard', 'FieldListSortAscending',
         'HasAutoFormat', 'InGridDropZones', 'PreserveFormatting', 'PrintTitles', 'RepeatItemsOnEachPrintedPage', 'ShowDrillIndicators',
         'ShowTableStyleColumnHeaders', 'ShowTableStyleColumnStripes', 'ShowTableStyleLastColumn', 'ShowTableStyleRowHeaders',
         'ShowTableStyleRowStripes', 'SortUsingCustomLists', 'SubtotalHiddenPageItems')
RANGES = ('ColumnRange', 'DataBodyRange', 'DataLabelRange', 'PageRange', 'PageRangeCells', 'RowRange', 'TableRange1', 'TableRange2')
FIELDS = ('ColumnFields', 'DataFields', 'HiddenFields', 'PageFields', 'PivotFields', 'RowFields', 'VisibleFields', 'CalculatedFields', 'PivotFormulas')
CUBE = ('AllocateChanges', 'CommitChanges', 'ConvertToFormulas False', 'DiscardChanges', 'RefreshDataSourceValues')


def _flag(flag: str) -> list[str]:
    return [f'x = Key(o.{flag}, book)', f'o.{flag} = Not o.{flag}: x = Key(o.{flag}, book)', f'o.{flag} = 0: x = Key(o.{flag}, book)',
            f'o.{flag} = "yes": x = Key(o.{flag}, book)']


MEMBERS = {'PivotTable': [
    *[call for flag in FLAGS for call in _flag(flag)],
    *[f'x = Key(o.{name}, book)' for name in RANGES],
    *[f'ws2.Rows(1).Insert: x = Key(o.{name}, book)' for name in ('TableRange1', 'DataBodyRange', 'RowRange')],
    *[f'x = Key(o.{name}, book) & o.{name}.Count' for name in FIELDS],
    'x = Key(o.ColumnFields(1), book)', 'x = Key(o.RowFields(1), book)', 'x = Key(o.DataFields(1), book)', 'x = Key(o.PivotFields("v"), book)',
    'o.PivotFields("v").Orientation = xlColumnField: x = Key(o.ColumnFields.Count, book) & " " & o.ColumnRange.Address',
    'o.PivotFields("k").Orientation = xlPageField: x = Key(o.PageFields.Count, book) & " " & o.PageRange.Address',
    'x = Key(o.ActiveFilters, book) & o.ActiveFilters.Count', 'x = Key(o.Slicers, book) & o.Slicers.Count',
    'x = Key(o.PivotCache, book) & " " & o.PivotCache.SourceType', 'x = Key(o.PivotColumnAxis, book)', 'x = Key(o.PivotRowAxis, book)',
    'x = Key(o.Version, book)',
    'x = TypeName(o.RefreshDate)', 'x = Key(o.RefreshName, book)',
    'x = Key(o.AlternativeText, book)', 'o.AlternativeText = "alt": x = Key(o.AlternativeText, book)', 'o.AlternativeText = 12: x = Key(o.AlternativeText, book)',
    'x = Key(o.Summary, book)', 'o.Summary = "s" & vbLf & "t": x = Key(o.Summary, book)', 'o.Summary = "": x = Key(o.Summary, book)',
    'x = Key(o.Tag, book)', 'o.Tag = "t1": x = Key(o.Tag, book)', 'o.Tag = 3.5: x = Key(o.Tag, book)',
    'o.ListFormulas: x = Key(ActiveSheet.Name = ws2.Name, book) & " " & wbT.Worksheets.Count',
    *[f'o.{call}: x = "done"' for call in CUBE],
    # Late bound through CallByName: typed, Nothing for a WorkbookConnection does not compile.
    'CallByName o, "ChangeConnection", VbMethod, Nothing: x = "done"',
]}
LISTS = base.Plan(REACHES, MEMBERS, (), {}, base.setup, base.books, base.DIMS, OUT, base.SAMPLE)


if __name__ == '__main__':
    base.main(LISTS)
