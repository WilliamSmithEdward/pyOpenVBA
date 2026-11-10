"""Held-out rounds over FormatCondition's PARTIAL rows whose code reaches no reported gap: AppliesTo, DateOperator, Font,
Formula1, Formula2, Interior, NumberFormat, Operator, Priority, PTCondition, ScopeType, SetFirstPriority, SetLastPriority,
StopIfTrue, Text and TextOperator. Each is read, set and read back where it sets, on the rule the table rounds' book holds
first over its cells, or on a text or date rule the call adds before it asks.

The book is the table rounds' (scripts/measure_table_pivot_rounds.py): a new book with one to four cell-value or
expression rules over D1:D5, the same saved and reopened, or the rules Excel saved in tests/fixtures/object_roundtrip.xlsx.

    python scripts/measure_format_condition_rounds.py predict <seed>
    python scripts/measure_format_condition_rounds.py measure <seed>

tests/test_excel_format_condition_rounds.py replays the rounds.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_table_pivot_rounds as base  # noqa: E402

OUT = base.FIXTURES / 'format_condition_rounds'
REACHES = {'FormatCondition': [*base.TABLES.reaches['FormatCondition'], 'wsL.Range(fcells).FormatConditions(wsL.Range(fcells).FormatConditions.Count)', 'wsL.Cells.FormatConditions(1)']}
TEXT = 'Set o = wsL.Range(fcells).FormatConditions.Add(xlTextString, String:="ab", TextOperator:={op}): '
DATE = 'Set o = wsL.Range(fcells).FormatConditions.Add(xlTimePeriod, DateOperator:={op}): '
PRIORITIES = 'x = Key(o.Priority, book) & " " & wsL.Range(fcells).FormatConditions(1).Priority & " " & wsL.Range(fcells).FormatConditions.Count'
MEMBERS = {'FormatCondition': [
    'x = Key(o.AppliesTo, book)', 'ws.Rows(1).Insert: x = Key(o.AppliesTo, book)', 'wsL.Range("D2").Delete xlShiftUp: x = Key(o.AppliesTo, book)',
    'x = Key(o.Formula1, book)', 'x = Key(o.Formula2, book)', 'o.Modify xlCellValue, xlBetween, "=1", "=9": x = Key(o.Formula2, book)',
    'x = Key(o.Operator, book)', 'x = Key(o.Priority, book)', 'x = Key(o.PTCondition, book)', 'x = Key(o.ScopeType, book)',
    'x = Key(o.StopIfTrue, book)', 'o.StopIfTrue = True: x = Key(o.StopIfTrue, book)', 'o.StopIfTrue = "no": x = Key(o.StopIfTrue, book)',
    'x = Key(o.NumberFormat, book)', 'o.NumberFormat = "0.00": x = Key(o.NumberFormat, book)', 'o.NumberFormat = "@": x = Key(o.NumberFormat, book)',
    'o.NumberFormat = "nonsense[": x = Key(o.NumberFormat, book)', 'o.NumberFormat = 12: x = Key(o.NumberFormat, book)',
    'x = Key(o.Font.Bold, book) & " " & o.Font.Color', 'o.Font.Bold = True: x = Key(o.Font.Bold, book)', 'o.Font.Color = RGB(200, 0, 0): x = Key(o.Font.Color, book)',
    'o.Font.Italic = True: o.Font.Size = 14: x = Key(o.Font.Italic, book) & " " & o.Font.Size', 'x = Key(o.Font, book)',
    'x = Key(o.Interior.Color, book) & " " & o.Interior.ColorIndex', 'o.Interior.Color = RGB(0, 200, 0): x = Key(o.Interior.Color, book)',
    'o.Interior.ColorIndex = 6: x = Key(o.Interior.Color, book)', 'o.Interior.Pattern = xlGray50: x = Key(o.Interior.Pattern, book)', 'x = Key(o.Interior, book)',
    f'o.Priority = 1: {PRIORITIES}', f'o.Priority = 99: {PRIORITIES}', f'o.Priority = 0: {PRIORITIES}', f'o.Priority = "2": {PRIORITIES}',
    f'o.SetFirstPriority: {PRIORITIES}', f'o.SetLastPriority: {PRIORITIES}', f'o.SetLastPriority: o.SetFirstPriority: {PRIORITIES}',
    'x = Key(o.Text, book)', 'x = Key(o.TextOperator, book)', 'x = Key(o.DateOperator, book)', 'o.ScopeType = xlSelectionScope: x = Key(o.ScopeType, book)',
    *[TEXT.format(op=op) + 'x = Key(o.Text, book) & " " & o.TextOperator & " " & o.Formula1' for op in ('xlContains', 'xlDoesNotContain', 'xlBeginsWith', 'xlEndsWith')],
    TEXT.format(op='xlContains') + 'o.Text = "zz": x = Key(o.Text, book) & " " & o.Formula1',
    TEXT.format(op='xlContains') + 'o.TextOperator = xlEndsWith: x = Key(o.TextOperator, book) & " " & o.Formula1',
    TEXT.format(op='xlContains') + 'x = Key(o.Operator, book) & " " & o.DateOperator',
    *[DATE.format(op=op) + 'x = Key(o.DateOperator, book) & " " & o.Formula1' for op in ('xlYesterday', 'xlToday', 'xlLast7Days', 'xlThisMonth', 'xlNextWeek')],
    DATE.format(op='xlToday') + 'o.DateOperator = xlTomorrow: x = Key(o.DateOperator, book) & " " & o.Formula1',
    DATE.format(op='xlToday') + 'x = Key(o.Text, book) & " " & o.TextOperator',
]}
LISTS = base.Plan(REACHES, MEMBERS, (), {}, base.setup, base.books, base.DIMS, OUT, base.SAMPLE)


if __name__ == '__main__':
    base.main(LISTS)
