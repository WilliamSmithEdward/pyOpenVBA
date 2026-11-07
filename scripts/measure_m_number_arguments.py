"""What Power Query makes of a count, an index or a number past what its functions take, evaluated in live Excel's own Power Query.

Each case is a query of its own, `let x = EXPRESSION in Text.From(x)`, loaded to a sheet and refreshed in the foreground. The families are
counts and positions (Text.Repeat, Text.Start, List.FirstN and the like), number arithmetic past a Double, and dates and durations past their
range. No count is large enough to build anything vast: each is either infinity or just past a 32-bit integer, where Power Query checks. No
file, server or web source is named.

    python scripts/measure_m_number_arguments.py

writes tests/fixtures/m_number_arguments.json, keeping each case as it finishes and asking only the rest.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/m_number_arguments.json'
CASES = [
    # Counts and positions
    'Text.Repeat("a", 1.8e308)', 'Text.Repeat("a", 2147483648)', 'Text.Repeat("a", -1)', 'Text.Repeat("ab", 2.5)',
    'Text.PadStart("a", 1.8e308)', 'Text.PadStart("a", 2147483648)', 'Text.PadEnd("a", 2147483648)', 'Text.PadStart("a", -1)',
    'Text.Start("abc", 1.8e308)', 'Text.Start("abc", 2147483648)', 'Text.Start("abc", -1)', 'Text.End("abc", 1.8e308)',
    'Text.Middle("abc", 1.8e308)', 'Text.Middle("abc", 1, 1.8e308)', 'Text.Middle("abc", 2147483648)', 'Text.Middle("abc", -1)',
    'Text.Middle("abc", 1, 2147483648)', 'Text.Range("abc", 1, 1.8e308)',
    'List.FirstN({1, 2}, 1.8e308)', 'List.FirstN({1, 2}, 2147483648)', 'List.LastN({1, 2}, 1.8e308)', 'List.Skip({1, 2}, 1.8e308)',
    'List.Range({1, 2, 3}, 1, 1.8e308)', 'List.Range({1, 2, 3}, 2147483648)', 'List.Range({1, 2, 3}, 1, 2147483648)',
    'Character.FromNumber(1.8e308)', 'Character.FromNumber(1e12)', 'Character.FromNumber(-1)', 'Character.FromNumber(1114112)',
    'Character.FromNumber(1114111)', 'Character.FromNumber(65.7)',
    'Number.Round(1.5, 1.8e308)', 'Number.Round(1.5, 1e12)', 'Number.Round(1.5, 16)', 'Number.Round(1.25, 2147483648)',
    'Number.Round(1.5, -1)', 'Number.Round(1e300, 2)', 'Number.Round(1.8e308)',
    'Date.AddDays(#date(2020, 1, 1), 1.8e308)', 'Date.AddDays(#date(2020, 1, 1), 1e12)', 'Date.AddDays(#date(2020, 1, 1), 3000000)',
    # Arithmetic past a Double
    'Number.Power(10, 400)', 'Number.Power(10, 1e300)', 'Number.Power(-8, 1/3)', 'Number.Power(0, -1)', 'Number.Power(0, 0)',
    'List.StandardDeviation({1e300, -1e300})', 'Number.Mod(1.8e308, 3)', 'Number.IntegerDivide(1.8e308, 3)', 'Number.Mod(5, 0)',
    'Number.IntegerDivide(5, 0)', 'Int64.From(1.8e308)', 'Int64.From(1e19)',
    # Dates and durations past their range
    'Duration.Days(#duration(1.8e308, 0, 0, 0))', '#duration(1e12, 0, 0, 0)', '#date(1e12, 1, 1)', '#date(10000, 1, 1)',
    # A count that is no whole number, and the 64-bit counts
    'List.FirstN({1, 2, 3}, 1.5)', 'Text.Start("abc", 1.5)', 'Date.AddDays(#date(2020, 1, 1), 1.5)', 'List.Range({1, 2, 3}, 0.5)',
    'List.LastN({1, 2}, 2147483648)', 'List.LastN({1, 2, 3}, 1.5)', 'List.Skip({1, 2, 3}, 1.5)', 'Text.Repeat("a", 2)',
    'Int64.From(9.2e18)', 'Number.Round(2.5)', 'Number.Round(1.2345, 2.5)',
    # Negative and past-the-end counts and positions, and the other rounding functions
    'Text.PadEnd("a", -1)', 'Text.End("abc", -1)', 'Text.Middle("abc", 1, -1)', 'Text.Middle("abc", 9)', 'Text.Start("abc", 9)',
    'Text.Insert("abc", 1.8e308, "x")', 'Text.Insert("abc", -1, "x")', 'Text.Insert("abc", 9, "x")', 'Text.Insert("abc", 1.5, "x")',
    'List.FirstN({1, 2}, -1)', 'List.Skip({1, 2}, -1)', 'List.Range({1, 2, 3}, -1)', 'List.Range({1, 2, 3}, 1, -1)', 'List.Range({1, 2, 3}, 9)',
    'List.Count({1..2.5})', 'Table.RowCount(Table.Repeat(#table({"a"}, {{1}}), 1.5))',
    'Number.RoundDown(1.5, 1.8e308)', 'Number.RoundUp(1.5, 1e12)', 'Number.RoundDown(1.25, 2.5)', 'Number.RoundUp(1e300, 2)',
    'Number.Mod(5, 1.8e308)', 'Number.IntegerDivide(5, 1.8e308)', 'Character.FromNumber(55296)', 'Number.Power(-10, 401)', 'Number.Power(-8, 3)',
    # What a walk of every function found: table counts, dates and times from numbers, date arithmetic, and the rest
    'Table.RowCount(Table.FirstN(#table({"a"}, {{1}}), 1.8e308))', 'Table.RowCount(Table.FirstN(#table({"a"}, {{1}}), -1))',
    'Table.RowCount(Table.FirstN(#table({"a"}, {{1}}), 1.5))', 'Table.RowCount(Table.LastN(#table({"a"}, {{1}}), 1.8e308))',
    'Table.RowCount(Table.Range(#table({"a"}, {{1}}), 1.8e308))', 'Table.RowCount(Table.Range(#table({"a"}, {{1}}), 0, 1.8e308))',
    'Table.RowCount(Table.RemoveRows(#table({"a"}, {{1}}), 1.8e308))', 'Table.RowCount(Table.Skip(#table({"a"}, {{1}}), 1.8e308))',
    'Table.RowCount(Table.Skip(#table({"a"}, {{1}}), -1))', 'Table.AddIndexColumn(#table({"a"}, {{1}}), "i", 1.8e308, 1)[i]{0}',
    'Table.AddIndexColumn(#table({"a"}, {{1}}), "i", 1.5, 1)[i]{0}', 'Table.AddIndexColumn(#table({"a"}, {{1}, {2}}), "i", 0, 1.8e308)[i]{1}',
    'Date.From(1.8e308)', 'Date.From(2147483648)', 'Date.From(-1)', 'Date.From(2958466)', 'DateTime.From(1.8e308)', 'DateTime.From(2147483648)',
    'DateTime.ToText(1.8e308)', 'Duration.Days(1.8e308)', 'Duration.Days(5)', 'Duration.ToText(1.8e308)', 'Time.From(1.8e308)', 'Time.From(0.5)',
    'Time.ToText(1.8e308)', '#time(1.8e308, 0, 0)', '#time(2147483648, 0, 0)', '#time(25, 0, 0)', '#time(1.5, 0, 0)',
    'Date.AddMonths(#date(2020, 1, 1), 1.8e308)', 'Date.AddMonths(#date(2020, 1, 1), 2147483648)', 'Date.AddMonths(#date(2020, 1, 1), 1.5)',
    'Date.AddMonths(#date(2020, 1, 1), 100000)', 'Date.AddYears(#date(2020, 1, 1), 1.8e308)', 'Date.AddYears(#date(2020, 1, 1), 10000)',
    'Date.DayOfWeek(#date(2020, 1, 1), 1.8e308)', 'Date.DayOfWeek(#date(2020, 1, 1), 9)', 'Number.IsEven(1.8e308)', 'Number.IsOdd(1.8e308)',
    'Number.IsEven(1.5)', 'Number.ToText(1e300, "F2")', 'Number.ToText(1, 5)', 'Json.Document("abc")', 'Json.Document("[1")',
    'List.Count(List.Numbers(1, -1))', 'List.Count(List.Numbers(1, -1.8e308))', 'List.Count(List.Numbers(1, 2.5))',
    'Table.ColumnCount(#table(2147483648, {}))', 'Table.ColumnCount(#table(1.8e308, {}))', 'Table.ColumnCount(#table(2.5, {}))',
    'Table.ColumnCount(#table(3, {}))',
]
TIMEOUT = 90.0
SOURCE = '''Function Probe(ByVal formula As String) As String
    Dim ws As Worksheet, lo As ListObject, e As Long, d As String, shown As String, q As Object
    Application.DisplayAlerts = False
    Set ws = Worksheets.Add
    For Each q In ThisWorkbook.Queries
        q.Delete
    Next
    ThisWorkbook.Queries.Add Name:="Q", Formula:=formula
    Set lo = ws.ListObjects.Add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Q;Extended Properties=""""", Destination:=ws.Range("$A$1"))
    lo.QueryTable.CommandType = xlCmdSql
    lo.QueryTable.CommandText = Array("SELECT * FROM [Q]")
    On Error Resume Next
    lo.QueryTable.Refresh BackgroundQuery:=False
    e = Err.Number
    d = Err.Description
    Err.Clear
    On Error GoTo 0
    If e = 0 Then shown = CStr(ws.Range("A2").Value) Else shown = "error " & e & ": " & d
    Probe = shown
    lo.Delete
    ThisWorkbook.Queries("Q").Delete
    ws.Delete
    Application.DisplayAlerts = True
End Function
'''


def formula(expression: str) -> str:
    return f'let x = {expression} in Text.From(x)'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig
    from pyvbaharness.results import HarnessError

    answers: dict[str, str] = json.loads(OUT.read_text(encoding='utf-8'))['answers'] if OUT.exists() else {}
    for case in CASES:
        if case in answers:
            continue
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                excel.new_document()
                result = excel.run_vba(SOURCE, 'Probe', (formula(case),), timeout=TIMEOUT, line_numbers=False)
            if result.dialogs:
                answers[case] = 'dialog: ' + ' '.join(result.dialogs[0].message.split())[:200]
            elif result.ok:
                answers[case] = str(result.value)
            else:
                # The harness's message names a screenshot on this machine, which no committed file may.
                answers[case] = f'ran past {TIMEOUT:.0f} s'
        except HarnessError as failed:
            answers[case] = f'harness: {str(failed)[:120]}'
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers}, indent=1) + '\n', encoding='utf-8',
                       newline='\n')
        print(ascii(case[:60]), '->', ascii(answers[case][:160]), flush=True)


if __name__ == '__main__':
    main()
