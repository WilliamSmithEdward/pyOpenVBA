"""SetSourceData over blocks with labels in their first rows and columns, one or several of each: which rows and
columns label the rest, the series Excel makes, each series' Formula, Name and XValues, and the chart part Excel
saves.

Each case makes a workbook, adds a chart through ChartObjects.Add while the sheet is empty, so that nothing is plotted
then, fills the case's block from A1, gives it to the chart with SetSourceData, by columns or by rows as the case
says, then reads the chart's series count and every series' Formula, Name and XValues, and saves it. The chart part
is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_multilevel_categories.py

writes tests/fixtures/chart_multilevel_categories.json, which tests/test_excel_chart_multilevel_categories.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_multilevel_categories.json'
#: A cell left out is None; a date is a (year, month, day) tuple.
Cell = str | float | tuple[int, int, int] | None
TWO: list[list[Cell]] = [['k', 'l', 'v'], ['x', 'a', 1], ['x', 'b', 2], ['y', 'c', 3], ['y', 'd', 4]]
GROUPED: list[list[Cell]] = [['k', 'l', 'v'], ['x', 'a', 1], [None, 'b', 2], ['y', 'c', 3], [None, 'd', 4]]
CASES: list[tuple[str, list[list[Cell]], str]] = [
    ('two text columns', TWO, 'xlColumns'),
    ('two text columns, groups once', GROUPED, 'xlColumns'),
    ('two text columns, empty corner', [[None, None, 'v'], *TWO[1:]], 'xlColumns'),
    ('two text columns, no header', TWO[1:], 'xlColumns'),
    ('two text columns, two series', [[*row, 10 + index] if index else [*row, 'w'] for index, row in enumerate(TWO)],
     'xlColumns'),
    ('three text columns', [['j', *TWO[0]], ['p', *TWO[1]], ['p', *TWO[2]], ['q', *TWO[3]], ['q', *TWO[4]]],
     'xlColumns'),
    ('text then dates', [['k', 'l', 'v'], ['x', (2026, 1, 1), 1], ['x', (2026, 1, 2), 2], ['y', (2026, 1, 3), 3]],
     'xlColumns'),
    ('text then numbers', [['k', 'l', 'v'], ['x', 5, 1], ['x', 6, 2], ['y', 7, 3]], 'xlColumns'),
    ('second text column half empty', [['k', 'l', 'v'], ['x', 'a', 1], ['x', None, 2], ['y', None, 3]],
     'xlColumns'),
    ('two text rows', [list(column) for column in zip(*TWO)], 'xlRows'),
    ('two text columns by rows', TWO, 'xlRows'),
    ('two text columns, missing PlotBy', TWO, ''),
    # Which blanks keep a second text column a level: none at all, or a text cell last.
    ('inner blank in the middle', [['k', 'l', 'v'], ['x', 'a', 1], ['x', None, 2], ['y', 'c', 3]], 'xlColumns'),
    ('inner blank last', [['k', 'l', 'v'], ['x', 'a', 1], [None, 'b', 2], ['y', None, 3]], 'xlColumns'),
    ('inner text only last', [['k', 'l', 'v'], ['x', None, 1], ['x', None, 2], ['y', 'c', 3]], 'xlColumns'),
    ('inner empty', [['k', 'l', 'v'], ['x', None, 1], ['x', None, 2], ['y', None, 3]], 'xlColumns'),
    ('outer only first', [['k', 'l', 'v'], ['x', 'a', 1], [None, 'b', 2], [None, 'c', 3]], 'xlColumns'),
    ('outer blank first', [['k', 'l', 'v'], [None, 'a', 1], ['x', 'b', 2], ['y', 'c', 3]], 'xlColumns'),
    ('one text column with a blank', [['k', 'v'], ['a', 1], [None, 2], ['c', 3]], 'xlColumns'),
    ('middle level grouped', [['j', 'k', 'l', 'v'], ['p', 'x', 'a', 1], [None, None, 'b', 2], ['q', 'y', 'c', 3]],
     'xlColumns'),
    ('dates then text', [['k', 'l', 'v'], [(2026, 1, 1), 'a', 1], [(2026, 1, 2), 'b', 2], [(2026, 1, 3), 'c', 3]],
     'xlColumns'),
    ('numbers then text', [['k', 'l', 'v'], [5, 'a', 1], [6, 'b', 2], [7, 'c', 3]], 'xlColumns'),
    ('text, number header', [['k', 2, 'v'], ['x', 'a', 1], ['x', 'b', 2], ['y', 'c', 3]], 'xlColumns'),
    # Whether the levels run to the last text cell of the bottom row, or stop at the first column without one.
    ('text, numbers, text', [['k', 'l', 'm', 'v'], ['x', 5, 'a', 1], ['x', 6, 'b', 2], ['y', 7, 'c', 3]],
     'xlColumns'),
    ('numbers ending in text', [['k', 'l', 'v'], ['x', 5, 1], ['x', 6, 2], ['y', 'c', 3]], 'xlColumns'),
    ('first column ends blank', [['k', 'v'], ['a', 1], ['b', 2], [None, 3]], 'xlColumns'),
    ('first column ends blank, text after', [['k', 'l', 'v'], ['a', 'x', 1], ['b', 'y', 2], [None, 'z', 3]],
     'xlColumns'),
    ('two header rows', [['k', 'l', 'v'], ['k', 'l', 'w'], ['x', 'a', 1], ['y', 'b', 2]], 'xlColumns'),
    ('two header rows, two series', [['k', 'l', 'v', 'u'], [None, None, 'w', 't'], ['x', 'a', 1, 4],
                                     ['y', 'b', 2, 5]], 'xlColumns'),
    ('header then a text row', [['k', 'v', 'u'], ['m', 'w', 3], ['x', 1, 4], ['y', 2, 5]], 'xlColumns'),
    ('all numbers', [[1, 2, 3], [4, 5, 6], [7, 8, 9]], 'xlColumns'),
    # Whether the header rows, likewise, run to the last text cell of the rightmost column.
    ('last header blank', [['k', 'v', None], ['a', 1, 4], ['b', 2, 5], ['c', 3, 6]], 'xlColumns'),
    ('last header a number', [['k', 'v', 9], ['a', 1, 4], ['b', 2, 5], ['c', 3, 6]], 'xlColumns'),
    ('first header blank', [['k', None, 'u'], ['a', 1, 4], ['b', 2, 5], ['c', 3, 6]], 'xlColumns'),
    ('one column with a header', [['v'], [1], [2]], 'xlColumns'),
    ('one column of text', [['v'], ['a'], [2]], 'xlColumns'),
    ('all text', [['k', 'v'], ['a', 'b'], ['c', 'd']], 'xlColumns'),
    ('blank first row', [[None, None, None], ['x', 'a', 1], ['y', 'b', 2]], 'xlColumns'),
    # Whether a blank top left corner sets how many rows and columns label the rest.
    ('blank corner cell, two text columns', [[None, 'l', 'v'], ['x', 'a', 1], ['y', 'b', 2]], 'xlColumns'),
    ('blank corner pair, numbers', [[None, None, 'v'], ['x', 5, 1], ['y', 6, 2]], 'xlColumns'),
    ('blank corner column, two header rows', [[None, 'u', 'v'], [None, 'w', 3], ['x', 1, 4], ['y', 2, 5]],
     'xlColumns'),
    ('blank corner block', [[None, None, 'v'], [None, None, 'w'], ['x', 'a', 1], ['y', 'b', 2]], 'xlColumns'),
    ('blank corner, short label column', [[None, 'v'], [None, 1], ['x', 2]], 'xlColumns'),
    ('blank first row, numbers', [[None, None], [1, 2], [3, 4]], 'xlColumns'),
    ('blank first column', [[None, 'v', 'u'], [None, 1, 4], [None, 2, 5]], 'xlColumns'),
]


def _literal(cell: Cell) -> str:
    if isinstance(cell, tuple):
        return f'DateSerial({cell[0]}, {cell[1]}, {cell[2]})'
    if isinstance(cell, str):
        return f'"{cell}"'
    return repr(cell)


def _probe() -> str:
    lines = ['Private Function Joined(ByVal values As Variant) As String',
             '    Dim i As Long, s As String', '    On Error GoTo Failed',
             '    For i = LBound(values) To UBound(values): s = s & "/" & TypeName(values(i)) & ":" & values(i): Next',
             '    Joined = "[" & LBound(values) & "]" & s', '    Exit Function', 'Failed:',
             '    Joined = "E" & Err.Number & s', 'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Chart, i As Long, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(200, 10, 300, 200).Chart',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (_, block, plot_by) in enumerate(CASES):
        lines.append(f'    Case {index}')
        for row, cells in enumerate(block, start=1):
            lines += [f'        ws.Cells({row}, {column}).Value = {_literal(cell)}'
                      for column, cell in enumerate(cells, start=1) if cell is not None]
        corner = f'ws.Cells({len(block)}, {len(block[0])}).Address'
        lines += [f'        ch.SetSourceData ws.Range("A1:" & {corner}){", " + plot_by if plot_by else ""}']
    lines += ['    End Select', '    s = ch.SeriesCollection.Count',
              '    For i = 1 To ch.SeriesCollection.Count',
              '        s = s & "|" & ch.SeriesCollection(i).Formula & " " & ch.SeriesCollection(i).Name & " "'
              ' & Joined(ch.SeriesCollection(i).XValues)',
              '    Next',
              '    Tried = "ok " & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & " " & s', '    Resume Saved',
              'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i, folder)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [chart_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, part] for (case, _, _), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        categories = re.findall(r'<c:cat>.*?</c:cat>', part)
        print(case, '->', answer, '||', categories[0][:400] if categories else '')


if __name__ == '__main__':
    main()
