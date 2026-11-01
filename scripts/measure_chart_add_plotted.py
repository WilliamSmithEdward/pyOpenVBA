"""What ChartObjects.Add plots, and drops, from a filled block, and how SetSourceData plots a block when no PlotBy is
given.

Each case makes a workbook, fills the case's block from A1, selects A1 or the whole block as the case says, adds a
chart through ChartObjects.Add, then gives it the block with SetSourceData and no PlotBy, and reads the series count
and the first series' Formula, and saves it. Add's own series are dropped, but Excel numbers SetSourceData's unique
ids past them, so the first unique id in the saved chart part says how many Add plotted.

    python scripts/measure_chart_add_plotted.py

writes tests/fixtures/chart_add_plotted.json, which tests/test_excel_chart_add_plotted.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_add_plotted.json'
Cell = str | float | None
CASES: list[tuple[str, list[list[Cell]], str]] = [
    ('three label columns, one series',
     [['j', 'k', 'l', 'v'], ['p', 'x', 'a', 1], [None, None, 'b', 2], ['q', 'y', 'c', 3]], 'A1'),
    ('text, numbers, text', [['k', 'l', 'm', 'v'], ['x', 5, 'a', 1], ['x', 6, 'b', 2], ['y', 7, 'c', 3]], 'A1'),
    ('three label columns, two series',
     [['j', 'k', 'l', 'v', 'u'], ['p', 'x', 'a', 1, 4], ['p', 'y', 'b', 2, 5], ['q', 'z', 'c', 3, 6]], 'A1'),
    ('label row and column, three by two',
     [['k', 'v', 'u', 't'], ['x', 1, 2, 3], ['y', 4, 5, 6]], 'A1'),
    ('blank corner pair', [[None, None, 'v'], ['x', 5, 1], ['y', 6, 2]], 'A1'),
    ('blank first row', [[None, None, None], ['x', 'a', 1], ['y', 'b', 2]], 'A1'),
    ('blank first row, selected', [[None, None, None], ['x', 'a', 1], ['y', 'b', 2]], 'block'),
    ('blank first row, numbers', [[None, None], [1, 2], [3, 4]], 'A1'),
    ('blank first column', [[None, 'v', 'u'], [None, 1, 4], [None, 2, 5]], 'A1'),
    ('blank first column, selected', [[None, 'v', 'u'], [None, 1, 4], [None, 2, 5]], 'block'),
    ('lone blank corner', [[None, None, 'v'], [None, None, 'w'], ['x', 'a', 1], ['y', 'b', 2]], 'A1'),
    ('lone blank corner, selected', [[None, None, 'v'], [None, None, 'w'], ['x', 'a', 1], ['y', 'b', 2]], 'block'),
    ('all numbers, wide', [[1, 2, 3, 4], [5, 6, 7, 8]], 'A1'),
    ('all numbers, tall', [[1, 2], [3, 4], [5, 6], [7, 8]], 'A1'),
    ('all numbers, square', [[1, 2], [3, 4]], 'A1'),
    ('labels, square data', [['k', 'v', 'u'], ['x', 1, 2], ['y', 3, 4]], 'A1'),
]


def _literal(cell: Cell) -> str:
    return f'"{cell}"' if isinstance(cell, str) else repr(cell)


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Chart, s As String, block As Range',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (_, block, selected) in enumerate(CASES):
        lines.append(f'    Case {index}')
        for row, cells in enumerate(block, start=1):
            lines += [f'        ws.Cells({row}, {column}).Value = {_literal(cell)}'
                      for column, cell in enumerate(cells, start=1) if cell is not None]
        lines.append(f'        Set block = ws.Range("A1", ws.Cells({len(block)}, {len(block[0])}))')
        lines.append('        block.Select' if selected == 'block' else '        ws.Range("A1").Select')
    lines += ['    End Select', '    Set ch = ws.ChartObjects.Add(200, 10, 300, 200).Chart',
              '    ch.SetSourceData block',
              '    s = ch.SeriesCollection.Count & "|" & ch.SeriesCollection(1).Formula',
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
        print(case, '->', answer, re.findall(r'uniqueId val="\{(\d+)\}', part))


if __name__ == '__main__':
    main()
