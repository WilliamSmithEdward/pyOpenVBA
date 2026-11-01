"""Range.ApplyOutlineStyles: what it answers, which cells take which outline style, and the cell styles, fonts and
formats the saved workbook holds.

Each case makes a workbook with 1 in A1:T20, groups rows or columns as the case says, applies outline styles to a
range and reads every cell's style name in A1:E6 (or down column A and along row 1 for the seven-level cases), then
saves it.

    python scripts/measure_outline_styles.py

writes tests/fixtures/outline_styles.json, which tests/test_excel_outline_styles.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/outline_styles.json'
SEVEN_ROWS = ': '.join(f'ws.Rows("{n + 1}:{15 - n}").Group' for n in range(1, 8))
SEVEN_COLUMNS = ': '.join(f'ws.Range(ws.Cells(1, {n + 1}), ws.Cells(1, {15 - n})).EntireColumn.Group'
                          for n in range(1, 8))
#: (statements, which cells are read: "block" for A1:E6, "column" for A1:A20, "row" for A1:T1)
CASES = [
    ('ws.Rows("2:4").Group: v = TypeName(ws.Range("A1:C6").ApplyOutlineStyles)', 'block'),
    ('ws.Rows("2:4").Group: ws.Rows("3:3").Group: ws.Range("A1:C6").ApplyOutlineStyles', 'block'),
    ('ws.Columns("B:C").Group: ws.Range("A1:E6").ApplyOutlineStyles', 'block'),
    ('ws.Range("A1:C6").ApplyOutlineStyles', 'block'),
    ('ws.Rows("2:4").Group: ws.Outline.SummaryRow = xlAbove: ws.Range("A1:C6").ApplyOutlineStyles', 'block'),
    ('ws.Rows("2:4").Group: ws.Range("B5").ApplyOutlineStyles', 'block'),
    ('ws.Columns("B:C").Group: ws.Outline.SummaryColumn = xlLeft: ws.Range("A1:E6").ApplyOutlineStyles', 'block'),
    (SEVEN_ROWS + ': ws.Range("A1:A20").ApplyOutlineStyles', 'column'),
    (SEVEN_COLUMNS + ': ws.Range("A1:T1").ApplyOutlineStyles', 'row'),
]
READS = {'block': '    For r = 1 To 6: For c = 1 To 5: s = s & ws.Cells(r, c).Style.Name & ";": Next: s = s & "/": Next',
         'column': '    For r = 1 To 20: s = s & ws.Cells(r, 1).Style.Name & ";": Next',
         'row': '    For c = 1 To 20: s = s & ws.Cells(1, c).Style.Name & ";": Next'}


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, s As String, r As Long, c As Long, v As Variant',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:T20").Value = 1', '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, read) in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statements}', READS[read].replace('    For', '        For', 1)]
    lines += ['    End Select', '    Tried = "ok " & v & "|" & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def style_parts(path: Path) -> list[str]:
    """styles.xml's fonts, cellStyleXfs, cellXfs and cellStyles."""
    with zipfile.ZipFile(path) as package:
        styles = package.read('xl/styles.xml').decode('utf-8')
    parts: list[str] = []
    for tag in ('fonts', 'cellStyleXfs', 'cellXfs', 'cellStyles'):
        found = re.search(f'<{tag}\\b.*?</{tag}>', styles, re.S)
        parts.append(found.group() if found else '')
    return parts


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [style_parts(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, answer, found]
                        for (statements, _), answer, found in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case[:70], '->', answer)


if __name__ == '__main__':
    main()
