"""Clearing whole rows on a sheet whose columns carry a format (issue #256): what the cleared rows show, and what a
save keeps for them, rows and cells, when every column is formatted (Cells.Font.Name) and when a few are.

Predictions written before measuring: with every column formatted, Excel gives each cleared row a row format of the
default (row s="0" customFormat="1") instead of a cell at every column, and the cleared rows show the default font;
with a few columns formatted, the measured rule of scripts/measure_row_formats.py holds, a cell at each formatted
column of each cleared row.

    python scripts/measure_clear_rows_formats.py

writes tests/fixtures/clear_rows_formats.json, which tests/test_excel_clear_rows_formats.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/clear_rows_formats.json'
READ = ('ws.Range("A1").Font.Name & "|" & ws.Range("B9").Font.Name & "|" & ws.Range("XFD12").Font.Name & "|" & '
        'ws.Range("C30").Font.Name & "|" & ws.UsedRange.Address & "|" & ws.Cells.SpecialCells(xlCellTypeLastCell).Address')
#: Each case: statements on the new sheet ws, and the file it saves.
CASES = [
    (['ws.Range("A10").Value = 1', 'ws.Cells.Font.Name = "Courier New"', 'ws.Range(ws.Rows(8), ws.Rows(12)).Clear'],
     'sheet_few.xlsx'),
    (['ws.Range("A10").Value = 1', 'ws.Cells.Font.Name = "Courier New"',
      'ws.Range(ws.Rows(8), ws.Rows(ws.Rows.Count)).Clear'], 'sheet_all.xlsx'),
    (['ws.Range("A10").Value = 1', 'ws.Cells.Font.Name = "Courier New"',
      'ws.Range(ws.Rows(8), ws.Rows(12)).ClearFormats'], 'sheet_formats.xlsx'),
    (['ws.Range("A10").Value = 1', 'ws.Columns("B:C").Font.Name = "Courier New"',
      'ws.Range(ws.Rows(8), ws.Rows(12)).Clear'], 'columns_few.xlsx'),
    (['ws.Range("A10").Value = 1', 'ws.Columns("B:C").Font.Name = "Courier New"',
      'ws.Range(ws.Rows(8), ws.Rows(ws.Rows.Count)).Clear'], 'columns_all.xlsx'),
    (['ws.Range("A10").Value = 1', 'ws.Columns("A:XFC").Font.Name = "Courier New"',
      'ws.Range(ws.Rows(8), ws.Rows(12)).Clear'], 'nearly_all.xlsx'),
    # Which formatted columns make Excel give the cleared rows a row format rather than cells.
    (['ws.Columns("A:A").Font.Name = "Courier New"', 'ws.Range(ws.Rows(8), ws.Rows(9)).Clear'], 'a.xlsx'),
    (['ws.Columns("A:B").Font.Name = "Courier New"', 'ws.Range(ws.Rows(8), ws.Rows(9)).Clear'], 'ab.xlsx'),
    (['ws.Columns("B:XFD").Font.Name = "Courier New"', 'ws.Range(ws.Rows(8), ws.Rows(9)).Clear'], 'b_xfd.xlsx'),
    (['ws.Columns("A:IU").Font.Name = "Courier New"', 'ws.Range(ws.Rows(8), ws.Rows(9)).Clear'], 'a_iu.xlsx'),
    (['ws.Columns("A:IV").Font.Name = "Courier New"', 'ws.Range(ws.Rows(8), ws.Rows(9)).Clear'], 'a_iv.xlsx'),
    (['ws.Columns("B:IW").Font.Name = "Courier New"', 'ws.Range(ws.Rows(8), ws.Rows(9)).Clear'], 'b_iw.xlsx'),
    (['ws.Columns("C:D").Font.Name = "Courier New"', 'ws.Columns("F:G").Font.Name = "Courier New"',
      'ws.Range(ws.Rows(8), ws.Rows(9)).Clear'], 'two_blocks.xlsx'),
    (['ws.Columns("A:Z").Font.Name = "Courier New"', 'ws.Range(ws.Rows(8), ws.Rows(9)).Clear'], 'a_z.xlsx'),
    (['ws.Columns("B:Z").Font.Name = "Courier New"', 'ws.Range(ws.Rows(8), ws.Rows(9)).Clear'], 'b_z.xlsx'),
    (['ws.Columns("A:C").Font.Name = "Courier New"', 'ws.Range(ws.Rows(8), ws.Rows(9)).Clear'], 'a_c.xlsx'),
    (['ws.Columns("B:D").Font.Name = "Courier New"', 'ws.Range(ws.Rows(8), ws.Rows(8)).Clear'], 'b_d_one.xlsx'),
    (['ws.Columns("B:C").Font.Name = "Courier New"', 'ws.Range(ws.Rows(8), ws.Rows(57)).Clear'], 'b_c_fifty.xlsx'),
    (['ws.Columns("B").Font.Name = "Courier New"', 'ws.Columns("D").Font.Name = "Courier New"',
      'ws.Columns("F").Font.Name = "Courier New"', 'ws.Range(ws.Rows(8), ws.Rows(9)).Clear'], 'three_apart.xlsx'),
    (['ws.Columns("B:E").Font.Name = "Courier New"', 'ws.Columns("C").Font.Name = "Arial"',
      'ws.Range(ws.Rows(8), ws.Rows(9)).Clear'], 'two_styles.xlsx'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        lines += [f'        Tried = {READ}', f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook']
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = Tried & "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim s As String, i As Long',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _sheet_summary(xml: str) -> dict[str, object]:
    """What the saved sheet keeps that matters here: its cols, the attributes of its rows, and its cells."""
    rows = re.findall(r'<row\b[^>]*?(?:/>|>)', xml)
    return {'cols': re.findall(r'<col\b[^>]*/>', xml), 'rows': rows[:12], 'row_count': len(rows),
            'cell_count': len(re.findall(r'<c\b', xml)), 'cells': re.findall(r'<c\b[^>]*?(?:/>|>)', xml)[:12],
            'dimension': re.findall(r'<dimension\b[^>]*/>', xml)}


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        sheets = {}
        for _, saved in CASES:
            with zipfile.ZipFile(Path(folder) / saved) as package:
                sheets[saved] = _sheet_summary(package.read('xl/worksheets/sheet1.xml').decode('utf-8'))
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, saved, answer] for (statements, saved), answer in zip(CASES, answers, strict=True)],
              'sheets': sheets}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, saved, answer in record['cases']:
        print(statements[-1], '->', answer)
        print('   ', json.dumps(sheets[saved])[:700])


if __name__ == '__main__':
    main()
