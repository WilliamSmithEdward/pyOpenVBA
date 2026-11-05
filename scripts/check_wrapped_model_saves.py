"""Live Excel reading files the model saved with wrapped rows: the height it reads, and whether it stays as the model wrote it; and the bytes against Excel's own save of the same steps. Needs Excel at 96 DPI.

    python scripts/check_wrapped_model_saves.py

Writes tests/fixtures/wrapped_model_saves.json. It uses no hyperlink, mail or send member.
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
from pyopenvba.apps.excel import ExcelApplication  # noqa: E402

OUT = ROOT / 'tests/fixtures/wrapped_model_saves.json'
LONG = 'one two three four five six seven'
#: The same steps go through the model and through Excel.
CASES: dict[str, list[str]] = {
    'wrapped': [],
    'wider after': ['ws.Columns(1).ColumnWidth = 40'],
    'narrower after': ['ws.Range("A1").Value = "one two"', 'ws.Columns(1).ColumnWidth = 4'],
    'big font': ['ws.Range("A1").Font.Size = 16'],
    'small font': ['ws.Range("A1").Font.Size = 8'],
    'bold': ['ws.Range("A1").Font.Bold = True'],
    'italic': ['ws.Range("A1").Font.Italic = True'],
    'indented': ['ws.Range("A1").IndentLevel = 2'],
    'custom': ['ws.Rows(1).RowHeight = 40'],
    'line feeds': ['ws.Range("A1").Value = "ab" & vbLf & "cd" & vbLf & "ef"'],
    'capped': ['ws.Range("A1").Value = String(3000, "a")'],
    'two cells': ['ws.Range("B1").WrapText = True', f'ws.Range("B1").Value = "{LONG} {LONG}"'],
    'formula after autofit': ['ws.Range("A2").Value = "x"', f'ws.Range("A3").Formula = "=""{LONG}"""', 'ws.Range("A3").WrapText = True', 'ws.Rows(3).AutoFit'],
    'wrap off': ['ws.Range("A1").WrapText = False'],
}


def source(steps: list[str]) -> str:
    return '\n'.join(['Public Function Go(ByVal path As String) As String', '    Dim wb As Workbook, ws As Object', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
                      '    ws.Range("A1").WrapText = True', f'    ws.Range("A1").Value = "{LONG}"', *[f'    {one}' for one in steps], '    Go = ws.Rows(1).RowHeight & "/" & ws.Rows(3).RowHeight', '    wb.SaveAs path, xlOpenXMLWorkbook', '    wb.Close False', '    Application.DisplayAlerts = True', 'End Function']) + '\n'


READER = '\n'.join(['Public Function Go(ByVal path As String) As String', '    Dim wb As Workbook', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Open(path)', '    Go = ActiveSheet.StandardHeight & "|" & wb.Worksheets(1).Rows(1).RowHeight & "/" & wb.Worksheets(1).Rows(3).RowHeight', '    wb.SaveAs path & ".again.xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Application.DisplayAlerts = True', 'End Function']) + '\n'


def row_tags(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        return re.findall(r'<row r="[0-9]+"[^>]*>', package.read('xl/worksheets/sheet1.xml').decode('utf-8'))


def main() -> None:
    records = {}
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        base = Path(folder)
        excel.new_document()
        for name, steps in CASES.items():
            app = ExcelApplication(sandbox=False)
            app.add_workbook()
            app.add_module(source(steps), name='Probe')
            ours = base / f'model {name}.xlsx'
            model_read = str(app.run('Probe.Go', str(ours)))
            theirs = base / f'excel {name}.xlsx'
            made = excel.run_vba(source(steps), 'Go', (str(theirs),), timeout=60.0, line_numbers=False)
            assert made.ok and not made.dialogs, (name, made)
            read = excel.run_vba(READER, 'Go', (str(ours),), timeout=60.0, line_numbers=False)
            assert read.ok and not read.dialogs, (name, read)
            height, opened = str(read.value).split('|')
            assert height == '15', f'Excel is not at 96 DPI: StandardHeight reads {height}'
            records[name] = {'model read': model_read, 'excel read': str(made.value), 'excel opening the model': opened, 'model rows': row_tags(ours), 'excel rows': row_tags(theirs), 'excel saving the model': row_tags(Path(f'{ours}.again.xlsx'))}
            print(name, model_read, str(made.value), opened, 'bytes alike' if row_tags(ours) == row_tags(theirs) else (row_tags(ours), row_tags(theirs)))
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': records}, indent=0) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
