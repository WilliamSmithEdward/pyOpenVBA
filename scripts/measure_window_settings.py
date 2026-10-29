"""Window members everyday macros set: DisplayOutline, DisplayRuler, DisplayWhitespace, GridlineColor and
GridlineColorIndex, SmallScroll, WindowNumber and Type; what each reads, takes and refuses, and what a save keeps
of it in the sheet's sheetView.

Each case runs on a workbook of its own, its window the active one; the
saved workbooks' first sheetView elements are recorded.

    python scripts/measure_window_settings.py

writes tests/fixtures/window_settings.json, which tests/test_excel_window_settings.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/window_settings.json'

#: Each case: statements run on the workbook's active window w, then what is read; and the file it saves, if any.
CASES = [
    ([], 'w.DisplayOutline & " " & w.DisplayRuler & " " & w.DisplayWhitespace & " " & w.GridlineColor & " " & '
         'w.GridlineColorIndex & " " & w.WindowNumber & " " & w.Type & " " & w.AutoFilterDateGrouping', ''),
    (['w.DisplayOutline = False', 'w.DisplayRuler = False', 'w.DisplayWhitespace = False'],
     'w.DisplayOutline & " " & w.DisplayRuler & " " & w.DisplayWhitespace', 'flags.xlsx'),
    (['w.GridlineColor = RGB(255, 0, 0)'], 'w.GridlineColor & " " & w.GridlineColorIndex', 'red.xlsx'),
    (['w.GridlineColor = RGB(10, 200, 30)'], 'w.GridlineColor & " " & w.GridlineColorIndex', 'green.xlsx'),
    (['w.GridlineColorIndex = 5'], 'w.GridlineColor & " " & w.GridlineColorIndex', 'index.xlsx'),
    (['w.GridlineColorIndex = 5', 'w.GridlineColorIndex = xlColorIndexAutomatic'],
     'w.GridlineColor & " " & w.GridlineColorIndex', 'automatic.xlsx'),
    (['w.GridlineColorIndex = 0'], 'w.GridlineColorIndex', ''),
    (['w.GridlineColorIndex = 57'], 'w.GridlineColorIndex', ''),
    (['w.GridlineColorIndex = xlColorIndexNone'], 'w.GridlineColorIndex', ''),
    (['w.GridlineColor = -1'], 'w.GridlineColor', ''),
    (['w.GridlineColor = 16777216'], 'w.GridlineColor', ''),
    (['w.DisplayOutline = "x"'], 'w.DisplayOutline', ''),
    (['w.SmallScroll Down:=3'], 'w.ScrollRow & " " & w.ScrollColumn', ''),
    (['w.SmallScroll Down:=3', 'w.SmallScroll Up:=10'], 'w.ScrollRow & " " & w.ScrollColumn', ''),
    (['w.SmallScroll ToRight:=2'], 'w.ScrollRow & " " & w.ScrollColumn', ''),
    (['w.SmallScroll Down:=5, Up:=2, ToRight:=4, ToLeft:=1'], 'w.ScrollRow & " " & w.ScrollColumn', ''),
    (['w.SmallScroll Down:=-2'], 'w.ScrollRow', ''),
    (['w.SmallScroll Down:=2.6'], 'w.ScrollRow', ''),
    (['w.SmallScroll Down:="x"'], 'w.ScrollRow', ''),
    (['w.ScrollRow = 1048575', 'w.SmallScroll Down:=5'], 'w.ScrollRow', ''),
    (['w.SmallScroll'], 'w.ScrollRow & " " & w.ScrollColumn', ''),
    (['Dim v As Variant', 'v = w.SmallScroll(Down:=1)'], 'TypeName(v)', ''),
    (['ActiveSheet.Range("A1:A3").Rows.Group', 'w.DisplayOutline = False'], 'w.DisplayOutline', 'grouped.xlsx'),
    (['w.DisplayOutline = False', 'ActiveWorkbook.Worksheets.Add', 'Set w = ActiveWindow'], 'w.DisplayOutline', ''),
    (['w.DisplayOutline = False', 'ActiveWorkbook.Worksheets.Add', 'ActiveWorkbook.Worksheets(2).Activate'],
     'w.DisplayOutline', ''),
    (['w.GridlineColorIndex = 3', 'ActiveWorkbook.Worksheets.Add'], 'w.GridlineColorIndex', ''),
    (['w.AutoFilterDateGrouping = False'], 'w.AutoFilterDateGrouping', 'dates.xlsx'),
    (['w.AutoFilterDateGrouping = False', 'ActiveWorkbook.Worksheets.Add'], 'ActiveWindow.AutoFilterDateGrouping', ''),
    (['w.AutoFilterDateGrouping = "x"'], 'w.AutoFilterDateGrouping', ''),
]


def _literal(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, w As Window', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set w = ActiveWindow', '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, read, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        lines.append(f'        Tried = {read}')
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim s As String, i As Long',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "|" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('|')
        views = {}
        for _, _, saved in CASES:
            if saved:
                with zipfile.ZipFile(Path(folder) / saved) as package:
                    sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
                    workbook = package.read('xl/workbook.xml').decode('utf-8')
                found = re.search(r'<sheetView\b.*?(?:/>|</sheetView>)', sheet, re.DOTALL)
                outline = re.search(r'<sheetPr\b.*?(?:/>|</sheetPr>)', sheet, re.DOTALL)
                book_view = re.search(r'<workbookView\b[^>]*>', workbook)
                views[saved] = {'sheetView': found.group(0) if found else '',
                                'sheetPr': outline.group(0) if outline else '',
                                'autoFilterDateGrouping': 'autoFilterDateGrouping="0"' in (book_view.group(0) if
                                                                                          book_view else '')}
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, read, saved, answer]
                        for (statements, read, saved), answer in zip(CASES, answers, strict=True)],
              'views': views}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, read, saved, answer in record['cases']:
        print(statements, '->', answer)
    for name, parts in views.items():
        print(name, parts)


if __name__ == '__main__':
    main()
