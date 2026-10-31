"""Measure Workbook.Colors and Workbook.ResetColors: the workbook's 56-colour palette.

Each case runs on a new workbook and answers as text; some save the workbook, and the palette its styles.xml holds
is recorded.

    python scripts/measure_workbook_colors.py

writes tests/fixtures/workbook_colors.json, which tests/test_excel_workbook_colors.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/workbook_colors.json'
#: Each case: statements on a new workbook wb with sheet ws; the expression read back; the file saved, if any.
CASES = [
    (['r = TypeName(wb.Colors) & ":" & LBound(wb.Colors) & ":" & UBound(wb.Colors)'], 'r', ''),
    (['r = TypeName(wb.Colors(3)) & ":" & wb.Colors(3) & ":" & wb.Colors(1) & ":" & wb.Colors(56)'], 'r', ''),
    (['r = wb.Colors(0)'], 'r', ''),
    (['r = wb.Colors(57)'], 'r', ''),
    (['r = wb.Colors(-1)'], 'r', ''),
    (['r = wb.Colors("3")'], 'r', ''),
    (['r = wb.Colors(3.6)'], 'r', ''),
    (['wb.Colors(3) = RGB(1, 2, 3)', 'r = wb.Colors(3) & ":" & TypeName(wb.Colors(3))'], 'r', 'one.xlsx'),
    (['ws.Range("A1").Interior.ColorIndex = 3', 'ws.Range("A2").Interior.Color = RGB(255, 0, 0)',
      'ws.Range("A3").Font.ColorIndex = 3', 'wb.Colors(3) = RGB(1, 2, 3)',
      'r = ws.Range("A1").Interior.Color & ":" & ws.Range("A1").Interior.ColorIndex & ":" & '
      'ws.Range("A2").Interior.Color & ":" & ws.Range("A2").Interior.ColorIndex & ":" & '
      'ws.Range("A3").Font.Color & ":" & ws.Range("A3").Font.ColorIndex'], 'r', 'cells.xlsx'),
    (['wb.Colors(3) = RGB(1, 2, 3)', 'wb.ResetColors', 'r = wb.Colors(3)'], 'r', 'reset.xlsx'),
    (['wb.Colors(3) = -1'], 'r', ''),
    (['wb.Colors(3) = &H1000000'], 'r', ''),
    (['wb.Colors(3) = "x"'], 'r', ''),
    (['wb.Colors(3) = "255"', 'r = wb.Colors(3)'], 'r', ''),
    (['wb.Colors(3) = 255.6', 'r = wb.Colors(3)'], 'r', ''),
    (['wb.Colors(0) = 255'], 'r', ''),
    (['wb.Colors(57) = 255'], 'r', ''),
    (['Dim a(1 To 56) As Long', 'For i = 1 To 56: a(i) = i * 1000: Next',
      'wb.Colors = a', 'r = wb.Colors(1) & ":" & wb.Colors(56)'], 'r', 'all.xlsx'),
    (['wb.Colors = Array(RGB(9, 9, 9), RGB(8, 8, 8), RGB(7, 7, 7))',
      'r = wb.Colors(1) & ":" & wb.Colors(2) & ":" & wb.Colors(3) & ":" & wb.Colors(4)'], 'r', 'some.xlsx'),
    (['wb.Colors = 255'], 'r', ''),
    (['wb.Colors = Array()', 'r = wb.Colors(1)'], 'r', ''),
    (['wb.Colors = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, '
      '26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, '
      '53, 54, 55, 56, 57, 58)', 'r = wb.Colors(1) & ":" & wb.Colors(56)'], 'r', ''),
    (['wb.Colors = Array(RGB(9, 9, 9), "x")', 'r = wb.Colors(1)'], 'r', ''),
    (['wb.Colors(3) = RGB(255, 0, 0)', 'wb.Colors(10) = RGB(255, 0, 0)', 'ws.Range("A1").Interior.Color = RGB(255, 0, 0)',
      'r = ws.Range("A1").Interior.ColorIndex'], 'r', ''),
    (['wb.Colors(3) = RGB(1, 2, 3)', 'ws.Range("A1").Interior.Color = RGB(255, 0, 0)',
      'r = ws.Range("A1").Interior.ColorIndex & ":" & ws.Range("A1").Interior.Color'], 'r', ''),
    (['wb.Colors(3) = RGB(1, 2, 3)', 'Dim other As Workbook', 'Set other = Workbooks.Add',
      'r = other.Colors(3)', 'other.Close False'], 'r', ''),
    (['wb.Colors(3) = RGB(255, 0, 0)', 'r = wb.Colors(3)'], 'r', 'same.xlsx'),
    (['ws.Range("A1").Interior.ColorIndex = 3', 'wb.Colors(3) = RGB(1, 2, 3)', 'ws.Range("A1").Copy ws.Range("B1")',
      'r = ws.Range("B1").Interior.Color & ":" & ws.Range("B1").Interior.ColorIndex'], 'r', ''),
    (['ws.Tab.ColorIndex = 3', 'wb.Colors(3) = RGB(1, 2, 3)', 'r = ws.Tab.Color & ":" & ws.Tab.ColorIndex'], 'r', ''),
    (['wb.Colors(3) = -1', 'r = wb.Colors(3)'], 'r', ''),
    (['wb.Colors(3) = &H1000000', 'r = wb.Colors(3)'], 'r', ''),
    (['wb.Colors(3) = &H1FF0102', 'r = wb.Colors(3)'], 'r', ''),
    (['wb.Colors(3) = "x"', 'r = wb.Colors(3)'], 'r', ''),
    (['wb.Colors(3) = "&HFF"', 'r = wb.Colors(3)'], 'r', ''),
    (['wb.Colors(3) = "65280"', 'r = wb.Colors(3)'], 'r', ''),
    (['wb.Colors(3) = True', 'r = wb.Colors(3)'], 'r', ''),
    (['wb.Colors(3) = Empty', 'r = wb.Colors(3)'], 'r', ''),
    (['wb.Colors(3) = 255.5', 'r = wb.Colors(3)'], 'r', ''),
    (['wb.Colors(3) = 254.5', 'r = wb.Colors(3)'], 'r', ''),
    (['wb.Colors(3) = Null'], 'r', ''),
    (['r = wb.Colors(2.5) & ":" & wb.Colors(3.5) & ":" & wb.Colors("4.9")'], 'r', ''),
    (['ws.Range("A1").Interior.Color = RGB(250, 0, 0)', 'r = ws.Range("A1").Interior.ColorIndex'], 'r', ''),
    (['wb.Colors(10) = RGB(9, 9, 9)', 'ws.Range("A1").Interior.Color = RGB(250, 0, 0)',
      'r = ws.Range("A1").Interior.ColorIndex'], 'r', ''),
    (['wb.Colors(10) = RGB(9, 9, 9)', 'ws.Range("A1").Interior.Color = RGB(255, 0, 0)',
      'r = ws.Range("A1").Interior.ColorIndex & ":" & ws.Range("A1").Font.ColorIndex'], 'r', ''),
    (['ws.Range("A1").Interior.Color = RGB(255, 0, 0)', 'wb.Colors(3) = RGB(1, 2, 3)', 'wb.ResetColors',
      'r = ws.Range("A1").Interior.ColorIndex'], 'r', ''),
    (['ws.Range("A1").Font.Color = RGB(255, 0, 0)', 'wb.Colors(3) = RGB(1, 2, 3)',
      'r = ws.Range("A1").Font.ColorIndex & ":" & ws.Range("A1").Font.Color'], 'r', ''),
    (['ws.Range("A1").Interior.Color = RGB(1, 2, 3)', 'wb.Colors(3) = RGB(1, 2, 3)',
      'r = ws.Range("A1").Interior.ColorIndex'], 'r', ''),
    (['wb.Colors(3) = RGB(1, 2, 3)', 'ws.Range("A1").Interior.Color = RGB(1, 2, 3)',
      'r = ws.Range("A1").Interior.ColorIndex'], 'r', 'chosen.xlsx'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, r As Variant, i As Long', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    r = "ok"', '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, read, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        lines.append(f'        Tried = CStr({read})')
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long',
              '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _run(source: str, folder: str, first: int, last: int) -> list[str] | None:
    """The answers of cases first to last in one Excel, or None where Excel itself fails (the RPC server dies)."""
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=300.0)
    if result.outcome == 'runner-error':
        return None
    assert result.ok and not result.dialogs, result
    return str(result.value).split('@@@')


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder:
        answers: list[str] = []
        for first in range(0, len(CASES), 8):
            last = min(first + 7, len(CASES) - 1)
            found = _run(source, folder, first, last)
            if found is None:
                # One of them takes Excel down: each in an Excel of its own, the one that crashes it said so.
                found = [(_run(source, folder, index, index) or ['Excel crashed'])[0]
                         for index in range(first, last + 1)]
            answers += found
        parts = {}
        for _, _, saved in CASES:
            if saved and (Path(folder) / saved).exists():
                with zipfile.ZipFile(Path(folder) / saved) as package:
                    styles = package.read('xl/styles.xml').decode('utf-8')
                found = re.search(r'<colors>.*?</colors>', styles, re.DOTALL)
                parts[saved] = found.group(0) if found else ''
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, read, saved, answer]
                        for (statements, read, saved), answer in zip(CASES, answers, strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, _, saved, answer in record['cases']:
        print(statements[-1][:70], '->', answer, '|', parts.get(saved, '')[:100])


if __name__ == '__main__':
    main()
