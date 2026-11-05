"""Wrapped text in a cell: how many lines Excel fits it into at each column width, pixel by pixel, for the rule behind a wrapped row's height. Needs Excel at 96 DPI, where
StandardHeight reads 15, and stops if the display changes.

    python scripts/measure_wrapped_rows.py

Writes tests/fixtures/wrapped_rows.json. It uses no hyperlink, mail or send member.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/wrapped_rows.json'
#: Single words first, to give each its room with nothing to break at; then what decides a break.
TEXTS = ['x', 'i', 'W', 'ab', 'abcd', 'abcdefgh', 'iiiiii', 'WWW', 'iW', 'Wi', 'ii', 'WW', 'aa', 'aaa', 'aaaa', 'aaaaaaaaaaaaaaaa', 'llllllll', 'mmmm', '0', '00', '0000', '00000000', 'Hello', 'world', 'Excel',
         'ab cd', 'ab  cd', 'one two', 'a b', 'a  b', 'a   b', ' ab', 'ab ', '  ab', 'ab  ', 'a-b', 'ab-cd', 'a/b', 'a.b', 'a,b', 'a b c d e f g h', 'one two three four five six seven', 'abcdefghijklmnopqrstuvwxyz', 'The quick brown fox jumps over the lazy dog']
FONTS = (('Calibri', 11), ('Aptos Narrow', 11), ('Arial', 10), ('Calibri', 16))
SOURCE = r'''Public Function Go(ByVal text As String, ByVal font As String, ByVal size As Double, ByVal most As Long) As String
    Dim ws As Object, r As String, wide As Long, used As Long
    Set ws = ActiveSheet
    ws.Cells.Clear
    With ws.Range("A1")
        .Value = "'" & text
        .WrapText = True
        .Font.Name = font
        .Font.Size = size
    End With
    For wide = 1 To most
        ws.Columns(1).ColumnWidth = wide / 7
        used = CLng(ws.Columns(1).Width * 4 / 3)
        ws.Rows(1).AutoFit
        r = r & "|" & used & "," & ws.Rows(1).RowHeight
    Next
    Go = ActiveSheet.StandardHeight & r
End Function
'''


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for font, size in FONTS:
            for text in TEXTS:
                most = min(60 + len(text) * 14, 420)
                result = excel.run_vba(SOURCE, 'Go', (text, font, size, most), timeout=120.0, line_numbers=False)
                assert result.ok and not result.dialogs, (text, font, result)
                height, *steps = str(result.value).split('|')
                assert height == '15', f'Excel is not at 96 DPI: StandardHeight reads {height}'
                curve: dict[int, float] = {}
                for step in steps:
                    wide, tall = step.split(',')
                    curve[int(wide)] = float(tall)
                records.append({'font': font, 'size': size, 'text': text, 'heights': sorted(curve.items())})
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': records}, indent=0) + '\n', encoding='utf-8')
    print(len(records), 'texts')


if __name__ == '__main__':
    main()
