"""Which palette slot Font.ColorIndex reads for a colour outside the palette: every grey, and random colours, on a cell.

    python scripts/measure_color_index_nearest.py

writes tests/fixtures/color_index_nearest.json as [colour, index] pairs in RRGGBB, which tests/test_excel_color_index_nearest.py replays.
"""
from __future__ import annotations

import json
import random
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/color_index_nearest.json'
SOURCE = '''Public Function Go(ByVal text As String) As String
    Dim c As Range, s As String, one As Variant, wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set c = wb.Worksheets(1).Range("D1")
    For Each one In Split(text, ",")
        c.Font.Color = CLng(one)
        s = s & c.Font.ColorIndex & ","
    Next
    wb.Close False
    Application.DisplayAlerts = True
    Go = s
End Function
'''


def main() -> None:
    random.seed(20261104)
    colours = [f'{g:02X}{g:02X}{g:02X}' for g in range(256)] + [f'{random.randrange(0x1000000):06X}' for _ in range(3000)]
    answers: list[list[object]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        for first in range(0, len(colours), 200):
            chunk = colours[first:first + 200]
            text = ','.join(str(int(one[4:6] + one[2:4] + one[:2], 16)) for one in chunk)
            done = excel.run_vba(SOURCE, 'Go', (text,), timeout=120.0, line_numbers=False)
            assert done.ok, done
            answers += [[colour, int(index)] for colour, index in zip(chunk, str(done.value).split(',')[:-1], strict=True)]
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'pairs': answers}) + '\n', encoding='utf-8')
    print(len(answers))


if __name__ == '__main__':
    main()
