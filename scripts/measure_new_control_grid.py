"""Where a new control's box lands over a grid of places and sizes, for the rule that rounds it to the display's pixels, and each kind's smallest size. Needs Excel at 96 DPI, where StandardHeight reads 15.

    python scripts/measure_new_control_grid.py

Writes tests/fixtures/new_control_grid.json. tests/test_excel_new_control_settings.py replays it. It uses no hyperlink, mail or send member.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/new_control_grid.json'
LEFTS = [round(10 + step * 0.1, 1) for step in range(11)] + [0, 0.1, 0.4, 0.8, 47.9, 48, 48.1, 95.9, 96, 100.3, 200.2, 300.7]
SIZES = [round(20 + step * 0.1, 1) for step in range(21)] + [1, 2, 3, 5.1, 10, 33.3, 90, 90.4, 90.5, 90.7, 91, 120, 200.3, 300.9]
TOPS = [round(10 + step * 0.1, 1) for step in range(11)] + [0, 0.1, 14.9, 15, 15.1, 29.9, 30, 50, 50.6, 60.4, 60.7, 61, 300.3]
SOURCE = r'''Public Function Go(ByVal text As String) As String
    Dim ws As Object, s As Object, r As String, line As Variant, n As Variant
    Set ws = ActiveSheet
    For Each line In Split(text, ";")
        n = Split(line, ",")
        Set s = ws.Shapes.AddFormControl(Val(n(4)), Val(n(0)), Val(n(1)), Val(n(2)), Val(n(3)))
        r = r & "|" & s.Left & "," & s.Top & "," & s.Width & "," & s.Height
        s.Delete
    Next
    Go = ActiveSheet.StandardHeight & r
End Function
'''


def main() -> None:
    boxes = [(left, 60, width, 30, 0) for left in LEFTS for width in SIZES] + [(100, top, 90, height, 0) for top in TOPS for height in SIZES]
    # Each kind's smallest size, and a list box's height, which goes by whole lines.
    boxes += [(100, 60, size, 60, kind) for kind in (0, 1, 2, 4, 5, 6, 7, 8, 9) for size in range(1, 41)]
    boxes += [(100, 60, 90, size, kind) for kind in (0, 1, 2, 4, 5, 6, 7, 8, 9) for size in range(1, 41)]
    boxes += [(100, top, 90, size, 6) for top in (50, 60, 61) for size in range(41, 131)]
    answers: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for start in range(0, len(boxes), 150):
            chunk = boxes[start:start + 150]
            result = excel.run_vba(SOURCE, 'Go', (';'.join(','.join(str(one) for one in box) for box in chunk),), timeout=120.0, line_numbers=False)
            assert result.ok and not result.dialogs, result
            height, *found = str(result.value).split('|')
            assert height == '15', f'Excel is not at 96 DPI: StandardHeight reads {height}'
            answers += found
    assert len(answers) == len(boxes)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'boxes': boxes, 'excel': answers}, indent=0) + '\n', encoding='utf-8')
    print(len(boxes), 'boxes')


if __name__ == '__main__':
    main()
