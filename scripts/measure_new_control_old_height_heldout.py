"""Held out: the rule fitted on tests/fixtures/new_control_old_height_steps.json, against fresh starting heights, tops and heights, into tests/fixtures/new_control_old_height_heldout.json.

Written before Excel was asked: a height set through a new form control's old object takes one EMU off the Shape's Top when it is taller than the control was and is a whole number of tenths of a
point; otherwise Top stays. Needs Excel at 96 DPI, checked at both ends of each call.
"""
from __future__ import annotations

import json
import random
from datetime import datetime, timezone
from fractions import Fraction
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURE = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures' / 'new_control_old_height_heldout.json'
SOURCE = r'''Public Function Go(ByVal kind As Long, ByVal top As Double, ByVal start As Double, ByVal heights As String) As String
    Dim ws As Object, s As Object, o As Object, r As String, before As String, one As Variant
    Set ws = ActiveSheet
    before = ws.StandardHeight
    For Each one In Split(heights, ",")
        Set s = ws.Shapes.AddFormControl(kind, 12, top, 100, start)
        Set o = ws.DrawingObjects(ws.DrawingObjects.Count)
        r = r & "|" & CStr(CDbl(s.Top) * 12700) & ":" & CStr(CDbl(s.Height))
        o.Height = CDbl(one)
        r = r & ":" & CStr(CDbl(s.Top) * 12700)
        s.Delete
    Next
    Go = before & "~" & ws.StandardHeight & r
End Function
'''


def loses(height: float, was: float) -> bool:
    exact = Fraction(str(height))
    return height > was and exact * 10 == int(exact * 10)


def main() -> None:
    chooser = random.Random(23)
    plans = []
    for kind in (0, 1, 2, 7):
        for top in (12.0, 31.0, 50.25):
            for start in (20.0, 39.75, 75.0):
                heights = [round(chooser.uniform(10, 110), chooser.choice((0, 1, 1, 2, 3))) for _ in range(14)]
                plans.append((kind, top, start, heights))
    cases = []
    right = total = 0
    wrong = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for kind, top, start, heights in plans:
            result = excel.run_vba(SOURCE, 'Go', (kind, top, start, ','.join(str(one) for one in heights)), timeout=120.0, line_numbers=False)
            assert result.ok and not result.dialogs, result
            checked, *rows = str(result.value).split('|')
            assert checked == '15~15', f'Excel is not at 96 DPI: {checked}'
            for height, row in zip(heights, rows, strict=True):
                made, was, after = row.split(':')
                lost = float(after) != float(made)
                total += 1
                if lost == loses(height, float(was)):
                    right += 1
                else:
                    wrong.append((kind, top, start, height, was, made, after))
                cases.append({'kind': kind, 'top': top, 'start': start, 'height': height, 'answer': row})
    FIXTURE.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': '15', 'source': SOURCE, 'cases': cases}, indent=0), encoding='utf-8', newline='\n')
    print(f'{right} of {total}')
    for one in wrong[:20]:
        print(one)


if __name__ == '__main__':
    main()
