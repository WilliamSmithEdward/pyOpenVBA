"""Which heights set through a new form control's old object take one EMU off the Shape's Top, into tests/fixtures/new_control_old_height_steps.json. Needs Excel at 96 DPI, checked at both ends of each call.

Written before Excel was asked, from tests/fixtures/new_control_old_heights.json, where 46, 45.6, 40 and 60 lost the EMU and 45.75, 39.75 and 30.1 did not, at every Top tried: a height that is a
whole number of fifths of a point loses it. The rival, NONE, says no rule over the height alone holds. Heights from 30 to 50 points in steps of 0.05.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from fractions import Fraction
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURE = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures' / 'new_control_old_height_steps.json'
HEIGHTS = [round(30 + step * 0.05, 2) for step in range(401)]
SOURCE = r'''Public Function Go(ByVal heights As String) As String
    Dim ws As Object, s As Object, o As Object, r As String, before As String, one As Variant
    Set ws = ActiveSheet
    before = ws.StandardHeight
    For Each one In Split(heights, ",")
        Set s = ws.Shapes.AddFormControl(xlButtonControl, 12, 12, 100, 40)
        Set o = ws.DrawingObjects(ws.DrawingObjects.Count)
        o.Height = CDbl(one)
        r = r & "|" & CStr(CDbl(s.Top) * 12700)
        s.Delete
    Next
    Go = before & "~" & ws.StandardHeight & r
End Function
'''


def loses(height: float) -> bool:
    return Fraction(str(height)) * 5 == int(Fraction(str(height)) * 5)


def main() -> None:
    found: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for start in range(0, len(HEIGHTS), 50):
            chunk = HEIGHTS[start:start + 50]
            result = excel.run_vba(SOURCE, 'Go', (','.join(str(one) for one in chunk),), timeout=120.0, line_numbers=False)
            assert result.ok and not result.dialogs, result
            heights, *tops = str(result.value).split('|')
            assert heights == '15~15', f'Excel is not at 96 DPI: {heights}'
            found += tops
    FIXTURE.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': '15', 'source': SOURCE, 'heights': HEIGHTS, 'tops': found}, indent=0), encoding='utf-8', newline='\n')
    right = sum((float(top) != 152400) == loses(height) for height, top in zip(HEIGHTS, found, strict=True))
    print(f'fifths of a point: {right} of {len(HEIGHTS)}')
    print('lost, not predicted:', [h for h, t in zip(HEIGHTS, found) if float(t) != 152400 and not loses(h)][:30])
    print('kept, predicted lost:', [h for h, t in zip(HEIGHTS, found) if float(t) == 152400 and loses(h)][:30])
    print('values seen:', sorted({t for t in found})[:6])


if __name__ == '__main__':
    main()
