"""What a shape's old object reads of a place or size set off the pixel, against what the Shape reads, into tests/fixtures/new_control_old_places.json. Needs Excel at 96 DPI, checked at both ends of each call.

Written before Excel was asked: a new form control's old object reads Left and Top at the nearest pixel, a half going up, as the saved file holds them, where the Shape reads what was set; Width and Height as the Shape.
"""
from __future__ import annotations

import json
import math
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURE = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures' / 'new_control_old_places.json'
MAKERS = {'button': 'ws.Shapes.AddFormControl(xlButtonControl, 12, 12, 100, 40)', 'check box': 'ws.Shapes.AddFormControl(xlCheckBox, 12, 12, 100, 40)', 'option button': 'ws.Shapes.AddFormControl(xlOptionButton, 12, 12, 100, 40)',
          'drop-down': 'ws.Shapes.AddFormControl(xlDropDown, 12, 12, 100, 40)', 'list box': 'ws.Shapes.AddFormControl(xlListBox, 12, 12, 100, 40)', 'group box': 'ws.Shapes.AddFormControl(xlGroupBox, 12, 12, 100, 40)',
          'label': 'ws.Shapes.AddFormControl(xlLabel, 12, 12, 100, 40)', 'spinner': 'ws.Shapes.AddFormControl(xlSpinner, 12, 12, 100, 40)', 'scroll bar': 'ws.Shapes.AddFormControl(xlScrollBar, 12, 12, 100, 40)',
          'rectangle': 'ws.Shapes.AddShape(1, 12, 12, 100, 40)', 'text box': 'ws.Shapes.AddTextbox(1, 12, 12, 100, 40)'}
VALUES = {'Left': (31, 30.1, 30.375, 30.6), 'Top': (31, 30.1, 30.375, 30.6), 'Width': (61, 60.6), 'Height': (46, 45.6)}
SOURCE = r'''Public Function Go() As String
    Dim ws As Object, s As Object, o As Object, r As String, before As String
    Set ws = ActiveSheet
    before = ws.StandardHeight
    Set s = %MAKE%
    Set o = ws.DrawingObjects(ws.DrawingObjects.Count)
    On Error Resume Next
    %SET%
    r = Err.Number & "|" & TypeName(o.Top) & ":" & o.Left & "," & o.Top & "," & o.Width & "," & o.Height & "|" & TypeName(s.Top) & ":" & s.Left & "," & s.Top & "," & s.Width & "," & s.Height
    Err.Clear
    On Error GoTo 0
    s.Delete
    Go = before & "~" & ws.StandardHeight & "~" & r
End Function
'''


def predicted_place(points: float) -> float:
    return math.floor(math.floor(points * 12700 + .5) / 12700 * 4 / 3 + .5) * .75


def main() -> None:
    cases = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name, make in MAKERS.items():
            for member, values in VALUES.items():
                for value in values:
                    for through in ('s', 'o'):
                        code = SOURCE.replace('%MAKE%', make).replace('%SET%', f'{through}.{member} = {value}')
                        result = excel.run_vba(code, 'Go', (), timeout=60.0, line_numbers=False)
                        assert result.ok and not result.dialogs, (name, member, value, result)
                        first, second, answer = str(result.value).split('~')
                        assert (first, second) == ('15', '15'), f'Excel is not at 96 DPI: {first}, {second}'
                        cases.append({'kind': name, 'set': f'{through}.{member} = {value}', 'answer': answer})
    FIXTURE.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': '15', 'source': SOURCE, 'makers': MAKERS, 'cases': cases}, indent=0), encoding='utf-8', newline='\n')
    right = total = 0
    wrong = []
    for case in cases:
        member, value = case['set'][2:].split(' = ')
        if case['kind'] in ('rectangle', 'text box') or member not in ('Left', 'Top'):
            continue
        total += 1
        old = case['answer'].split('|')[1].split(':')[1].split(',')
        if float(old[0 if member == 'Left' else 1]) == predicted_place(float(value)):
            right += 1
        else:
            wrong.append(case)
    print(f'old object place: {right} of {total}', wrong[:6])
    for case in cases:
        if case['kind'] in ('button', 'rectangle', 'text box', 'spinner') and (case['set'].endswith(('= 30.6', '= 60.6', '= 45.6')) or case['kind'] == 'button'):
            print(case['kind'], case['set'], case['answer'])


if __name__ == '__main__':
    main()
