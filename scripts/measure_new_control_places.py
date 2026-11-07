"""Where a place or size set on a new form control lands, straight after the set and from the saved file, into tests/fixtures/new_control_places.json. Needs Excel at 96 DPI, checked at both ends of each call.

The values split two rules for what the set does, written below before Excel was asked: floor, back to a pixel, and nearest, with a half going up. Each is scored for a size as read straight after, and for a place and a size as read from the file.
"""
from __future__ import annotations

import json
import math
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURE = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures' / 'new_control_places.json'
KINDS = ('xlButtonControl', 'xlCheckBox', 'xlOptionButton', 'xlDropDown', 'xlListBox', 'xlGroupBox', 'xlLabel', 'xlSpinner', 'xlScrollBar')
PLACES = (30.1, 30.2, 30.3, 30.375, 30.4, 30.5, 30.6, 30.7, 30.9, 31.0, 31.125, 31.2, 45.125, 45.375, 45.625)
SIZES = (60.1, 60.2, 60.375, 60.4, 60.6, 60.9, 61.125, 61.5, 45.125, 45.375, 45.6, 46.0)
MEMBERS = ('Left', 'Top', 'Width', 'Height')
SOURCE = r'''Public Function Go(ByVal kind As Long, ByVal member As String, ByVal value As Double, ByVal path As String) As String
    Dim wb As Workbook, s As Object, r As String, before As String, after As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    before = wb.Worksheets(1).StandardHeight
    Set s = wb.Worksheets(1).Shapes.AddFormControl(kind, 12, 12, 100, 40)
    Select Case member
        Case "Left": s.Left = value
        Case "Top": s.Top = value
        Case "Width": s.Width = value
        Case "Height": s.Height = value
    End Select
    r = TypeName(s.Left) & ":" & s.Left & "," & s.Top & "," & s.Width & "," & s.Height
    after = wb.Worksheets(1).StandardHeight
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(path)
    Set s = wb.Worksheets(1).Shapes(1)
    r = r & "|" & s.Left & "," & s.Top & "," & s.Width & "," & s.Height
    wb.Close False
    Application.DisplayAlerts = True
    Go = before & "~" & after & "~" & r
End Function
'''
CONSTANTS = {'xlButtonControl': 0, 'xlCheckBox': 1, 'xlDropDown': 2, 'xlListBox': 6, 'xlGroupBox': 4, 'xlLabel': 5, 'xlOptionButton': 7, 'xlScrollBar': 8, 'xlSpinner': 9}


def floor_rule(points: float) -> float:
    return math.floor(points * 4 / 3) * .75


def nearest_rule(points: float) -> float:
    return math.floor(points * 4 / 3 + .5) * .75


def main() -> None:
    cases = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for kind in KINDS:
            for member in MEMBERS:
                for value in PLACES if member in ('Left', 'Top') else SIZES:
                    path = str(Path(folder) / f'{len(cases)}.xlsx')
                    result = excel.run_vba(SOURCE, 'Go', (CONSTANTS[kind], member, value, path), timeout=60.0, line_numbers=False)
                    assert result.ok and not result.dialogs, (kind, member, value, result)
                    first, second, answer = str(result.value).split('~')
                    assert (first, second) == ('15', '15'), f'Excel is not at 96 DPI: {first}, {second}'
                    cases.append({'kind': kind, 'member': member, 'value': value, 'answer': answer})
    FIXTURE.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': '15', 'source': SOURCE, 'cases': cases}, indent=0), encoding='utf-8', newline='\n')
    for rule in (floor_rule, nearest_rule):
        right, total = 0, 0
        wrong = []
        for case in cases:
            at = MEMBERS.index(case['member'])
            read, saved = case['answer'].split('|')
            read_values, saved_values = read.split(':')[1].split(','), saved.split(',')
            if case['kind'] in ('xlSpinner', 'xlScrollBar', 'xlListBox') and at >= 2:
                continue
            total += 1
            if float(saved_values[at]) == rule(case['value']):
                right += 1
            else:
                wrong.append((case['kind'][2:8], case['member'], case['value'], saved_values[at], read_values[at]))
        print(rule.__name__, right, 'of', total, 'from the file; wrong', wrong[:10])
    for case in cases:
        if case['kind'] in ('xlSpinner', 'xlScrollBar', 'xlListBox') and case['member'] in ('Width', 'Height'):
            print(case['kind'][2:8], case['member'], case['value'], case['answer'])


if __name__ == '__main__':
    main()
