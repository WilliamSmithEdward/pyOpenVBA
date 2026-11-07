"""Why a height set through a new form control's old object leaves the Shape's Top an EMU short, into tests/fixtures/new_control_old_heights.json. Needs Excel at 96 DPI, checked at both ends of each call.

Rival rules, written before Excel was asked. ALWAYS: Top loses one EMU whatever the place and height. PLACE: it depends on the Top alone. HEIGHT: it depends on the height set. BOTTOM: Top is worked
back from the bottom edge at the pixel, so it moves by what the snapped height leaves over. Each case reads Top in EMUs before and after, and after a save and reopen.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURE = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures' / 'new_control_old_heights.json'
TOPS = (12, 10.5, 30, 31, 7.2, 100)
HEIGHTS = (46, 45.6, 45.75, 40, 39.75, 60, 30.1)
SOURCE = r'''Public Function Go(ByVal kind As Long, ByVal top As Double, ByVal tall As Double, ByVal through As String, ByVal path As String) As String
    Dim wb As Workbook, s As Object, o As Object, r As String, before As String, after As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    before = wb.Worksheets(1).StandardHeight
    Set s = wb.Worksheets(1).Shapes.AddFormControl(kind, 12, top, 100, 40)
    Set o = wb.Worksheets(1).DrawingObjects(1)
    r = CStr(CDbl(s.Top) * 12700) & "," & CStr(CDbl(s.Height) * 12700)
    If through = "o" Then o.Height = tall Else s.Height = tall
    r = r & "|" & CStr(CDbl(s.Top) * 12700) & "," & CStr(CDbl(s.Height) * 12700) & "," & o.Top & "," & o.Height
    after = wb.Worksheets(1).StandardHeight
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(path)
    Set s = wb.Worksheets(1).Shapes(1)
    r = r & "|" & CStr(CDbl(s.Top) * 12700) & "," & CStr(CDbl(s.Height) * 12700)
    wb.Close False
    Application.DisplayAlerts = True
    Go = before & "~" & after & "~" & r
End Function
'''


def main() -> None:
    cases = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for kind, name in ((0, 'button'), (1, 'check box')):
            for top in TOPS:
                for tall in HEIGHTS:
                    for through in ('o', 's'):
                        path = str(Path(folder) / f'{len(cases)}.xlsx')
                        result = excel.run_vba(SOURCE, 'Go', (kind, top, tall, through, path), timeout=60.0, line_numbers=False)
                        assert result.ok and not result.dialogs, (name, top, tall, result)
                        first, second, answer = str(result.value).split('~')
                        assert (first, second) == ('15', '15'), f'Excel is not at 96 DPI: {first}, {second}'
                        cases.append({'kind': name, 'top': top, 'height': tall, 'through': through, 'answer': answer})
    FIXTURE.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': '15', 'source': SOURCE, 'cases': cases}, indent=0), encoding='utf-8', newline='\n')
    for case in cases:
        made, after, saved = case['answer'].split('|')
        print(case['kind'][:6], case['through'], case['top'], case['height'], '| made', made, '| after', after, '| saved', saved)


if __name__ == '__main__':
    main()
