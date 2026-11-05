"""A new form control through Shape: where its box lands, what it reads from the first, and what each setting reads and leaves in the file across a save and a reopen.

    python scripts/measure_new_control_settings.py

A control's box is rounded to the display's pixels, so where it lands is measured only with Excel at 96 DPI, where StandardHeight reads 15; at any other, those cases wait for
a later run. Each run records the StandardHeight it saw. The settings are measured at either. It uses no hyperlink, mail or send member.

Writes tests/fixtures/new_control_settings.json and, at 96 DPI, tests/fixtures/new_control_boxes.json. tests/test_excel_new_control_settings.py replays them.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
#: XlFormControl, as each is called here.
KINDS = {'Button': 0, 'CheckBox': 1, 'DropDown': 2, 'GroupBox': 4, 'Label': 5, 'ListBox': 6, 'OptionButton': 7, 'ScrollBar': 8, 'Spinner': 9}
#: Left, top, width, height. Each pair of the last four is split so that rounding the width alone and rounding the right edge differ.
BOXES = ((100, 50, 120, 60), (10, 60, 90, 20), (11, 61, 91, 21), (10, 60, 90.5, 20.5), (10.4, 60.4, 90.4, 20.4), (10.7, 60.7, 90.7, 20.7), (100.3, 50.6, 33.3, 17.7), (5.1, 5.1, 5.1, 5.1))
SETTINGS = ('s.AlternativeText = "Alt"', 's.Title = "Ttl"', 's.Rotation = 30', 's.Visible = msoFalse', 's.Placement = xlMoveAndSize', 's.Placement = xlMove', 's.Placement = xlFreeFloating', 's.Decorative = msoTrue', 's.Locked = msoFalse',
            's.LockAspectRatio = msoTrue', 's.Left = 33', 's.Top = 44.4', 's.Width = 55.5', 's.Height = 33.3')
READS = ('s.Name', 's.Left', 's.Top', 's.Width', 's.Height', 's.Placement', 's.AlternativeText', 's.Title', 's.Visible', 's.Locked', 's.LockAspectRatio', 's.Rotation', 's.Decorative', 's.Type', 's.FormControlType', 's.AutoShapeType', 's.Connector', 's.ConnectionSiteCount',
         's.ZOrderPosition', 's.HorizontalFlip', 's.VerticalFlip', 's.Fill.Visible', 's.Line.Visible', 's.OnAction', 's.TopLeftCell.Address', 's.BottomRightCell.Address', 's.ID', 's.Parent.Name')


def _block(name: str, form: str) -> list[str]:
    lines = [f'Private Function {name}(s As Object) As String', '    Dim v As Variant, r As String', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', f'    v = {read}', '    If Err.Number Then', '        r = r & "|E"' + (' & Err.Number' if form == 'value' else ''), '    Else', '        r = r & "|" & ' + ('CStr(v)' if form == 'value' else 'TypeName(v)'), '    End If']
    return [*lines, f'    {name} = r', 'End Function', '']


def source(kind: int, box: tuple[float, ...], statement: str) -> str:
    """One case. The answer is what reads before the save, then after the reopen, each with its types."""
    left, top, width, height = box
    lines = [*_block('Reads', 'value'), *_block('TypesOf', 'type'), 'Public Function Go(ByVal folder As String) As String', '    Dim wb As Workbook, ws As Object, s As Object, r As String, name As String', '    Application.DisplayAlerts = False',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', f'    Set s = ws.Shapes.AddFormControl({kind}, {left}, {top}, {width}, {height})', '    name = s.Name', '    On Error Resume Next',
             f'    {statement}' if statement else "    ' Nothing is set.", '    If Err.Number Then r = "E" & Err.Number & "!"', '    Err.Clear', '    On Error GoTo 0', '    r = r & Reads(s) & "##" & TypesOf(s) & "@@"', '    wb.SaveAs folder & "\\saved.xlsx", xlOpenXMLWorkbook', '    wb.Close False',
             '    Set wb = Workbooks.Open(folder & "\\saved.xlsx")', '    r = r & Reads(wb.Worksheets(1).Shapes(name)) & "##" & TypesOf(wb.Worksheets(1).Shapes(name))', '    wb.Close False', '    Application.DisplayAlerts = True', '    Go = r', 'End Function']
    return '\n'.join(lines) + '\n'


def parts_of(path: Path) -> dict[str, str]:
    """What the file holds of the one control: its drawing anchor, the sheet's entry, and its VML."""
    with zipfile.ZipFile(path) as package:
        found: dict[str, str] = {'drawing': package.read('xl/drawings/drawing1.xml').decode('utf-8')}
        for part in package.namelist():
            if part.endswith('.vml'):
                found['vml'] = package.read(part).decode('utf-8')
            if part.startswith('xl/worksheets/sheet') and part.endswith('.xml'):
                control = re.search(r'<controls>.*?</controls>', package.read(part).decode('utf-8'), re.S)
                if control:
                    found['controls'] = control.group()
            if part.startswith('xl/ctrlProps/'):
                found['properties'] = package.read(part).decode('utf-8')
    return found


def _measured(excel: ExcelSession, folder: str, cases: list[tuple[str, tuple[float, ...], str]]) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    for name, box, statement in cases:
        result = excel.run_vba(source(KINDS[name], box, statement), 'Go', (folder,), timeout=60.0, line_numbers=False)
        assert result.ok and not result.dialogs, (name, box, statement, result)
        records.append({'kind': name, 'box': list(box), 'statement': statement, 'answer': str(result.value), **parts_of(Path(folder) / 'saved.xlsx')})
    return records


def _written(stem: str, height: str, records: list[dict[str, object]]) -> None:
    (FIXTURES / f'{stem}.json').write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': height, 'kinds': KINDS, 'reads': READS, 'cases': records}, indent=0) + '\n', encoding='utf-8')
    print(stem, len(records), 'cases;', sum(str(one['answer']).startswith('E') for one in records), 'raised')


def main() -> None:
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        checked = excel.run_vba('Public Function Height() As String\n    Height = CStr(ActiveSheet.StandardHeight)\nEnd Function\n', 'Height', (), timeout=30.0, line_numbers=False)
        assert checked.ok, checked
        height = str(checked.value)
        _written('new_control_settings', height, _measured(excel, folder, [(name, BOXES[0], statement) for name in KINDS for statement in ('', *SETTINGS)]))
        if height == '15':
            boxes = _measured(excel, folder, [(name, box, '') for name in KINDS for box in BOXES])
            # The display can change while this runs, with a remote session coming or going.
            again = excel.run_vba('Public Function Height() As String\n    Height = CStr(ActiveSheet.StandardHeight)\nEnd Function\n', 'Height', (), timeout=30.0, line_numbers=False)
            assert again.ok and str(again.value) == '15', f'The display changed during the run: StandardHeight reads {again.value}'
            _written('new_control_boxes', height, boxes)
        else:
            print(f'StandardHeight reads {height}, not 15: the boxes wait for Excel at 96 DPI')


if __name__ == '__main__':
    main()
