"""ControlFormat's DropDownLines, Enabled, ListIndex, LockedText and PrintObject: what a drop-down, a list box, a
check box and a button answer before and after each setting, the errors, and the control's saved ctrlProp part and
VML client data.

Each case makes a workbook with J1:J3 holding a, b and c, adds the case's form control at (10, 10, 90, 60), gives a
drop-down or list box J1:J3 as its list, runs the case's statement through cf, then reads the five properties, each
as its value or "E" and the error number, and saves it.

    python scripts/measure_control_format_more.py

writes tests/fixtures/control_format_more.json, which tests/test_excel_control_format_more.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/control_format_more.json'
SETTINGS = ['', 'cf.Enabled = False', 'cf.LockedText = False', 'cf.PrintObject = False',
            'cf.Enabled = False: cf.Enabled = True']
LISTS = ['cf.DropDownLines = 3', 'cf.DropDownLines = 0', 'cf.DropDownLines = -1', 'cf.DropDownLines = 40000',
         'cf.ListIndex = 2', 'cf.ListIndex = 0', 'cf.ListIndex = 4', 'cf.Value = 3']
CASES = [(kind, statement) for kind in ('xlDropDown', 'xlListBox') for statement in SETTINGS + LISTS] + [
    (kind, statement) for kind in ('xlCheckBox', 'xlButtonControl')
    for statement in SETTINGS + ['cf.DropDownLines = 3', 'cf.ListIndex = 1']]
NAMES = ('DropDownLines', 'Enabled', 'ListIndex', 'LockedText', 'PrintObject')


def _probe() -> str:
    lines = ['Private Function Safe(ByVal target As Object, ByVal name As String) As String',
             '    On Error GoTo Failed', '    Safe = CallByName(target, name, VbGet)', '    Exit Function', 'Failed:',
             '    Safe = "E" & Err.Number', 'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal kind As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, cf As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("J1:J3").Value = Application.Transpose(Array("a", "b", "c"))',
             '    Set cf = ws.Shapes.AddFormControl(kind, 10, 10, 90, 60).ControlFormat',
             '    If kind = xlDropDown Or kind = xlListBox Then cf.ListFillRange = "$J$1:$J$3"',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (_, statement) in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {statement}'] if statement else [])
    lines += ['    End Select',
              '    s = ' + ' & "," & '.join(f'Safe(cf, "{name}")' for name in NAMES),
              '    Tried = "ok " & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim s As String',
              '    Application.DisplayAlerts = False']
    for index, (kind, _) in enumerate(CASES):
        lines.append(f'    s = s & "@@@" & Tried({index}, {kind}, folder)')
    lines += ['    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def control_parts(path: Path) -> list[str]:
    """The control's ctrlProp part and its VML shape's client data, as saved."""
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        found = [package.read(name).decode('utf-8') for name in names if name.startswith('xl/ctrlProps/')]
        for name in names:
            if name.endswith('.vml'):
                found += re.findall(r'<x:ClientData\b.*?</x:ClientData>', package.read(name).decode('utf-8'), re.S)
    return found


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        saved = [control_parts(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[f'{kind}: {statement}', answer, found]
                        for (kind, statement), answer, found in zip(CASES, answers, saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, found in record['cases']:
        props = re.search(r'<formControlPr\b[^>]*/?>', found[0]).group(0) if found else ''
        print(case, '->', answer, '||', props)


if __name__ == '__main__':
    main()
