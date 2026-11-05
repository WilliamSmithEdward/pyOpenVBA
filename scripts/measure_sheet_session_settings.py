"""Worksheet settings across Excel's own save and reopen. Rival rules: kept in the file, so a model that loses them is wrong; or for the session only, so it is right. Hidden and shown, under a temporary folder.

    python scripts/measure_sheet_session_settings.py

writes tests/fixtures/sheet_session_settings.json, which tests/test_excel_sheet_save_reopen.py reads for what Excel loses.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/sheet_session_settings.json'
ROWS = [('ScrollArea', '"A1:C3"'), ('EnableCalculation', 'False'), ('DisplayPageBreaks', 'True'), ('EnableAutoFilter', 'True'), ('EnableOutlining', 'True'), ('EnablePivotTable', 'True')]


def source() -> str:
    text = ('Public Function Tried(ByVal which As Long, ByVal path As String, ByVal shown As Boolean) As String\n    Dim wb As Workbook, ws As Object, a As String, b As String, was As Boolean\n'
            '    Application.DisplayAlerts = False\n    was = Application.Visible\n    Application.Visible = shown\n    Set wb = Workbooks.Add(xlWBATWorksheet)\n    Set ws = wb.Worksheets(1)\n'
            '    ws.Range("A1").Value = 1\n    On Error Resume Next\n    Select Case which\n')
    for at, (member, value) in enumerate(ROWS):
        text += f'    Case {at}\n        ws.{member} = {value}\n        a = TypeName(ws.{member}) & " " & ws.{member} & " E" & Err.Number\n        Err.Clear\n'
    text += '    End Select\n    On Error GoTo 0\n    wb.SaveAs path, xlOpenXMLWorkbook\n    wb.Close False\n    Set wb = Workbooks.Open(path, 0)\n    Set ws = wb.Worksheets(1)\n    On Error Resume Next\n    Select Case which\n'
    for at, (member, _) in enumerate(ROWS):
        text += f'    Case {at}\n        b = TypeName(ws.{member}) & " " & ws.{member} & " E" & Err.Number\n'
    return text + '    End Select\n    On Error GoTo 0\n    wb.Close False\n    Application.Visible = was\n    Tried = a & "|" & b\nEnd Function\n'


#: Top on a new control, read through its old object, which the sweep saw read 31 from the model's set and 30.75 from its save. Rival rules: Excel reads 31 until a save, so it changes there too; or it
#: reads the snapped 30.75 straight away, so the model is wrong. It is the second, though Shape.Top reads 31 after the same set.
CONTROLS = {'Buttons': '10, 10, 100, 40', 'CheckBoxes': '10, 10, 100, 20', 'OptionButtons': '10, 10, 100, 20', 'DropDowns': '10, 10, 100, 20', 'ListBoxes': '10, 10, 100, 60'}


def control_source(name: str, box: str) -> str:
    return (f'Public Function Tried(ByVal path As String) As String\n    Dim wb As Workbook, s As Object, a As String\n    Application.DisplayAlerts = False\n    Set wb = Workbooks.Add(xlWBATWorksheet)\n'
            f'    Set s = wb.Worksheets(1).{name}.Add({box})\n    s.Top = 31\n    a = s.Top\n    wb.SaveAs path, xlOpenXMLWorkbook\n    wb.Close False\n    Set wb = Workbooks.Open(path, 0)\n'
            f'    Tried = wb.Worksheets(1).StandardHeight & "|" & a & "|" & wb.Worksheets(1).{name}(1).Top\n    wb.Close False\nEnd Function\n')


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases: list[dict[str, object]] = []
    controls: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        for shown in (False, True):
            for at, (member, value) in enumerate(ROWS):
                done = excel.run_vba(source(), 'Tried', (at, str(Path(held) / f'r{at}{shown}.xlsx'), shown), timeout=40.0, line_numbers=False)
                assert done.ok, done
                before, _, after = str(done.value).partition('|')
                cases.append({'member': member, 'value': value, 'shown': shown, 'before': before, 'after': after})
                print(member, shown, before, '|', after, flush=True)
        for name, box in CONTROLS.items():
            done = excel.run_vba(control_source(name, box), 'Tried', (str(Path(held) / f'{name}.xlsx'),), timeout=40.0, line_numbers=False)
            assert done.ok, done
            height, before, after = str(done.value).split('|')
            assert height == '15', f'StandardHeight {height}: geometry would not be at 96 DPI'
            controls.append({'collection': name, 'before': before, 'after': after})
            print(name, before, '|', after, flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': cases, 'controls': controls}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
