"""What Excel writes when a setting changes on a form control a macro has just made: for each kind and setting, a new workbook gets one
control through its old collection's Add, the setting is changed, and the workbook is saved; every part that differs from the same
workbook saved with nothing changed is kept, before and after, with what the set answers, what the control reads in the session, and what
it reads after a reopen.

Rival rules, written before measuring: a new control writes a setting where a loaded one does (tests/fixtures/loaded_control_edits.json),
or a new control writes it nowhere and reads it only in the session, or the set is refused. It reads no size or place, so the display's
DPI does not enter, and it names no clipboard, mail, print or crypto member.

    python scripts/measure_new_control_edits.py
    python scripts/measure_new_control_edits.py model

The first writes tests/fixtures/new_control_edits.json, which tests/test_excel_new_control_edits.py replays. The second has the model make each setting
and save, and records what Excel reads of the model's file as excel_reads_model.
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/new_control_edits.json'
#: Each kind's settings and the values tried, as VBA spells them.
SETTINGS: dict[str, tuple[tuple[str, str], ...]] = {
    'Buttons': (('AddIndent', 'True'), ('DefaultButton', 'True'), ('CancelButton', 'True'), ('DismissButton', 'True'), ('HelpButton', 'True'),
                ('HorizontalAlignment', 'xlCenter'), ('HorizontalAlignment', 'xlLeft'), ('HorizontalAlignment', 'xlDistributed'),
                ('VerticalAlignment', 'xlCenter'), ('VerticalAlignment', 'xlTop'), ('VerticalAlignment', 'xlBottom'),
                ('Orientation', 'xlUpward'), ('Orientation', 'xlHorizontal'), ('ReadingOrder', 'xlRTL'), ('ReadingOrder', 'xlContext'),
                ('PhoneticAccelerator', '"Q"'), ('AutoScaleFont', 'True'), ('Formula', '"=$A$21"'), ('AutoSize', 'True')),
    'CheckBoxes': (('Display3DShading', 'True'), ('PhoneticAccelerator', '"Q"')),
    'OptionButtons': (('Display3DShading', 'True'), ('PhoneticAccelerator', '"Q"')),
    'DropDowns': (('Display3DShading', 'True'),),
    'ListBoxes': (('Display3DShading', 'True'),),
}
#: Where each kind's control goes, as the made book of scripts/measure_old_collection_reads.py puts its first.
BOXES = {'Buttons': '100, 70, 60, 30', 'CheckBoxes': '180, 70, 60, 30', 'DropDowns': '260, 70, 60, 20', 'ListBoxes': '420, 70, 60, 50',
         'OptionButtons': '500, 70, 60, 30'}
SOURCE = r'''Private Function Show(ByVal v As Variant) As String
    If IsObject(v) Then
        Show = "obj " & TypeName(v)
    ElseIf IsNull(v) Then
        Show = "null"
    Else
        Show = TypeName(v) & " " & CStr(v)
    End If
End Function

Private Function ReadOf(ByVal o As Object, ByVal member As String) As String
    On Error Resume Next
    Dim v As Variant
    v = CallByName(o, member, VbGet)
    If Err.Number Then ReadOf = "E" & Err.Number Else ReadOf = Show(v)
End Function

Public Function Row(ByVal saved As String, ByVal kind As String, ByVal box As String, ByVal member As String, ByVal value As Variant, ByVal change As Boolean) As String
    Dim wb As Workbook, ws As Worksheet, o As Object, e As Long, s As String, b() As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A20").Value = "one": ws.Range("A21").Value = "two": ws.Range("A22").Value = "three"
    b = Split(box, ", ")
    Set o = CallByName(CallByName(ws, kind, VbGet), "Add", VbMethod, CDbl(b(0)), CDbl(b(1)), CDbl(b(2)), CDbl(b(3)))
    If change Then
        On Error Resume Next
        CallByName o, member, VbLet, value
        e = Err.Number
        Err.Clear
        On Error GoTo 0
    End If
    s = ReadOf(o, member)
    wb.SaveAs saved, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(saved, 0)
    Set o = CallByName(wb.Worksheets(1), kind, VbMethod, 1)
    s = s & "|reopened " & ReadOf(o, member)
    wb.Close False
    Row = "E" & e & "|" & s
End Function
'''
#: VBA constants the values name, as a Variant must carry them.
CONSTANTS = {'xlRTL': -5004, 'xlContext': -5002, 'xlLeft': -4131, 'xlCenter': -4108, 'xlDistributed': -4117, 'xlTop': -4160, 'xlBottom': -4107,
             'xlUpward': -4171, 'xlHorizontal': -4128}


def value_of(text: str) -> object:
    if text in ('True', 'False'):
        return text == 'True'
    if text.startswith('"'):
        return text.strip('"')
    if text in CONSTANTS:
        return CONSTANTS[text]
    return int(text)


def parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name).decode('utf-8', 'replace') for name in package.namelist()
                if re.search(r'ctrlProp|vmlDrawing|drawings/drawing|worksheets/sheet', name)}


READ_BACK = r'''Public Function ReadBack(ByVal path As String, ByVal kind As String, ByVal member As String) As String
    Dim wb As Workbook, o As Object, v As Variant, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set o = CallByName(wb.Worksheets(1), kind, VbMethod, 1)
    On Error Resume Next
    v = CallByName(o, member, VbGet)
    If Err.Number Then s = "E" & Err.Number Else s = TypeName(v) & " " & CStr(v)
    On Error GoTo 0
    wb.Close False
    ReadBack = s
End Function
'''


def model_saves() -> None:
    """Each setting the model makes, saved by the model; Excel opens the model's file and reads the setting: excel_reads_model."""
    from pyvbaharness import ExcelSession, HarnessConfig

    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    record = json.loads(OUT.read_text(encoding='utf-8'))
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        for number, case in enumerate(record['cases']):
            saved = Path(held) / f'model{number}.xlsx'
            app = ExcelApplication(sandbox=False)
            app.add_workbook()
            app.add_module(record['source'], name='Probe')
            try:
                app.run('Probe.Row', str(saved), case['kind'], record['boxes'][case['kind']], case['member'], value_of(str(case['value'])), True)
            except VBAUnsupportedError:
                case.pop('excel_reads_model', None)
                continue
            ran = excel.run_vba(READ_BACK, 'ReadBack', (str(saved), case['kind'], case['member']), timeout=60.0, line_numbers=False)
            case['excel_reads_model'] = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
            print(case['kind'], case['member'], case['value'], case['excel_reads_model'], flush=True)
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        home = Path(held)
        excel.new_document()
        for kind, settings in SETTINGS.items():
            base = home / f'{kind}_base.xlsx'
            ran = excel.run_vba(SOURCE, 'Row', (str(base), kind, BOXES[kind], 'Name', '', False), timeout=60.0, line_numbers=False)
            assert ran.ok and not ran.dialogs, ran
            before = parts(base)
            for number, (member, text) in enumerate(settings):
                saved = home / f'{kind}_{number}.xlsx'
                ran = excel.run_vba(SOURCE, 'Row', (str(saved), kind, BOXES[kind], member, value_of(text), True), timeout=60.0, line_numbers=False)
                answer = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
                changed = parts(saved) if saved.exists() else {}
                differ = {name: {'before': before.get(name, ''), 'after': part} for name, part in changed.items() if before.get(name) != part}
                cases.append({'kind': kind, 'member': member, 'value': text, 'excel': answer, 'parts': differ})
                print(kind, member, text, answer, sorted(differ), flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'boxes': BOXES, 'cases': cases}, indent=0) + '\n',
                   encoding='utf-8')


if __name__ == '__main__':
    if sys.argv[1:] == ['model']:
        model_saves()
    else:
        main()
