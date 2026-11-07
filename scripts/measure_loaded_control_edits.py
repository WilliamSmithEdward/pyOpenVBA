"""What Excel writes when a setting changes on a form control loaded from a file: for each kind and setting, the workbook Excel saved for
tests/fixtures/old_collection_members.json is opened, the first control of the kind changed, and saved; every part that differs from the
same workbook saved with nothing changed is kept, before and after, with what the control reads after a reopen.

Rival rules, written before measuring: a setting lives in the control's ctrlProp part, or in its VML shape, or in both; a setting at its
default is left out of the part, or written with its default.

    python scripts/measure_loaded_control_edits.py

writes tests/fixtures/loaded_control_edits.json, which tests/test_excel_loaded_control_edits.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/loaded_control_edits.json'
BOOK = ROOT / 'tests/fixtures/old_collections.xlsx'
KINDS = ('Buttons', 'CheckBoxes', 'OptionButtons', 'DropDowns', 'ListBoxes')
#: Each setting and the value it is set to, one that differs from what the first control of each kind holds.
SETTINGS = {
    'Enabled': 'False', 'LockedText': 'False', 'PrintObject': 'False', 'Display3DShading': 'True', 'DropDownLines': '4', 'PhoneticAccelerator': '"Q"',
    'Accelerator': '"K"', 'AddIndent': 'True', 'ReadingOrder': 'xlRTL', 'HorizontalAlignment': 'xlLeft', 'VerticalAlignment': 'xlTop',
    'Orientation': 'xlUpward', 'AutoSize': 'True', 'Caption': '"Fresh"', 'Locked': 'False', 'Placement': 'xlFreeFloating', 'Visible': 'False',
    'DefaultButton': 'True', 'CancelButton': 'True', 'DismissButton': 'True', 'HelpButton': 'True', 'MultiSelect': 'xlSimple',
}
SOURCE = r'''Public Function Row(ByVal path As String, ByVal saved As String, ByVal kind As String, ByVal member As String, ByVal value As Variant, ByVal change As Boolean) As String
    Dim wb As Workbook, o As Object, e As Long, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set o = CallByName(wb.Worksheets(1), kind, VbMethod, 1)
    If change Then
        On Error Resume Next
        CallByName o, member, VbLet, value
        e = Err.Number
        Err.Clear
        On Error GoTo 0
    End If
    wb.SaveAs saved, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(saved, 0)
    Set o = CallByName(wb.Worksheets(1), kind, VbMethod, 1)
    On Error Resume Next
    s = TypeName(CallByName(o, member, VbGet)) & " " & CStr(CallByName(o, member, VbGet))
    If Err.Number Then s = "E" & Err.Number
    On Error GoTo 0
    wb.Close False
    Row = "E" & e & "|reopened " & s
End Function
'''
#: VBA constants the values name, as a Variant must carry them.
CONSTANTS = {'xlRTL': -5004, 'xlLeft': -4131, 'xlTop': -4160, 'xlUpward': -4171, 'xlFreeFloating': 3, 'xlSimple': 1}


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


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        home = Path(held)
        with ExcelSession(HarnessConfig(exclusive=False)) as excel:
            excel.new_document()
            bases: dict[str, dict[str, str]] = {}
            for kind in KINDS:
                book, saved = home / f'{kind}_base_in.xlsx', home / f'{kind}_base.xlsx'
                book.write_bytes(BOOK.read_bytes())
                ran = excel.run_vba(SOURCE, 'Row', (str(book), str(saved), kind, 'Name', '', False), timeout=60.0, line_numbers=False)
                assert ran.ok, ran
                bases[kind] = parts(saved)
            for kind in KINDS:
                for member, text in SETTINGS.items():
                    book, saved = home / f'{kind}_{member}_in.xlsx', home / f'{kind}_{member}.xlsx'
                    book.write_bytes(BOOK.read_bytes())
                    ran = excel.run_vba(SOURCE, 'Row', (str(book), str(saved), kind, member, value_of(text), True), timeout=60.0, line_numbers=False)
                    answer = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
                    changed = parts(saved) if saved.exists() else {}
                    differ = {name: {'before': bases[kind].get(name, ''), 'after': text_} for name, text_ in changed.items() if bases[kind].get(name) != text_}
                    cases.append({'kind': kind, 'member': member, 'value': text, 'excel': answer, 'parts': differ})
                    print(kind, member, answer, sorted(differ), flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'cases': cases}, indent=0) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
