"""What a loaded button's HorizontalAlignment, VerticalAlignment, Orientation and ReadingOrder write at each value, and what a loaded
button, check box and option button read of PhoneticAccelerator straight after a set. The workbook Excel saved for
tests/fixtures/old_collection_members.json is opened, the first control of the kind changed and saved, late bound; every part that
differs from the same workbook saved with nothing changed is kept, before and after, with what the control reads in the session and
after a reopen.

Rival rules, written before measuring: each alignment is one DrawingML bodyPr or pPr attribute and one VML ClientData child, the inset on
the side it aligns to dropped (as xlLeft and xlTop were found to); or each value writes its own set. PhoneticAccelerator holds its value
in the session and the file loses it, as the buttons' roles do, or it reads none back at all, as Accelerator does.

    python scripts/measure_button_text.py

writes tests/fixtures/button_text.json, which tests/test_excel_loaded_control_edits.py replays.
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
from measure_loaded_control_edits import BOOK, parts  # noqa: E402

OUT = ROOT / 'tests/fixtures/button_text.json'
CONSTANTS = {'xlLeft': -4131, 'xlCenter': -4108, 'xlRight': -4152, 'xlJustify': -4130, 'xlDistributed': -4117, 'xlTop': -4160,
             'xlBottom': -4107, 'xlHorizontal': -4128, 'xlUpward': -4171, 'xlDownward': -4170, 'xlVertical': -4166, 'xlRTL': -5004,
             'xlLTR': -5003, 'xlContext': -5002}
BUTTON = {
    'HorizontalAlignment': ('xlLeft', 'xlCenter', 'xlRight', 'xlJustify', 'xlDistributed'),
    'VerticalAlignment': ('xlTop', 'xlCenter', 'xlBottom', 'xlJustify', 'xlDistributed'),
    'Orientation': ('xlHorizontal', 'xlUpward', 'xlDownward', 'xlVertical'),
    'ReadingOrder': ('xlRTL', 'xlLTR', 'xlContext'),
    'PhoneticAccelerator': ('"Q"',),
}
CASES = ([('Buttons', member, value) for member, values in BUTTON.items() for value in values]
         + [(kind, 'PhoneticAccelerator', '"Q"') for kind in ('CheckBoxes', 'OptionButtons')])
SOURCE = r'''Public Function Row(ByVal path As String, ByVal saved As String, ByVal kind As String, ByVal member As String, ByVal value As Variant, ByVal change As Boolean) As String
    Dim wb As Workbook, o As Object, e As Long, s As String, during As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set o = CallByName(wb.Worksheets(1), kind, VbMethod, 1)
    If change Then
        On Error Resume Next
        CallByName o, member, VbLet, value
        e = Err.Number
        Err.Clear
        during = TypeName(CallByName(o, member, VbGet)) & " " & CStr(CallByName(o, member, VbGet))
        If Err.Number Then during = "E" & Err.Number
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
    Row = "E" & e & "|during " & during & "|reopened " & s
End Function
'''


def value_of(text: str) -> object:
    return text.strip('"') if text.startswith('"') else CONSTANTS[text]


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        folder = Path(held)
        bases: dict[str, dict[str, str]] = {}
        for kind in sorted({kind for kind, _, _ in CASES}):
            book, saved = folder / 'in.xlsx', folder / f'base_{kind}.xlsx'
            book.write_bytes(BOOK.read_bytes())
            ran = excel.run_vba(SOURCE, 'Row', (str(book), str(saved), kind, 'Visible', True, False), timeout=60.0, line_numbers=False)
            assert ran.ok and not ran.dialogs, ran.outcome
            bases[kind] = parts(saved)
        for at, (kind, member, text) in enumerate(CASES):
            book, saved = folder / 'in.xlsx', folder / f'out_{at}.xlsx'
            book.write_bytes(BOOK.read_bytes())
            ran = excel.run_vba(SOURCE, 'Row', (str(book), str(saved), kind, member, value_of(text), True), timeout=60.0, line_numbers=False)
            answer = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
            after = parts(saved) if saved.exists() else {}
            changed = {name: {'before': bases[kind].get(name, ''), 'after': part} for name, part in after.items() if bases[kind].get(name) != part}
            cases.append({'kind': kind, 'member': member, 'value': text, 'excel': answer, 'parts': changed})
            print(kind, member, text, answer, sorted(changed), flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'cases': cases}, indent=1) + '\n',
                   encoding='utf-8')


if __name__ == '__main__':
    main()
