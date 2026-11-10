"""What Excel writes when a setting changes on a drawing loaded from a file: a rectangle, an oval, an arc and a text box from the workbook
Excel saved for tests/fixtures/old_collection_members.json. The first item of the kind is changed, late bound, and the workbook saved;
what changed in every part that differs from the same workbook saved with nothing changed is kept, tag by tag, with what the item reads
in the session and after a reopen.

Rival rules, written before measuring: a drawing keeps each setting in its DrawingML alone, as the shapes it is (bodyPr, pPr, prstGeom,
effectLst, clientData), and Excel patches only that; or Excel writes the whole shape again. A text setting on an arc, which has no text
frame, is refused (438) or taken and lost.

Safety (the Supervisor's rules, 2026-11-07): one Excel through live.lock (an exclusive session), no mail, print, export, encryption,
signing, certificate, IRM, sensitivity-label or identity member, and no window touched. AutoSize changes a shape's size, so the
default row height is read at both ends of the run and nothing is written unless it is 15.

    python scripts/measure_loaded_drawing_edits.py [kind ...]

writes the kinds named (all by default) into tests/fixtures/loaded_drawing_edits.json, keeping the kinds already there.
"""
from __future__ import annotations

import difflib
import json
import re
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
from measure_loaded_control_edits import BOOK, parts  # noqa: E402

OUT = ROOT / 'tests/fixtures/loaded_drawing_edits.json'
KINDS = ('Rectangles', 'Ovals', 'Arcs', 'TextBoxes')
CONSTANTS = {'xlLeft': -4131, 'xlCenter': -4108, 'xlRight': -4152, 'xlJustify': -4130, 'xlDistributed': -4117, 'xlTop': -4160,
             'xlBottom': -4107, 'xlHorizontal': -4128, 'xlUpward': -4171, 'xlDownward': -4170, 'xlVertical': -4166, 'xlRTL': -5004,
             'xlLTR': -5003, 'xlContext': -5002, 'xlFreeFloating': 3, 'xlMove': 2, 'xlMoveAndSize': 1}
#: Each setting and the values it is set to.
SETTINGS: dict[str, tuple[str, ...]] = {
    'HorizontalAlignment': ('xlLeft', 'xlCenter', 'xlRight', 'xlJustify', 'xlDistributed'),
    'VerticalAlignment': ('xlTop', 'xlCenter', 'xlBottom', 'xlJustify', 'xlDistributed'),
    'Orientation': ('xlHorizontal', 'xlUpward', 'xlDownward', 'xlVertical'),
    'ReadingOrder': ('xlRTL', 'xlLTR', 'xlContext'),
    'AddIndent': ('True',), 'LockedText': ('False',), 'RoundedCorners': ('True',), 'Shadow': ('True',), 'AutoSize': ('True',),
    'Caption': ('"Fresh"',), 'Locked': ('False',), 'Placement': ('xlFreeFloating', 'xlMove'), 'Visible': ('False',), 'PrintObject': ('False',),
    'Enabled': ('False',), 'Formula': ('"=$A$1"',), 'Accelerator': ('"K"',),
}
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

Public Function RowHeight() As String
    RowHeight = CStr(ActiveSheet.StandardHeight)
End Function
'''


def changes(before: str, after: str) -> list[list[object]]:
    """What changed from one part to the other, tag by tag and text by text: each change's kind, what was there and what is."""
    one, two = (re.split(r'(?<=>)|(?=<)', text) for text in (before, after))
    return [[tag, one[i1:i2], two[j1:j2]] for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, one, two, autojunk=False).get_opcodes()
            if tag != 'equal']


def value_of(text: str) -> object:
    if text in ('True', 'False'):
        return text == 'True'
    if text.startswith('"'):
        return text.strip('"')
    return CONSTANTS[text]


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    kinds = tuple(sys.argv[1:]) or KINDS
    assert all(kind in KINDS for kind in kinds), kinds
    record = json.loads(OUT.read_text(encoding='utf-8')) if OUT.exists() else {'source': SOURCE, 'cases': []}
    kept = [one for one in record['cases'] if one['kind'] not in kinds]
    cases: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        height = excel.run_vba(SOURCE, 'RowHeight', (), timeout=60.0, line_numbers=False)
        assert height.ok and str(height.value) == '15', f'StandardHeight {height.value}: not recording geometry'
        folder = Path(held)
        for kind in kinds:
            book, saved = folder / 'in.xlsx', folder / f'base_{kind}.xlsx'
            book.write_bytes(BOOK.read_bytes())
            ran = excel.run_vba(SOURCE, 'Row', (str(book), str(saved), kind, 'Visible', True, False), timeout=60.0, line_numbers=False)
            assert ran.ok and not ran.dialogs, ran.outcome
            base = parts(saved)
            for member, values in SETTINGS.items():
                for text in values:
                    book.write_bytes(BOOK.read_bytes())
                    out = folder / f'out_{kind}_{member}_{len(cases)}.xlsx'
                    ran = excel.run_vba(SOURCE, 'Row', (str(book), str(out), kind, member, value_of(text), True), timeout=60.0, line_numbers=False)
                    answer = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
                    after = parts(out) if out.exists() else {}
                    changed = {name: changes(base.get(name, ''), part) for name, part in after.items() if base.get(name) != part}
                    cases.append({'kind': kind, 'member': member, 'value': text, 'excel': answer, 'parts': changed})
                    print(kind, member, text, answer, sorted(changed), flush=True)
        height = excel.run_vba(SOURCE, 'RowHeight', (), timeout=60.0, line_numbers=False)
        assert height.ok and str(height.value) == '15', f'StandardHeight {height.value} at the end: not recording'
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'cases': kept + cases}, indent=1) + '\n',
                   encoding='utf-8')


if __name__ == '__main__':
    main()
