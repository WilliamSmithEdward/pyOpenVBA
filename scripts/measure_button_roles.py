"""Whether a button's role (DefaultButton, CancelButton, HelpButton) clears the same role on the sheet's other buttons, on the workbook Excel
saved for tests/fixtures/old_collection_members.json, typed and late, with no save: tests/fixtures/old_collection_lets.json shows a set on
the whole Buttons collection leaving the role on the last button alone.

Rival rules, written before measuring:
- exclusive: setting the role on a button clears it on the sheet's other buttons, so item 1 then item 2 reads False;True, and item 2 then
  item 1 reads True;False;
- last only: a button keeps its role whatever the others do (True;True both ways), and the collection set alone reaches the last item.
Setting False on the holder, and the role across two different roles (Default on item 1, Cancel on item 2), are read too.

Safety (the Supervisor's rules): one Excel through live.lock (an exclusive session), no save, no window touched, nothing geometric.

    python scripts/measure_button_roles.py

writes tests/fixtures/button_roles.json.
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/button_roles.json'
BOOK = ROOT / 'tests/fixtures/old_collections.xlsx'
ROLES = ('DefaultButton', 'CancelButton', 'HelpButton')
#: Each case: the sets in order, as (item, role, value).
STEPS: dict[str, tuple[tuple[int, str, bool], ...]] = {
    'one then two': ((1, '{r}', True), (2, '{r}', True)),
    'two then one': ((2, '{r}', True), (1, '{r}', True)),
    'one then off': ((1, '{r}', True), (1, '{r}', False)),
    'one alone': ((1, '{r}', True),),
}
#: Across roles: does one role clear another?
ACROSS: dict[str, tuple[tuple[int, str, bool], ...]] = {
    'default one, cancel two': ((1, 'DefaultButton', True), (2, 'CancelButton', True)),
    'default one, cancel one': ((1, 'DefaultButton', True), (1, 'CancelButton', True)),
}
DRIVER = r'''Public Function Row(ByVal path As String) As String
    Dim wb As Workbook, ws As Worksheet, b As {declared}, e As String, s As String, i As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0, True)
    Set ws = wb.Worksheets(1)
    On Error Resume Next
{steps}
    For i = 1 To ws.Buttons.Count
        Set b = ws.Buttons(i)
        s = s & ";" & b.DefaultButton & "," & b.CancelButton & "," & b.HelpButton
    Next
    On Error GoTo 0
    wb.Close False
    Row = Mid(e, 2) & "|" & Mid(s, 2)
End Function
'''
STEP = '''    Set b = ws.Buttons({item})
    Err.Clear
    b.{role} = {value}
    e = e & ";" & Err.Number'''


def source(steps: tuple[tuple[int, str, bool], ...], role: str, binding: str) -> str:
    lines = [STEP.format(item=item, role=name.replace('{r}', role), value=value) for item, name, value in steps]
    return DRIVER.replace('{declared}', 'Excel.Button' if binding == 'Typed' else 'Object').replace('{steps}', '\n'.join(lines))


def cases() -> list[dict[str, object]]:
    found: list[dict[str, object]] = []
    for binding in ('Typed', 'Late'):
        for role in ROLES:
            for name, steps in STEPS.items():
                found.append({'case': name, 'role': role, 'binding': binding, 'source': source(steps, role, binding)})
        for name, steps in ACROSS.items():
            found.append({'case': name, 'role': '', 'binding': binding, 'source': source(steps, '', binding)})
    return found


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    found = cases()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        book = Path(held) / 'book.xlsx'
        for case in found:
            book.write_bytes(BOOK.read_bytes())
            ran = excel.run_vba(str(case['source']), 'Row', (str(book),), timeout=60.0, line_numbers=False)
            case['excel'] = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
            print(case['binding'], case['role'], case['case'], case['excel'], flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': found}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    sys.exit(main())
