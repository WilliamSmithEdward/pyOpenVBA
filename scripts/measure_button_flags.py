"""What a button's AddIndent, DefaultButton, CancelButton, DismissButton and HelpButton read straight after a set, on the first button of the
workbook Excel saved for tests/fixtures/old_collection_members.json and on a new one, late bound, and after a save and a reopen.

Rival rules, written before measuring: the set holds in the session and only the file loses it, or it never reads back. Found: the four
button roles read True in the session and False after a reopen, as Excel reads none back from the file; AddIndent reads False throughout.

    python scripts/measure_button_flags.py

writes tests/fixtures/button_flags.json, which tests/test_excel_loaded_control_edits.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/button_flags.json'
BOOK = ROOT / 'tests/fixtures/old_collections.xlsx'
MEMBERS = ('AddIndent', 'DefaultButton', 'CancelButton', 'DismissButton', 'HelpButton')
SOURCE = r'''Public Function Row(ByVal path As String, ByVal saved As String, ByVal member As String, ByVal fresh As Boolean) As String
    Dim wb As Workbook, o As Object, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    If fresh Then
        Set o = wb.Worksheets(1).Buttons.Add(300, 300, 60, 30)
    Else
        Set o = wb.Worksheets(1).Buttons(1)
    End If
    s = "before " & CallByName(o, member, VbGet)
    On Error Resume Next
    CallByName o, member, VbLet, True
    s = s & "|set E" & Err.Number
    Err.Clear
    s = s & "|after " & CallByName(o, member, VbGet) & "|first " & CallByName(wb.Worksheets(1).Buttons(1), member, VbGet)
    On Error GoTo 0
    wb.SaveAs saved, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(saved, 0)
    s = s & "|reopened " & CallByName(wb.Worksheets(1).Buttons(wb.Worksheets(1).Buttons.Count), member, VbGet)
    wb.Close False
    Row = s
End Function
'''


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        for fresh in (False, True):
            for member in MEMBERS:
                book, saved = Path(held) / 'in.xlsx', Path(held) / f'out_{member}_{fresh}.xlsx'
                book.write_bytes(BOOK.read_bytes())
                ran = excel.run_vba(SOURCE, 'Row', (str(book), str(saved), member, fresh), timeout=60.0, line_numbers=False)
                answer = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
                cases.append({'member': member, 'new': fresh, 'excel': answer})
                print('new' if fresh else 'loaded', member, answer, flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'cases': cases}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
