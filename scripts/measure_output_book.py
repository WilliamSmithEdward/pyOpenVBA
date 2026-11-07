"""The workbook tests/test_excel_output_hooks.py replays on, as Excel saves it with a chart sheet, and what PrintOut returns, bound late: rival rules are True always, or False when BeforePrint cancels. Bound
typed on a Workbook or Sheets, a return value does not compile.

Every PrintOut passes PrintToFile:=True under a temporary folder, so nothing goes to paper.

    python scripts/measure_output_book.py

writes tests/fixtures/output_members.xlsx and tests/fixtures/print_returns.json.
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
from fixture_workbook import strip_save_path  # noqa: E402
from measure_output_members import MAKER  # noqa: E402

BOOK = ROOT / 'tests/fixtures/output_members.xlsx'
OUT = ROOT / 'tests/fixtures/print_returns.json'
OWNERS = {'Worksheet': 'wb.Worksheets(1)', 'Workbook': 'wb', 'Range': 'wb.Worksheets(1).Range("A1:B3")', 'Chart': 'wb.Charts(1)', 'Sheets': 'wb.Sheets'}


def source(owner: str, cancel: bool) -> str:
    return f'''Public Function Row(ByVal path As String, ByVal folder As String) As String
    Dim wb As Workbook, o As Object, r As Variant, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    wb.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString "Private Sub Workbook_BeforePrint(Cancel As Boolean)" & vbLf & "Cancel = {cancel}" & vbLf & "End Sub"
    Application.EnableEvents = True
    Set o = {OWNERS[owner]}
    r = o.PrintOut(PrintToFile:=True, PrToFileName:=folder & "\\p.prn")
    Application.EnableEvents = False
    s = TypeName(r) & " " & CStr(r)
    wb.Close False
    Row = s
End Function
'''


def main() -> None:
    answers: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
        saved = Path(folder) / 'output_members.xlsx'
        with ExcelSession(HarnessConfig(exclusive=False)) as excel:
            excel.new_document()
            made = excel.run_vba(MAKER, 'Make', (str(saved),), timeout=60.0, line_numbers=False)
            assert made.ok, made
        BOOK.write_bytes(saved.read_bytes())
        strip_save_path(BOOK)
        for owner in OWNERS:
            for cancel in (False, True):
                # An Excel of its own for each: a file form left up by one would stop the next.
                with ExcelSession(HarnessConfig(exclusive=False)) as excel:
                    excel.new_document()
                    done = excel.run_vba(source(owner, cancel), 'Row', (str(saved), folder), timeout=45.0, line_numbers=False)
                    answer = str(done.value) if done.ok else f'failed {done.outcome} {[one.message[:80] for one in done.dialogs]}'
                answers.append({'owner': owner, 'cancel': cancel, 'answer': answer})
                print(owner, cancel, answer, flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': answers}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
