"""Workbook.FullNameURLEncoded beside FullName: for an unsaved workbook, and for workbooks saved under names with
spaces, # and %, and letters beyond ASCII.

Each case makes a workbook, saves it under the case's name in a temporary folder unless the case is unsaved, and
reads both properties with the folder replaced by "<folder>", so the fixture holds no machine path.

    python scripts/measure_full_name_url_encoded.py

writes tests/fixtures/full_name_url_encoded.json, which tests/test_excel_full_name_url_encoded.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/full_name_url_encoded.json'
NAMES = ['', 'plain.xlsx', 'with space.xlsx', 'hash#and%percent.xlsx', 'caf' + chr(233) + '.xlsx', 'a&b+c.xlsx']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, s As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Select Case which']
    for index, name in enumerate(NAMES):
        lines.append(f'    Case {index}')
        if name:
            spelled = ''.join(ch if ord(ch) < 128 else f'" & ChrW({ord(ch)}) & "' for ch in name)
            lines.append(f'        wb.SaveAs folder & "\\{spelled}", xlOpenXMLWorkbook')
    lines += ['    End Select',
              '    s = Replace(wb.FullName, folder, "<folder>") & "|" & Replace(wb.FullNameURLEncoded, folder, '
              '"<folder>")',
              '    s = s & "|" & (wb.FullName = wb.FullNameURLEncoded)', '    wb.Close False', '    Tried = s',
              'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(NAMES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [list(pair) for pair in zip(NAMES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
    for name, answer in record['cases']:
        print(repr(name), '->', answer)


if __name__ == '__main__':
    main()
