"""Collections a new workbook holds nothing in: each one's TypeName, Count, Item(1) and Item("x") as a value or "E"
and the error number, Parent's TypeName and Application's TypeName, on a new workbook of one sheet.

    python scripts/measure_empty_collections.py

writes tests/fixtures/empty_collections.json, which tests/test_excel_empty_collections.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/empty_collections.json'
COLLECTIONS = ['ws.QueryTables', 'ws.OLEObjects', 'ws.NamedSheetViews', 'wb.XmlMaps', 'wb.XmlNamespaces',
               'wb.CustomViews', 'wb.PublishObjects', 'wb.SlicerCaches', 'wb.Charts', 'wb.Excel4MacroSheets',
               'wb.Excel4IntlMacroSheets', 'Application.Charts']
READS = ['TypeName({c})', '{c}.Count', 'TypeName({c}.Item(1))', 'TypeName({c}.Item("x"))', 'TypeName({c}.Parent)',
         'TypeName({c}.Application)', 'TypeName({c}(1))']


def _probe() -> str:
    lines = ['Public Function Probe() As String', '    Dim wb As Workbook, ws As Worksheet, s As String, t As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)']
    for collection in COLLECTIONS:
        lines.append('    s = s & "@@@"')
        for read in READS:
            expression = read.format(c=collection)
            lines += ['    t = ""', '    On Error Resume Next', f'    t = CStr({expression})',
                      '    If Err.Number Then t = "E" & Err.Number', '    Err.Clear', '    On Error GoTo 0',
                      '    s = s & "|" & t']
    lines += ['    wb.Close False', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=120.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'reads': READS,
              'cases': [list(pair) for pair in zip(COLLECTIONS, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for collection, answer in record['cases']:
        print(collection, '->', answer)


if __name__ == '__main__':
    main()
