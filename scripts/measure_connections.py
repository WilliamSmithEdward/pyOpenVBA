"""Workbook.Connections on a new workbook and on a copy of tests/fixtures/power_query/loaded_to_sheet.xlsx, whose
Power Query connection is an OLE DB one: the count, the lookups, and what the connection and its OLEDBConnection read.

    python scripts/measure_connections.py

writes tests/fixtures/connections.json, which tests/test_excel_connections.py replays.
"""
from __future__ import annotations

import json
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/connections.json'
FIXTURE = 'power_query/loaded_to_sheet.xlsx'
READS = ['c.Name', 'c.Description', 'c.Type', 'TypeName(c.OLEDBConnection)', 'c.OLEDBConnection.Connection',
         'c.OLEDBConnection.CommandText', 'c.OLEDBConnection.CommandType', 'c.OLEDBConnection.BackgroundQuery',
         'c.OLEDBConnection.RefreshOnFileOpen', 'c.OLEDBConnection.EnableRefresh', 'c.RefreshWithRefreshAll',
         'c.InModel', 'TypeName(c.Parent)', 'c.OLEDBConnection.SavePassword', 'c.OLEDBConnection.MaintainConnection',
         'c.OLEDBConnection.RefreshPeriod', 'c.OLEDBConnection.IsConnected', 'TypeName(c.ODBCConnection)',
         'TypeName(c.ModelConnection)', 'TypeName(c.TextConnection)', 'TypeName(c.WorksheetDataConnection)',
         'c.OLEDBConnection.SourceConnectionFile', 'c.OLEDBConnection.ServerCredentialsMethod',
         'c.OLEDBConnection.AlwaysUseConnectionFile', 'c.OLEDBConnection.LocalConnection',
         'c.OLEDBConnection.UseLocalConnection', 'TypeName(c.OLEDBConnection.Parent)',
         'wb.Connections(wb.Connections(1).Name).Name', 'wb.Connections("CONNECTION").Name',
         'TypeName(wb.Connections(0))', 'TypeName(wb.Connections(9))', 'TypeName(wb.Connections("zz"))',
         'TypeName(wb.Connections.Parent)', 'wb.Connections.Count']


def _probe() -> str:
    lines = ['Private Function Read(ByVal wb As Object, ByVal c As Object, ByVal which As Long) As String',
             '    Dim v As Variant', '    On Error GoTo Failed', '    Select Case which']
    for index, read in enumerate(READS):
        lines += [f'    Case {index}', f'        v = {read}']
    lines += ['    End Select', '    Read = TypeName(v) & ":" & CStr(v)', '    Exit Function', 'Failed:',
              '    Read = "E" & Err.Number', 'End Function', '',
              'Public Function Probe(ByVal path As String) As String', '    Dim wb As Workbook, s As String, i As Long',
              '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
              '    s = wb.Connections.Count', '    wb.Close False',
              '    Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True)',
              f'    For i = 0 To {len(READS) - 1}', '        s = s & "~" & Read(wb, wb.Connections(1), i)', '    Next',
              '    wb.Close False', '    Application.DisplayAlerts = True', '    Probe = s', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        # Excel reads a copy, so that nothing it does touches the fixture.
        copy = Path(folder) / 'book.xlsx'
        shutil.copy(ROOT / 'tests/fixtures' / FIXTURE, copy)
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (str(copy),), timeout=300.0)
        assert result.ok and not result.dialogs, result
    new, *answers = str(result.value).split('~')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'file': FIXTURE, 'source': source,
              'new_count': new, 'reads': [[read, answer] for read, answer in zip(READS, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    print('new', new)
    for read, answer in record['reads']:
        print(read, '->', answer)


if __name__ == '__main__':
    main()
