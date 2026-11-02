"""More Application settings, read only: what each reads as Excel starts, and the read-only flags' errors when set to
what they read.

Nothing here changes the measuring user's profile. The writable settings are read and never set: some apply only once
Excel starts again, so a setting cannot be read back to check it was put back, and an earlier run that set
AlwaysUseClearType could not restore it. Their setters wait for a way to measure them that leaves the profile alone.
A read-only flag is set to the value it reads, which changes nothing either way.

    python scripts/measure_application_settings_more.py

writes tests/fixtures/application_settings_more.json, which tests/test_excel_application_settings_more.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/application_settings_more.json'
#: Writable settings, read as Excel starts.
WRITABLE = ['DisplayClipboardWindow', 'DisplayDocumentActionTaskPane', 'DisplayDocumentInformationPanel',
            'DisplayExcel4Menus', 'DisplayRecentFiles', 'DisplayXMLSourcePane', 'ShowDevTools', 'ShowStartupDialog',
            'UseClusterConnector', 'CommandUnderlines', 'DataEntryMode', 'DefaultSaveFormat']
READ_ONLY = ['CanPlaySounds', 'CanRecordSounds', 'MathCoprocessorAvailable', 'MouseAvailable', 'Ready',
             'ArbitraryXMLSupportAvailable', 'IsSandboxed', 'RecordRelative']


def _probe() -> str:
    lines = ['Private Function Fetch(ByVal name As String) As String',
             '    On Error GoTo Failed', '    Fetch = CStr(CallByName(Application, name, VbGet))', '    Exit Function',
             'Failed:', '    Fetch = "E" & Err.Number', 'End Function', '',
             'Private Function Same(ByVal name As String) As String',
             '    On Error GoTo Failed',
             '    CallByName Application, name, VbLet, CallByName(Application, name, VbGet)',
             '    Same = "ok"', '    Exit Function', 'Failed:', '    Same = "E" & Err.Number', 'End Function', '',
             'Public Function Probe() As String', '    Dim s As String']
    lines += [f'    s = s & "@@@start=" & Fetch("{name}")' for name in WRITABLE]
    lines += [f'    s = s & "@@@start=" & Fetch("{name}") & "|set:" & Same("{name}")' for name in READ_ONLY]
    lines += ['    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=120.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [list(pair) for pair in zip([*WRITABLE, *READ_ONLY], answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for name, answer in record['cases']:
        print(name, '->', answer)


if __name__ == '__main__':
    main()
