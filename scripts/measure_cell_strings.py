"""Replay native text/formula persistence and UTF-16 unit controls."""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

from pyvbaharness import ExcelSession


def main() -> None:
    folder = Path(__file__).resolve().parents[1] / 'tests/fixtures/cell_strings'
    path = folder / 'native.xlsx'
    fixture = folder / 'controls.json'
    record = json.loads(fixture.read_text(encoding='utf-8'))
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\n'
                               'VersionBuild=Application.Version & "|" & Application.Build\nEnd Function',
                               'VersionBuild', line_numbers=False)
        assert version.ok and not version.dialogs, version
        result = excel.run_vba(record['source'], 'Probe', line_numbers=False)
        assert result.ok and not result.dialogs and result.value == record['result'], result
        saved = excel.run_vba('Sub SaveNative()\nThisWorkbook.SaveAs "' + str(path) + '",xlOpenXMLWorkbook\nEnd Sub',
                             'SaveNative', line_numbers=False)
        assert saved.ok and not saved.dialogs, saved
        excel.open_document(path)
        opened = excel.run_vba(record['inspection_source'], 'Snapshot', line_numbers=False)
        assert opened.ok and not opened.dialogs and opened.value == record['result'], opened
        for case in record['unicode_probes']:
            excel.open_document(path)
            checked = excel.run_vba(case['source'], 'Probe', line_numbers=False)
            assert checked.ok and not checked.dialogs and checked.value == case['result'], (case['name'], checked)
        record['excel_version_build'] = version.value
    record['measured_at'] = dt.datetime.now(dt.timezone.utc).isoformat()
    fixture.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
