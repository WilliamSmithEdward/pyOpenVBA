"""Replay Currency assignment precision against native Excel without instrumentation."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession


def main() -> None:
    path = Path(__file__).resolve().parents[1] / 'tests/fixtures/currency_writes.json'
    record = json.loads(path.read_text(encoding='utf-8'))
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\n'
                               'VersionBuild=Application.Version & "|" & Application.Build\nEnd Function',
                               'VersionBuild', line_numbers=False)
        assert version.ok and not version.dialogs, version
        record['excel_version_build'] = version.value
        for case in record['probes']:
            excel.new_document()
            result = excel.run_vba(case['source'], 'Probe', line_numbers=False)
            assert result.ok and result.value == case['result'] and not result.dialogs, (case['name'], result)
            print(case['name'], 'matched', flush=True)
    record['measured_at'] = datetime.now(timezone.utc).isoformat()
    path.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
