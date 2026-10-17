"""Reproduce independent native conditional-formatting API controls."""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

from pyvbaharness import ExcelSession


def main() -> None:
    directory = Path(__file__).resolve().parents[1] / 'tests/fixtures/conditional'
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\n'
                               'VersionBuild=Application.Version & "|" & Application.Build\nEnd Function',
                               'VersionBuild', line_numbers=False)
        assert version.ok and not version.dialogs, version
        for path in sorted(directory.glob('*.json')):
            record = json.loads(path.read_text(encoding='utf-8'))
            for case in record['probes']:
                excel.new_document()
                result = excel.run_vba(case['source'], 'Probe', line_numbers=False)
                assert result.ok and not result.dialogs and result.value == case['result'], (case['name'], result)
                print(path.stem, case['name'], 'matched', flush=True)
            record['excel_version_build'] = version.value
            record['measured_at'] = dt.datetime.now(dt.timezone.utc).isoformat()
            path.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
