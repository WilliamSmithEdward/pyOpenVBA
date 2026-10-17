"""Recheck VBA GetPivotData Range results and errors in isolated native Excel."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession


def main() -> None:
    root = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivots'
    path = root / 'getdata.json'
    record = json.loads(path.read_text(encoding='utf-8'))
    for filename in dict.fromkeys(case['file'] for case in record['probes']):
        with ExcelSession() as excel:
            excel.open_document(root / filename)
            version = excel.run_vba('Function VersionBuild() As String\n'
                                   'VersionBuild=Application.Version & "|" & Application.Build\nEnd Function',
                                   'VersionBuild', line_numbers=False)
            assert version.ok and not version.dialogs, version
            record['excel_version_build'] = version.value
            for case in (case for case in record['probes'] if case['file'] == filename):
                procedure = case['source'].split('(')[0].split()[-1]
                result = excel.run_vba(case['source'], procedure, line_numbers=False)
                assert result.ok and not result.dialogs, (filename, case['arguments'], result)
                assert result.value == case['result'], (filename, case['arguments'], result.value, case['result'])
            print(filename, 'all method controls matched', flush=True)
    record['measured_at'] = datetime.now(timezone.utc).isoformat()
    path.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
