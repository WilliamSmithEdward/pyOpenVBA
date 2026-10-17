"""Verify GETPIVOTDATA controls in native Excel through Evaluate and cells.

Golden reports were created by Excel; advanced setup macros are retained in
the fixture record. This runner reopens those saved files, does not refresh
their reports, and verifies both evaluation surfaces without instrumentation.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession


def main() -> None:
    root = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivots'
    path = root / 'pivotdata.json'
    record = json.loads(path.read_text(encoding='utf-8'))
    for filename in dict.fromkeys(case['file'] for case in record['probes']):
        with ExcelSession() as excel:
            excel.open_document(root / filename)
            version = excel.run_vba('Function VersionBuild() As String\n'
                                   'VersionBuild=Application.Version & "|" & Application.Build\nEnd Function',
                                   'VersionBuild', line_numbers=False)
            assert version.ok and not version.dialogs, version
            record['excel_version_build'] = version.value
            for i, case in enumerate(case for case in record['probes'] if case['file'] == filename):
                name = case['source'].split('(')[0].split()[-1]
                result = excel.run_vba(case['source'], name, line_numbers=False)
                assert result.ok and not result.dialogs, (filename, case['formula'], result)
                assert result.value == case['result'], (filename, case['formula'], result.value, case['result'])
                address = 'Z20' if filename.startswith('native-advanced-') else 'J20'
                escaped = case['formula'].replace('"', '""')
                source = f'Function CellProbe{i}() As String\nRange("{address}").Formula="{escaped}"\n'
                source += f'CellProbe{i}=CStr(Range("{address}").Value2)\nEnd Function'
                result = excel.run_vba(source, f'CellProbe{i}', line_numbers=False)
                assert result.ok and not result.dialogs, (filename, case['formula'], result)
                assert result.value == case['cell_result'], (filename, case['formula'], result.value, case['cell_result'])
            print(filename, 'all Evaluate and cell controls matched', flush=True)
    record['measured_at'] = datetime.now(timezone.utc).isoformat()
    path.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
    for name in ('pivotdata-arrays.json', 'pivotdata-formula-text.json'):
        controls = json.loads((root / name).read_text(encoding='utf-8'))
        with ExcelSession() as excel:
            excel.open_document(root / controls['file'])
            for case in controls['probes']:
                procedure = case['source'].split('(')[0].split()[-1]
                result = excel.run_vba(case['source'], procedure, line_numbers=False)
                assert result.ok and not result.dialogs, (case['formula'], result)
                assert result.value == case['result'], (case['formula'], result.value, case['result'])
        print(name, len(controls['probes']), 'controls matched', flush=True)


if __name__ == '__main__':
    main()
