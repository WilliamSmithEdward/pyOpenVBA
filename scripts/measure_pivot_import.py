"""Re-measure imported PivotTable controls using an isolated native Excel session.

The seed workbook was authored by Excel with Region/Amount source columns,
five records (East 10, West 20, East 30, West 40, East 50), a Region row
field and an Amount sum data field named Total, placed at E3. No VBA is
stored in the seed. The explicit handlers below measure runtime errors;
the harness does not inject handlers or line numbers.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession


def main() -> None:
    root = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivots'
    path = root / 'import.json'
    record = json.loads(path.read_text(encoding='utf-8'))
    with ExcelSession() as excel:
        excel.open_document(root / 'native-simple.xlsx')
        version = excel.run_vba('Function VersionBuild() As String\n'
                               'VersionBuild=Application.Version & "|" & Application.Build\nEnd Function',
                               'VersionBuild', line_numbers=False)
        assert version.ok and not version.dialogs, version
        record['excel_version_build'] = version.value
        for case in record['probes']:
            name = case['source'].split('(')[0].split()[-1]
            result = excel.run_vba(case['source'], name, line_numbers=False)
            assert result.ok and not result.dialogs, (case['expression'], result)
            expected = case['result'].replace('pivot-preview.xlsx', 'native-simple.xlsx')
            assert result.value == expected, (case['expression'], result.value, expected)
            case['result'] = result.value
            print(case['expression'], result.value, flush=True)
        record['measured_at'] = datetime.now(timezone.utc).isoformat()
    path.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
    axes_path = root / 'axes.json'
    variants = json.loads(axes_path.read_text(encoding='utf-8'))
    for variant in variants:
        with ExcelSession() as excel:
            excel.open_document(root / variant['file'])
            for case in variant['probes']:
                name = case['source'].split('(')[0].split()[-1]
                result = excel.run_vba(case['source'], name, line_numbers=False)
                assert result.ok and not result.dialogs, (variant['file'], case['expression'], result)
                assert result.value == case['result'], (variant['file'], case['expression'], result.value, case['result'])
            variant['measured_at'] = datetime.now(timezone.utc).isoformat()
            print(variant['file'], len(variant['probes']), 'controls matched', flush=True)
    axes_path.write_text(json.dumps(variants, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
