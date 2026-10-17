"""Replay native PivotTable page, field-label and field-data ranges."""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

from pyvbaharness import ExcelSession


def main() -> None:
    folder = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivots'
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\n'
                               'VersionBuild=Application.Version & "|" & Application.Build\nEnd Function',
                               'VersionBuild', line_numbers=False)
        assert version.ok and not version.dialogs, version
        for name in ('field_ranges.json', 'imported_field_ranges.json', 'retained_fields.json'):
            path = folder / name
            record = json.loads(path.read_text(encoding='utf-8'))
            for case in record['probes']:
                if 'file' in case:
                    excel.open_document(folder / case['file'])
                else:
                    excel.new_document()
                entrypoint = 'FieldRanges' if 'file' in case else 'Probe'
                for class_name, source in case.get('classes', {}).items():
                    excel.add_module(class_name, source, kind='class')
                result = excel.run_vba(case['source'], entrypoint, line_numbers=False)
                assert result.ok and not result.dialogs and result.value == case['result'], (case.get('name', case.get('file')), result)
                print(case.get('name', case.get('file')), 'matched', flush=True)
            record['excel_version_build'] = version.value
            record['measured_at'] = dt.datetime.now(dt.timezone.utc).isoformat()
            path.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
