"""Reproduce native controls; ChrW keeps non-ASCII text out of VBE source conversion."""
from __future__ import annotations

import datetime as dt
import json
from argparse import ArgumentParser
from pathlib import Path

from pyvbaharness import ExcelSession


def main() -> None:
    directory = Path(__file__).resolve().parents[1] / 'tests/fixtures/conditional'
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--fixture', action='append', choices=sorted(path.name for path in directory.glob('*.json')),
                        help='Replay one fixture; repeat to select several. The default replays all fixtures.')
    options = parser.parse_args()
    paths = [directory / name for name in options.fixture] if options.fixture else sorted(directory.glob('*.json'))
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\n'
                               'VersionBuild=Application.Version & "|" & Application.Build\nEnd Function',
                               'VersionBuild', line_numbers=False)
        assert version.ok and not version.dialogs, version
        for path in paths:
            record = json.loads(path.read_text(encoding='utf-8'))
            for case in record['probes']:
                if case['result'] == 'NATIVE-OPEN-ERROR':
                    print(path.stem, case['name'], 'recorded native-open limitation; not replayed', flush=True)
                    continue
                assert case['source'].isascii(), (case['name'], 'Use ChrW UTF-16 units for non-ASCII VBA strings')
                if case.get('phase') == 'after-native-save-reopen':
                    excel.open_document(directory / case['file'])
                else:
                    excel.new_document()
                result = excel.run_vba(case['source'], 'Probe', line_numbers=False)
                assert result.ok and not result.dialogs and result.value == case['result'], (case['name'], result)
                print(path.stem, case['name'], 'matched', flush=True)
            record['excel_version_build'] = version.value
            record['measured_at'] = dt.datetime.now(dt.timezone.utc).isoformat()
            path.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
