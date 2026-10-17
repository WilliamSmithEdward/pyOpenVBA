"""Remeasure syntax and normal calls beside invalid unused VBA bodies."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession

ROOT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_semantics'


def main() -> None:
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild', line_numbers=False)
        assert version.ok and not version.dialogs, version
    for name in ('body_demand', 'ampersand_syntax'):
        path = ROOT / (name + '.json')
        record = json.loads(path.read_text(encoding='utf-8'))
        record['excel_version_build'] = version.value
        for case in record['probes']:
            with ExcelSession() as excel:
                excel.new_document()
                excel.add_module('Harness', case.get('input_source', case['source']), line_numbers=False)
                if name == 'body_demand':
                    assert excel.run_raw('Harness.Probe', timeout=20) == case['result'], case['name']
                    if 'selected_result' in case:
                        assert excel.run_raw('Harness.Unused', timeout=20) == case['selected_result'], case['name']
                else:
                    compiled = excel.compile_project(watch_seconds=20)
                    assert compiled.outcome in ('accepted', 'rejected'), (case['name'], compiled)
                    assert compiled.ok != bool(case.get('compile_error')), (case['name'], compiled)
                    if compiled.ok:
                        assert excel.run_raw('Harness.Probe', timeout=20) == case['result'], case['name']
                    else:
                        assert compiled.dialog is not None
                        case['compiler_message'] = compiled.dialog.message
            if name == 'body_demand':
                # A separate session avoids conflating a normal run with an
                # explicit whole-project compilation of its unused routine.
                with ExcelSession() as excel:
                    excel.new_document()
                    excel.add_module('Harness', case.get('input_source', case['source']), line_numbers=False)
                    compiled = excel.compile_project(watch_seconds=20)
                    expected = 'accepted' if 'selected_result' in case else 'rejected'
                    assert compiled.outcome == expected, (case['name'], compiled)
                    if compiled.dialog is not None:
                        case['compiler_message'] = compiled.dialog.message
            case['measured_at'] = datetime.now(timezone.utc).isoformat()
            print(name, case['name'], case.get('result', 'compile-error'), flush=True)
        record['measured_at'] = datetime.now(timezone.utc).isoformat()
        temporary = path.with_suffix('.json.tmp')
        temporary.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
        temporary.replace(path)


if __name__ == '__main__':
    main()
