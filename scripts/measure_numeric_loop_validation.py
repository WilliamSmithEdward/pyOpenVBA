"""Measure native compiler acceptance of numeric For storage targets."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession


def main() -> None:
    path = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_semantics/numeric_loop_validation.json'
    record = json.loads(path.read_text(encoding='utf-8'))
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild=Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild', line_numbers=False)
        assert version.ok and not version.dialogs, version
        record['excel_version_build'] = version.value
        for case in record['probes']:
            excel.new_document()
            for module in case.get('modules', []):
                excel.add_module(module['name'], module['source'], kind=module['kind'], line_numbers=False)
            excel.add_module('Harness', case['source'], line_numbers=False)
            if 'demand_result' in case:
                assert excel.run_raw('Harness.Probe', timeout=20) == case['demand_result'], case['name']
                # Compile the invalid routine in a fresh project after the
                # independent normal-call demand control.
                excel.new_document()
                excel.add_module('Harness', case['source'], line_numbers=False)
            compiled = excel.compile_project(watch_seconds=20)
            assert compiled.outcome in ('accepted', 'rejected'), (case['name'], compiled)
            if 'result' in case or 'compile_error' in case:
                assert compiled.ok != bool(case.get('compile_error')), (case['name'], compiled)
            if compiled.ok:
                actual = excel.run_raw('Harness.Probe', timeout=20)
                if 'result' in case:
                    assert actual == case['result'], (case['name'], actual, case['result'])
                case['result'] = actual
            else:
                assert compiled.dialog is not None
                case.update(compile_error=True, compiler_message=compiled.dialog.message)
            case['measured_at'] = datetime.now(timezone.utc).isoformat()
            print(case['name'], case.get('result', 'compile-error'), flush=True)
            record['measured_at'] = case['measured_at']
            temporary = path.with_suffix('.json.tmp')
            temporary.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
            temporary.replace(path)


if __name__ == '__main__':
    main()
