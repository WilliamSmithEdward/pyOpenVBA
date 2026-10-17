"""Remeasure native DefType defaults and explicit compiler outcomes."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_semantics/default_types.json'


def main() -> None:
    record = json.loads(OUT.read_text(encoding='utf-8'))
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild', line_numbers=False)
        assert version.ok and not version.dialogs, version
        record['excel_version_build'] = version.value
        for case in record['probes']:
            excel.new_document()
            for module in case.get('modules', []):
                excel.add_module(module['name'], module['source'], kind=module['kind'], line_numbers=False)
            excel.add_module('Harness', case['source'], line_numbers=False)
            if case.get('compile_mode') == 'on-demand':
                assert excel.run_raw('Harness.Probe', timeout=20) == case['result'], case['name']
                case['measured_at'] = datetime.now(timezone.utc).isoformat()
                print(case['name'], case['result'], flush=True)
                continue
            compiled = excel.compile_project(watch_seconds=20)
            assert compiled.outcome in ('accepted', 'rejected'), (case['name'], compiled)
            assert compiled.ok != bool(case.get('compile_error')), (case['name'], compiled)
            case['compiler_signal'] = compiled.signal
            if case.get('compile_error'):
                assert compiled.dialog is not None, (case['name'], compiled)
                case['compiler_message'] = compiled.dialog.message
            else:
                result = excel.run_macro('Harness.Probe', timeout=20)
                assert result.ok and all(dialog.action == 'note:vbe-window-visible' for dialog in result.dialogs), (case['name'], result)
                assert result.value == case['result'], (case['name'], case['result'], result.value)
                case['result'] = result.value
            case['measured_at'] = datetime.now(timezone.utc).isoformat()
            print(case['name'], case.get('result', 'compile-error'), flush=True)
    record['measured_at'] = datetime.now(timezone.utc).isoformat()
    temporary = OUT.with_suffix('.json.tmp')
    temporary.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
    temporary.replace(OUT)


if __name__ == '__main__':
    main()
