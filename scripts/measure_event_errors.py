"""Remeasure host and custom event errors without injected error handlers."""
from __future__ import annotations

import json
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/event_errors.json'


def main() -> None:
    record = json.loads(OUT.read_text(encoding='utf-8'))
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild', line_numbers=False)
        assert version.ok and not version.dialogs, version
        record['excel_version_build'] = version.value
        for case in record['probes']:
            install = ['Sub Install()', 'Dim text As String, part As Object']
            modules = [(None, module['name'], module['source']) for module in case.get('modules', [])]
            if 'sink' in case:
                modules += [(None, 'ExcelSink', case['sink']), ('Worksheets(1).CodeName', '', case['sheet_handler']),
                            ('"ThisWorkbook"', '', case['book_handler'])]
            for owner, name, code in modules:
                install += ['text = ""']
                if owner is None:
                    install += ['Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)', f'part.Name = "{name}"']
                else:
                    install += [f'Set part = ActiveWorkbook.VBProject.VBComponents({owner})']
                install += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in code.splitlines()]
                install += ['part.CodeModule.AddFromString text']
            install += ['End Sub']
            excel.new_document()
            excel.add_module('EventLog', record['helpers'], line_numbers=False)
            installed = excel.run_vba('\n'.join(install), 'Install', line_numbers=False)
            assert installed.ok and not installed.dialogs, installed
            result = excel.run_vba(case['source'], 'Probe', timeout=20.0, line_numbers=False)
            if case['unhandled']:
                # Excel's End button is the harness's defined response to this
                # expected unhandled-error dialog. Persistent cells retain the
                # trace showing that the caller and later listeners did not run.
                assert result.outcome == 'runner-error' and len(result.dialogs) == 1, result
                dialog = result.dialogs[0]
                assert dialog.classification == 'runtime-error' and "'9'" in dialog.message, result
            else:
                assert result.ok and not result.dialogs, result
            read = excel.run_vba(record['read'], 'ReadState', line_numbers=False)
            assert read.ok and not read.dialogs, read
            case['result'] = read.value
            case['native_run'] = asdict(result)
            print(case['name'], read.value, flush=True)
    record['measured_at'] = datetime.now(timezone.utc).isoformat()
    record['instrumentation'] = 'none'
    temporary = OUT.with_suffix('.json.tmp')
    temporary.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
    temporary.replace(OUT)


if __name__ == '__main__':
    main()
