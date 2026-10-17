"""Remeasure bounded Calculate-handler reentry and edits in native Excel."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/calculation_reentry.json'


def main() -> None:
    record = json.loads(OUT.read_text(encoding='utf-8'))
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
        record['excel_version_build'] = version.value
        for case in record['probes']:
            install = ['Sub Install()', 'Dim text As String']
            for owner, code in [('Worksheets(1).CodeName', case['sheet_handler']), ('"ThisWorkbook"', record['book_handler'])]:
                install += ['text = ""']
                install += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in code.splitlines()]
                install += [f'ActiveWorkbook.VBProject.VBComponents({owner}).CodeModule.AddFromString text']
            install += ['End Sub']
            excel.new_document()
            excel.add_module('EventLog', record['helpers'], line_numbers=False)
            installed = excel.run_vba('\n'.join(install), 'Install')
            assert installed.ok and not installed.dialogs, installed
            excel.add_module('Harness', case['source'], line_numbers=False)
            case['result'] = excel.run_raw('Harness.Probe')
            print(case['name'], case['result'], flush=True)
        assert not any(event.get('dialogs') for event in excel.events)
    record['measured_at'] = datetime.now(timezone.utc).isoformat()
    temporary = OUT.with_suffix('.json.tmp')
    temporary.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
    temporary.replace(OUT)


if __name__ == '__main__':
    main()
