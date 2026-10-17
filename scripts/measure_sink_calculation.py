"""Remeasure cached values and events when attaching Excel WithEvents listeners."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/sink_calculation.json'


def main() -> None:
    record = json.loads(OUT.read_text(encoding='utf-8'))
    install = ['Sub Install()', 'Dim part As Object, text As String',
               'Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)', 'part.Name = "ExcelSink"', 'text = ""']
    install += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in record['sink'].splitlines()]
    install += ['part.CodeModule.AddFromString text', 'End Sub']
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
        record['excel_version_build'] = version.value
        for case in record['probes']:
            excel.new_document()
            excel.add_module('EventLog', record['helpers'], line_numbers=False)
            result = excel.run_vba('\n'.join(install), 'Install')
            assert result.ok and not result.dialogs, result
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
