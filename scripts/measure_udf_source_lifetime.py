"""Remeasure the committed UDF source-edit inputs through separate native macro calls."""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/udf_source_lifetime.json'
READ_SAVED = '''Function ReadSaved() As String
Dim cell As Range
For Each cell In Worksheets("Target").Range("B1:C1")
ReadSaved = ReadSaved & "|" & CStr(cell.Value2) & ":" & cell.Text & ":" & cell.Formula
Next cell
End Function
Sub CalculateNow()
Application.Calculate
End Sub
Sub SetSaveFlag()
Application.CalculateBeforeSave = False
End Sub
'''


def save_native_files(excel: ExcelSession, record: dict) -> None:
    folder = OUT.parent / 'udf_source_files'
    folder.mkdir(exist_ok=True)
    files = []
    for name in ['after_replace_manual_input', 'after_insert_manual_input', 'after_remove_manual_input',
                 'after_replace_auto_calculate', 'after_insert_auto_calculate', 'after_remove_auto_calculate']:
        case = next(case for case in record['probes'] if case['name'] == name)
        source = case['macro'] + '\n' + READ_SAVED
        excel.new_document()
        excel.add_module('Harness', source, line_numbers=False)
        for step in case['steps']:
            result = excel.run_raw('Harness.' + step['entry'])
            assert not excel.events[-1].get('dialogs') and result == step['result'], result
        path = folder / (name + '.xlsm')
        save = ('Sub SaveFixture()\nApplication.DisplayAlerts = False\nApplication.CalculateBeforeSave = False\n'
                'ActiveWorkbook.SaveAs "' + str(path).replace('"', '""') + '", 52\nEnd Sub')
        result = excel.run_vba(save, 'SaveFixture')
        assert result.ok and not result.dialogs, result
        excel.open_document(path)
        read = excel.run_raw('Harness.ReadSaved')
        excel.run_raw('Harness.CalculateNow')
        after = excel.run_raw('Harness.ReadSaved')
        assert not any(event.get('dialogs') for event in excel.events), excel.events[-1]
        files.append({'name': name, 'file': path.relative_to(OUT.parent).as_posix(),
                      'result': read, 'calculated': after,
                      'calculation': -4135 if '_manual_' in name else -4105})
    record['files'] = files
    record['read_saved'] = READ_SAVED


def main() -> None:
    record = json.loads(OUT.read_text(encoding='utf-8'))
    selected = {name for name in sys.argv[1:] if not name.startswith('--')}
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
        if selected:
            assert version.value == record.get('excel_version_build'), version
        record['excel_version_build'] = version.value
        for case in record['probes']:
            if '--files-only' in sys.argv:
                break
            if selected and case['name'] not in selected:
                continue
            excel.new_document()
            excel.add_module('Harness', case['macro'], line_numbers=False)
            for step in case['steps']:
                result = excel.run_raw('Harness.' + step['entry'], timeout=30.0)
                assert not excel.events[-1].get('dialogs'), (case['name'], step['entry'], result)
                step['result'] = result
            print(case['name'], [step['result'] for step in case['steps']], flush=True)
        save_native_files(excel, record)
        record['measured_at'] = datetime.now(timezone.utc).isoformat()
    temporary = OUT.with_suffix('.json.tmp')
    temporary.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
    temporary.replace(OUT)


if __name__ == '__main__':
    main()
