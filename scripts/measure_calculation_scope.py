"""Remeasure calculation scope and deferred events through direct native macro calls."""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/calculation_scope.json'
SAVED = '''Function SavedValues() As String
SavedValues = CStr(Worksheets("Left").Range("A1").Value2) & ":" & CStr(Worksheets("Left").Range("B1").Value2) & ":" & CStr(Worksheets("Right").Range("B1").Value2) & ":" & CStr(Worksheets("Blank").Range("A1").Value2)
End Function
Sub Recalculate()
Application.Calculate
End Sub
Function CalculateTrace() As String
Application.EnableEvents = True
Log = ""
Application.Calculate
CalculateTrace = Log & "|" & SavedValues()
Application.EnableEvents = False
End Function
Sub SetSaveFlag()
Application.CalculateBeforeSave = False
End Sub
Sub SaveNative(ByVal target As String)
Application.DisplayAlerts = False
SetSaveFlag
ActiveWorkbook.SaveAs target, 52
End Sub
'''


def main() -> None:
    record = json.loads(OUT.read_text(encoding='utf-8'))
    selected = {arg for arg in sys.argv[1:] if not arg.startswith('--')}
    repeats = next((int(arg.split('=', 1)[1]) for arg in sys.argv[1:] if arg.startswith('--repeat=')), 1)
    assert repeats > 0
    install = ['Sub Install()', 'Application.EnableEvents = False', 'Dim text As String']
    for owner, code in [('"ThisWorkbook"', record['book_handler'])] + [
            (f'Worksheets("{name}").CodeName', record['sheet_handler']) for name in ['Left', 'Right', 'Blank']]:
        install += ['text = ""']
        install += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in code.splitlines()]
        install += [f'ActiveWorkbook.VBProject.VBComponents({owner}).CodeModule.AddFromString text']
    install += ['End Sub']
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
        record['excel_version_build'] = version.value
        for case in record['probes']:
            if '--files-only' in sys.argv:
                break
            if selected and case['name'] not in selected:
                continue
            if '--udfs' in sys.argv and not case['name'].startswith(('scope_stable_', 'scope_volatile_')):
                continue
            for repeat in range(repeats):
                excel.new_document()
                prepared = excel.run_vba(record['prepare'], 'Prepare')
                assert prepared.ok and not prepared.dialogs, prepared
                excel.add_module('EventLog', record['helpers'], line_numbers=False)
                installed = excel.run_vba('\n'.join(install), 'Install')
                assert installed.ok and not installed.dialogs, installed
                excel.add_module('Harness', case['macro'], line_numbers=False)
                result = excel.run_raw('Harness.Probe')
                assert not excel.events[-1].get('dialogs'), (case['name'], result)
                if repeats > 1 or 'results' in case:
                    case['results'] = list(dict.fromkeys(case.get('results', [case['result']]) + [result]))
                    case.setdefault('observations', []).append({'result': result, 'repeat': repeat,
                        'measured_at': datetime.now(timezone.utc).isoformat()})
                case['result'] = result
                print(case['name'], result, flush=True)
        if '--files-only' in sys.argv:
            folder = OUT.parent / 'calculation_scope_files'
            folder.mkdir(exist_ok=True)
            files = []
            for name in ['scope_ordinary_-4135_True_left', 'scope_stable_-4135_True_left',
                         'scope_forward_-4135_True_range_pair', 'notifications_-4135_entry_none']:
                case = next(case for case in record['probes'] if case['name'] == name)
                source = case['macro'].split('\nLog = ""\nApplication.Calculate\nProbe = Probe &', 1)[0]
                source += '\nApplication.EnableEvents = False\nEnd Function\n' + SAVED
                excel.new_document()
                prepared = excel.run_vba(record['prepare'], 'Prepare')
                assert prepared.ok and not prepared.dialogs, prepared
                excel.add_module('EventLog', record['helpers'], line_numbers=False)
                installed = excel.run_vba('\n'.join(install), 'Install')
                assert installed.ok and not installed.dialogs, installed
                excel.add_module('Harness', source, line_numbers=False)
                setup = excel.run_raw('Harness.Probe')
                before_save = excel.run_raw('Harness.SavedValues')
                path = folder / (name + '.xlsm')
                excel.run_raw('Harness.SaveNative', str(path))
                excel.open_document(path)
                before = excel.run_raw('Harness.SavedValues')
                after = excel.run_raw('Harness.CalculateTrace')
                files.append({'name': name, 'file': path.relative_to(OUT.parent).as_posix(),
                              'macro': source, 'setup': setup, 'before_save': before_save,
                              'before': before, 'after': after})
                print(name, before_save, before, after, flush=True)
            record['files'] = files
        assert not any(event.get('dialogs') for event in excel.events)
    record['measured_at'] = datetime.now(timezone.utc).isoformat()
    temporary = OUT.with_suffix('.json.tmp')
    temporary.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
    temporary.replace(OUT)


if __name__ == '__main__':
    main()
