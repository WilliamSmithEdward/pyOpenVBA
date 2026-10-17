"""Remeasure calculation and formula-entry inputs, including native XLSM persistence."""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/full_calculation.json'
READ_SAVED = '''Function ReadSaved() As String
Dim cell As Range
For Each cell In Worksheets("Target").Range("B1:C1")
ReadSaved = ReadSaved & "|" & CStr(cell.Value2) & ":" & cell.Text & ":" & cell.Formula
Next cell
End Function
Sub FullNow()
Application.CalculateFull
End Sub
Sub RebuildNow()
Application.CalculateFullRebuild
End Sub
Sub SetSaveFlag()
Application.CalculateBeforeSave = False
Worksheets("Target").Range("A1").Value = 9
End Sub
'''


def main() -> None:
    record = json.loads(OUT.read_text(encoding='utf-8'))
    selected = {name for name in sys.argv[1:] if not name.startswith('--')}
    repeats = next((int(arg.split('=', 1)[1]) for arg in sys.argv[1:] if arg.startswith('--repeat=')), 1)
    if repeats < 1:
        raise ValueError('--repeat must be positive')
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
        if selected:
            assert version.value == record['excel_version_build']
        record['excel_version_build'] = version.value
        for case in record['probes']:
            if '--files-only' in sys.argv:
                break
            if selected and case['name'] not in selected:
                continue
            for repeat in range(repeats):
                excel.new_document()
                excel.add_module('Harness', case['macro'], line_numbers=False)
                result = excel.run_raw('Harness.Probe', timeout=30.0)
                assert not excel.events[-1].get('dialogs'), (case['name'], result)
                case['result'] = result
                if 'results' in case:
                    case['results'] = list(dict.fromkeys(case['results'] + [result]))
                    case.setdefault('observations', []).append({'result': result, 'repeat': repeat,
                        'execution': 'direct Application.Run', 'measured_at': datetime.now(timezone.utc).isoformat()})
                print(case['name'], result, flush=True)
        files = []
        folder = OUT.parent / 'full_calculation_files'
        folder.mkdir(exist_ok=True)
        for name in ['full_clean_udf_manual_CalculateFull', 'full_chained_udfs_manual_CalculateFullRebuild',
                     'full_indirect_udf_manual_CalculateFullRebuild', 'full_independent_udfs_auto_CalculateFull',
                     'full_dirty_udf_manual_CalculateFull', 'full_ordinary_auto_CalculateFullRebuild']:
            case = next(case for case in record['probes'] if case['name'] == name)
            excel.new_document()
            # Persist pure results: native scheduling of stateful independent
            # UDFs varies, and the trace probes separately retain that evidence.
            source = case['macro'].replace('Tick = value * 10 + Calls', 'Tick = value * 10') + '\n' + READ_SAVED
            excel.add_module('Harness', source, line_numbers=False)
            setup = excel.run_raw('Harness.Probe')
            excel.run_raw('Harness.SetSaveFlag')
            path = folder / (name + '.xlsm')
            save = 'Sub SaveFixture()\nApplication.DisplayAlerts = False\nApplication.CalculateBeforeSave = False\nActiveWorkbook.SaveAs "' + str(path) + '",52\nEnd Sub'
            saved = excel.run_vba(save, 'SaveFixture')
            assert saved.ok and not saved.dialogs, saved
            excel.open_document(path)
            before = excel.run_raw('Harness.ReadSaved')
            excel.run_raw('Harness.FullNow')
            after_full = excel.run_raw('Harness.ReadSaved')
            excel.run_raw('Harness.RebuildNow')
            after_rebuild = excel.run_raw('Harness.ReadSaved')
            assert not any(event.get('dialogs') for event in excel.events), excel.events[-1]
            files.append({'name': name, 'file': path.relative_to(OUT.parent).as_posix(), 'result': before,
                          'full': after_full, 'rebuild': after_rebuild,
                          'macro': source, 'setup': setup,
                          'calculation': -4135 if '_manual_' in name else -4105})
        record['files'] = files
        record['read_saved'] = READ_SAVED
        record['measured_at'] = datetime.now(timezone.utc).isoformat()
    temporary = OUT.with_suffix('.json.tmp')
    temporary.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
    temporary.replace(OUT)


if __name__ == '__main__':
    main()
