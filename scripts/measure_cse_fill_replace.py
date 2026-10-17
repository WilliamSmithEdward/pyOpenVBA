"""Measure CSE AutoFill, Replace and duplicate-removal behavior and return values."""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_cse_array_edits import macro as base_macro

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/cse_fill_replace.json'
OPERATIONS: dict[str, tuple[str, str]] = {
    'fill_down': ('', 'ws.Range("B2:C4").AutoFill(Destination:=ws.Range("B2:C7"))'),
    'fill_partial': ('', 'ws.Range("B2:C4").AutoFill(Destination:=ws.Range("B2:C6"))'),
    'fill_short': ('', 'ws.Range("B2:C4").AutoFill(Destination:=ws.Range("B2:C5"))'),
    'fill_right': ('', 'ws.Range("B2:C4").AutoFill(Destination:=ws.Range("B2:E4"))'),
    'fill_right_partial': ('', 'ws.Range("B2:C4").AutoFill(Destination:=ws.Range("B2:F4"))'),
    'fill_up': ('ws.Range("B2:C4").Cut Destination:=ws.Range("B5")', 'ws.Range("B5:C7").AutoFill(Destination:=ws.Range("B2:C7"))'),
    'fill_left': ('ws.Range("B2:C4").Cut Destination:=ws.Range("E2")', 'ws.Range("E2:F4").AutoFill(Destination:=ws.Range("C2:F4"))'),
    'fill_child': ('', 'ws.Range("B2").AutoFill(Destination:=ws.Range("B2:F2"))'),
    'fill_mixed': ('ws.Range("A1").Value = 1', 'ws.Range("A1:D4").AutoFill(Destination:=ws.Range("A1:D8"))'),
    'fill_over_array': ('ws.Range("A2:A4").Value = 9', 'ws.Range("A2:A4").AutoFill(Destination:=ws.Range("A2:C4"), Type:=xlFillCopy)'),
    'fill_over_other_array': ('ws.Range("F2:G4").FormulaArray = "=1"', 'ws.Range("B2:C4").AutoFill(Destination:=ws.Range("B2:G4"))'),
    'fill_invalid_type': ('', 'ws.Range("B2:C4").AutoFill(Destination:=ws.Range("B2:C7"), Type:=99)'),
    'formats_from_child': ('ws.Range("B2").Font.Bold = True', 'ws.Range("B2").AutoFill(Destination:=ws.Range("B2:F2"), Type:=xlFillFormats)'),
    'replace_whole': ('', 'ws.Range("B2:C4").Replace(What:="ROW", Replacement:="COLUMN", LookAt:=xlPart)'),
    'replace_anchor': ('', 'ws.Range("B2").Replace(What:="ROW", Replacement:="COLUMN", LookAt:=xlPart)'),
    'replace_child': ('', 'ws.Range("B3").Replace(What:="ROW", Replacement:="COLUMN", LookAt:=xlPart)'),
    'replace_partial_anchor': ('', 'ws.Range("B2:B3").Replace(What:="ROW", Replacement:="COLUMN", LookAt:=xlPart)'),
    'replace_child_value': ('', 'ws.Range("B3").Replace(What:="21", Replacement:="99", LookAt:=xlWhole)'),
    'replace_missing': ('', 'ws.Range("B2:C4").Replace(What:="never", Replacement:="x")'),
    'replace_constant_formula': ('', 'ws.Range("B2:C4").Replace(What:="=ROW(A1:A3)*10+COLUMN(A1:B1)", Replacement:="=7", LookAt:=xlWhole)'),
    'replace_constant_value': ('', 'ws.Range("B2:C4").Replace(What:="=ROW(A1:A3)*10+COLUMN(A1:B1)", Replacement:="7", LookAt:=xlWhole)'),
    'replace_invalid': ('', 'ws.Range("B2:C4").Replace(What:="ROW", Replacement:=")", LookAt:=xlPart)'),
    'duplicates_unique': ('', 'ws.Range("B2:C4").RemoveDuplicates(Columns:=1, Header:=xlNo)'),
    'duplicates_constant': ('ws.Range("B2:C4").FormulaArray = "=1"', 'ws.Range("B2:C4").RemoveDuplicates(Columns:=1, Header:=xlNo)'),
    'duplicates_omitted_columns': ('', 'ws.Range("B2:C4").RemoveDuplicates()'),
    'duplicates_bad_columns': ('', 'ws.Range("B2:C4").RemoveDuplicates(Columns:=9)'),
    'duplicates_bad_header': ('', 'ws.Range("B2:C4").RemoveDuplicates(Columns:=1, Header:=9)'),
    'duplicates_single': ('', 'ws.Range("B3").RemoveDuplicates(Columns:=1, Header:=xlNo)'),
    'fill_up_partial': ('ws.Range("B2:C4").FormulaArray = "=ROW(J10:J12)*10+COLUMN(J10:K10)"\nApplication.Calculate\nws.Range("B2:C4").Cut Destination:=ws.Range("B5")', 'ws.Range("B5:C7").AutoFill(Destination:=ws.Range("B4:C7"))'),
    'fill_left_partial': ('ws.Range("B2:C4").FormulaArray = "=ROW(J10:J12)*10+COLUMN(J10:K10)"\nApplication.Calculate\nws.Range("B2:C4").Cut Destination:=ws.Range("E2")', 'ws.Range("E2:F4").AutoFill(Destination:=ws.Range("D2:F4"))'),
    'fill_partial_destination': ('ws.Range("A2:A4").Value = 9', 'ws.Range("A2:A4").AutoFill(Destination:=ws.Range("A2:B4"), Type:=xlFillCopy)'),
    'fill_partial_destination_formats': ('ws.Range("A2").Font.Bold = True', 'ws.Range("A2:A4").AutoFill(Destination:=ws.Range("A2:B4"), Type:=xlFillFormats)'),
    'fill_plain_formula': ('ws.Range("B2:C4").ClearContents\nws.Range("B2").Formula = "=ROW(A1)"\nApplication.Calculate', 'ws.Range("B2").AutoFill(Destination:=ws.Range("B2:B5"))'),
    'replace_anchor_value': ('', 'ws.Range("B2").Replace(What:="=ROW(A1:A3)*10+COLUMN(A1:B1)", Replacement:="7", LookAt:=xlWhole)'),
    'replace_long_formula': ('', 'ws.Range("B2:C4").Replace(What:="10", Replacement:="' + 'SUM(' + ','.join(['1'] * 115) + ')' + '", LookAt:=xlPart)'),
    'replace_long_argument': ('', 'ws.Range("B2:C4").Replace(What:="10", Replacement:="' + 'SUM(' + ','.join(['1'] * 130) + ')' + '", LookAt:=xlPart)'),
    'duplicates_header_array': ('ws.Range("B2:C4").ClearContents\nws.Range("B2:C2").FormulaArray = "=COLUMN(A1:B1)"\nws.Range("B3:C4").Value = 1', 'ws.Range("B2:C4").RemoveDuplicates(Columns:=1, Header:=xlYes)'),
}
for _kind in range(11):
    OPERATIONS['fill_type_' + str(_kind)] = ('ws.Range("B2").Font.Bold = True\nws.Range("B2").NumberFormat = "0.00"',
                                           f'ws.Range("B2:C4").AutoFill(Destination:=ws.Range("B2:C7"), Type:={_kind})')
OPERATIONS['read_long_formula_array'] = ('Call ' + OPERATIONS['replace_long_formula'][1], 'ws.Range("B2").FormulaArray')


def macro(setup: str, expression: str, *, manual: bool, calculate: bool) -> str:
    operation = 'Call ' + expression if '.RemoveDuplicates(' in expression else 'outcome = ' + expression
    source = base_macro(setup + '\n' + operation, manual=manual, calculate=calculate)
    source = source.replace('Dim ws As Worksheet, n As Long', 'Dim ws As Worksheet, n As Long, outcome As Variant')
    source = source.replace('Probe = CStr(n) & Mid(ReadSaved(),2)', 'Probe = CStr(n) & ":" & ResultText(outcome) & Mid(ReadSaved(),2)')
    source = source.replace('& ":" & ArrayAddress(cell)', '& ":" & ArrayAddress(cell) & ":" & cell.Font.Bold & ":" & cell.NumberFormat')
    return '''Function ResultText(value As Variant) As String
ResultText = TypeName(value)
If IsNull(value) Or IsArray(value) Then Exit Function
ResultText = ResultText & ":" & CStr(value)
End Function
''' + source


def event_macro(name: str, *, manual: bool) -> str:
    handler = ('Private Sub Worksheet_Change(ByVal Target As Range)\nTrace = Trace & "C:" & Target.Address(False,False) & ";"\nEnd Sub\n'
               'Private Sub Worksheet_Calculate()\nTrace = Trace & "K;"\nEnd Sub')
    quoted = ' & vbCrLf & '.join('"' + line.replace('"', '""') + '"' for line in handler.splitlines())
    setup, expression = OPERATIONS[name]
    setup = ('Dim project As Object\nSet project = ws.Parent.VBProject\n'
             'project.VBComponents(ws.CodeName).CodeModule.AddFromString ' + quoted
             + '\nIf Err.Number <> 0 Then\nProbe = "setup failed:" & Err.Number\nExit Function\nEnd If\n'
             'Trace = "installed;"\nApplication.EnableEvents = True\n' + setup)
    source = 'Public Trace As String\n' + macro(setup, expression, manual=manual, calculate=False)
    return source.replace('Probe = CStr(n) & ":" & ResultText(outcome) & Mid(ReadSaved(),2)',
                          'Probe = CStr(n) & ":" & ResultText(outcome) & Mid(ReadSaved(),2) & "|Trace:" & Trace')


def readback(excel: ExcelSession, record: dict) -> None:
    """Calculation after reopening can differ from calculation before saving a clipped array."""
    for case in record['files']:
        excel.new_document()
        mode = excel.run_vba('Sub SetMode()\nApplication.Calculation = ' + str(case['calculation']) + '\nEnd Sub', 'SetMode')
        assert mode.ok and not mode.dialogs, mode
        excel.open_document(OUT.parent / case['file'])
        excel.add_module('Readback', record['read_saved'] + '\nSub CalculateNow()\nApplication.Calculate\nEnd Sub', line_numbers=False)
        before = excel.run_raw('Readback.ReadSaved')
        assert before == case['result'], case['name']
        excel.run_raw('Readback.CalculateNow')
        case['after_calculate'] = excel.run_raw('Readback.ReadSaved')
        assert not any(event.get('dialogs') for event in excel.events)
        print(case['name'], 'reopened calculation recorded', flush=True)
    record['readback_measured_at'] = datetime.now(timezone.utc).isoformat()


def main() -> None:
    selected = [name for name in sys.argv[1:] if not name.startswith('--')]
    previous = json.loads(OUT.read_text(encoding='utf-8')) if (selected or '--files-only' in sys.argv or '--events-only' in sys.argv or '--readback-only' in sys.argv) and OUT.exists() else None
    probes = previous['probes'] if previous is not None else []
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
        if previous is not None:
            assert previous['excel_version_build'] == version.value
        if '--readback-only' in sys.argv:
            assert previous is not None
            readback(excel, previous)
            temporary = OUT.with_suffix('.json.tmp')
            temporary.write_text(json.dumps(previous, indent=2) + '\n', encoding='utf-8')
            temporary.replace(OUT)
            return
        for name in ([] if '--files-only' in sys.argv or '--events-only' in sys.argv else selected or OPERATIONS):
            setup, expression = OPERATIONS[name]
            for mode, manual, calculate in [('auto', False, False), ('manual', True, False), ('calculated', True, True)]:
                excel.new_document()
                source = macro(setup, expression, manual=manual, calculate=calculate)
                result = excel.run_vba(source, 'Probe', timeout=30.0)
                assert result.ok and not result.dialogs, (name, mode, result)
                probes = [probe for probe in probes if probe['name'] != name + '_' + mode]
                probes.append({'name': name + '_' + mode, 'macro': source, 'result': result.value})
                print(name, mode, result.value, flush=True)
                temporary = OUT.with_suffix('.json.tmp')
                temporary.write_text(json.dumps({'excel_version_build': version.value, 'measured_at': datetime.now(timezone.utc).isoformat(),
                                                'probes': probes}, indent=2) + '\n', encoding='utf-8')
                temporary.replace(OUT)
        if '--events-only' in sys.argv or not selected and '--files-only' not in sys.argv:
            for operation in ['replace_whole', 'replace_child', 'replace_constant_value', 'fill_down']:
                for manual in [False, True]:
                    name = 'events_' + operation + ('_manual' if manual else '_auto')
                    excel.new_document()
                    source = event_macro(operation, manual=manual)
                    result = excel.run_vba(source, 'Probe', timeout=30.0)
                    assert result.ok and not result.dialogs and '|Trace:installed;' in str(result.value), (name, result)
                    probes = [probe for probe in probes if probe['name'] != name]
                    probes.append({'name': name, 'macro': source, 'result': result.value})
                    print(name, result.value, flush=True)
        files = []
        folder = OUT.parent / 'cse_fill_replace_files'
        folder.mkdir(exist_ok=True)
        helper = macro('', 'True', manual=False, calculate=False).split('Function Probe()')[0]
        for name in ['fill_down_auto', 'fill_partial_manual', 'fill_up_partial_manual', 'fill_left_partial_calculated', 'fill_left_partial_manual',
                     'fill_mixed_auto', 'fill_type_3_auto', 'fill_type_4_manual', 'fill_over_array_auto',
                     'replace_whole_manual', 'replace_constant_value_auto', 'replace_long_formula_auto',
                     'duplicates_header_array_manual']:
            probe = next(probe for probe in probes if probe['name'] == name)
            excel.new_document()
            result = excel.run_vba(probe['macro'], 'Probe', timeout=30.0)
            assert result.ok and not result.dialogs, result
            path = folder / (name + '.xlsx')
            source = ('Sub SaveProbe()\nApplication.DisplayAlerts = False\nApplication.CalculateBeforeSave = False\n'
                      'ActiveWorkbook.SaveAs "' + str(path).replace('"', '""') + '", xlOpenXMLWorkbook\nEnd Sub')
            saved = excel.run_vba(source, 'SaveProbe', timeout=30.0)
            assert saved.ok and not saved.dialogs, saved
            read = excel.run_vba(helper, 'ReadSaved', timeout=30.0)
            assert read.ok and not read.dialogs, read
            files.append({'name': name, 'file': path.relative_to(OUT.parent).as_posix(), 'result': read.value,
                          'calculation': -4105 if name.endswith('_auto') else -4135})
        record = {'excel_version_build': version.value, 'measured_at': datetime.now(timezone.utc).isoformat(),
                  'probes': probes, 'read_saved': helper, 'files': files}
        readback(excel, record)
        temporary = OUT.with_suffix('.json.tmp')
        temporary.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
        temporary.replace(OUT)


if __name__ == '__main__':
    main()
