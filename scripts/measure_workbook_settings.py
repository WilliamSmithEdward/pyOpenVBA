"""Measure workbook settings a macro reads and sets: the default table, pivot, slicer and timeline styles, the pivot
field list and chart field buttons, CreateBackup, InactiveListBorderVisible, UpdateRemoteReferences,
ChartDataPointTrack, AccuracyVersion, Final, HasPassword, and the criteria options CaseSensitive,
UseWholeCellCriteria and UseWildcards.

Each case runs on a new workbook and answers as text; some save it, and workbook.xml's workbookPr and calcPr,
styles.xml's tableStyles and the custom properties are recorded.

    python scripts/measure_workbook_settings.py

writes tests/fixtures/workbook_settings.json, which tests/test_excel_workbook_settings.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/workbook_settings.json'
SETTINGS = ['DefaultTableStyle', 'DefaultPivotTableStyle', 'DefaultSlicerStyle', 'DefaultTimelineStyle',
            'ShowPivotTableFieldList', 'ShowPivotChartActiveFields', 'CreateBackup', 'InactiveListBorderVisible',
            'UpdateRemoteReferences', 'ChartDataPointTrack', 'AccuracyVersion', 'Final', 'HasPassword',
            'CaseSensitive', 'UseWholeCellCriteria', 'UseWildcards']
#: Each case: statements on a new workbook wb with sheet ws; the expression read back; the file saved, if any.
CASES: list[tuple[list[str], str, str]] = []
for name in SETTINGS:
    CASES.append(([f'r = TypeName(wb.{name}) & ":" & CStr(wb.{name})'], 'r', ''))
for name in ('ShowPivotTableFieldList', 'ShowPivotChartActiveFields', 'CreateBackup', 'InactiveListBorderVisible',
             'UpdateRemoteReferences', 'ChartDataPointTrack', 'CaseSensitive', 'UseWholeCellCriteria',
             'UseWildcards'):
    CASES.append(([f'wbo.{name} = Not wb.{name}', f'r = CStr(wb.{name})'], 'r', f'{name}.xlsx'))
    CASES.append(([f'wbo.{name} = "x"', f'r = CStr(wb.{name})'], 'r', ''))
    CASES.append(([f'wbo.{name} = 0', f'r = CStr(wb.{name})'], 'r', ''))
for name, values in (('DefaultTableStyle', ['"TableStyleMedium9"', '"tablestylelight1"', '"NoSuch"', '""',
                                            'wb.TableStyles("TableStyleDark1")', '5']),
                     ('DefaultPivotTableStyle', ['"PivotStyleMedium2"', '"NoSuch"', '""']),
                     ('DefaultSlicerStyle', ['"SlicerStyleDark1"', '"NoSuch"']),
                     ('DefaultTimelineStyle', ['"TimeSlicerStyleDark1"', '"NoSuch"'])):
    for value in values:
        stem = re.sub(r'\W', '', value) or 'blank'
        CASES.append(([f'wbo.{name} = {value}', f'r = TypeName(wb.{name}) & ":" & CStr(wb.{name})'], 'r',
                      f'{name}_{stem}.xlsx'))
CASES += [
    (['wbo.DefaultTableStyle = "TableStyleLight1"', 'ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes).Name '
      '= "T"', 'r = ws.ListObjects(1).TableStyle'], 'r', ''),
    (['wbo.AccuracyVersion = 1', 'r = CStr(wb.AccuracyVersion)'], 'r', 'AccuracyVersion.xlsx'),
    (['wbo.AccuracyVersion = 2'], 'r', ''),
    (['wbo.HasPassword = True'], 'r', ''),
    (['wbo.Password = "a"', 'r = CStr(wb.HasPassword)'], 'r', ''),
    (['ws.Range("A1:A3").Value = Application.Transpose(Array("ab", "AB", "a*"))', 'wbo.CaseSensitive = True',
      'ws.Range("B1").Formula = "=COUNTIF(A1:A3,""ab"")"', 'r = ws.Range("B1").Value'], 'r', ''),
    (['ws.Range("A1:A3").Value = Application.Transpose(Array("ab", "AB", "a*"))', 'wbo.UseWildcards = False',
      'ws.Range("B1").Formula = "=COUNTIF(A1:A3,""a*"")"', 'r = ws.Range("B1").Value'], 'r', ''),
    (['ws.Range("A1:A3").Value = Application.Transpose(Array("ab", "abc", "a"))', 'wbo.UseWholeCellCriteria = False',
      'ws.Range("B1").Formula = "=COUNTIF(A1:A3,""a"")"', 'r = ws.Range("B1").Value'], 'r', ''),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, wbo As Object, ws As Worksheet, r As Variant',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set wbo = wb',
             '    Set ws = wb.Worksheets(1)', '    r = "ok"', '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, read, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        lines.append(f'        Tried = CStr({read})')
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        workbook = package.read('xl/workbook.xml').decode('utf-8')
        styles = package.read('xl/styles.xml').decode('utf-8')
        custom = package.read('docProps/custom.xml').decode('utf-8') if 'docProps/custom.xml' in \
            package.namelist() else ''
    found = {name: (re.search(rf'<{name}\b[^>]*/?>', text) or [''])[0]
             for name, text in (('workbookPr', workbook), ('calcPr', workbook), ('tableStyles', styles))}
    found['custom'] = custom
    return found


def _run(source: str, folder: str, first: int, last: int) -> list[str] | None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=300.0)
    if result.outcome == 'runner-error':
        return None
    assert result.ok and not result.dialogs, result
    return str(result.value).split('@@@')


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder:
        answers: list[str] = []
        for first in range(0, len(CASES), 30):
            last = min(first + 29, len(CASES) - 1)
            found = _run(source, folder, first, last)
            if found is None:
                found = [(_run(source, folder, index, index) or ['Excel crashed'])[0]
                         for index in range(first, last + 1)]
            answers += found
        parts = {saved: _parts(Path(folder) / saved) for _, _, saved in CASES
                 if saved and (Path(folder) / saved).exists()}
        blank = _blank(folder)
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, read, saved, answer]
                        for (statements, read, saved), answer in zip(CASES, answers, strict=True)],
              'parts': parts, 'blank': blank}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, _, saved, answer in record['cases']:
        print(statements[0][:60], '->', answer)
    print('blank', blank)


def _blank(folder: str) -> dict[str, str]:
    """The same parts of a new workbook saved as it is, to tell what each setting changed."""
    source = ('Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String\n'
              '    Dim wb As Workbook\n    Set wb = Workbooks.Add(xlWBATWorksheet)\n'
              '    wb.SaveAs folder & "\\blank.xlsx", xlOpenXMLWorkbook\n    wb.Close False\n    Probe = "ok"\n'
              'End Function\n')
    _run(source, folder, 0, 0)
    return _parts(Path(folder) / 'blank.xlsx')


if __name__ == '__main__':
    main()
