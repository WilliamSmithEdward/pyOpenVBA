"""Shape.AlternativeText, Title, Decorative, Rotation, LockAspectRatio and Locked changed on a shape a workbook
opened with: what the shape reads after the change, and the anchor the saved drawing holds.

Excel saves two workbooks with a rectangle: plain.xlsx, and set.xlsx, whose rectangle already has alternative text,
a title, the decorative flag, a 30 degree rotation, a locked aspect ratio and Locked off. Both are copied into
tests/fixtures/loaded_shape/ without their save path. Each case opens one, changes its shape and saves it.

No case turns the shape far enough to move its anchor: Excel works the new anchor out from the columns and rows as
the measuring machine's display draws them, which at its 144 DPI are not the 96 DPI ones the model keeps.
tests/fixtures/shape_properties.json measures that rule where the two agree.

    python scripts/measure_loaded_shape_properties.py

writes tests/fixtures/loaded_shape/cases.json, which tests/test_excel_loaded_shape_properties.py replays.
"""
from __future__ import annotations

import json
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from fixture_workbook import copy_saved

FOLDER = Path(__file__).resolve().parents[1] / 'tests/fixtures/loaded_shape'
READ = ('"|" & p.AlternativeText & "," & p.Title & "," & p.Decorative & "," & p.Rotation & "," & '
        'p.LockAspectRatio & "," & p.Locked')
SET = ('p.AlternativeText = "D": p.Title = "T": p.Decorative = True: p.Rotation = 30: '
       'p.LockAspectRatio = msoTrue: p.Locked = False')
#: (base workbook, statements)
CASES = [
    ('plain', ''),
    ('plain', 'p.AlternativeText = "A box"'),
    ('plain', 'p.Title = "Box"'),
    ('plain', 'p.AlternativeText = "x & <y>": p.Title = "T"'),
    ('plain', 'p.Decorative = True'),
    ('plain', 'p.Rotation = 30'),
    ('plain', 'p.IncrementRotation 400'),
    ('plain', 'p.LockAspectRatio = msoTrue'),
    ('plain', 'p.Locked = False'),
    ('set', ''),
    ('set', 'p.AlternativeText = ""'),
    ('set', 'p.Title = ""'),
    ('set', 'p.Decorative = False'),
    ('set', 'p.Rotation = 0'),
    ('set', 'p.LockAspectRatio = msoFalse'),
    ('set', 'p.Locked = True'),
    ('set', 'p.AlternativeText = "New": p.Title = "Other"'),
]


def _probe() -> str:
    lines = ['Public Function MakeBases(ByVal folder As String) As String',
             '    Dim wb As Workbook, p As Object, i As Long', '    Application.DisplayAlerts = False',
             '    For i = 0 To 1', '        Set wb = Workbooks.Add(xlWBATWorksheet)',
             '        Set p = wb.Worksheets(1).Shapes.AddShape(msoShapeRectangle, 100, 50, 120, 60)',
             f'        If i = 1 Then {SET}',
             '        wb.SaveAs folder & "\\" & IIf(i = 0, "plain", "set") & ".xlsx", xlOpenXMLWorkbook',
             '        wb.Close False', '    Next', '    Application.DisplayAlerts = True', '    MakeBases = "ok"',
             'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal base As String, ByVal folder As String) As String',
             '    Dim wb As Workbook, p As Object, s As String',
             '    Set wb = Workbooks.Open(folder & "\\" & base & ".xlsx")', '    Set p = wb.Worksheets(1).Shapes(1)',
             '    On Error GoTo Failed', f'    s = {READ}', '    Select Case which']
    for index, (_, statements) in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {statements}'] if statements else [])
    lines += ['    End Select', f'    Tried = "ok" & s & {READ}', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & s', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim s As String',
              '    Application.DisplayAlerts = False']
    for index, (base, _) in enumerate(CASES):
        lines.append(f'    s = s & "@@@" & Tried({index}, "{base}", folder)')
    lines += ['    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def anchor(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.search(r'<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>', xml, re.S)
    assert found is not None
    return found.group()


def main() -> None:
    source = _probe()
    FOLDER.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        made = excel.run_vba(source, 'MakeBases', (folder,), timeout=120.0)
        assert made.ok and not made.dialogs, made
        for base in ('plain', 'set'):
            copy_saved(Path(folder) / f'{base}.xlsx', FOLDER / f'{base}.xlsx')
            # Each case opens the committed copy, so Excel and the model start from the same bytes.
            shutil.copyfile(FOLDER / f'{base}.xlsx', Path(folder) / f'{base}.xlsx')
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        anchors = [anchor(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[base, statements, answer, found]
                        for (base, statements), answer, found in zip(CASES, answers, anchors, strict=True)]}
    (FOLDER / 'cases.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for base, statements, answer, _ in record['cases']:
        print(base, statements, '->', answer)


if __name__ == '__main__':
    main()
