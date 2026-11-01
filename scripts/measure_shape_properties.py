"""Shape.AlternativeText, Title, Decorative, Rotation, IncrementLeft, IncrementTop, IncrementRotation,
LockAspectRatio and Locked: what a shape reads before and after each change, the errors bad values raise, and the
shape markup the saved drawing holds.

Each case makes a workbook, adds a rectangle with Shapes.AddShape and runs the case's statements on it, then saves
it. The drawing's shape element is recorded with its creationId, which is random, cut out.

    python scripts/measure_shape_properties.py

writes tests/fixtures/shape_properties.json, which tests/test_excel_shape_properties.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_properties.json'
READ = ('"|" & p.AlternativeText & "," & p.Title & "," & p.Decorative & "," & p.Rotation & "," & p.Left & "," & '
        'p.Top & "," & p.LockAspectRatio & "," & p.Locked')
CASES = [
    '',
    'p.AlternativeText = "A box"',
    'p.Title = "Box title"',
    'p.AlternativeText = "Desc & <more>": p.Title = "T"',
    'p.AlternativeText = "x": p.AlternativeText = ""',
    'p.Decorative = True',
    'p.Decorative = msoTrue',
    'p.Rotation = 45',
    'p.Rotation = -30',
    'p.Rotation = 400',
    'p.Rotation = 12.5',
    'p.IncrementRotation 10: p.IncrementRotation 15',
    'p.IncrementLeft 12: p.IncrementTop -4.5',
    'p.IncrementLeft -500',
    'p.LockAspectRatio = msoTrue',
    'p.LockAspectRatio = msoTrue: p.LockAspectRatio = msoFalse',
    'p.Locked = False',
    'p.Locked = False: p.Locked = True',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, p As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set p = ws.Shapes.AddShape(msoShapeRectangle, 100, 50, 120, 60)', '    On Error GoTo Failed',
             f'    s = {READ}', '    Select Case which']
    for index, statements in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {statements}'] if statements else [])
    lines += ['    End Select', f'    Tried = "ok" & s & {READ}', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & s', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def shape_markup(path: Path) -> str:
    """The drawing's one anchor, its creationId cut out."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.search(r'<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>', xml, re.S)
    assert found is not None
    return re.sub(r'(<a16:creationId\b[^>]*\bid=")\{[0-9A-F-]{36}\}', r'\1{ID}', found.group())


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [shape_markup(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, answer, part] for statements, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
