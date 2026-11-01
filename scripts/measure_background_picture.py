"""Worksheet.SetBackgroundPicture: what it answers, and the picture element, relationships, media parts and default
content types the saved workbook holds.

Each case makes a workbook with 1 in A1, gives its sheet (or sheets) a background picture as the case says, and
saves it. The picture is a 2 by 2 red PNG this script writes. Each saved sheet is recorded from its pageMargins on,
with its relationships, and the workbook's media parts, Default content types and whether the media holds the PNG's
own bytes.

    python scripts/measure_background_picture.py

writes tests/fixtures/background_picture.json, which tests/test_excel_background_picture.py replays.
"""
from __future__ import annotations

import json
import re
import struct
import tempfile
import zipfile
import zlib
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/background_picture.json'
CASES = [
    'ws.SetBackgroundPicture folder & "\\bg.png"',
    'ws.SetBackgroundPicture folder & "\\bg.png": ws.SetBackgroundPicture ""',
    'ws.SetBackgroundPicture folder & "\\missing.png"',
    'ws.SetBackgroundPicture folder & "\\bg.png": wb.Worksheets.Add(After:=ws).SetBackgroundPicture folder & "\\bg.png"',
    'ws.Range("A1").AddComment "n": ws.SetBackgroundPicture folder & "\\bg.png"',
    'ws.SetBackgroundPicture ""',
    'v = TypeName(ws.SetBackgroundPicture(folder & "\\bg.png"))',
]


def red_png() -> bytes:
    """A 2 by 2 red PNG."""
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data) & 0xFFFFFFFF)
    raw = b''.join(b'\x00' + b'\xff\x00\x00' * 2 for _ in range(2))
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 2, 2, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(raw, 9)) + chunk(b'IEND', b''))


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Object, v As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    ws.Range("A1").Value = 1', '    On Error GoTo Failed',
             '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', '    Tried = "ok " & v', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number',
              '    Resume Saved', 'End Function', '', 'Public Function Probe(ByVal folder As String) As String',
              '    Dim i As Long, s As String', '    Application.DisplayAlerts = False',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i, folder)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def saved(path: Path, picture: bytes) -> dict[str, object]:
    """The saved workbook's sheets from their pageMargins on, their relationships, its media and Defaults."""
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        sheets = sorted(name for name in names if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', name))
        found: dict[str, object] = {}
        for sheet in sheets:
            text = package.read(sheet).decode('utf-8')
            rels = sheet.replace('worksheets/', 'worksheets/_rels/') + '.rels'
            found[sheet] = [text[text.find('<pageMargins'):],
                            package.read(rels).decode('utf-8') if rels in names else '']
        media = sorted(name for name in names if name.startswith('xl/media/'))
        found['media'] = [[name, package.read(name) == picture] for name in media]
        found['defaults'] = re.findall(r'<Default\b[^>]*/>', package.read('[Content_Types].xml').decode('utf-8'))
    return found


def main() -> None:
    source = _probe()
    picture = red_png()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        (Path(folder) / 'bg.png').write_bytes(picture)
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=200.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        files = [saved(Path(folder) / f'c{index}.xlsx', picture) for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, found] for case, answer, found in zip(CASES, answers, files, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, found in record['cases']:
        print(case, '->', answer, '|', found['media'], found['defaults'])


if __name__ == '__main__':
    main()
