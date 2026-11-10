"""Worksheet.TextBoxes, Pictures, Rectangles, Ovals, Lines and Arcs, and DrawingObjects, replayed from live Excel
(scripts/measure_old_drawings.py, tests/fixtures/old_drawings.json)."""
from __future__ import annotations

import json
import os
import re
import struct
import zipfile
import zlib
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/old_drawings.json').read_text(encoding='utf-8'))
READS: list[str] = RECORD['reads']
#: Each of these reports itself: what they write is not worked out, or was measured once. Copy and Cut failed in the probe with the clipboard away.
#: A linked drawing's Font reads a colour by a rule not worked out (Formula is replayed in test_excel_drawing_text_links.py).
#: xlArrowHeadStyleTriangle is no constant: it sets Empty, which Excel took as Open, from one case (the measured styles are replayed in
#: test_excel_line_arrowheads.py).
UNBUILT = ('ws.Buttons.Add', 'h.Copy', 'h.Cut', 'h.Formula =', 'h.ArrowHeadStyle = xlArrowHeadStyleTriangle', 'h.RoundedCorners =', 'h.AutoSize = True', 'ws.DrawingObjects.Left', 'h.Characters(1, 2).Font',
           'Set h = ws.Lines.Add(10, 10, 110, 50): Set o = h.Duplicate', 'Set h = ws.Arcs.Add(10, 10, 110, 50): Set o = h.Duplicate')
PICTURE_UNBUILT = ('h.Interior', 'h.Border', 'Set o = h.Duplicate', 'h.Shadow')
#: Reading these goes through the display, whose scale here snaps them to a pixel after some reads, and neither Interior.ColorIndex nor Enabled gives more.
SNAPPED = {'h.Left', 'h.Top', 'h.Width', 'h.Height'}
#: A text box's own colours and dash go through what a shape's Line refuses on one.
#: Giving an arc a fill, and taking its line away, write what the shape's Fill does not.
ARC_UNBUILT = ('h.Interior.ColorIndex = 3', 'h.Interior.Color = 255', 'h.Border.LineStyle = xlNone')
TEXT_BOX_UNBUILT = ('h.Interior.ColorIndex = 3', 'h.Border.ColorIndex = 3', 'h.Border.Color = 255', 'h.Border.LineStyle = xlDash')
CASES = [case for case in RECORD['cases'] if not any(one in case[0] for one in UNBUILT)
         and not ('ws.TextBoxes.Add' in case[0] and any(one in case[0] for one in TEXT_BOX_UNBUILT))
         and not ('ws.Arcs.Add' in case[0] and any(one in case[0] for one in ARC_UNBUILT))
         and not ('Pictures.Insert' in case[0] and any(one in case[0] for one in PICTURE_UNBUILT))]
IDS = [f'{index}: {case[0][:100]}' for index, case in enumerate(CASES)]


def _png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 2, 2, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(b'\0\xff\0\0\xff\0\0' * 2)) + chunk(b'IEND', b'')


def _source(statements: str, folder: str) -> str:
    made = statements.replace('folder & "\\', f'"{folder}{os.sep}')
    lines = ['Public Function Go() As String', '    Dim ws As Object, h As Object, o As Object, v As Variant, s As String, t As String, s2 As Variant', '    Set ws = ActiveSheet']
    lines += ['    On Error Resume Next' if 'Err.Clear' in statements else '    On Error GoTo Failed', f'    {made}', '    s = "ok"', 'Measured:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    t = ""', f'    t = {read}', '    If Err.Number Then t = "E" & Err.Number', '    s = s & "~" & t']
    lines += ['    Go = s', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function']
    return '\n'.join(lines) + '\n'


def _steady(part: str) -> str:
    return re.sub(r'<a:extLst>.*?</a:extLst>', '', part)


def _replayed(case: list[Any], tmp_path: Path) -> tuple[str, list[str], list[str]]:
    (tmp_path / 'p.png').write_bytes(_png())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(case[0], str(tmp_path)), name='Old')
    answer = str(app.run('Old.Go'))
    path = tmp_path / 'b.xlsx'
    app.save(path)
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8') if 'xl/drawings/drawing1.xml' in names else ''
    return answer, re.findall(r'<xdr:(?:sp|cxnSp|pic)\b.*?</xdr:(?:sp|cxnSp|pic)>', xml, re.DOTALL), names


def _close(mine: str, theirs: str) -> bool:
    try:
        return abs(float(mine) - float(theirs)) <= 0.75
    except ValueError:
        return mine == theirs


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_case_as_excel_answered(case: list[Any], tmp_path: Path) -> None:
    answer, parts, _ = _replayed(case, tmp_path)
    names = ['status', *READS]
    mine = dict(zip(names, answer.split('~'), strict=True))
    theirs = dict(zip(names, case[1].split('~'), strict=True))
    for read in SNAPPED:
        assert _close(mine.pop(read), theirs.pop(read)), read
    assert mine == theirs
    assert [_steady(part) for part in parts] == [_steady(part) for part in case[2]]


def test_a_picture_inserted_is_linked_to_its_file_not_kept(tmp_path: Path) -> None:
    case = next(one for one in RECORD['cases'] if one[0] == 'Set h = ws.Pictures.Insert(folder & "\\p.png")')
    _, _, names = _replayed(case, tmp_path)
    assert not [name for name in names if 'media' in name]
    with zipfile.ZipFile(tmp_path / 'b.xlsx') as package:
        rels = package.read('xl/drawings/_rels/drawing1.xml.rels').decode('utf-8')
        types = package.read('[Content_Types].xml').decode('utf-8')
    # The target holds this machine's own path, spelled its own way on another system.
    assert rels.replace(str(tmp_path) + os.sep, '<folder>\\') == RECORD['picture_relationships']
    assert 'Extension="png"' not in types
