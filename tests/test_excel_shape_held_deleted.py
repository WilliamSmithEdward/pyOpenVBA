"""A shape, a ShapeRange or a part of a shape held after the shape has gone, replayed from live Excel
(scripts/measure_shape_held_deleted.py, tests/fixtures/shape_held_deleted.json)."""
from __future__ import annotations

import json
import struct
import zlib
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_held_deleted.json').read_text(encoding='utf-8'))
#: The clipboard failed in the probe, so Cut and Copy deleted nothing there; Selection was another sheet's.
NOT_DELETING = ('sh.Cut', 'sh.Copy', 'Selection.Delete')
#: These are not built: the old control collections (excel-old-controls has them), DrawingObjects.Delete, PictureEffects, ThreeD, and an Office type's
#: answer to a member it never had, which reports itself.
UNBUILT = ('ws.Buttons', 'ws.CheckBoxes', 'PictureEffects', 'ws.Shapes.Range(Array(1))', 'ws.DrawingObjects', 'sh.ThreeD')
CASES = [case for case in RECORD['cases'] if not any(one in case[2] for one in NOT_DELETING) and not any(one in ' '.join([case[0], case[1][0], case[2]]) for one in UNBUILT)
         and case[1][:2] != ['sh.Fill', 'h.NoSuchMember']]
IDS = [f'{index}: {case[0][:40]} | {case[1][0]} | {case[2]}' for index, case in enumerate(CASES)]


def _png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 2, 2, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(b'\0\xff\0\0\xff\0\0' * 2)) + chunk(b'IEND', b'')


def _go(case: list[Any], folder: Path) -> Any:
    made, (held, read, write, call), gone = case[0], case[1], case[2]
    made = made.replace('folder & "\\p.png"', f'"{folder / "p.png"}"')
    if made.startswith('With'):
        made = '\n    '.join(made.split(': '))
    lines = ['Public Function Go() As String', '    Dim ws As Object, sh As Object, h As Object, o As Object, s2 As Variant, v As Variant, s As String',
             '    Set ws = ActiveWorkbook.Worksheets(1)', '    ActiveWorkbook.Worksheets.Add After:=ws', '    On Error Resume Next', f'    {made}', f'    Set h = {held}',
             '    s = "made " & Err.Number: Err.Clear', f'    {gone}', '    s = s & "~gone " & Err.Number: Err.Clear',
             '    v = "": v = TypeName(h): s = s & "~" & v & "/" & Err.Number: Err.Clear', '    v = "": v = (h Is Nothing): s = s & "~" & v & "/" & Err.Number: Err.Clear']
    lines += [f'    v = "": v = {read}: s = s & "~" & Err.Number: Err.Clear' if read else '    s = s & "~"']
    lines += [f'    {one}: s = s & "~" & Err.Number: Err.Clear' if one else '    s = s & "~"' for one in (write, call)]
    lines += ['    Go = s', 'End Function']
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('\n'.join(lines) + '\n', name='Held')
    return app.run('Held.Go')


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_case_as_excel_answered(case: list[Any], tmp_path: Path) -> None:
    (tmp_path / 'p.png').write_bytes(_png())
    assert _go(case, tmp_path) == case[3]


def test_a_control_format_held_after_its_delete_raises_1004() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Function Go() As Long\n    Dim sh As Object, h As Object, v As Variant\n    Set sh = ActiveSheet.Shapes.AddFormControl(0, 10, 10, 100, 40)\n'
                   '    Set h = sh.ControlFormat\n    sh.Delete\n    On Error Resume Next\n    v = h.Enabled\n    Go = Err.Number\nEnd Function\n', name='Held')
    assert app.run('Held.Go') == 1004
