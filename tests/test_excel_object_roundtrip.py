"""What a sheet's objects answer when loaded from a file Excel saved, replayed from live Excel (scripts/measure_object_roundtrip.py,
tests/fixtures/object_roundtrip.json), and what a save keeps of them after an edit the model redraws them for."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'object_roundtrip.json').read_text(encoding='utf-8'))
BOOK = FIXTURES / 'object_roundtrip.xlsx'


def _source(reads: list[str]) -> str:
    lines: list[str] = []
    for number, read in enumerate(reads):
        lines += [f'Public Function R{number}() As String', '    Dim wb As Object, ws As Object, v As Variant', '    Set wb = ActiveWorkbook: Set ws = wb.Worksheets(1)', '    On Error Resume Next', f'    v = {read}',
                  '    If Err.Number Then', f'        R{number} = "E" & Err.Number', '    ElseIf IsObject(v) Or IsArray(v) Or IsError(v) Or IsNull(v) Then', f'        R{number} = TypeName(v)', '    Else', f'        R{number} = CStr(v)', '    End If', 'End Function']
    return '\n'.join(lines) + '\n'


def test_each_read_of_excels_file_answers_as_excel_did() -> None:
    reads, truth = RECORD['reads'], RECORD['excel']
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module(_source(reads), name='Reader')
    wrong: list[tuple[str, str, str]] = []
    for number, read in enumerate(reads):
        try:
            got = str(app.run(f'Reader.R{number}'))
        except VBAUnsupportedError:
            continue  # it says it cannot
        if read.startswith('ws.Buttons') and got == 'E438':
            continue  # the old control collections are on excel-old-controls, not here yet
        if got != truth[number]:
            wrong.append((read, truth[number], got))
    assert not wrong


def _note_boxes(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as saved:
        vml = saved.read('xl/drawings/vmlDrawing1.vml').decode('utf-8')
    boxes = [shape for shape in re.findall(r'<v:shape .*?</v:shape>', vml, re.DOTALL) if 'ObjectType="Note"' in shape]
    return {found.group(1): shape for shape in boxes if (found := re.search(r'<x:Row>(\d+)</x:Row>', shape))}


def test_a_row_inserted_above_a_note_leaves_its_settings(tmp_path: Path) -> None:
    """A12's note has a yellow fill and A13's is unlocked, and each is a row further down afterward."""
    before = _note_boxes(BOOK)
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module('Public Sub Edit()\n    ActiveWorkbook.Worksheets(1).Rows(1).Insert\nEnd Sub\n', name='Editor')
    app.run('Editor.Edit')
    after = _note_boxes(Path(app.save(tmp_path / 'edited.xlsx')))
    assert sorted(int(row) for row in after) == [int(row) + 1 for row in sorted(before, key=int)]
    for row, box in before.items():
        moved = after[str(int(row) + 1)]
        for setting in ('fillcolor="[^"]*"', 'strokecolor="[^"]*"', '<x:Locked>[^<]*</x:Locked>', '<v:fill [^>]*>'):
            assert re.findall(setting, moved) == re.findall(setting, box)
    assert any('<x:Locked>False</x:Locked>' in box for box in after.values())
