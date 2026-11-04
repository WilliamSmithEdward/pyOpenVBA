"""The sheet's old control collections, replayed from live Excel (scripts/measure_old_controls.py,
tests/fixtures/old_controls.json)."""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/old_controls.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Refused, which each case below checks: setting what a new control only reads, a control's Border and Interior, a size a
#: control picks for itself, changes on a sheet whose drawing objects are protected, a button's font and duplicating a control,
#: whose writing to the file is not measured.
REFUSED = ('o.Duplicate', 'o.Font.', 'TypeName(o.Characters.Font) & " " & o.Characters.Font.Bold', 'o.AutoSize =', 'o.HorizontalAlignment =', 'o.VerticalAlignment =', 'o.Orientation =', 'o.Display3DShading =', 'TypeName(o.Border)', 'TypeName(o.Interior)', 'Add(48, 96, 0, 0)', 'ws.Protect: o.Left')
#: Measured and not replayed: a list box's height, which Excel fits to whole rows at this display's DPI, and a shape's place
#: after a row goes in above it, in rows of this display's height; and DrawingObjects, which still answers as Shapes.
LEFT_OUT = ('ListBoxes.Add(Left:=', 'TypeName(ws.DrawingObjects(1))', 'ws.Rows(1).Insert: r = o.Top')


def _kept(name: str) -> bool:
    return not any(text in name for text in LEFT_OUT)


def _height_fitted(name: str) -> bool:
    return 'ListBoxes' in name and name.endswith('r = o.Height')


def _refused(index: int) -> bool:
    name, answer = RECORD['cases'][index]
    if not any(text in name for text in REFUSED) or not _kept(name) or answer.startswith('Err 438'):
        return False
    # Only a button has a font of its own, and the others' Characters answer theirs as Null.
    if 'Add(48, 96, 0, 0)' in name:
        return 'ws.Buttons' not in name and 'ws.DropDowns' not in name
    return 'ws.Buttons' in name or 'Font' not in name


REFUSALS = [index for index in range(len(NAMES)) if _refused(index)]
KEPT = [index for index, name in enumerate(NAMES) if index not in REFUSALS and _kept(name) and not _height_fitted(name)]


def _run(index: int) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='OldProbe')
    return str(app.run('OldProbe.Probe', index))


def _same_book(text: str) -> str:
    """A macro reads with its workbook's name, which is the session's."""
    return re.sub(r'Book\d+!', 'Book!', text)


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(index: int) -> None:
    assert _same_book(_run(index)) == _same_book(RECORD['cases'][index][1])


@pytest.mark.parametrize('index', REFUSALS, ids=[NAMES[index] for index in REFUSALS])
def test_what_is_not_built_says_so(index: int) -> None:
    with pytest.raises(VBAUnsupportedError):
        _run(index)


def _recorded(body: str) -> object:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'''Public Function Recorded() As String
    ActiveSheet.Buttons.Add(48, 96, 72, 24).Select
{body}
End Function
''', name='Recorder')
    return app.run('Recorder.Recorded')


def test_what_the_macro_recorder_writes() -> None:
    assert _recorded('''    Selection.OnAction = "Macro1"
    Selection.Characters.Text = "Run it"
    Recorded = ActiveSheet.Shapes(1).TextFrame.Characters.Text & "|" & ActiveSheet.Buttons(1).Caption & "|" & ActiveSheet.Shapes(1).OnAction''') == 'Run it|Run it|Book1!Macro1'


def test_the_font_block_the_recorder_writes_says_it_is_not_built() -> None:
    with pytest.raises(VBAUnsupportedError):
        _recorded('''    With Selection.Characters(Start:=1, Length:=3).Font
        .Name = "Arial"
    End With''')


def test_a_control_is_saved_and_opened_again(tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('''Public Sub Make(ByVal path As String)
    Dim o As Object
    Set o = ActiveSheet.Buttons.Add(48, 96, 72, 24)
    o.OnAction = "Macro1"
    o.Caption = "Run"
    ActiveSheet.CheckBoxes.Add 48, 144, 72, 24
    ActiveSheet.OptionButtons.Add(48, 192, 72, 24).Enabled = False
    ActiveSheet.DropDowns.Add(48, 240, 72, 24).AddItem "a"
    ActiveSheet.ListBoxes.Add(48, 288, 72, 48).AddItem "b"
    ActiveWorkbook.SaveAs path, xlOpenXMLWorkbookMacroEnabled
End Sub
''', name='Make')
    path = tmp_path / 'made.xlsm'
    app.run('Make.Make', str(path))
    again = ExcelApplication(sandbox=False)
    again.open_workbook(str(path))
    again.add_module('''Public Function Read() As String
    Dim ws As Worksheet
    Set ws = Workbooks(1).Worksheets(1)
    Read = ws.Buttons.Count & ws.CheckBoxes.Count & ws.OptionButtons.Count & ws.DropDowns.Count & ws.ListBoxes.Count & " " & ws.Buttons(1).Name & " " & ws.Shapes.Count
End Function
''', name='Read')
    assert again.run('Read.Read') == '11111 Button 1 5'
