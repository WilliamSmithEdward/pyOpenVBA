"""Shapes through the sweep in tests/set_save_reopen.py: new ones and ones in a file, and each setting that can be claimed of them."""
from __future__ import annotations

from pathlib import Path

import pytest
from set_save_reopen import Case, Target, cases, check

BOOK = Path(__file__).parent / 'fixtures/loaded_shape_text.xlsx'
_NEW = 'ws.Shapes.Add'


def _new(name: str, kind: str, made: str) -> Target:
    return Target(name, f'Set s = {_NEW}{made}: s.Name = "Target"', 'Set s = ws.Shapes("Target")', kind=kind)


def _loaded(name: str, kind: str) -> Target:
    return Target(f'loaded {name}', f'Set s = ws.Shapes("{name}")', f'Set s = ws.Shapes("{name}")', BOOK, kind)


TARGETS = [_new('rectangle', 'text', 'Shape(1, 10, 10, 100, 40): s.TextFrame.Characters.Text = "Hello world"'), _new('text box', 'text', 'Textbox(1, 10, 10, 100, 40): s.TextFrame.Characters.Text = "Hello world"'),
           _new('connector', 'line', 'Connector(1, 10, 10, 100, 40)'), _new('line', 'line', 'Line(10, 10, 100, 40)'), _new('check box', 'control', 'FormControl(xlCheckBox, 10, 10, 100, 40)'),
           _new('button', 'control', 'FormControl(xlButtonControl, 10, 10, 100, 40)'), _new('drop-down', 'control', 'FormControl(xlDropDown, 10, 10, 100, 40)'), _new('spinner', 'control', 'FormControl(xlSpinner, 10, 10, 100, 40)'),
           _loaded('Plain', 'text'), _loaded('Box', 'text'), _loaded('Runs', 'text'), _loaded('Paras', 'text'), _loaded('Clipped', 'text'), _loaded('Joined', 'line'), _loaded('Elbow', 'line'), _loaded('Line', 'line'),
           _loaded('Pic', 'picture'), _loaded('Tuned', 'picture'), _loaded('Clear', 'picture')]
_GENERAL = {'s.AlternativeText': ['"Alt"'], 's.Left': ['33'], 's.Top': ['44'], 's.Width': ['111'], 's.Height': ['66'], 's.Rotation': ['30'], 's.Visible': ['msoFalse'], 's.LockAspectRatio': ['msoTrue'],
            's.Placement': ['xlFreeFloating'], 's.Decorative': ['msoTrue'], 's.Locked': ['False'], 's.Fill.Visible': ['msoFalse'], 's.Fill.ForeColor.RGB': ['255'], 's.Fill.Transparency': ['0.3'], 's.Line.Visible': ['msoFalse'], 's.Line.ForeColor.RGB': ['255'],
            's.Line.Weight': ['3'], 's.Line.DashStyle': ['4'], 's.DrawingObject.Interior.Color': ['255'], 's.DrawingObject.Border.Color': ['255'], 's.DrawingObject.Border.Weight': ['xlThick', 'xlThin', 'xlHairline', 'xlMedium'],
            's.DrawingObject.Border.LineStyle': ['xlDash'], 's.DrawingObject.Text': ['"Xy"']}
_TEXT = {'s.TextFrame2.MarginLeft': ['9'], 's.TextFrame2.WordWrap': ['msoFalse'], 's.TextFrame2.AutoSize': ['2'], 's.TextFrame2.Orientation': ['3'], 's.TextFrame2.VerticalAnchor': ['4'], 's.TextFrame2.HorizontalAnchor': ['2'], 's.TextFrame2.NoTextRotation': ['msoTrue'],
         's.TextFrame2.Column.Number': ['2'], 's.TextFrame.HorizontalAlignment': ['xlHAlignRight'], 's.TextFrame.VerticalAlignment': ['xlVAlignBottom'], 's.TextFrame.HorizontalOverflow': ['0'], 's.TextFrame.VerticalOverflow': ['2'], 's.TextFrame.ReadingOrder': ['xlRTL'],
         's.TextFrame.Characters.Text': ['"New"'], 's.TextFrame.Characters.Font.Bold': ['True'], 's.TextFrame.Characters.Font.Size': ['15'], 's.TextFrame.Characters.Font.Name': ['"Arial"'], 's.TextFrame.Characters.Font.Color': ['255'], 's.TextFrame.Characters.Font.Underline': ['2', '-4119'],
         's.TextFrame2.TextRange.Font.UnderlineStyle': ['3'], 's.TextFrame2.TextRange.Font.Strike': ['2'], 's.TextFrame2.TextRange.Font.BaselineOffset': ['0.2'], 's.TextFrame2.TextRange.Font.Fill.Transparency': ['0.2'], 's.TextFrame2.TextRange.Font.Line.Weight': ['2'],
         's.TextFrame2.TextRange.ParagraphFormat.Alignment': ['2'], 's.TextFrame2.TextRange.ParagraphFormat.LeftIndent': ['9'], 's.TextFrame2.TextRange.ParagraphFormat.IndentLevel': ['2'], 's.TextFrame2.TextRange.ParagraphFormat.SpaceWithin': ['2'], 's.TextFrame2.TextRange.ParagraphFormat.Bullet.Visible': ['-1']}
_ONLY = {'s.ConnectorFormat.Type': ['2'], 's.PictureFormat.Brightness': ['0.7'], 's.PictureFormat.Contrast': ['0.3'], 's.PictureFormat.ColorType': ['2'], 's.PictureFormat.TransparentBackground': ['0'], 's.PictureFormat.TransparencyColor': ['255']}
SETTINGS = {**_GENERAL, **_TEXT, **_ONLY}
KINDS: dict[str, tuple[str, ...]] = {**{path: ('text',) for path in _TEXT}, 's.ConnectorFormat.Type': ('line',), **{path: ('picture',) for path in _ONLY if 'Picture' in path}}
#: What a save still loses. It may only shrink.
_NEW_CONTROLS = ('check box', 'button', 'drop-down', 'spinner')
_CONTROL_SETTINGS = ('s.AlternativeText = "Alt"', 's.Rotation = 30', 's.Visible = msoFalse', 's.LockAspectRatio = msoTrue', 's.Placement = xlFreeFloating', 's.Decorative = msoTrue', 's.Locked = False')
ACCEPTED: dict[str, str] = {f'{control}: {setting}': "a new form control's drawing and VML are written without it" for control in _NEW_CONTROLS for setting in _CONTROL_SETTINGS}


@pytest.mark.parametrize('case', cases(TARGETS, SETTINGS, KINDS), ids=str)
def test_each_setting_comes_back_from_the_file(case: Case, tmp_path: Path) -> None:
    check(case, tmp_path, ACCEPTED)
