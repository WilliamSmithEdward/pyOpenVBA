"""FormControl.set_font gives a control a font of its own, as the designer's Font dialog does (issue #248).

The control's TextProps is the one the designer writes for a control it adds (tests/fixtures/form_designer.json),
with the font given in place of the form's: face, height in whole twips, charset, effects and weight. Excel reads
back each font as set, Meiryo UI bold in the Japanese charset among them (checked live when this was written:
UserForms.Add, then each control's Font).
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile
from pyopenvba._oforms_records import FormParseError


def _reloaded(tmp_path: Path, kind: str, **font: Any) -> dict[str, object]:
    target = tmp_path / 'fonts.xlsm'
    with ExcelFile.create_new(target) as workbook:
        form = workbook.add_form('frmFonts')
        control = form.add_control(kind, 'Probe')
        control.set_font(**font)
        workbook.save()
    with ExcelFile(target) as workbook:
        record = next(one for one in workbook.forms() if one.name == 'frmFonts').control('Probe').record
        assert record is not None and record.text_props is not None
        return record.text_props.properties()


def test_a_japanese_bold_font(tmp_path: Path) -> None:
    stored = _reloaded(tmp_path, 'CommandButton', name='Meiryo UI', size=10, bold=True, charset=128)
    assert stored['FontName'] == 'Meiryo UI'
    assert (stored['FontHeight'], stored['FontCharSet'], stored['FontWeight']) == (200, 128, 700)
    # Bold, with fAutoColor as the designer sets it beside any effect; a button centres its text.
    assert stored['FontEffects'] == 1 | 1 << 30
    assert stored['ParagraphAlign'] == 3


def test_styles_and_a_fractional_size(tmp_path: Path) -> None:
    stored = _reloaded(tmp_path, 'Label', name='Arial', size=8.25, italic=True, underline=True, strikethrough=True)
    assert (stored['FontName'], stored['FontHeight'], stored['FontCharSet']) == ('Arial', 165, 0)
    assert stored['FontEffects'] == 2 | 4 | 8 | 1 << 30
    assert 'FontWeight' not in stored and 'ParagraphAlign' not in stored


def test_a_control_without_a_font_of_its_own_is_refused(tmp_path: Path) -> None:
    with ExcelFile.create_new(tmp_path / 'image.xlsm') as workbook:
        image = workbook.add_form('frmImage').add_control('Image', 'Picture')
        with pytest.raises(FormParseError, match='no font'):
            image.set_font('Arial', 10)
