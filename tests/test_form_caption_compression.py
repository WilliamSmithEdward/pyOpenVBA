"""A string property is stored compressed, one byte a character, only while every character fits one byte (issue #247).

A new control already holds a compressed default caption. Text beyond U+00FF set over it was stored compressed in the
project's code page, which Excel shows garbled (Japanese 実行 as Às); it is now stored as UTF-16. Text that fits one
byte keeps the compression the VBE gives it.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from pyopenvba import ExcelFile


def _caption_after_save(tmp_path: Path, text: str) -> tuple[str, bool]:
    target = tmp_path / 'form.xlsm'
    with ExcelFile.create_new(target) as workbook:
        form = workbook.add_form('frmTest', caption='Monthly')
        form.add_control('CommandButton', 'btnOK').set_property('Caption', text)
        workbook.save()
    with ExcelFile(target) as workbook:
        control = next(form for form in workbook.forms() if form.name == 'frmTest').control('btnOK')
        assert control.record is not None
        stored = control.record.strings['Caption']
        return stored.text, stored.compressed


@pytest.mark.parametrize('text', ['実行', 'Ünder 実', 'Ω'])
def test_text_beyond_one_byte_is_stored_as_utf16(tmp_path: Path, text: str) -> None:
    assert _caption_after_save(tmp_path, text) == (text, False)


@pytest.mark.parametrize('text', ['Run', 'Café'])
def test_text_of_one_byte_characters_stays_compressed(tmp_path: Path, text: str) -> None:
    assert _caption_after_save(tmp_path, text) == (text, True)
