"""A workbook keeps the spelling of the folder it was given, measured in Excel with an 8.3 name: FullName and Path both answered
C:\\...\\A_LONG~1, on SaveAs and on Open, where the long name is the same folder."""
from __future__ import annotations

import ctypes
import sys
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication


def _short(folder: Path) -> str:
    if sys.platform != 'win32':
        pytest.skip('8.3 names are Windows')
    buffer = ctypes.create_unicode_buffer(600)
    ctypes.windll.kernel32.GetShortPathNameW(str(folder), buffer, 600)  # type: ignore[attr-defined]
    if buffer.value.casefold() == str(folder).casefold():
        pytest.skip('this volume keeps no 8.3 names')
    return buffer.value


def test_a_workbook_keeps_the_folder_as_it_was_given(tmp_path: Path) -> None:
    folder = tmp_path / 'a_long_folder_name_for_short_paths'
    folder.mkdir()
    short = _short(folder)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'''Function Probe() As String
    Dim o As Object, f As String, s As String
    f = "{short}\\saved.xlsx"
    Set o = Workbooks.Add(xlWBATWorksheet)
    o.SaveAs f, xlOpenXMLWorkbook
    s = (o.FullName = f) & "|" & (o.Path = "{short}")
    o.Close False
    Set o = Workbooks.Open(f)
    s = s & "|" & (o.FullName = f) & "|" & (o.Path = "{short}")
    Set o = Workbooks.Open("{folder}\\saved.xlsx")
    s = s & "|" & (o.FullName = f) & "|" & Workbooks.Count
    o.Close False
    Probe = s
End Function''', name='Harness')
    found = str(app.run('Harness.Probe'))
    assert found.startswith('True|True|True|True|True|')
