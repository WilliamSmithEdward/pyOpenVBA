"""Window members everyday macros set, replayed from live Excel (scripts/measure_window_settings.py,
tests/fixtures/window_settings.json): DisplayOutline, DisplayRuler, DisplayWhitespace, GridlineColor and
GridlineColorIndex, SmallScroll, WindowNumber, Type and AutoFilterDateGrouping, and what a save keeps of each."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/window_settings.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def test_each_case_as_excel_answered(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='WindowProbe')
    got = str(app.run('WindowProbe.Probe', str(tmp_path))).split('|')
    assert got == [answer for _, _, _, answer in RECORD['cases']]
    for name, wanted in RECORD['views'].items():
        with zipfile.ZipFile(tmp_path / name) as package:
            sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
            workbook = package.read('xl/workbook.xml').decode('utf-8')
        view = re.search(r'<sheetView\b.*?(?:/>|</sheetView>)', sheet, re.DOTALL)
        outline = re.search(r'<sheetPr\b.*?(?:/>|</sheetPr>)', sheet, re.DOTALL)
        book_view = re.search(r'<workbookView\b[^>]*>', workbook)
        assert view is not None and view.group(0) == wanted['sheetView'], name
        assert (outline.group(0) if outline else '') == wanted['sheetPr'], name
        assert ('autoFilterDateGrouping="0"' in (book_view.group(0) if book_view else '')) \
            is wanted['autoFilterDateGrouping'], name


def test_the_settings_read_back(tmp_path: Path) -> None:
    """A file's sheetView and workbookView give the settings back as Excel saved them."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub Save(ByVal path As String)\nWith ActiveWindow\n.DisplayOutline = False\n.DisplayRuler = False\n'
                   '.GridlineColorIndex = 5\n.AutoFilterDateGrouping = False\nEnd With\n'
                   'ActiveWorkbook.SaveAs path, xlOpenXMLWorkbook\nEnd Sub\n', name='Saving')
    app.run('Saving.Save', str(tmp_path / 'set.xlsx'))
    reopened = ExcelApplication.open(tmp_path / 'set.xlsx')
    reopened.add_module('Function Read() As String\nWith ActiveWindow\nRead = .DisplayOutline & " " & .DisplayRuler & '
                        '" " & .DisplayWhitespace & " " & .GridlineColorIndex & " " & .GridlineColor & " " & '
                        '.AutoFilterDateGrouping\nEnd With\nEnd Function\n', name='Reading')
    assert reopened.run('Reading.Read') == 'False False True 5 16711680 False'
