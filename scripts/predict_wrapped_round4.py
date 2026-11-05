"""The fourth held-out round for wrapped rows, from the built model and committed before Excel is asked: a line feed at a text's very end, a hyphen at its end, and the first letter's hyphen, in columns from a pixel up.

    python scripts/predict_wrapped_round4.py

Writes tests/fixtures/wrapped_round4_predictions.json.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
from pyopenvba.apps.excel import ExcelApplication  # noqa: E402
from pyopenvba.apps.excel._dimensions import row_points_read  # noqa: E402

OUT = ROOT / 'tests/fixtures/wrapped_round4_predictions.json'
SETTINGS = (('Calibri', 11, False, False, 0), ('Aptos Narrow', 11, False, False, 0), ('Calibri', 15, False, False, 0), ('Aptos Narrow', 9, False, True, 0), ('Calibri', 13, True, False, 1))
TEXTS = ('end\n', 'two words\n', 'x\n', 'abc-\n', 'a b\n\n', 'wide letters MMM\n', 'word-', 'long word-', 'ab-', 'abc--', 'x- ', 'trailing- ', 'WWWW-', 'iiii-', 'stop- go', 'm-n', 'k-lm-n', 'p-qr', 'z-', '-w', 'u-v-w-x', 'W-W', 'i-i', 'M-MM-M M-M')
COLUMNS = tuple(range(1, 50)) + tuple(range(50, 200, 7))
SOURCE = ('Public Sub Make(ByVal text As String, ByVal font As String, ByVal size As Double, ByVal bold As Boolean, ByVal italic As Boolean, ByVal indent As Long)\n    With ActiveSheet.Range("A1")\n'
          '        .Value = "\'" & Replace(text, "~", vbLf)\n        .WrapText = True\n        .Font.Name = font\n        .Font.Size = size\n        .Font.Bold = bold\n        .Font.Italic = italic\n        .IndentLevel = indent\n    End With\nEnd Sub\n')


def main() -> None:
    cases = []
    for font, size, bold, italic, indent in SETTINGS:
        for text in TEXTS:
            app = ExcelApplication()
            app.add_workbook()
            app.add_module(SOURCE, name='Probe')
            app.run('Probe.Make', text.replace('\n', '~'), font, size, bold, italic, indent)
            dims = app.workbook.sheets_[0].dims
            predicted: dict[str, float] = {}
            for column in range(1, 220):
                dims.set_column_width(1, column)
                dims.autofit_row(1)
                predicted[str(column)] = row_points_read(dims.row_quarters(1))
            cases.append({'font': font, 'size': size, 'bold': bold, 'italic': italic, 'indent': indent, 'text': text, 'predicted': predicted})
    OUT.write_text(json.dumps({'columns': COLUMNS, 'cases': cases}, indent=0) + '\n', encoding='utf-8')
    print(len(cases), 'curves')


if __name__ == '__main__':
    main()
