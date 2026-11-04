"""How Excel spells a number in a cell's <v> when it saves: digits below 1E+16, then 15 or 17 significant digits with an E.

    python scripts/measure_saved_numbers.py

writes tests/fixtures/saved_numbers.json, which tests/test_excel_saved_numbers.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/saved_numbers.json'
VALUES = ['999999999999999', '1E+15', '1000000000000001', '9007199254740991#', '1E+16', '12345678901234567#', '1E+17',
          '123456789012345678#', '1E+18', '1E+19', '1E+20', '1E+21', '1.5E+21', '1E+22', '1.23456789012345E+25', '1E+100',
          '1.5E+300', '1E+308', '1.79769313486231E+308', '-1E+15', '-1E+16', '-1E+20', '-1E+21', '-1E+308', '1E-5',
          '1E-300', '1.5E-10', '0.1', '1.1', '0.30000000000000004#', '1/3', '2/3', '1E+14', '123456.789', '-1.1', '0', '1',
          '-1', '20.1', '19.9', '15.2', '0.001', '0.0001', '0.00001', '1234567.1234567', '0.5', '0.01', '0.02', '0.0125',
          '0.00999', '0.005', '0.0025', '0.0012345', '0.002', '-0.001', '0.1234567']


def main() -> None:
    with tempfile.TemporaryDirectory() as folder:
        path = str(Path(folder) / 'numbers.xlsx')
        writes = ''.join(f'    ws.Cells({row}, 1).Value = {value}\n' for row, value in enumerate(VALUES, 1))
        source = ('Public Function Probe(ByVal path As String) As String\n    Dim b As Object, ws As Object\n'
                  f'    Set b = Workbooks.Add(xlWBATWorksheet)\n    Set ws = b.Worksheets(1)\n{writes}'
                  '    b.SaveAs path, xlOpenXMLWorkbook\n    b.Close False\nEnd Function\n')
        with ExcelSession(HarnessConfig(lock_wait_s=600.0, exclusive=False)) as excel:
            excel.new_document()
            result = excel.run_vba(source, 'Probe', (path,), timeout=60.0)
            assert result.ok and not result.dialogs, result
        with zipfile.ZipFile(path) as package:
            sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = dict(re.findall(r'<c r="A(\d+)"[^>]*><v>([^<]*)</v>', sheet))
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(),
                               'spelled': [found[str(row)] for row in range(1, len(VALUES) + 1)]}, indent=0) + '\n',
                   encoding='utf-8')


if __name__ == '__main__':
    main()
