"""Which names Excel's Names.Add takes: those that start like an R1C1 reference, with every ending, and others that
read as a cell or start oddly.

Each candidate is added to the workbook and deleted again; a name Excel
refuses raises error 1004.

    python scripts/measure_name_validity.py

writes tests/fixtures/name_validity.json, which tests/test_excel_name_validity.py
replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/name_validity.json'

PREFIXES = ['R', 'C', 'RC', 'CR', 'R1', 'C1', 'R1C1', 'Rx', 'Cx', 'Ra', 'Ca', 'r', 'rc', 'R12', 'C12', 'RR', 'CC',
            'Z', 'S', 'Z1S1', 'L', 'K', 'L1C1']
SUFFIXES = ['', '_', 'x', '1', 'x1', '_1', 'xy', '1x', '.', '.1', '?', '?1', '\\']
OTHERS = ['Rate1', 'Region2', 'Cost1', 'Cat', 'Row1', 'Col1', 'Rc', 'rC1', 'Rb12c', 'Cab1', 'Rxx1', 'Ra1', 'R1C1C1',
          'RC1C', 'Rcolumn1', 'Cell1', 'A1x', 'A1_1', 'AB1C', 'XFD1x', 'XFE1', 'A0', 'A1048577', 'AAAA1', 'a?b', '_A1',
          '_3rd', '3rd_', '.dot_', '_.dot', 'x?', '?x', 'R1C.', 'RC1.', 'R1C1_', 'A1_']
CANDIDATES = [prefix + suffix for prefix in PREFIXES for suffix in SUFFIXES] + OTHERS

MACRO = '''Public Function Took(ByVal nm As String) As String
    On Error GoTo Failed
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="=1"
    ThisWorkbook.Names(nm).Delete
    Took = "1"
    Exit Function
Failed:
    Took = "0"
End Function
'''


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        taken: dict[str, bool] = {}
        for name in CANDIDATES:
            result = excel.run_vba(MACRO, 'Took', (name,), timeout=60.0)
            assert result.ok, result
            taken[name] = result.value == '1'
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'taken': taken}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    print(sum(taken.values()), 'of', len(taken), 'taken')


if __name__ == '__main__':
    main()
