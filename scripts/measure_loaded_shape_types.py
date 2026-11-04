"""The VBA type of each answer in tests/fixtures/loaded_shape_text.json and tests/fixtures/loaded_shape_foreign.json, whose answers are text and so do not tell a
Long 11 from a Double 11.

    python scripts/measure_loaded_shape_types.py

Adds a "types" list to each. tests/test_excel_loaded_shape_text.py replays it.
"""
from __future__ import annotations

import json
import shutil
import tempfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
#: A module past about 15,000 lines times out on the way in.
CHUNK = 1500


def reader(reads: list[str]) -> str:
    lines: list[str] = []
    names: list[str] = []
    for start in range(0, len(reads), 60):
        name = f'Part{start}'
        names.append(name)
        lines += [f'Private Function {name}(ws As Object) As String', '    Dim v As Variant, s As String', '    On Error Resume Next']
        for read in reads[start:start + 60]:
            lines += ['    s = s & "@@"', '    Err.Clear', f'    v = {read}', '    If Err.Number Then', '        s = s & "E"', '    Else', '        s = s & TypeName(v)', '    End If']
        lines += [f'    {name} = s', 'End Function']
    lines += ['Public Function Opened(ByVal path As String) As String', '    Dim wb As Workbook, s As String', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Open(path)']
    lines += [f'    s = s & {name}(wb.Worksheets(1))' for name in names]
    lines += ['    wb.Close False', '    Application.DisplayAlerts = True', '    Opened = s', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for stem in ('loaded_shape_text', 'loaded_shape_foreign'):
            path = FIXTURES / f'{stem}.json'
            record = json.loads(path.read_text(encoding='utf-8'))
            copy = Path(folder) / f'{stem}.xlsx'
            shutil.copyfile(FIXTURES / f'{stem}.xlsx', copy)
            types: list[str] = []
            for start in range(0, len(record['reads']), CHUNK):
                result = excel.run_vba(reader(record['reads'][start:start + CHUNK]), 'Opened', (str(copy),), timeout=280.0, line_numbers=False)
                assert result.ok and not result.dialogs, result
                types += str(result.value).split('@@')[1:]
            assert len(types) == len(record['reads'])
            record['types'] = types
            path.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
            print(stem, len(types))


if __name__ == '__main__':
    main()
