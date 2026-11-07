"""Where each held-out string in tests/fixtures/wrapped_heldout_predictions.json stays on one line in live Excel, read after the predictions were committed. Needs Excel at 96 DPI, checked at both ends of each chunk.

    python scripts/measure_wrapped_heldout.py

Writes tests/fixtures/wrapped_heldout.json. It uses no hyperlink, mail or send member.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
SOURCE = r'''Public Function Go(ByVal text As String, ByVal font As String, ByVal size As Double) As String
    Dim ws As Object, wide As Long, low As Long, high As Long, one As Double, under As Long, over As Long
    Set ws = ActiveSheet
    ws.Cells.Clear
    With ws.Range("A1")
        .Value = "'" & text
        .WrapText = True
        .Font.Name = font
        .Font.Size = size
    End With
    ws.Columns(1).ColumnWidth = 255
    ws.Rows(1).AutoFit
    one = ws.Rows(1).RowHeight
    low = 0: high = 255 * 7
    Do While high - low > 1
        wide = (low + high) \ 2
        ws.Columns(1).ColumnWidth = wide / 7
        ws.Rows(1).AutoFit
        If ws.Rows(1).RowHeight > one Then low = wide Else high = wide
    Loop
    ws.Columns(1).ColumnWidth = low / 7
    under = CLng(ws.Columns(1).Width * 4 / 3)
    ws.Columns(1).ColumnWidth = high / 7
    over = CLng(ws.Columns(1).Width * 4 / 3)
    Go = ActiveSheet.StandardHeight & "|" & under & "|" & over & "|" & one
End Function
'''


def main() -> None:
    record = json.loads((FIXTURES / 'wrapped_heldout_predictions.json').read_text(encoding='utf-8'))
    answers = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for case in record['cases']:
            result = excel.run_vba(SOURCE, 'Go', (case['text'], case['font'], case['size']), timeout=60.0, line_numbers=False)
            assert result.ok and not result.dialogs, (case, result)
            height, under, over, one = str(result.value).split('|')
            assert height == '15', f'Excel is not at 96 DPI: StandardHeight reads {height}'
            answers.append({'under': int(under), 'over': int(over), 'one_line': float(one)})
    (FIXTURES / 'wrapped_heldout.json').write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers}, indent=0) + '\n', encoding='utf-8')
    wrong = [(case['font'], case['size'], case['gdi'], case['predicted'], one['over']) for case, one in zip(record['cases'], answers, strict=True) if case['predicted'] != one['over']]
    print(len(answers), 'measured;', len(wrong), 'off prediction')
    for row in wrong[:60]:
        print(row)


if __name__ == '__main__':
    main()
