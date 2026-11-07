"""Live Excel against tests/fixtures/wrapped_round4_predictions.json, read after they were committed. Needs Excel at 96 DPI, checked in each call.

    python scripts/measure_wrapped_round4.py

Writes tests/fixtures/wrapped_round4.json. It uses no hyperlink, mail or send member.
"""
from __future__ import annotations

import collections
import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
SOURCE = r'''Public Function Go(ByVal text As String, ByVal font As String, ByVal size As Double, ByVal bold As Boolean, ByVal italic As Boolean, ByVal indent As Long, ByVal columns As String) As String
    Dim ws As Object, r As String, one As Variant
    Set ws = ActiveSheet
    ws.Cells.Clear
    With ws.Range("A1")
        .Value = "'" & Replace(text, "~", vbLf)
        .WrapText = True
        .Font.Name = font
        .Font.Size = size
        .Font.Bold = bold
        .Font.Italic = italic
        .IndentLevel = indent
    End With
    For Each one In Split(columns, ",")
        ws.Columns(1).ColumnWidth = Val(one) / 7
        ws.Rows(1).AutoFit
        r = r & "|" & CLng(ws.Columns(1).Width * 4 / 3) & "," & ws.Rows(1).RowHeight
    Next
    Go = ActiveSheet.StandardHeight & r
End Function
'''


def main() -> None:
    record = json.loads((FIXTURES / 'wrapped_round4_predictions.json').read_text(encoding='utf-8'))
    columns = ','.join(str(one) for one in record['columns'])
    curves = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for case in record['cases']:
            result = excel.run_vba(SOURCE, 'Go', (case['text'].replace('\n', '~'), case['font'], case['size'], case['bold'], case['italic'], case['indent'], columns), timeout=200.0, line_numbers=False)
            assert result.ok and not result.dialogs, (case['text'], result)
            height, *steps = str(result.value).split('|')
            assert height == '15', f'Excel is not at 96 DPI: StandardHeight reads {height}'
            curves.append({one.split(',')[0]: float(one.split(',')[1]) for one in steps})
    (FIXTURES / 'wrapped_round4.json').write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'curves': curves}, indent=0) + '\n', encoding='utf-8')
    good = bad = 0
    missed: collections.Counter[str] = collections.Counter()
    settings: collections.Counter[str] = collections.Counter()
    samples: dict[str, tuple[str, float, float]] = {}
    for case, curve in zip(record['cases'], curves, strict=True):
        for column, tall in curve.items():
            got = case['predicted'][column]
            if abs(got - tall) < 1e-6:
                good += 1
            else:
                bad += 1
                key = f"{case['font']} {case['size']} b{int(case['bold'])} i{int(case['italic'])} indent{case['indent']}"
                settings[key] += 1
                missed[f"{key} {case['text']!r}"] += 1
                samples.setdefault(f"{key} {case['text']!r}", (column, tall, got))
    print('good', good, 'off', bad)
    for key, count in settings.most_common():
        print('SET', count, key)
    for key, count in missed.most_common(30):
        print(count, key, samples[key])


if __name__ == '__main__':
    main()
