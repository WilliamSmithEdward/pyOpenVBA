"""Whether Now and Time carry anything below a second, read in live Excel into a Date and into a Variant, at moments a fraction of a second apart.

The clock moves, so what is recorded is how far each read lies past its whole second, in millionths of a second, rounded. Excel answers 0 for
every one; a model that kept the clock's fraction would answer it.

    python scripts/measure_now_whole_seconds.py

writes tests/fixtures/now_whole_seconds.json.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/now_whole_seconds.json'
SAMPLES = 7
SOURCE = '''Private Function Past(ByVal serial As Double) As Long
    Dim seconds As Double
    seconds = serial * 86400#
    Past = CLng((seconds - Int(seconds + 0.0000005)) * 1000000#)
End Function
Function Probe() As String
    Dim i As Long, s As String, t As Single, d As DECLARED, e As DECLARED
    For i = 1 To SAMPLES
        t = Timer
        Do While Timer - t < 0.13
        Loop
        d = Now
        e = Time
        s = s & Past(CDbl(d)) & "," & Past(CDbl(e)) & "," & Past(CDbl(Now) - Fix(CDbl(Now))) & "," & (Second(Now) = Int(Second(Now))) & ";"
    Next
    Probe = s
End Function
'''


def macro(declared: str) -> str:
    return SOURCE.replace('DECLARED', declared).replace('SAMPLES', str(SAMPLES))


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    answers: dict[str, str] = {}
    for declared in ('Date', 'Variant'):
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(macro(declared), 'Probe', timeout=60.0, line_numbers=False)
        if not result.ok or result.dialogs:
            raise SystemExit(f'{declared} stopped: {result.message}')
        answers[declared] = str(result.value)
        print(declared, answers[declared], flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'samples': SAMPLES, 'answers': answers}, indent=1) + '\n',
                   encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
