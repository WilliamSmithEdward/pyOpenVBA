"""Measure the text Err.Description and Error() give for every error number from 1
to 1000, which interpreter/_values.py keeps as ERROR_TEXT.

Each probe raises a block of numbers under On Error Resume Next and records
``n=Err.Description|Error(n)`` for each. Each probe is a module whose
``Function Probe() As String`` the fixture records (see vba_probe_fixture.py).
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_error_text.json'
BLOCKS = [(1, 100), (101, 300), (301, 500), (501, 1000)]


def body(low: int, high: int) -> str:
    return (f'Dim n As Long, s As String\nOn Error Resume Next\nFor n = {low} To {high}\nErr.Clear\nErr.Raise n\n'
            's = s & n & "=" & Err.Description & "|" & Error(n) & ";"\nNext\nProbe = s')


CASES = {f"descriptions_{low}_{high}": body(low, high) for low, high in BLOCKS}


def macro(name: str) -> str:
    return '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT)


if __name__ == '__main__':
    main()
