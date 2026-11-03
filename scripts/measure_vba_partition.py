"""QBColor and Partition, each case alone so that an error is recorded as the answer.

    python scripts/measure_vba_partition.py

writes tests/fixtures/vba_partition.json, which tests/test_vba_partition.py replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_partition.json'
QB = [str(n) for n in range(16)] + ['16', '-1', '1.5', '2.5', '"3"', '"x"', 'Null', 'Empty', 'True', '15.5']
PARTITION = ['5, 0, 100, 10', '0, 0, 100, 10', '9, 0, 100, 10', '10, 0, 100, 10', '100, 0, 100, 10', '101, 0, 100, 10', '-1, 0, 100, 10', '1000, 0, 100, 10',
             '50, 0, 99, 5', '99, 0, 99, 5', '95, 0, 99, 5', '100, 0, 99, 5', '-1, 0, 99, 5', '5, 0, 9, 1', '10, 0, 9, 1', '-1, 0, 9, 1', '5, 1, 10, 3', '10, 1, 10, 3', '1, 1, 10, 3',
             '-5, -10, 10, 5', '-10, -10, 10, 5', '10, -10, 10, 5', '11, -10, 10, 5', '-11, -10, 10, 5', '0, 0, 0, 1', '1, 0, 0, 1', '-1, 0, 0, 1', '7, 5, 9, 100', '5, 5, 9, 100',
             '10, 5, 9, 100', '4, 5, 9, 100', '5, 0, 10000, 100', '10000, 0, 10000, 100', '10001, 0, 10000, 100', '5, 0, 100000, 1000', '5.5, 0, 10, 2', '5.4, 0, 10, 2', '2.5, 0, 10, 1',
             '3.5, 0, 10, 1', '5, 0.5, 10, 2', '5, 0, 10.5, 2', '5, 0, 10, 2.5', '5, 0, 10, 1.5', '5, 0, 10, 0', '5, 0, 10, -1', '5, 10, 0, 1', '5, 5, 5, 1', '5, -1, 10, 1', '"5", "0", "10", "2"',
             '"x", 0, 10, 2', 'Null, 0, 10, 2', '5, Null, 10, 2', '5, 0, Null, 2', '5, 0, 10, Null', 'Empty, 0, 10, 2', 'True, -5, 5, 1', '2147483647, 0, 2147483647, 1000000',
             '5, 0, 2147483647, 1000000', '5, -2147483647, 2147483647, 1000', '5, 0, 2147483648#, 1000', '99999999999#, 0, 10, 1']
CASES = {
    **{f'qbcolor_{index}': f's = TypeName(QBColor({text})) & "|" & QBColor({text})' for index, text in enumerate(QB)},
    **{f'partition_{index}': f'v = Partition({text})\ns = TypeName(v) & "[" & v & "]"' for index, text in enumerate(PARTITION)},
}


def macro(name: str) -> str:
    return "Function Probe() As String\nDim s As String, v As Variant\n" + CASES[name] + "\nProbe = s\nEnd Function\n"


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
