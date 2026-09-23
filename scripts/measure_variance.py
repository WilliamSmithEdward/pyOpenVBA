"""The exact doubles Excel's SUM, AVERAGE, VAR, VARP, STDEV and STDEVP answer, bit for bit.

Excel shows fifteen digits and snaps a difference near zero to zero, so
neither the display nor a subtraction pins its last bits. This probe
writes each set of numbers to a sheet, has each function work on the
range, and reads the answer's eight bytes through LSet into two Longs.
The sets are fixed cases and seeded random ones: small decimals, large
offsets, mixed signs and magnitudes.

    python scripts/measure_variance.py

writes tests/fixtures/variance.json, which tests/test_formula_variance.py replays.
"""

from __future__ import annotations

import json
import random
import struct
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "variance.json"
#: Each function measured, and its formula over the range {r}.
FORMULAS = {"SUM": "SUM({r})", "AVERAGE": "AVERAGE({r})", "VAR": "VAR({r})", "VARP": "VARP({r})",
            "STDEV": "STDEV({r})", "STDEVP": "STDEVP({r})", "DEVSQ": "DEVSQ({r})", "SUMSQ": "SUMSQ({r})",
            "SUMIF": 'SUMIF({r},""<>zzz"")', "AVERAGEIF": 'AVERAGEIF({r},""<>zzz"")',
            "SUMPRODUCT": "SUMPRODUCT({r},{r})"}
FUNCTIONS = tuple(FORMULAS)

HELPER = '''Private Type Bits
    Value As Double
End Type

Private Type Halves
    Low As Long
    High As Long
End Type

Private Function Hex64(v As Variant) As String
    Dim b As Bits, h As Halves
    If IsError(v) Then
        Hex64 = "E" & CStr(CLng(v))
        Exit Function
    End If
    b.Value = CDbl(v)
    LSet h = b
    Hex64 = Right$("00000000" & Hex$(h.High), 8) & Right$("00000000" & Hex$(h.Low), 8)
End Function
'''


def data_sets() -> list[list[float]]:
    fixed = [[0.1, 0.2, 0.3], [0.3, 0.2, 0.1], [1.0, 6.0, 7.0, 8.0], [10.1, 10.2, 10.3], [100.1, 100.2, 100.3],
             [0.1, 0.2, 0.3, 0.7, 1.1, 1.3], [2.5, 3.5, 4.5, 9.5], [1e9 + 1, 1e9 + 2, 1e9 + 3],
             [1e15 + 1, 1e15 + 2, 1e15 + 3], [0.1, 0.7], [1.1, 2.2, 3.3], [-0.5, 0.25, 1e-3, 7.75]]
    rng = random.Random(20260923)
    for _ in range(40):
        size = rng.randint(2, 10)
        scale = rng.choice((1e-3, 0.1, 1.0, 10.0, 1e3, 1e6))
        offset = rng.choice((0.0, 0.0, 1.0, 100.0, 1e6, 1e9))
        # Numbers as a person types them: a few decimals, so the doubles carry binary noise.
        fixed.append([round(offset + rng.uniform(-1, 1) * scale, rng.randint(1, 6)) for _ in range(size)])
    # Where Excel turns from one pass to two: small variances round 1e-7, and spreads whose share of the
    # sum of squares -- the one-pass formula's cancellation -- steps from 0.004 to 0.12, at three scales.
    base = [rng.uniform(-1, 1) for _ in range(6)]
    for variance in (1e-9, 3e-8, 6e-8, 9e-8, 1.1e-7, 1.5e-7, 3e-7, 1e-6):
        fixed.append([float("%.6g" % (v * variance ** 0.5 * 1.7)) for v in base])
    for share in (0.004, 0.006, 0.01, 0.02, 0.03, 0.05, 0.07, 0.09, 0.11):
        for scale in (1e-3, 1.0, 1e3):
            spread = (share / (1 - share)) ** 0.5
            fixed.append([float("%.8g" % ((1 + v * spread) * scale)) for v in base])
    return fixed


def main() -> None:
    sets = data_sets()
    lines = [HELPER]
    for index, values in enumerate(sets):
        # One function a set: a single procedure holding them all is too large for VBA.
        lines += [f"Private Function Set{index}(ws As Object) As String", "ws.Cells.ClearContents"]
        for row, value in enumerate(values, 1):
            lines.append(f"ws.Cells({row}, 1).Value = {value!r}")
        for column, name in enumerate(FUNCTIONS, 3):
            formula = FORMULAS[name].format(r=f"A1:A{len(values)}")
            lines.append(f'ws.Cells(1, {column}).Formula = "={formula}"')
        # What the cells hold, so a typed number is checked too.
        cells = " & \",\" & ".join(f"Hex64(ws.Cells({row}, 1).Value)" for row in range(1, len(values) + 1))
        answers = " & \",\" & ".join(f"Hex64(ws.Cells(1, {column}).Value)" for column in range(3, 3 + len(FUNCTIONS)))
        lines += [f'Set{index} = {cells} & ";" & {answers}', "End Function"]
    lines += ["Public Function Build() As String", "Dim ws As Object, out As String",
              "Set ws = ActiveWorkbook.Worksheets(1)",
              *(f'out = out & Set{index}(ws) & "|"' for index in range(len(sets))), "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(lines) + "\n", "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    records = []
    for values, answer in zip(sets, str(result.value).split("|")[: len(sets)], strict=True):
        cells, answers = answer.split(";")
        stored = [struct.unpack(">d", bytes.fromhex(one))[0] for one in cells.split(",")]
        records.append({"typed": values, "stored": stored,
                        "answers": dict(zip(FUNCTIONS, answers.split(","), strict=True))})
    OUT.write_text(json.dumps({"formulas": FORMULAS, "sets": records}, indent=1) + "\n", encoding="utf-8")
    print(len(records), "sets")


if __name__ == "__main__":
    main()
