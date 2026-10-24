"""Which number format codes Excel refuses, for a conditional format's NumberFormat and a cell's alike.

measure_conditional_number_formats.py found Excel refusing one printable
character at a time in some places and not others. These codes try the
rules that answer suggests on whole codes: other characters outside
ASCII, a colon or a stroke where no digit is yet, fill and skip-width
characters with nothing to repeat, a text section that is not the last,
brackets left open, a unit after a space, exponents, and fractions of a
second. Each is a rule's NumberFormat and a cell's on its own row.

    python scripts/measure_number_format_refusals.py

writes tests/fixtures/conditional_number_formats/refusals.json, which
tests/test_excel_conditional_number_formats.py replays.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from measure_conditional_number_formats import vba  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "conditional_number_formats" / "refusals.json"

CODES = [
    # Characters outside ASCII: letters, signs, currencies and others.
    *("0" + char for char in "°µ²ü年¥™§½Ωßñ¢©"
      "«·×€£é₹฿"),
    "°0", "@°", "d°m", "£0", "@£", "d£m",
    # A colon or a stroke before a digit, after one, between them, and alone.
    "0:0", ":", "/", "0/8", "#/8", "0 /?", "? /?", "#:##", ":#", "$:0", ":@", "d:0", "0/0", "0 /", "?/ ?",
    "# ?/? :", ": # ?/?",
    # Fill and skip-width characters with nothing after them inside a code, and escapes.
    "0*;0", "0_;0", "0\\;0", "*", "_", "\\", "0* ", "0_ ", "0*", "@*-", "0;0*", "0;0_",
    # Text sections: twice, not last, among four and five sections, and with literals.
    "@;@", "0;@;0", "@;0", "0;0;0;@", "0;0;0;@;0", "0;0;0;0", "@ @", "@,", "@$", "@:", '"x"@', "@-", "@ x",
    "@ A", "@ AM/PM", "@ General", "General@", "General;@",
    # Brackets.
    "0[", "[0", "[Red", "[Red]", "0]", "[>0]0", "[x]0", "[$-409]", "[$-409]0", "[Red];0", "[h]0", "[h]", "[mm]",
    # Units after a space, and General beside others.
    "0 m", "0.00 m", "0 d", "m 0", "General m", "General 0", "0 General", "0 x", "0 K", "0 L",
    # Exponents.
    "E+0", "0E+", "0E+0", "0.0e+0", "#E-0", "0E0", "dE+m", "0.0E+0E+0", "0E-0",
    # The letters that are refused alone.
    "0 N", "Gen", "0 n", '0 "N"', "0\\N",
    # Fractions of a second, and digits in a date.
    "ss.0", "s.0", "d.0", "h.0", "mm.0", "[s].0", "ss.00", "ss.000", "ss.0000", "s.00 0", "h 0",
]
# Every Latin-1 and Latin Extended-A letter after a digit, since two of them, and no other character outside ASCII
# tried above, are refused; those not already listed.
CODES += [code for code in ("0" + chr(point) for point in range(0xC0, 0x180) if chr(point).isalpha())
          if code not in CODES]
# Each of those in a date and after the text placeholder, and a few other scripts.
CODES += ["dém", "@é", "dñm", "@ñ", "0α", "0д", "0א", "0ا", "0あ",
          "0ก", "0Ñ", "0É"]
# Accented letters whose base letter is a date code, in a date: whether each reads as that code.
CODES += ["d" + char + "m" for char in "ÝýèÈŠšĥĤĝĎďÿ"] + ["Ý", "ý", "è", "Š", "ĥ", "ĝ", "ď", "Ým", "dÝ"]
CODES = list(dict.fromkeys(CODES))


#: Codes to a procedure, which VBA refuses to compile past a size, each in a workbook of its own, as a workbook
#: holds about 250 custom formats.
BATCH = 100


def code() -> str:
    batches = range(0, len(CODES), BATCH)
    lines = ["Public Function Probe() As String", "Dim wb As Object, out As String", "Application.DisplayAlerts = False",
             *(line for start in batches for line in (
                 "Set wb = Workbooks.Add(xlWBATWorksheet)", f"out = out & Batch{start}(wb.Worksheets(1))",
                 "wb.Close False")),
             "Probe = out", "End Function"]
    for start in batches:
        lines += [f"Private Function Batch{start}(ws As Object) As String",
                  "Dim f As Object, out As String, a As String, b As String", "On Error Resume Next"]
        for row, one in enumerate(CODES[start:start + BATCH], start + 1):
            lines += ["Err.Clear", f'Set f = ws.Cells({row}, 2).FormatConditions.Add(xlCellValue, xlGreater, "0")',
                      f"f.NumberFormat = {vba(one)}",
                      'If Err.Number <> 0 Then a = "E" & Err.Number Else a = f.NumberFormat', "Err.Clear",
                      f"ws.Cells({row}, 1).NumberFormat = {vba(one)}",
                      f'If Err.Number <> 0 Then b = "E" & Err.Number Else b = ws.Cells({row}, 1).NumberFormat',
                      'out = out & a & "<=>" & b & "<<>>"']
        lines += ["On Error GoTo 0", f"Batch{start} = out", "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    assert len(set(CODES)) == len(CODES), "a code is listed twice"
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code(), "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    shown = [one.split("<=>") for one in str(result.value).split("<<>>")[:-1]]
    assert len(shown) == len(CODES) and all(len(one) == 2 for one in shown)
    record = {"codes": [{"code": one, "rule": rule, "cell": cell} for one, (rule, cell) in zip(CODES, shown, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for one in record["codes"]:
        print(f"{ascii(one['code']):22} rule {ascii(one['rule']):22} cell {ascii(one['cell'])}")


if __name__ == "__main__":
    main()
