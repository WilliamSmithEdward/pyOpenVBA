"""Which escaped characters Range.NumberFormat keeps escaped.

Excel parses a number format and writes it back, and a character escaped
with a backslash does not always stay escaped: \\T\\R\\U\\E reads back as
T\\RU\\E. The probe sets formats that escape each printable character
after a digit, before one, alone and twice, and a few whole words, and
reads NumberFormat back.

    python scripts/measure_format_escapes.py

writes tests/fixtures/format_escapes.json, which
tests/test_excel_format_escapes.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "format_escapes.json"

CHARACTERS = [chr(code) for code in range(33, 127)]
WORDS = ["\\T\\R\\U\\E", "\\F\\A\\L\\S\\E", "\\A\\B\\C", "\\a\\b\\c", "\\x\\y\\z", "0\\ \\k\\g", "0\\E\\+", "0\\e\\-0",
         "\\N\\o", "\\G\\B", "\\a\\/\\p", "\\A\\M\\/\\P\\M", "0\\.0", "\\d\\a\\y\\s", "0\\,0", "\\[0\\]", "\\_0", "\\*0",
         "\\?0", "\\#0", "0\\%", "\\@", "\\B\\2", "\\b\\1", "\\g\\g", "\\e\\e", "\\r\\R", "\\n\\N", "\\w\\W", "\\q\\Q"]


def codes() -> list[str]:
    out: list[str] = []
    for char in CHARACTERS:
        escaped = "\\" + char
        out += ["0" + escaped, escaped + "0", escaped, "0" + escaped + escaped]
    return out + WORDS


def vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def main() -> None:
    wanted = codes()
    # A workbook holds only so many custom formats, a little over 200, so each batch takes a new one.
    lines = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String, codes As Variant",
             "Dim i As Long", "Application.DisplayAlerts = False", "On Error Resume Next"]
    for index in range(0, len(wanted), 40):
        chunk = wanted[index:index + 40]
        lines += ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
                  "codes = Array(" + ", ".join(vba_text(code) for code in chunk) + ")"]
        lines += ["For i = 0 To UBound(codes)", "    Err.Clear", '    ws.Range("A1").NumberFormat = "General"',
                  '    ws.Range("A1").NumberFormat = codes(i)',
                  '    If Err.Number <> 0 Then out = out & "E" & Err.Number & "|~|" Else '
                  'out = out & ws.Range("A1").NumberFormat & "|~|"', "Next", "wb.Close False"]
    lines += ["Probe = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(lines) + "\n", "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = str(result.value).split("|~|")[: len(wanted)]
    OUT.write_text(json.dumps({"codes": dict(zip(wanted, answers, strict=True))}, indent=1) + "\n", encoding="utf-8")
    for code, answer in zip(wanted, answers, strict=True):
        if code != answer:
            print(f"{code!r:14} -> {answer!r}")


if __name__ == "__main__":
    main()
