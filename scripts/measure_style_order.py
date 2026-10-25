"""How Excel lists, finds and refuses cell styles named beyond ASCII: accents, ligatures, Greek and Cyrillic, case.

A macro adds a style under each of NAMES, in an order no sort gives, built from ChrW since a module's text cannot
hold them, and keeps the error each Add raised; looks each of FOUND up by name, keeping the error and the name the
style it found has; lists the Styles collection; and saves the workbook, whose styles.xml lists the cellStyle
elements in the order the record keeps. Each name comes back as its characters' codes, in hex.

    python scripts/measure_style_order.py

writes tests/fixtures/style_order.json, which tests/test_excel_sort_text.py replays.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "style_order.json"
#: Names a macro adds a style under: spellings of a word that differ in accents, case, ligatures and hyphens alone.
NAMES = ["zeta", "éclair", "Ärger", "eclair", "co-op", "резюме", "coop",
         "αβγ", "naïve", "Éclair", "straße", "strasse", "Æsop", "aesop", "naive",
         "Резюме", "œuvre", "oeuvre", "İstanbul", "istanbul",
         "Ωmega", "Zoë", "zoe", "co‐op", "re'sume", "resume", "ΑΒΓ",
         "Ångström", "angstrom", "ǅemal", "ǆemal", "Normal", "€uro", "–dash", "-dash",
         "dash", "arger", "ARGER"]
#: Names a macro looks a style up by.
FOUND = ["ÉCLAIR", "ECLAIR", "STRASSE", "STRAẞE", "ÆSOP", "AESOP", "ŒUVRE", "OEUVRE", "ISTANBUL",
         "İSTANBUL", "РЕЗЮМЕ", "ΑΒΓ", "ZOË", "ZOE",
         "ǄEMAL", "CO-OP", "COOP", "ÄRGER", "normal"]


def vba_text(text: str) -> str:
    return " & ".join(f"ChrW({ord(character)})" for character in text)


def code() -> str:
    lines = ["Private Function Coded(text As String) As String",
             "    Dim i As Long, out As String",
             "    For i = 1 To Len(text)",
             '        out = out & Hex(AscW(Mid(text, i, 1)) And &HFFFF&) & "."',
             "    Next",
             "    Coded = out",
             "End Function",
             "",
             "Public Function Measure(path As String) As String",
             "    Dim wb As Object, out As String, i As Long, found As String",
             "    Set wb = Workbooks.Add(xlWBATWorksheet)",
             "    On Error Resume Next"]
    for name in NAMES:
        lines += ["    Err.Clear", f"    wb.Styles.Add {vba_text(name)}", '    out = out & Err.Number & ","']
    lines += ['    out = out & "|"']
    for name in FOUND:
        lines += ["    Err.Clear", '    found = ""', f"    found = wb.Styles({vba_text(name)}).Name",
                  '    out = out & Err.Number & ":" & Coded(found) & ","']
    lines += ["    On Error GoTo 0",
              '    out = out & "|"',
              "    For i = 1 To wb.Styles.Count",
              '        out = out & Coded(wb.Styles(i).Name) & ","',
              "    Next",
              "    Application.DisplayAlerts = False",
              "    wb.SaveAs path, 51",
              "    wb.Close False",
              "    Measure = out",
              "End Function"]
    return "\n".join(lines) + "\n"


def decoded(coded: str) -> str:
    return "".join(chr(int(part, 16)) for part in coded.split(".") if part)


def saved_order(path: Path) -> list[str]:
    """The names of the cellStyle elements of a saved workbook's styles.xml, in their order."""
    import html

    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
    return [html.unescape(found) for found in re.findall(r'<cellStyle name="([^"]*)"', styles)]


def main() -> None:
    source = code()
    with tempfile.TemporaryDirectory() as folder:
        saved = Path(folder) / "style_order.xlsx"
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(source, "Measure", [str(saved)], timeout=600.0)
            assert result.ok, f"{result.outcome}: {result.message} {result.error}"
        written = saved_order(saved)
    added, found, listed = str(result.value).split("|")
    record = {"names": NAMES, "found": FOUND, "code": source,
              "added": [int(number) for number in added.split(",") if number],
              "lookups": [[int(number), decoded(name)] for number, name in
                          (one.split(":") for one in found.split(",") if one)],
              "listed": [decoded(name) for name in listed.split(",") if name],
              "saved": written}
    OUT.write_text(json.dumps(record, indent=0, ensure_ascii=True) + "\n", encoding="utf-8")
    print("added:", record["added"])
    print("lookups:", [(number, ascii(name)) for number, name in record["lookups"]])
    print("listed:", [ascii(name) for name in record["listed"]])
    print("saved:", [ascii(name) for name in written])


if __name__ == "__main__":
    main()
