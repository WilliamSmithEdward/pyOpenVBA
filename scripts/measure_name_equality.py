"""Which sheet names and defined names Excel takes for one: case, accents, ligatures, a sharp s, Greek and Cyrillic.

A macro names a new sheet each of SHEETS in turn, in an order no sort gives, keeping the error each name raised, and
looks each of SHEETS_FOUND up in the Worksheets collection; then adds a workbook name each of NAMES, each referring
to its place in the list, and looks each of NAMES_FOUND up in the Names collection. The sheets and the names come
back in the collections' order, a name with what it refers to. Each name comes back as its characters' codes, in
hex, since a module's text cannot hold them and the macro builds them from ChrW.

    python scripts/measure_name_equality.py

writes tests/fixtures/name_equality.json, which tests/test_excel_name_equality.py replays. Two more probes
(rules_code) tell the candidate rules apart over PAIRS, and over MARKS, the pairs with a combining accent.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "name_equality.json"
#: Names a macro gives sheets: spellings of a word that differ in case, accents, ligatures and hyphens alone.
SHEETS = ["\u00c6sop", "aesop", "stra\u00dfe", "STRASSE", "\u0153uvre", "oeuvre", "\u0130stanbul", "istanbul",
          "\u00e9clair", "Eclair", "\u00c9CLAIR", "co-op", "coop", "\u0440\u0435\u0437\u044e\u043c\u0435",
          "\u0420\u0415\u0417\u042e\u041c\u0415", "\u03b1\u03b2\u03b3", "\u0391\u0392\u0393", "\u01c5emal",
          "\u01c6emal", "Zo\u00eb", "zoe", "na\u00efve", "NAIVE", "\u00c5ngstr\u00f6m", "angstrom"]
#: Names a macro looks a sheet up by.
SHEETS_FOUND = ["AESOP", "\u00c6SOP", "STRASSE", "STRA\u1e9eE", "OEUVRE", "\u0152UVRE", "ISTANBUL", "\u0130STANBUL",
                "ECLAIR", "\u00c9CLAIR", "COOP", "CO-OP", "\u0420\u0415\u0417\u042e\u041c\u0415", "\u0391\u0392\u0393",
                "\u01c4EMAL", "ZO\u00cb", "ZOE", "NA\u00cfVE", "\u00c5NGSTR\u00d6M"]
#: Names a macro adds to the workbook.
NAMES = ["\u00c6sop", "aesop", "stra\u00dfe", "STRASSE", "\u0153uvre", "oeuvre", "\u0130stanbul", "istanbul",
         "\u00e9clair", "Eclair", "\u00c9CLAIR", "co.op", "coop", "\u0440\u0435\u0437\u044e\u043c\u0435",
         "\u0420\u0415\u0417\u042e\u041c\u0415", "\u03b1\u03b2\u03b3", "\u0391\u0392\u0393", "\u01c5emal",
         "\u01c6emal", "Zo\u00eb", "zoe", "na\u00efve", "NAIVE", "\u00c5ngstr\u00f6m", "angstrom", "a_b", "ab"]
#: Names a macro looks a workbook name up by.
NAMES_FOUND = ["AESOP", "\u00c6SOP", "STRASSE", "STRA\u1e9eE", "OEUVRE", "\u0152UVRE", "ISTANBUL", "\u0130STANBUL",
               "ECLAIR", "\u00c9CLAIR", "COOP", "CO.OP", "\u0420\u0415\u0417\u042e\u041c\u0415",
               "\u0391\u0392\u0393", "\u01c4EMAL", "ZO\u00cb", "ZOE", "NA\u00cfVE", "\u00c5NGSTR\u00d6M", "A_B"]


#: Pairs of spellings that one rule or another takes for one: (first, second, what tells them apart).
PAIRS = [("K", "\u212a", "Kelvin sign: lower case"), ("I", "\u0131", "dotless i: upper case"),
         ("S", "\u017f", "long s: upper case"), ("\u03c3", "\u03c2", "final sigma: upper case"),
         ("\u00c5", "\u212b", "angstrom sign: lower case"), ("\u00df", "\u1e9e", "capital sharp s: lower case"),
         ("\u03bc", "\u00b5", "micro sign: upper case"), ("\u03b8", "\u03d1", "theta symbol: upper case"),
         ("\u03a9", "\u2126", "ohm sign: lower case"), ("\u01c4", "\u01c5", "titlecase DZ: either case"),
         ("i", "\u0130", "dotted capital I: linguistic lower case"), ("ae", "\u00e6", "ash: word sort"),
         ("ss", "\u00df", "sharp s: word sort"), ("e", "\u00e9", "an accent: none"),
         ("\u00e6ss", "ae\u00df", "ash and ss against ae and sharp s: as long")]
#: Pairs with a combining accent, which the model does not read in a name.
MARKS = [("\u00e9", "e\u0301", "a precomposed accent against a combining one"),
         ("e\u0301", "E\u0301", "a combining accent in either case")]
#: A sheet named with the ash alone, that formulas then call by its letters.
LONE = "q1\u00c6sop"


def vba_text(text: str) -> str:
    return " & ".join(f"ChrW({ord(character)})" for character in text)


def rules_code(pairs: list[tuple[str, str, str]], macro: str, lone: bool) -> str:
    """For each pair: a sheet named with the first spelling and one renamed to the second, each holding its place in
    A1, and both looked up; two more named the other way round, second spelling first, looked up by the first and
    read through Evaluate; a workbook name added under each spelling, referring to 10 times the pair's place plus 1
    and plus 2; and a name under the first spelling alone, that a cell's formula calls by the second. With lone,
    a sheet named with the ash alone, that formulas and the Worksheets collection call by its letters."""
    lines = ["Private Function Coded(text As String) As String",
             "    Dim i As Long, out As String",
             "    For i = 1 To Len(text)",
             '        out = out & Hex(AscW(Mid(text, i, 1)) And &HFFFF&) & "."',
             "    Next",
             "    Coded = out",
             "End Function",
             "",
             "Private Function Placed(wb As Object, text As String, place As Long) As Long",
             "    Dim ws As Object",
             "    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
             "    On Error Resume Next",
             "    ws.Name = text",
             "    Placed = Err.Number",
             "    On Error GoTo 0",
             "    If Placed <> 0 Then",
             "        Application.DisplayAlerts = False",
             "        ws.Delete",
             "        Application.DisplayAlerts = True",
             "    Else",
             "        ws.Range(\"A1\").Value = place",
             "    End If",
             "End Function",
             "",
             "Private Function Found(wb As Object, text As String) As String",
             "    On Error Resume Next",
             "    Found = Coded(wb.Worksheets(text).Name)",
             "    If Err.Number <> 0 Then Found = \"E\" & Err.Number",
             "End Function",
             "",
             "Private Function NameFound(wb As Object, text As String) As String",
             "    On Error Resume Next",
             "    NameFound = Coded(wb.Names(text).Name) & \"=\" & Coded(wb.Names(text).RefersTo)",
             "    If Err.Number <> 0 Then NameFound = \"E\" & Err.Number",
             "End Function",
             "",
             "Private Function Shown(v As Variant) As String",
             "    If IsError(v) Then",
             "        Shown = \"E\" & CStr(CLng(v))",
             "    Else",
             "        Shown = CStr(v)",
             "    End If",
             "End Function",
             "",
             f"Public Function {macro}() As String",
             "    Dim wb As Object, out As String, i As Long, first As String, second As String, ws As Object",
             "    Set wb = Workbooks.Add(xlWBATWorksheet)",
             "    Set ws = wb.Worksheets(1)"]
    for place, (first, second, _) in enumerate(pairs, start=1):
        lines += [f"    first = {vba_text(first)}", f"    second = {vba_text(second)}",
                  f'    out = out & Placed(wb, "s{place}" & first, {place * 10 + 1}) & ";"',
                  f'    out = out & Placed(wb, "s{place}" & second, {place * 10 + 2}) & ";"',
                  f'    out = out & Found(wb, "s{place}" & first) & ";" & Found(wb, "s{place}" & second) & ";"',
                  f'    out = out & Placed(wb, "r{place}" & second, {place * 10 + 3}) & ";"',
                  f'    out = out & Placed(wb, "r{place}" & first, {place * 10 + 4}) & ";"',
                  f'    out = out & Found(wb, "r{place}" & first) & ";"',
                  f"    out = out & Shown(Application.Evaluate(\"'r{place}\" & first & \"'!A1\")) & \";\"",
                  "    On Error Resume Next",
                  f'    wb.Names.Add Name:="n{place}" & first, RefersTo:="={place * 10 + 1}"',
                  '    out = out & Err.Number & ";"',
                  "    Err.Clear",
                  f'    wb.Names.Add Name:="n{place}" & second, RefersTo:="={place * 10 + 2}"',
                  '    out = out & Err.Number & ";"',
                  "    Err.Clear",
                  f'    wb.Names.Add Name:="f{place}" & first, RefersTo:="={place}"',
                  "    On Error GoTo 0",
                  f'    out = out & NameFound(wb, "n{place}" & first) & ";" & NameFound(wb, "n{place}" & second)',
                  '    out = out & "|"']
    if lone:
        lines += [f'    out = out & Placed(wb, {vba_text(LONE)}, 777) & ";"',
                  "    out = out & Shown(Application.Evaluate(\"'q1AESOP'!A1\")) & \";\"",
                  "    out = out & Shown(Application.Evaluate(\"'q1aesop'!A1\")) & \";\"",
                  '    out = out & Found(wb, "q1AESOP") & "|"']
    # The names are listed before a formula calls one no name has, for which Excel adds a hidden _xlfn.SINGLE.
    lines += ["    For i = 1 To wb.Names.Count",
              '        out = out & Coded(wb.Names(i).Name) & "=" & Coded(wb.Names(i).RefersTo) & ","',
              "    Next",
              '    out = out & "|"',
              "    On Error Resume Next"]
    for place, (_, second, _) in enumerate(pairs, start=1):
        lines += ["    Err.Clear", f"    second = {vba_text(second)}",
                  f'    ws.Range("B{place}").Formula = "=f{place}" & second',
                  '    out = out & Err.Number & ";"',
                  f'    out = out & Shown(ws.Range("B{place}").Value) & ";" & Coded(ws.Range("B{place}").Formula) & ","']
    lines += ["    On Error GoTo 0",
              "    wb.Close False",
              f"    {macro} = out",
              "End Function"]
    return "\n".join(lines) + "\n"


def code() -> str:
    lines = ["Private Function Coded(text As String) As String",
             "    Dim i As Long, out As String",
             "    For i = 1 To Len(text)",
             '        out = out & Hex(AscW(Mid(text, i, 1)) And &HFFFF&) & "."',
             "    Next",
             "    Coded = out",
             "End Function",
             "",
             "Private Function Named(wb As Object, ws As Object, text As String) As Long",
             "    On Error Resume Next",
             "    ws.Name = text",
             "    Named = Err.Number",
             "    If Err.Number <> 0 Then",
             "        Application.DisplayAlerts = False",
             "        ws.Delete",
             "        Application.DisplayAlerts = True",
             "    End If",
             "End Function",
             "",
             "Public Function Measure() As String",
             "    Dim wb As Object, out As String, i As Long, found As String, ws As Object",
             "    Set wb = Workbooks.Add(xlWBATWorksheet)",
             "    On Error Resume Next"]
    for name in SHEETS:
        lines += ["    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'    out = out & Named(wb, ws, {vba_text(name)}) & ","']
    lines += ['    out = out & "|"']
    for name in SHEETS_FOUND:
        lines += ["    Err.Clear", '    found = ""', f"    found = wb.Worksheets({vba_text(name)}).Name",
                  '    out = out & Err.Number & ":" & Coded(found) & ","']
    lines += ['    out = out & "|"',
              "    For i = 1 To wb.Worksheets.Count",
              '        out = out & Coded(wb.Worksheets(i).Name) & ","',
              "    Next",
              '    out = out & "|"']
    for place, name in enumerate(NAMES, start=1):
        lines += ["    Err.Clear", f'    wb.Names.Add Name:={vba_text(name)}, RefersTo:="={place}"',
                  '    out = out & Err.Number & ","']
    lines += ['    out = out & "|"']
    for name in NAMES_FOUND:
        lines += ["    Err.Clear", '    found = ""',
                  f'    found = Coded(wb.Names({vba_text(name)}).Name) & ":" & '
                  f'Coded(wb.Names({vba_text(name)}).RefersTo)',
                  '    out = out & Err.Number & ":" & found & ","']
    lines += ['    out = out & "|"',
              "    For i = 1 To wb.Names.Count",
              '        out = out & Coded(wb.Names(i).Name) & ":" & Coded(wb.Names(i).RefersTo) & ","',
              "    Next",
              "    On Error GoTo 0",
              "    wb.Close False",
              "    Measure = out",
              "End Function"]
    return "\n".join(lines) + "\n"


def decoded(coded: str) -> str:
    return "".join(chr(int(part, 16)) for part in coded.split(".") if part)


def parse(value: str) -> dict[str, object]:
    """The macro's answer as the record keeps it."""
    added, sheets_found, sheets, names_added, names_found, names = value.split("|")
    return {
        "sheets_added": [int(number) for number in added.split(",") if number],
        "sheets_found": [[int(number), decoded(name)] for number, name in
                         (one.split(":") for one in sheets_found.split(",") if one)],
        "sheets": [decoded(name) for name in sheets.split(",") if name],
        "names_added": [int(number) for number in names_added.split(",") if number],
        "names_found": [[int(parts[0]), decoded(parts[1]) if len(parts) > 1 else "",
                         decoded(parts[2]) if len(parts) > 2 else ""]
                        for parts in (one.split(":") for one in names_found.split(",") if one)],
        "names": [[decoded(name), decoded(refers)] for name, refers in
                  (one.split(":") for one in names.split(",") if one)],
    }


def _shown(coded: str) -> str:
    """A name the probe sent coded, or the error it sent as E and its number."""
    return coded if coded.startswith("E") else decoded(coded)


def _named(coded: str) -> list[str]:
    return [coded] if coded.startswith("E") else [decoded(part) for part in coded.split("=")]


def parse_rules(value: str, lone: bool) -> dict[str, object]:
    """What a rules probe saw for each pair, for the lone sheet where it has one, the workbook's names, and each
    pair's formula."""
    *pairs, names, formulas = value.split("|")
    alone = pairs.pop() if lone else ""
    seen: list[dict[str, object]] = []
    for one, written in zip(pairs, (one for one in formulas.split(",") if one), strict=True):
        part, formula = one.split(";"), written.split(";")
        seen.append({"sheet_first": int(part[0]), "sheet_second": int(part[1]), "found_first": _shown(part[2]),
                     "found_second": _shown(part[3]), "reversed_second": int(part[4]),
                     "reversed_first": int(part[5]), "reversed_found": _shown(part[6]), "evaluated": part[7],
                     "name_first": int(part[8]), "name_second": int(part[9]),
                     "name_found_first": _named(part[10]), "name_found_second": _named(part[11]),
                     "formula_error": int(formula[0]), "formula_value": formula[1], "formula": decoded(formula[2])})
    found: dict[str, object] = {"rules": seen, "names": [[decoded(part) for part in one.split("=")]
                                                         for one in names.split(",") if one]}
    if lone:
        placed, upper, lower, sheet = alone.split(";")
        found["lone"] = {"placed": int(placed), "evaluated_upper": upper, "evaluated_lower": lower,
                         "found": _shown(sheet)}
    return found


def main() -> None:
    source = code()
    rules, marks = rules_code(PAIRS, "Rules", lone=True), rules_code(MARKS, "Marks", lone=False)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, "Measure", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
        ruled = excel.run_vba(rules, "Rules", timeout=600.0)
        assert ruled.ok, f"{ruled.outcome}: {ruled.message} {ruled.error}"
        marked = excel.run_vba(marks, "Marks", timeout=600.0)
        assert marked.ok, f"{marked.outcome}: {marked.message} {marked.error}"
    record: dict[str, object] = {"sheet_names": SHEETS, "sheet_lookups": SHEETS_FOUND, "names": NAMES,
                                 "name_lookups": NAMES_FOUND, "code": source, "raw": str(result.value)}
    record.update(parse(str(result.value)))
    record["rules"] = {"pairs": PAIRS, "code": rules, "raw": str(ruled.value), **parse_rules(str(ruled.value), True)}
    record["marks"] = {"pairs": MARKS, "code": marks, "raw": str(marked.value),
                       **parse_rules(str(marked.value), False)}
    OUT.write_text(json.dumps(record, indent=0, ensure_ascii=True) + "\n", encoding="utf-8")
    for key in ("sheets_added", "sheets_found", "sheets", "names_added", "names_found", "names"):
        print(key, ascii(record[key]))
    for part in ("rules", "marks"):
        probe: dict[str, Any] = record[part]  # type: ignore[assignment]
        for pair, seen in zip(probe["pairs"], probe["rules"], strict=True):
            print(pair[2], ascii(seen))
        print(part, "lone", ascii(probe.get("lone")))
        print(part, "names", ascii(probe["names"]))


if __name__ == "__main__":
    main()
