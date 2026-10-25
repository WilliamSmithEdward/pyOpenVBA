"""Names.Add in live Excel: a name whose RefersTo is a structured reference, and a name added to the workbook's
Names that a sheet already has a name of its own for.

Each case makes a workbook of its own: two sheets, Data with a table Sales over A1:B3 and Other beside it, the sheet
the case names active. Its steps add names to the workbook's or a sheet's Names, an error's number and description
noted where one raises one, and it reads back every name: its Name, RefersTo, RefersToR1C1 and Visible, in the
Names collection's order. The workbook is saved, opened again and read once more.

    python scripts/measure_names_add.py

writes tests/fixtures/names_add/names_add.json, which tests/test_excel_names_add.py replays.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "names_add"

#: Structured references and their neighbours, each a RefersTo tried with the table's own sheet active and another.
REFERENCES = {"column": "=Sales[Qty]", "table": "=Sales", "all": "=Sales[#All]", "headers": "=Sales[[#Headers],[Qty]]",
              "data": "=Sales[#Data]", "this_row": "=Sales[@Qty]", "sum": "=SUM(Sales[Qty])",
              "two_columns": "=Sales[[Qty]:[Price]]", "cell": "=Data!$D$4"}

CASES: dict[str, tuple[str, list[str]]] = {}
for key, text in REFERENCES.items():
    literal = '"' + text.replace('"', '""') + '"'
    CASES[f"refers_{key}_on_table_sheet"] = ("Data", [f'wb.Names.Add "N_{key}", {literal}'])
    CASES[f"refers_{key}_on_other_sheet"] = ("Other", [f'wb.Names.Add "N_{key}", {literal}'])
CASES.update({
    "r1c1_column_on_table_sheet": ("Data", ['wb.Names.Add Name:="N", RefersToR1C1:="=Sales[Qty]"']),
    "local_refers_column_on_table_sheet": ("Data", ['wb.Worksheets("Data").Names.Add "N", "=Sales[Qty]"']),
    "local_refers_column_on_other_sheet": ("Other", ['wb.Worksheets("Data").Names.Add "N", "=Sales[Qty]"']),
    "other_local_refers_column_on_table_sheet": ("Data", ['wb.Worksheets("Other").Names.Add "N", "=Sales[Qty]"']),
    "refers_local_column_on_table_sheet": ("Data", ['wb.Names.Add Name:="N", RefersToLocal:="=Sales[Qty]"']),
    # Which name an unqualified Add of the workbook's Names reaches.
    "active_local_then_book": ("Other", ['wb.Worksheets("Other").Names.Add "Twin", "=Other!$A$1"',
                                         'wb.Names.Add "Twin", "=Data!$A$3"']),
    "other_local_then_book": ("Other", ['wb.Worksheets("Data").Names.Add "Twin", "=Data!$A$1"',
                                        'wb.Names.Add "Twin", "=Data!$A$3"']),
    "book_then_active_local_then_book": ("Other", ['wb.Names.Add "Twin", "=Data!$A$2"',
                                                   'wb.Worksheets("Other").Names.Add "Twin", "=Other!$A$1"',
                                                   'wb.Names.Add "Twin", "=Data!$A$3"']),
    "book_then_book": ("Other", ['wb.Names.Add "Twin", "=Data!$A$2"', 'wb.Names.Add "twin", "=Data!$A$3"']),
    "qualified_inactive": ("Other", ['wb.Names.Add "Data!Twin", "=Data!$A$1"', 'wb.Names.Add "Data!Twin", "=Data!$A$2"']),
    "qualified_quoted": ("Data", ["wb.Names.Add \"'Other'!Twin\", \"=Other!$A$1\""]),
    "active_local_hidden_by_book": ("Other", ['wb.Worksheets("Other").Names.Add "Twin", "=Other!$A$1"',
                                              'wb.Names.Add "Twin", "=Data!$A$3", False']),
    "active_local_then_sheet_add": ("Other", ['wb.Worksheets("Other").Names.Add "Twin", "=Other!$A$1"',
                                              'wb.Worksheets("Data").Names.Add "Twin", "=Data!$A$3"']),
    "active_local_lookup": ("Other", ['wb.Names.Add "Twin", "=Data!$A$2"',
                                      'wb.Worksheets("Other").Names.Add "Twin", "=Other!$A$1"',
                                      'notes = notes & wb.Names("Twin").RefersTo & ":" & '
                                      'TypeName(wb.Names("Twin").Parent) & ";"']),
})
#: Which name each lookup reaches: the workbook's Names and a sheet's, by a bare name and a qualified one.
LOOKUPS = ['notes = notes & "[book]" & wb.Names("Twin").RefersTo & ":" & TypeName(wb.Names("Twin").Parent) & ";"',
           'notes = notes & "[other]" & wb.Worksheets("Other").Names("Twin").RefersTo & ";"',
           'notes = notes & "[data]" & wb.Worksheets("Data").Names("Twin").RefersTo & ";"',
           'notes = notes & "[qualified]" & wb.Names("Data!Twin").RefersTo & ";"',
           'notes = notes & "[count]" & wb.Names.Count & ":" & wb.Worksheets("Other").Names.Count & ";"']
CASES.update({
    "lookup_active_local_only": ("Other", ['wb.Worksheets("Other").Names.Add "Twin", "=Other!$A$1"', *LOOKUPS]),
    "lookup_other_local_only": ("Other", ['wb.Worksheets("Data").Names.Add "Twin", "=Data!$A$1"', *LOOKUPS]),
    "lookup_book_and_other_local": ("Other", ['wb.Names.Add "Twin", "=Data!$A$2"',
                                              'wb.Worksheets("Data").Names.Add "Twin", "=Data!$A$1"', *LOOKUPS]),
    "lookup_book_and_active_local": ("Other", ['wb.Names.Add "Twin", "=Data!$A$2"',
                                               'wb.Worksheets("Other").Names.Add "Twin", "=Other!$A$1"', *LOOKUPS]),
    "lookup_all_three": ("Other", ['wb.Names.Add "Twin", "=Data!$A$2"',
                                   'wb.Worksheets("Other").Names.Add "Twin", "=Other!$A$1"',
                                   'wb.Worksheets("Data").Names.Add "Twin", "=Data!$A$1"', *LOOKUPS]),
    # The first tab named after the others, Zed (see FIRST_TAB): is the first sheet's own name found first, or the
    # one first by name?
    "lookup_first_tab_last_by_name": ("Data", [
        'wb.Worksheets("Zed").Names.Add "Twin", "=Zed!$A$1"', 'wb.Worksheets("Other").Names.Add "Twin", "=Other!$A$1"',
        'wb.Names.Add "Lone", "=Data!$A$2"', 'wb.Worksheets("Other").Names.Add "Lone", "=Other!$A$2"',
        'notes = notes & "[twin]" & wb.Names("Twin").RefersTo & ";"',
        'notes = notes & "[lone]" & wb.Names("Lone").RefersTo & ";"']),
    "lookup_first_tab_none": ("Other", [
        'wb.Worksheets("Other").Names.Add "Twin", "=Other!$A$1"', 'wb.Worksheets("Data").Names.Add "Twin", "=Data!$A$1"',
        'notes = notes & "[twin]" & wb.Names("Twin").RefersTo & ";"', 'wb.Worksheets("Zed").Activate',
        'notes = notes & "[zed active]" & wb.Names("Twin").RefersTo & ";"']),
})

#: Cases whose workbook has a third sheet, Zed, as its first tab, made before the table as the others are.
FIRST_TAB = {"lookup_first_tab_last_by_name", "lookup_first_tab_none"}

HELPER = '''
Private Function Listed(wb As Object) As String
    Dim n As Object, out As String
    For Each n In wb.Names
        out = out & n.Name & "=" & n.RefersTo & "|" & n.RefersToR1C1 & "|" & n.Visible & ";"
    Next
    Listed = out
End Function
'''


def case_code(index: int, name: str, active: str, steps: list[str]) -> str:
    lines = [f"Private Function Make{index}(folder As String) As String",
             "    Dim wb As Object, ws As Object, notes As String, out As String",
             "    Set wb = Workbooks.Add(xlWBATWorksheet)", "    Set ws = wb.Worksheets(1)", '    ws.Name = "Data"',
             '    wb.Worksheets.Add After:=ws', '    wb.Worksheets(2).Name = "Other"',
             *(['    wb.Worksheets.Add Before:=ws', '    wb.Worksheets(1).Name = "Zed"'] if name in FIRST_TAB else []),
             '    ws.Range("A1").Value = "Qty"', '    ws.Range("B1").Value = "Price"', '    ws.Range("A2:B3").Value = 1',
             '    ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes).Name = "Sales"',
             f'    wb.Worksheets("{active}").Activate', "    On Error Resume Next"]
    for step in steps:
        lines += ["    Err.Clear", f"    {step}",
                  '    If Err.Number <> 0 Then notes = notes & "E" & Err.Number & ":" & Err.Description & " "']
    # Separators no name, reference or error description holds: a structured reference may hold @.
    lines += ["    On Error GoTo 0", f'    out = "{name}{{{{name}}}}" & notes & "{{{{names}}}}" & Listed(wb) & "{{{{reopened}}}}"',
              f'    wb.SaveAs folder & "\\{name}.xlsx", 51', "    wb.Close False",
              f'    Set wb = Workbooks.Open(folder & "\\{name}.xlsx")', "    out = out & Listed(wb)", "    wb.Close False",
              f'    Make{index} = out & "{{{{case}}}}"', "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    import tempfile

    functions = "".join(case_code(index, name, active, steps)
                        for index, (name, (active, steps)) in enumerate(CASES.items()))
    calls = "".join(f"    out = out & Make{index}(folder)\n" for index in range(len(CASES)))
    code = (HELPER + functions + "Public Function Build(folder As String) As String\n    Dim out As String\n"
            "    Application.DisplayAlerts = False\n" + calls + "    Build = out\nEnd Function\n")
    folder = tempfile.mkdtemp()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", (folder,), timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    cases = []
    for one in str(result.value).split("{{case}}"):
        if "{{name}}" in one:
            name, rest = one.split("{{name}}", 1)
            session, reopened = rest.split("{{reopened}}")
            notes, listed = session.split("{{names}}", 1)
            cases.append({"name": name, "active": CASES[name][0], "steps": CASES[name][1], "notes": notes,
                          "names": listed, "reopened": reopened})
    OUT.mkdir(parents=True, exist_ok=True)
    record = {"helper": HELPER, "build": code, "cases": cases}
    (OUT / "names_add.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for case in cases:
        print(case["name"], "|", ascii(case["notes"][:120]), "|", ascii(case["names"][:200]))


if __name__ == "__main__":
    sys.stdout.reconfigure(errors="backslashreplace")  # type: ignore[union-attr]
    main()
