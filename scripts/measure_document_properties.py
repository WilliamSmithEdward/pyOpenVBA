"""Workbook.BuiltinDocumentProperties and CustomDocumentProperties in live Excel, and the docProps parts a save writes.

Each case makes a workbook of its own, sets, adds, deletes or links its
properties, an error's number noted where a step raises one, and reads
back every property's Name, Type and Value. The workbook is saved, its
docProps parts, package relationships and content types kept, opened again
and read once more; two cases save the opened file again, and one whose
file Excel cannot open is not opened. A workbook Excel saved, its docProps
parts taken out, is opened, read and saved too. A Value equal to the
author a new workbook has reads <author>, and a date within five seconds
of the case's start reads now, so the fixture names no one and no day; the
measurement writes <author> for that name in the parts too, and <now> for
a time of the run. Excel writes a date in UTC by the machine's time zone,
whose rules Windows gives for each year, recorded beside the cases.

    python scripts/measure_document_properties.py

writes tests/fixtures/document_properties/: the workbooks as Excel saved
them and document_properties.json, which
tests/test_excel_document_properties.py replays.
"""

from __future__ import annotations

import datetime as dt
import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "document_properties"
B = 'wb.BuiltinDocumentProperties'
C = 'wb.CustomDocumentProperties'

#: name -> the steps, made on a new workbook ``wb`` whose first sheet is ``ws``.
CASES: dict[str, list[str]] = {
    "new": [],
    "set_text": [f'{B}("Title").Value = "A title"', f'{B}("Subject").Value = "A subject"',
                 f'{B}("Author").Value = "Someone"', f'{B}("Keywords").Value = "a;b"',
                 f'{B}("Comments").Value = "Notes"', f'{B}("Template").Value = "Plain"',
                 f'{B}("Last author").Value = "Another"', f'{B}("Revision number").Value = "7"',
                 f'{B}("Category").Value = "Cat"', f'{B}("Manager").Value = "Boss"', f'{B}("Company").Value = "Co"',
                 f'{B}("Hyperlink base").Value = "http://example.com/"', f'{B}("Content status").Value = "Draft"',
                 f'{B}("Language").Value = "en-US"', f'{B}("Document version").Value = "1.0"',
                 f'{B}("Content type").Value = "Kind"', f'{B}("Format").Value = "Fmt"',
                 f'{B}("Application name").Value = "Other"'],
    "set_numbers": [f'{B}("Total editing time").Value = 5', f'{B}("Number of pages").Value = 3',
                    f'{B}("Security").Value = 2', f'{B}("Number of words").Value = 9',
                    f'{B}("Number of bytes").Value = 100'],
    "set_dates": [f'{B}("Last print date").Value = DateSerial(2024, 1, 2) + TimeSerial(3, 4, 5)',
                  f'{B}("Creation date").Value = DateSerial(2023, 6, 7) + TimeSerial(8, 9, 10)',
                  f'{B}("Last save time").Value = DateSerial(2022, 2, 3)'],
    "set_coerced": [f'{B}("Title").Value = 5', f'{B}("Subject").Value = True', f'{B}("Comments").Value = 1.5',
                    f'{B}("Number of pages").Value = "7"', f'{B}("Total editing time").Value = 2.6',
                    f'{B}("Last print date").Value = "1/2/2024"', f'{B}("Number of pages").Value = "abc"',
                    f'{B}("Last print date").Value = "abc"', f'{B}("Security").Value = 3000000000#'],
    "set_empty": [f'{B}("Title").Value = "x"', f'{B}("Title").Value = Empty', f'{B}("Subject").Value = Null',
                  f'{B}("Number of pages").Value = Empty', f'{B}("Last print date").Value = Empty'],
    "set_long": [f'{B}("Title").Value = String(300, "t")', f'{B}("Comments").Value = String(300, "c")'],
    "builtin_items": [f'notes = notes & {B}("title").Name & ":"', f'notes = notes & {B}(9).Name & ":"',
                      f'notes = notes & {B}("Nope").Name', f'notes = notes & {B}(0).Name', f'notes = notes & {B}(35).Name',
                      f'notes = notes & TypeName({B}) & ":" & TypeName({B}(1)) & ":" & TypeName({B}.Parent) & ":" & '
                      f'TypeName({B}(1).Parent) & ":" & {B}.Creator & ":" & TypeName({B}.Application)',
                      f'For Each p In {B}: n = n + 1: Next: notes = notes & ":each=" & n',
                      f'notes = notes & ":" & {B}("Title").LinkToContent & ":" & {B}("Title").LinkSource',
                      f'{B}("Title").Delete', f'{B}("Title").Name = "X"', f'{B}("Title").Type = 1',
                      f'{B}("Title").LinkToContent = True'],
    "custom_types": [f'Set p = {C}.Add(Name:="Client", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="Acme")',
                     'notes = notes & TypeName(p) & ":" & p.Name & ":" & p.LinkToContent & ":" & TypeName(p.Parent)',
                     f'{C}.Add Name:="Count", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:=42',
                     f'{C}.Add Name:="Flag", LinkToContent:=False, Type:=msoPropertyTypeBoolean, Value:=True',
                     f'{C}.Add Name:="Off", LinkToContent:=False, Type:=msoPropertyTypeBoolean, Value:=False',
                     f'{C}.Add Name:="When", LinkToContent:=False, Type:=msoPropertyTypeDate, '
                     'Value:=DateSerial(2024, 1, 2) + TimeSerial(3, 4, 5)',
                     f'{C}.Add Name:="Day", LinkToContent:=False, Type:=msoPropertyTypeDate, Value:=DateSerial(2024, 5, 6)',
                     f'{C}.Add Name:="Ratio", LinkToContent:=False, Type:=msoPropertyTypeFloat, Value:=1.5',
                     f'{C}.Add Name:="Third", LinkToContent:=False, Type:=msoPropertyTypeFloat, Value:=1 / 3',
                     f'{C}.Add Name:="Negative", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:=-7',
                     f'{C}.Add Name:="Empty text", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=""',
                     f'{C}.Add Name:="Long", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=String(300, "x")',
                     f'{C}.Add Name:="Escaped", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="<a & b>"'],
    "custom_coerced": [f'{C}.Add Name:="Text", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=5',
                       f'{C}.Add Name:="Number", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:="17"',
                       f'{C}.Add Name:="Rounded", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:=2.5',
                       f'{C}.Add Name:="Truth", LinkToContent:=False, Type:=msoPropertyTypeBoolean, Value:="True"',
                       f'{C}.Add Name:="Number truth", LinkToContent:=False, Type:=msoPropertyTypeBoolean, Value:=0',
                       f'{C}.Add Name:="Text date", LinkToContent:=False, Type:=msoPropertyTypeDate, Value:="1/2/2024"',
                       f'{C}.Add Name:="Text float", LinkToContent:=False, Type:=msoPropertyTypeFloat, Value:="2.5"'],
    "custom_errors": [f'{C}.Add Name:="Client", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="Acme"',
                      f'{C}.Add Name:="client", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="Dup"',
                      f'{C}.Add Name:="Bad", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:="abc"',
                      f'{C}.Add Name:="NoValue", LinkToContent:=False, Type:=msoPropertyTypeString',
                      f'{C}.Add Name:="Big", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:=3000000000#',
                      f'{C}.Add Name:="Kind", LinkToContent:=False, Type:=9, Value:="x"',
                      f'{C}.Add Name:="", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="x"',
                      f'{C}.Add Name:=String(300, "n"), LinkToContent:=False, Type:=msoPropertyTypeString, Value:="x"',
                      f'{C}.Add Name:="Title", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="custom title"',
                      f'notes = notes & {C}("Nope").Name', f'notes = notes & {C}(0).Name'],
    "custom_changes": [f'{C}.Add Name:="Count", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:=42',
                       f'{C}.Add Name:="Flag", LinkToContent:=False, Type:=msoPropertyTypeBoolean, Value:=True',
                       f'{C}.Add Name:="Ratio", LinkToContent:=False, Type:=msoPropertyTypeFloat, Value:=1.5',
                       f'{C}.Add Name:="Client", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="Acme"',
                       f'{C}("Count").Value = "17"', f'{C}("Count").Value = "abc"', f'{C}("Count").Value = 2.6',
                       f'{C}("Flag").Type = msoPropertyTypeString', f'{C}("Client").Type = msoPropertyTypeNumber',
                       f'{C}("Ratio").Name = "Fraction"', f'{C}("Fraction").Name = "Count"',
                       f'{C}("Fraction").Value = DateSerial(2024, 1, 2)', f'{C}("Client").Value = "Beta"'],
    "custom_delete": [f'{C}.Add Name:="One", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="1"',
                      f'{C}.Add Name:="Two", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="2"',
                      f'{C}.Add Name:="Three", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="3"',
                      f'{C}("Two").Delete', f'notes = notes & {C}.Count & ":" & {C}(2).Name',
                      f'{C}.Add Name:="Four", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="4"'],
    "custom_link": ['ws.Range("B2").Value = "linked value"', 'ws.Range("C3").Value = 12', 'wb.Names.Add "Src", "=Sheet1!$B$2"',
                    'wb.Names.Add "Num", "=Sheet1!$C$3"',
                    f'{C}.Add Name:="Linked", LinkToContent:=True, Type:=msoPropertyTypeString, LinkSource:="Src"',
                    f'{C}.Add Name:="Linked number", LinkToContent:=True, Type:=msoPropertyTypeNumber, LinkSource:="Num"',
                    f'notes = notes & {C}("Linked").LinkSource & ":" & {C}("Linked").LinkToContent',
                    f'{C}.Add Name:="Nowhere", LinkToContent:=True, Type:=msoPropertyTypeString, LinkSource:="Missing"',
                    f'{C}.Add Name:="No source", LinkToContent:=True, Type:=msoPropertyTypeString',
                    'ws.Range("B2").Value = "changed value"'],
    "names_listed": ['ws.Name = "Data"', 'wb.Worksheets.Add After:=ws', 'wb.Worksheets(2).Name = "Second"',
                     'wb.Worksheets.Add After:=wb.Worksheets(2)', 'wb.Worksheets(3).Visible = xlSheetHidden',
                     'wb.Names.Add "Here", "=Data!$A$1"', 'wb.Names.Add "Constant", "=5"',
                     'wb.Names.Add "Formula", "=SUM(Data!$A$1:$A$3)"', 'wb.Worksheets(2).Names.Add "Local", "=Second!$B$2"',
                     'wb.Names.Add "Secret", "=Data!$C$3", Visible:=False', 'ws.PageSetup.PrintArea = "$A$1:$B$4"',
                     'ws.Range("A1:B3").Value = 1', 'ws.Range("A1:B3").AutoFilter 1, "1"'],
    "reopened_saved_again": [f'{B}("Title").Value = "first"'],
    "set_counts": [f'{B}("Number of characters").Value = 11', f'{B}("Number of lines").Value = 12',
                   f'{B}("Number of paragraphs").Value = 13', f'{B}("Number of slides").Value = 14',
                   f'{B}("Number of notes").Value = 15', f'{B}("Number of hidden Slides").Value = 16',
                   f'{B}("Number of multimedia clips").Value = 17',
                   f'{B}("Number of characters (with spaces)").Value = 18', f'{B}("Total editing time").Value = 19',
                   f'{B}("Number of pages").Value = 20', f'{B}("Number of words").Value = 21',
                   f'{B}("Template").Value = "T"', f'{B}("Format").Value = "F"', f'{B}("Manager").Value = "M"',
                   f'{B}("Company").Value = "C"', f'{B}("Hyperlink base").Value = "H"'],
    "set_text_kinds": [f'{B}("Title").Value = DateSerial(2024, 1, 2)', f'{B}("Subject").Value = False',
                       f'{B}("Keywords").Value = 1234567890123#', f'{B}("Comments").Value = CCur(2.5)'],
    "builtin_add": [f'{B}.Add Name:="Extra", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="x"'],
    "custom_bad_kinds": [f'{C}.Add Name:="Truth", LinkToContent:=False, Type:=msoPropertyTypeBoolean, Value:="abc"',
                         f'{C}.Add Name:="Day", LinkToContent:=False, Type:=msoPropertyTypeDate, Value:="abc"',
                         f'{C}.Add Name:="Float", LinkToContent:=False, Type:=msoPropertyTypeFloat, Value:="abc"',
                         f'{C}.Add Name:="Nothing", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=Empty',
                         f'{C}.Add Name:="Null", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=Null',
                         f'{C}.Add Name:="Date text", LinkToContent:=False, Type:=msoPropertyTypeString, '
                         'Value:=DateSerial(2024, 1, 2)',
                         f'{C}.Add Name:="Positional", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:=1'],
    "names_sorted": ['ws.Name = "My Data"', 'wb.Worksheets.Add After:=ws', 'wb.Worksheets(2).Name = "Beta"',
                     'wb.Names.Add "zeta", "=\'My Data\'!$A$1"', 'wb.Names.Add "Alpha", "=\'My Data\'!$A$2"',
                     'wb.Names.Add "beta", "=Beta!$A$3"', 'wb.Worksheets(2).Names.Add "Gamma", "=Beta!$B$2"',
                     'ws.Names.Add "alpha2", "=\'My Data\'!$B$5"', 'wb.Names.Add "Errored", "=#REF!"',
                     'wb.Names.Add "Union", "=\'My Data\'!$A$1,\'My Data\'!$C$3"'],
}

CASES["custom_all_deleted"] = [
    f'{C}.Add Name:="One", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="1"']

#: Local times across the year, before 1970, across the 2007 change to the United States' daylight saving rules,
#: past 3000, before 1601 (where a FILETIME stops) and in the hour a clock skips or repeats, each a custom date.
ZONE_DATES = {
    "Z1950": "DateSerial(1950, 7, 1) + TimeSerial(12, 0, 0)", "Z1969": "DateSerial(1969, 7, 20) + TimeSerial(12, 0, 0)",
    "Z1970": "DateSerial(1970, 1, 1)", "Z1970s": "DateSerial(1970, 7, 1) + TimeSerial(12, 0, 0)",
    "Z2006a": "DateSerial(2006, 3, 20) + TimeSerial(12, 0, 0)", "Z2006b": "DateSerial(2006, 10, 30) + TimeSerial(12, 0, 0)",
    "Z1899": "DateSerial(1899, 7, 1) + TimeSerial(12, 0, 0)", "Z1800": "DateSerial(1800, 1, 1)",
    "Z1601": "DateSerial(1601, 1, 2)", "Z100": "DateSerial(100, 1, 1)",
    "Z3500": "DateSerial(3500, 7, 1) + TimeSerial(12, 0, 0)", "Z9999": "DateSerial(9999, 12, 31) + TimeSerial(15, 0, 0)",
    "Z9999b": "DateSerial(9999, 12, 31) + TimeSerial(16, 0, 0)", "Zneg": "CDate(-1.25 - 0.5 / 86400)",
    "Zneg2": "CDate(-0.75 - 0.4 / 86400)", "Zneg3": "CDate(-1)", "Zneg4": "CDate(-181.5)",
    "Zbefore": "DateSerial(2026, 3, 8) + TimeSerial(1, 59, 59)", "Zskipped": "DateSerial(2026, 3, 8) + TimeSerial(2, 30, 0)",
    "Zafter": "DateSerial(2026, 3, 8) + TimeSerial(3, 0, 0)", "Zfall0": "DateSerial(2026, 11, 1) + TimeSerial(0, 59, 59)",
    "Zrepeated": "DateSerial(2026, 11, 1) + TimeSerial(1, 30, 0)", "Zfall2": "DateSerial(2026, 11, 1) + TimeSerial(2, 0, 0)",
    "Znoon": "TimeSerial(12, 0, 0)", "Zhalf": "DateSerial(2024, 1, 2) + 0.5 / 86400",
}
CASES["dates_zones"] = [
    f'{B}("Last print date").Value = DateSerial(1950, 7, 1) + TimeSerial(12, 0, 0)',
    f'{B}("Creation date").Value = CDate(-1.25)',
    *(f'{C}.Add Name:="{name}", LinkToContent:=False, Type:=msoPropertyTypeDate, Value:={value}'
      for name, value in ZONE_DATES.items())]
CASES["dates_early"] = [
    f'{B}("Last print date").Value = DateSerial(100, 1, 1)',
    f'{B}("Creation date").Value = DateSerial(1600, 12, 31) + TimeSerial(23, 0, 0)',
    f'{B}("Last print date").Value = DateSerial(1601, 1, 1)',
    f'{C}.Add Name:="First", LinkToContent:=False, Type:=msoPropertyTypeDate, Value:=DateSerial(1601, 1, 1)',
    f'{C}.Add Name:="Second", LinkToContent:=False, Type:=msoPropertyTypeDate, Value:=DateSerial(1601, 1, 2)',
    f'{C}.Add Name:="Float", LinkToContent:=False, Type:=msoPropertyTypeFloat, Value:=-109205.5',
    f'{C}("Float").Type = msoPropertyTypeDate',
    f'{C}.Add Name:="Float2", LinkToContent:=False, Type:=msoPropertyTypeFloat, Value:=-5.25',
    f'{C}("Float2").Type = msoPropertyTypeDate',
    f'{C}.Add Name:="Num", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:=-3',
    f'{C}("Num").Type = msoPropertyTypeDate',
    f'{C}.Add Name:="Back", LinkToContent:=False, Type:=msoPropertyTypeDate, Value:=CDate(-2.5)',
    f'{C}("Back").Type = msoPropertyTypeFloat',
    f'{C}.Add Name:="Reset", LinkToContent:=False, Type:=msoPropertyTypeDate, Value:=CDate(-2.5)',
    f'{C}("Reset").Value = {C}("Reset").Value']
CASES["dates_overflow"] = [
    f'{B}("Creation date").Value = DateSerial(9999, 12, 31) + TimeSerial(23, 59, 59)',
    f'{C}.Add Name:="Late", LinkToContent:=False, Type:=msoPropertyTypeDate, '
    'Value:=DateSerial(9999, 12, 31) + TimeSerial(23, 59, 59)',
    f'{C}.Add Name:="Edge", LinkToContent:=False, Type:=msoPropertyTypeDate, '
    'Value:=DateSerial(9999, 12, 31) + TimeSerial(16, 0, 0)']
#: Cases whose saved file Excel cannot open again: they are read once and saved, not opened.
UNOPENED = {"dates_overflow"}
CASES["names_more"] = [
    'ws.Name = "Main"', 'wb.Worksheets.Add After:=ws', 'wb.Worksheets(2).Name = "Bob\'s"',
    'wb.Worksheets.Add After:=wb.Worksheets(2)', 'wb.Worksheets(3).Name = "Hidden one"',
    'wb.Worksheets(3).Visible = xlSheetHidden', 'wb.Names.Add "Here", "=Main!$A$1"',
    'wb.Names.Add "ThreeD", "=\'Main:Hidden one\'!$A$1"', 'wb.Names.Add "Relative", "=Main!A1"',
    'wb.Names.Add "Indirect", "=Here"', 'wb.Names.Add "Dynamic", "=OFFSET(Main!$A$1,0,0,2,2)"',
    'wb.Names.Add "OnHidden", "=\'Hidden one\'!$B$2"', 'wb.Names.Add "Quoted", "=\'Bob\'\'s\'!$C$3"',
    'wb.Worksheets(2).Names.Add "Local", "=\'Bob\'\'s\'!$D$4"',
    'wb.Worksheets(3).Names.Add "HiddenLocal", "=\'Hidden one\'!$E$5"', 'ws.PageSetup.PrintTitleRows = "$1:$2"',
    'wb.Names.Add "Dup", "=Main!$F$6"', 'ws.Names.Add "Dup", "=Main!$G$7"', 'wb.Names.Add "WholeCol", "=Main!$A:$A"',
    'wb.Names.Add "Arr", "={1,2,3}"', 'wb.Names.Add "Txt", "=""text"""',
    'wb.Names.Add "Inter", "=Main!$A$1:$C$3 Main!$B$2:$D$4"', 'wb.Names.Add "Ranged", "=Main!$A$1:Main!$B$2"',
    'wb.Names.Add "_under", "=Main!$H$1"', 'wb.Names.Add "a_b", "=Main!$H$2"', 'wb.Names.Add "ab", "=Main!$H$3"',
    'wb.Names.Add "a1x", "=Main!$H$4"', 'wb.Names.Add "A2x", "=Main!$H$5"', 'wb.Names.Add "B.c", "=Main!$H$6"',
    'wb.Names.Add "Zed", "=Main!$H$7"', 'wb.Names.Add ChrW(233) & "clair", "=Main!$H$8"',
    'wb.Names.Add "eclair", "=Main!$H$9"', 'wb.Names.Add "Eb", "=Main!$H$10"', 'wb.Names.Add "a", "=Main!$H$11"',
    'wb.Names.Add "a\\b", "=Main!$H$12"', 'wb.Names.Add "a.b", "=Main!$H$13"']
CASES["chart_sheet"] = ['ws.Range("A1:B3").Value = 1', 'wb.Charts.Add After:=ws', 'wb.Worksheets.Add After:=wb.Sheets(2)',
                        'wb.Names.Add "Here", "=Sheet1!$A$1"']
CASES["saved_flag"] = [
    'wb.Saved = True', f'{B}("Title").Value = "t"', 'notes = notes & wb.Saved & ":"',
    'wb.Saved = True', f'{C}.Add Name:="S", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="s"',
    'notes = notes & wb.Saved & ":"', 'wb.Saved = True', f'{C}("S").Value = "u"', 'notes = notes & wb.Saved & ":"',
    'wb.Saved = True', f'{C}("S").Name = "R"', 'notes = notes & wb.Saved & ":"',
    'wb.Saved = True', f'{C}("R").Delete', 'notes = notes & wb.Saved & ":"',
    'wb.Saved = True', f'notes = notes & {B}("Title").Value & wb.Saved & ":"',
    'wb.Saved = True', f'{B}("Title").Value = "t"', 'notes = notes & wb.Saved']
CASES["defaults"] = [
    f'{B}("Title") = "by default"', f'notes = notes & {B}("Title") & ":"',
    f'{C}.Add "Positional", False, msoPropertyTypeNumber, 5', f'notes = notes & {C}("Positional") & ":"',
    f'notes = notes & {B}(1).Creator & ":" & TypeName({B}(1).Application) & ":" & TypeName({C}(1).Parent)',
    f'{C}.Add "TypeLess", False, , "x"', f'{C}.Add "NoLink"', f'{C}.Add Name:="Var", LinkToContent:=False, Value:=5',
    f'{C}.Add Name:="VarBool", LinkToContent:=False, Value:=True',
    f'{C}.Add Name:="VarDouble", LinkToContent:=False, Value:=2.5',
    f'{C}.Add Name:="VarDate", LinkToContent:=False, Value:=DateSerial(2024, 1, 2)',
    f'{C}.Add Name:="VarText", LinkToContent:=False, Value:="t"']
CASES["custom_link_changes"] = [
    'ws.Range("B2").Value = "src"', 'ws.Range("C3").Value = 12', 'wb.Names.Add "Src", "=Sheet1!$B$2"',
    'wb.Names.Add "Num", "=Sheet1!$C$3"',
    f'{C}.Add Name:="L", LinkToContent:=True, Type:=msoPropertyTypeString, LinkSource:="Src"',
    f'{C}("L").LinkSource = "Num"', f'notes = notes & {C}("L").LinkSource & ":"',
    f'{C}.Add Name:="U", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="u"',
    f'notes = notes & "[" & {C}("U").LinkSource & "]:"', f'{C}("U").LinkSource = "Src"',
    f'notes = notes & {C}("U").LinkToContent & ":"',
    f'{C}.Add Name:="M", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="m"',
    f'{C}("M").LinkToContent = True', f'notes = notes & {C}("M").LinkToContent & ":"',
    f'{C}.Add Name:="K", LinkToContent:=True, Type:=msoPropertyTypeString, LinkSource:="Src"',
    f'{C}("K").LinkToContent = False', f'notes = notes & {C}("K").LinkToContent & ":" & {C}("K").Value & ":"',
    f'{C}.Add Name:="V", LinkToContent:=True, Type:=msoPropertyTypeString, LinkSource:="Src"',
    f'{C}("V").Value = "set"', f'notes = notes & {C}("V").Value & ":"',
    f'{C}.Add Name:="Both", LinkToContent:=True, Type:=msoPropertyTypeString, Value:="given", LinkSource:="Src"',
    f'notes = notes & {C}("Both").Value & ":"',
    f'{C}.Add Name:="W", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="w"',
    f'{C}("W").LinkSource = "Src"', f'{C}("W").LinkToContent = True',
    f'notes = notes & {C}("W").LinkToContent & ":" & {C}("W").Value & ":"',
    f'{C}.Add Name:="Typeless", LinkToContent:=True, LinkSource:="Src"',
    f'{C}.Add LinkToContent:=False, Type:=msoPropertyTypeString, Value:="nameless"',
    f'{C}("L").LinkSource = "Nowhere"', f'notes = notes & {C}("L").LinkSource & ":" & {C}("L").Value']
CASES["custom_stale"] = [
    f'{C}.Add Name:="X", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="x"', f'Set p = {C}("X")',
    f'{C}("X").Delete', 'notes = notes & p.Name & ":"', 'notes = notes & p.Value & ":"', 'p.Value = "y"',
    f'notes = notes & {C}.Count & ":"', 'p.Delete',
    f'{C}.Add Name:="Y", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="y"',
    f'For Each p In {C}: n = n + 1: Next: notes = notes & "each=" & n & ":"', f'Set p = {B}("Title")',
    'p.Value = "through a reference"', 'notes = notes & TypeName(p) & ":" & p.Value']
CASES["names_tables"] = [
    'ws.Name = "Zeta"', 'wb.Worksheets.Add After:=ws', 'wb.Worksheets(2).Name = "Alpha"',
    'ws.Range("A1:B3").Value = 1', 'ws.Range("A1").Value = "Qty"', 'ws.Range("B1").Value = "Price"',
    'ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes).Name = "Sales"',
    'wb.Names.Add "Whole", "=Sales"', 'wb.Names.Add "Cell", "=Zeta!$D$4"', 'wb.Names.Add "Twin", "=Zeta!$A$3"',
    'wb.Worksheets(2).Names.Add "Twin", "=Alpha!$A$1"', 'ws.Names.Add "Twin", "=Zeta!$A$2"']
#: How Names.Add takes a structured reference and an unqualified name the active sheet has a name of its own for:
#: each name and what it refers to, read back.
CASES["names_add_context"] = [
    'ws.Name = "Zeta"', 'ws.Range("A1:B3").Value = 1', 'ws.Range("A1").Value = "Qty"', 'ws.Range("B1").Value = "Price"',
    'ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes).Name = "Sales"',
    'wb.Names.Add "OnTablesSheet", "=Sales[Qty]"', 'wb.Worksheets.Add After:=ws', 'wb.Worksheets(2).Name = "Alpha"',
    'wb.Names.Add "OnOtherSheet", "=Sales[Qty]"', 'wb.Worksheets(2).Names.Add "Twin", "=Alpha!$A$1"',
    'wb.Names.Add "Twin", "=Zeta!$A$3"',
    'For Each p In wb.Names: notes = notes & p.Name & "=" & p.RefersTo & ";": Next']
#: The cells custom_link_kinds links to, each named K<label> (a name like N1 is a cell's address, which Names refuses).
LINKED_CELLS = {"Date": 'Value = DateSerial(2024, 1, 2)', "Bool": 'Value = True', "Err": 'Formula = "=NA()"',
                "Empty": "", "Float": 'Value = 1.5', "Text": 'Value = "text"', "Fmt": 'Value = 42',
                "Formula": 'Formula = "=1+1"', "Big": 'Value = 3000000000#', "Whole": 'Value = 7'}
CASES["custom_link_kinds"] = [
    *(f'ws.Range("D{row}").{action}' for row, action in enumerate(LINKED_CELLS.values(), start=1) if action),
    'ws.Range("D7").NumberFormat = "0.00"', 'ws.Range("B20").Value = "first"', 'ws.Range("C20").Value = "second"',
    *(f'wb.Names.Add "K{label}", "=Sheet1!$D${row}"' for row, label in enumerate(LINKED_CELLS, start=1)),
    'wb.Names.Add "Pair", "=Sheet1!$B$20:$C$21"',
    *(f'{C}.Add Name:="L{label}", LinkToContent:=True, Type:=msoPropertyTypeString, LinkSource:="K{label}"'
      for label in LINKED_CELLS),
    f'{C}.Add Name:="Pair", LinkToContent:=True, Type:=msoPropertyTypeString, LinkSource:="Pair"',
    f'{C}.Add Name:="AsDate", LinkToContent:=True, Type:=msoPropertyTypeDate, LinkSource:="KDate"',
    f'{C}.Add Name:="AsBool", LinkToContent:=True, Type:=msoPropertyTypeBoolean, LinkSource:="KBool"',
    f'{C}.Add Name:="AsNumber", LinkToContent:=True, Type:=msoPropertyTypeNumber, LinkSource:="KFloat"',
    f'{C}.Add Name:="AsFloat", LinkToContent:=True, Type:=msoPropertyTypeFloat, LinkSource:="KText"',
    f'{C}.Add Name:="AsWhole", LinkToContent:=True, Type:=msoPropertyTypeNumber, LinkSource:="KWhole"',
    f'notes = notes & TypeName({C}("AsDate").Value) & ":" & TypeName({C}("AsBool").Value) & ":" & '
    f'TypeName({C}("AsFloat").Value)']
CASES["custom_text"] = [
    f'{C}.Add Name:="A ""quoted"" name", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="q"',
    f'{C}.Add Name:="Bob\'s", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="b"',
    f'{C}.Add Name:="<tag>&", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="t"',
    f'{C}.Add Name:="  spaced  ", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="s"',
    f'{C}.Add Name:=ChrW(233), LinkToContent:=False, Type:=msoPropertyTypeString, Value:="e"',
    f'{C}.Add Name:="Lines", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="a" & vbLf & "b"',
    f'{C}.Add Name:="Return", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="a" & vbCr & "b"',
    f'{C}.Add Name:="Tab", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="a" & vbTab & "b"',
    f'{C}.Add Name:="Control", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="a" & Chr(1) & "b"',
    f'{C}.Add Name:="Escape", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="a_x0041_b"',
    f'{C}.Add Name:="Pair", LinkToContent:=False, Type:=msoPropertyTypeString, '
    'Value:="a" & ChrW(&HD83D) & ChrW(&HDE00) & "b"',
    f'{C}.Add Name:="Quote", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="say ""hi"" \'x\'"',
    f'{B}("Title").Value = "t" & vbLf & "u" & Chr(1) & "v"', f'{B}("Subject").Value = "a_x0041_b"',
    f'{B}("Keywords").Value = "a" & ChrW(&HD83D) & ChrW(&HDE00)', f'{B}("Comments").Value = "say ""hi"" \'x\'"',
    f'{B}("Company").Value = "a" & vbCr & vbLf & "b"']
CASES["session_after_save"] = [
    'ws.Range("B2").Value = "v"', 'wb.Names.Add "Src", "=Sheet1!$B$2"',
    f'{C}.Add Name:="Linked", LinkToContent:=True, Type:=msoPropertyTypeString, LinkSource:="Src"',
    f'{B}("Last author").Value = "Another"', f'{B}("Application name").Value = "Other"', f'{B}("Security").Value = 2',
    f'{B}("Content type").Value = "Kind"', 'wb.SaveAs folder & "\\session_after_save_inner.xlsx", 51']
CASES["copied_out"] = [
    f'{B}("Title").Value = "Source title"', f'{B}("Creation date").Value = DateSerial(2020, 1, 1)',
    f'{C}.Add Name:="SC", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="sc"',
    'Set p = wb', 'ws.Copy', 'Set wb = ActiveWorkbook', 'Set ws = wb.Worksheets(1)', 'p.Close False']

#: Cases whose workbook, once opened again, takes these steps and is saved a second time, as <name>_again.xlsx.
AGAIN: dict[str, list[str]] = {
    "reopened_saved_again": [],
    "custom_all_deleted": [f"{C}(1).Delete"],
}

HELPER = '''Public Author As String
Public Started As Date

Private Function Shown(p As Object) As String
    Dim v As Variant
    On Error GoTo Failed
    v = p.Value
    If VarType(v) = vbDate Then
        If Abs(CDbl(v) - CDbl(Started)) < 5# / 86400# Then
            Shown = "Date:now"
            Exit Function
        End If
    End If
    If VarType(v) = vbString And Author <> "" Then
        If v = Author Then
            Shown = "String:<author>"
            Exit Function
        End If
    End If
    Shown = TypeName(v) & ":" & CStr(v)
    Exit Function
Failed:
    Shown = "E" & Err.Number
End Function

Private Function Listed(props As Object) As String
    Dim p As Object, out As String, i As Long
    out = props.Count & "["
    For i = 1 To props.Count
        Set p = props(i)
        out = out & p.Name & "|" & p.Type & "|" & Shown(p) & "|" & p.LinkToContent & ";"
    Next
    Listed = out & "]"
End Function

Private Function Dump(wb As Object) As String
    Dump = Listed(wb.BuiltinDocumentProperties) & "#" & Listed(wb.CustomDocumentProperties)
End Function
'''

BUILD = HELPER + '''{functions}
Public Function Build(folder As String) As String
    Dim wb As Object, out As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Author = wb.BuiltinDocumentProperties("Author").Value
    wb.Close False
{cases}
    Build = out
End Function
'''


def vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def case_code(index: int, name: str, steps: list[str]) -> str:
    """A function of its own for one case: its workbook made, changed, read, saved, opened and read again."""
    lines = [f"Private Function Make{index}(folder As String) As String",
             "    Dim wb As Object, ws As Object, p As Object, notes As String, n As Long, out As String",
             "    Started = Now",
             "    Set wb = Workbooks.Add(xlWBATWorksheet)", "    Set ws = wb.Worksheets(1)", "    On Error Resume Next"]
    for step in steps:
        lines += ["    Err.Clear", f"    {step}",
                  '    If Err.Number <> 0 Then notes = notes & "E" & Err.Number & ":" & Err.Description & " "']
    lines += ["    On Error GoTo 0",
              f'    out = "{name}^" & notes & "@" & Dump(wb) & "@@"',
              f'    wb.SaveAs folder & "\\{name}.xlsx", 51', "    wb.Close False"]
    if name in UNOPENED:
        return "\n".join([*lines, f'    Make{index} = out & "~"', "End Function"]) + "\n"
    lines += [f'    Set wb = Workbooks.Open(folder & "\\{name}.xlsx")', '    out = out & Dump(wb)']
    if name in AGAIN:
        lines += [f"    {step}" for step in AGAIN[name]]
        lines += ['    out = out & "@@" & Dump(wb)', f'    wb.SaveAs folder & "\\{name}_again.xlsx", 51']
    lines += ["    wb.Close False", '    Make{0} = out & "~"'.format(index)]
    return "\n".join([*lines, "End Function"]) + "\n"


def parts(path: Path, author: str, run: dt.datetime) -> dict[str, str]:
    """The docProps parts, the package relationships and the content types, the author written <author> and a
    time of the run <now>."""
    out: dict[str, str] = {}
    with zipfile.ZipFile(path) as package:
        for name in sorted(package.namelist()):
            if name.startswith("docProps/") or name in ("_rels/.rels", "[Content_Types].xml"):
                out[name] = anonymous(package.read(name).decode("utf-8"), author, run)
    return out


def anonymous(text: str, author: str, run: dt.datetime) -> str:
    """The author written <author>, and each W3CDTF time within ten minutes of the run written <now>."""
    if author:
        text = text.replace(f">{author}<", "><author><")

    def stamp(found: re.Match[str]) -> str:
        try:
            when = dt.datetime.strptime(found.group(), "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=dt.timezone.utc)
        except ValueError:
            # Excel writes year 10000 as 0000, which no datetime holds.
            return found.group()
        return "<now>" if abs((when - run).total_seconds()) < 600 else found.group()

    return re.sub(r"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ", stamp, text)


#: A workbook Excel saved, its docProps parts taken out with their relationships and content types, which Excel then
#: opens, reads, saves as no_docprops_again.xlsx and opens again.
OPENED = HELPER + '''
Public Function Opened(folder As String) As String
    Dim wb As Object, out As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Author = wb.BuiltinDocumentProperties("Author").Value
    wb.Close False
    Started = Now
    Set wb = Workbooks.Open(folder & "\\no_docprops.xlsx")
    out = "no_docprops^@" & Dump(wb) & "@@"
    wb.SaveAs folder & "\\no_docprops_again.xlsx", 51
    wb.Close False
    Set wb = Workbooks.Open(folder & "\\no_docprops_again.xlsx")
    out = out & Dump(wb)
    wb.Close False
    Opened = out & "~"
End Function
'''


def zone_rules() -> list[dict[str, object]]:
    """This machine's time zone as Windows gives it for each year Excel can write, 1601 to 10000, in runs of years
    that share a rule: what Excel turns a local time into UTC by."""
    import ctypes
    from ctypes import wintypes

    class SYSTEMTIME(ctypes.Structure):
        _fields_ = [(field, wintypes.WORD) for field in
                    ("wYear", "wMonth", "wDayOfWeek", "wDay", "wHour", "wMinute", "wSecond", "wMilliseconds")]

    class TIME_ZONE_INFORMATION(ctypes.Structure):
        _fields_ = [("Bias", wintypes.LONG), ("StandardName", wintypes.WCHAR * 32), ("StandardDate", SYSTEMTIME),
                    ("StandardBias", wintypes.LONG), ("DaylightName", wintypes.WCHAR * 32),
                    ("DaylightDate", SYSTEMTIME), ("DaylightBias", wintypes.LONG)]

    def times(found: SYSTEMTIME) -> list[int]:
        return [found.wYear, found.wMonth, found.wDayOfWeek, found.wDay, found.wHour, found.wMinute, found.wSecond,
                found.wMilliseconds]

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.GetTimeZoneInformationForYear.argtypes = [wintypes.USHORT, ctypes.c_void_p,
                                                     ctypes.POINTER(TIME_ZONE_INFORMATION)]
    runs: list[dict[str, object]] = []
    for year in range(1601, 10001):
        info = TIME_ZONE_INFORMATION()
        assert kernel.GetTimeZoneInformationForYear(year, None, ctypes.byref(info)), ctypes.get_last_error()
        rule = {"bias": info.Bias, "standard_bias": info.StandardBias, "daylight_bias": info.DaylightBias,
                "standard": times(info.StandardDate), "daylight": times(info.DaylightDate)}
        if runs and {key: value for key, value in runs[-1].items() if key != "years"} == rule:
            runs[-1]["years"] = [runs[-1]["years"][0], year]  # type: ignore[index]
        else:
            runs.append({"years": [year, year], **rule})
    return runs


def without_docprops(source: Path, target: Path) -> None:
    """A copy of a saved workbook with docProps/core.xml and app.xml gone, and their relationships and content
    types."""
    with zipfile.ZipFile(source) as package:
        items = [(info, package.read(info.filename)) for info in package.infolist()]
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as package:
        for info, data in items:
            if info.filename.startswith("docProps/"):
                continue
            if info.filename in ("_rels/.rels", "[Content_Types].xml"):
                text = data.decode("utf-8")
                text = re.sub(r'<(?:Relationship|Override)\b[^>]*docProps/[^>]*/>', "", text)
                data = text.encode("utf-8")
            package.writestr(info, data)


def records(value: str, name: str) -> list[dict[str, object]]:
    """Each case's notes and properties as the macro listed them, the author written <author>."""
    found: list[dict[str, object]] = []
    for one in value.split("~"):
        if "^" not in one:
            continue
        label, rest = one.split("^", 1)
        session, reopened, *again = rest.split("@@")
        found.append({"name": label, "steps": CASES.get(label, []), "again_steps": AGAIN.get(label),
                      "session": session.replace(name, "<author>"), "reopened": reopened.replace(name, "<author>"),
                      "again": again[0].replace(name, "<author>") if again else None})
    return found


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    functions = "".join(case_code(index, name, steps) for index, (name, steps) in enumerate(CASES.items()))
    calls = "".join(f"    out = out & Make{index}(folder)\n" for index in range(len(CASES)))
    code = BUILD.replace("{functions}", functions).replace("{cases}", calls)
    run = dt.datetime.now(dt.timezone.utc)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        author = excel.run_vba('Public Function Who() As String\nDim wb As Object\nSet wb = Workbooks.Add\n'
                               'Who = wb.BuiltinDocumentProperties("Author").Value\nwb.Close False\nEnd Function\n',
                               "Who", timeout=120.0)
        assert author.ok, author.message
        result = excel.run_vba(code, "Build", (str(OUT),), timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
        without_docprops(OUT / "new.xlsx", OUT / "no_docprops.xlsx")
        opened = excel.run_vba(OPENED, "Opened", (str(OUT),), timeout=300.0)
        assert opened.ok, f"{opened.outcome}: {opened.message} {opened.error}"
    name = str(author.value)
    record_cases = records(str(result.value), name) + records(str(opened.value), name)
    files: dict[str, dict[str, str]] = {}
    for path in sorted(OUT.glob("*.xlsx")):
        strip_save_path(path)
        files[path.stem] = parts(path, name, run)
        # The workbooks keep only what the parts hold, which names no one once written <author>.
        _anonymize_file(path, name)
    record = {"build": code.replace(name, "<author>"), "opened": OPENED, "zone": zone_rules(), "cases": record_cases,
              "files": files}
    text = json.dumps(record, indent=1, ensure_ascii=False) + "\n"
    assert name not in text
    (OUT / "document_properties.json").write_text(text, encoding="utf-8")
    for case in record_cases:
        print(case["name"], ascii(case["session"][:300]))


def _anonymize_file(path: Path, author: str) -> None:
    """Rewrite a saved workbook with the author's name written <author> in its docProps parts."""
    if not author:
        return
    with zipfile.ZipFile(path) as package:
        items = [(info, package.read(info.filename)) for info in package.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for info, data in items:
            if info.filename.startswith("docProps/"):
                data = data.replace(f">{author}<".encode(), b"><author><")
            package.writestr(info, data)


if __name__ == "__main__":
    main()
