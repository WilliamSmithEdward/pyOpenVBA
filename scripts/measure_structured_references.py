"""Structured references: how Excel reads, shows, saves and works out a formula that names part of a table.

One workbook is built in live Excel from VBA: Table1 (Name, Qty, Price
over three rows), a table named Odd whose headers hold every character
that has to be escaped or bracketed, a table named Inner to write
formulas into, and a defined name over a table column. Then:

* every formula in CASES is written to its cell with Range.Formula (or
  FormulaR1C1), and whether Excel took it, what Formula and FormulaR1C1
  read back and the value it gave are recorded;
* a formula is written to a block, copied and filled;
* Range and Evaluate are asked to read structured references;
* the workbook is saved, and the <f> Excel wrote for every cell is read
  out of the package;
* the table is renamed and its headers rewritten step by step, and every
  formula is read again after each step.

    python scripts/measure_structured_references.py

writes tests/fixtures/structured_references/ -- the two saved workbooks
and structured.json -- which tests/test_excel_structured_references.py
replays.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
FOLDER = ROOT / "tests" / "fixtures" / "structured_references"

#: The headers of the table named Odd: each needs an escape or brackets, or shows that it does not.
ODD = ["Unit Price", "Q#", "It's", "[x]", "@home", "a@b", "a,b", "a.b", "Total $", "a-b", "a&b", "a:b"]

#: The sheets and tables every case runs against; ws is the first sheet and ws2 the second.
BUILD = "\n".join([
    'ws.Range("A1:C1").Value = Array("Name", "Qty", "Price")',
    'ws.Range("A2:C2").Value = Array("apple", 3, 1.5)',
    'ws.Range("A3:C3").Value = Array("pear", 5, 2.25)',
    'ws.Range("A4:C4").Value = Array("plum", 7, 0.5)',
    'ws.ListObjects.Add xlSrcRange, ws.Range("A1:C4"), , xlYes',
    # A value starting with @ is read as a formula, and refused, unless an apostrophe keeps it text.
    'ws.Range("H1:S1").Value = Array(' + ", ".join('"' + ("'" if name.startswith("@") else "") + name + '"'
                                                    for name in ODD) + ")",
    'ws.Range("H2:S2").Value = Array(' + ", ".join(str(number) for number in range(1, len(ODD) + 1)) + ")",
    'ws.ListObjects.Add(xlSrcRange, ws.Range("H1:S2"), , xlYes).Name = "Odd"',
    'ws.Range("A20:C20").Value = Array("Qty", "Price", "Note")',
    'ws.Range("A21:C21").Value = Array(1, 10, "a")',
    'ws.Range("A22:C22").Value = Array(2, 20, "b")',
    'ws.Range("A23:C23").Value = Array(3, 30, "c")',
    'ws.Range("A24:C24").Value = Array(4, 40, "d")',
    'ws.ListObjects.Add(xlSrcRange, ws.Range("A20:C24"), , xlYes).Name = "Inner"',
    'ws.Range("AM1:AN1").Value = Array("a", "b c")',
    'ws.Range("AM2:AN2").Value = Array(1, 2)',
    'ws.ListObjects.Add(xlSrcRange, ws.Range("AM1:AN2"), , xlYes).Name = "One"',
    'ws.Range("AP1:AQ1").Value = Array("x y", "z")',
    'ws.Range("AP2:AQ2").Value = Array(1, 4)',
    'ws.Range("AP3:AQ3").Value = Array(2, 5)',
    'ws.Range("AP4:AQ4").Value = Array(3, 6)',
    'ws.ListObjects.Add(xlSrcRange, ws.Range("AP1:AQ4"), , xlYes).Name = "Many"',
    'ws.Range("A30:C30").Value = Array("Qty", "Price", "Note")',
    *[f'ws.Range("A{row}:C{row}").Value = Array({row - 30}, {(row - 30) * 10}, "n{row - 30}")'
      for row in range(31, 43)],
    'ws.ListObjects.Add(xlSrcRange, ws.Range("A30:C42"), , xlYes).Name = "Big"',
])

#: Each case: its name, the sheet (1 or 2), the cell, the formula, and whether it is written as FormulaR1C1.
CASES: list[tuple[str, int, str, str, bool]] = [
    ("column", 1, "Z1", "=SUM(Table1[Qty])", False),
    ("case", 1, "Z2", "=SUM(table1[qty])", False),
    ("data_column", 1, "Z3", "=SUM(Table1[[#Data],[Qty]])", False),
    ("all_column", 1, "Z4", "=SUM(Table1[[#All],[Qty]])", False),
    ("header_cell", 1, "Z5", "=Table1[[#Headers],[Price]]", False),
    ("headers", 1, "Z6", "=COUNTA(Table1[#Headers])", False),
    ("all", 1, "Z7", "=ROWS(Table1[#All])", False),
    ("data", 1, "Z8", "=ROWS(Table1[#Data])", False),
    ("empty", 1, "Z9", "=ROWS(Table1[])", False),
    ("bare", 1, "Z10", "=ROWS(Table1)", False),
    ("span", 1, "Z11", "=SUM(Table1[[Qty]:[Price]])", False),
    ("span_reversed", 1, "Z12", "=SUM(Table1[[Price]:[Qty]])", False),
    ("headers_data", 1, "Z13", "=ROWS(Table1[[#Headers],[#Data]])", False),
    ("data_totals", 1, "Z14", "=ROWS(Table1[[#Data],[#Totals]])", False),
    ("totals", 1, "Z15", "=Table1[#Totals]", False),
    ("headers_data_column", 1, "Z16", "=ROWS(Table1[[#Headers],[#Data],[Qty]])", False),
    ("implicit_off", 1, "Z17", "=Table1[Qty]", False),
    ("unknown_column", 1, "Z18", "=SUM(Table1[Nope])", False),
    ("unknown_table", 1, "Z19", "=SUM(Nope[Qty])", False),
    ("unqualified_outside", 1, "Z20", "=SUM([Qty])", False),
    ("inner_spaces", 1, "Z21", "=SUM(Table1[ Qty ])", False),
    ("spaced_items", 1, "Z22", "=SUM(Table1[ [#All] , [Qty] ])", False),
    ("item_case", 1, "Z23", "=ROWS(Table1[#all])", False),
    ("range_operator", 1, "Z24", "=SUM(Table1[Qty]:Table1[Price])", False),
    ("index", 1, "Z25", "=INDEX(Table1[Qty],2)", False),
    ("match", 1, "Z26", '=MATCH("pear",Table1[Name],0)', False),
    ("vlookup", 1, "Z27", '=VLOOKUP("plum",Table1,3,FALSE)', False),
    ("vlookup_all", 1, "Z28", '=VLOOKUP("plum",Table1[#All],3,FALSE)', False),
    ("fixed_column", 1, "Z29", "=SUM(Table1[[Qty]:[Qty]])", False),
    ("all_span", 1, "Z30", "=SUM(Table1[[#All],[Qty]:[Price]])", False),
    ("headers_totals", 1, "Z31", "=ROWS(Table1[[#Headers],[#Totals]])", False),
    ("all_data", 1, "Z32", "=ROWS(Table1[[#All],[#Data]])", False),
    ("r1c1_written", 1, "Z34", "=SUM(Table1[Qty])+R1C1", True),
    ("sheet_qualified", 1, "Z35", "=SUM(Sheet1!Table1[Qty])", False),
    ("in_text", 1, "Z36", '="Table1[Qty]"', False),
    ("two_tables", 1, "Z38", "=SUM(Table1[Qty])+SUM(Inner[Qty])", False),
    ("union", 1, "Z39", "=SUM((Table1[Qty],Table1[Price]))", False),
    ("intersection", 1, "Z40", "=SUM(Table1[Qty] Table1[#Data])", False),
    ("odd_space", 1, "Z41", "=SUM(Odd[[Unit Price]])", False),
    ("odd_space_single", 1, "Z42", "=SUM(Odd[Unit Price])", False),
    ("odd_hash", 1, "Z43", "=SUM(Odd[[Q'#]])", False),
    ("odd_hash_bare", 1, "Z44", "=SUM(Odd[Q#])", False),
    ("odd_quote", 1, "Z45", "=SUM(Odd[[It''s]])", False),
    ("odd_brackets", 1, "Z46", "=SUM(Odd[['[x']]])", False),
    ("odd_at_start", 1, "Z47", "=SUM(Odd[['@home]])", False),
    ("odd_at_middle", 1, "Z48", "=SUM(Odd[[a@b]])", False),
    ("odd_comma", 1, "Z49", "=SUM(Odd[[a,b]])", False),
    ("odd_period", 1, "Z50", "=SUM(Odd[[a.b]])", False),
    ("odd_dollar", 1, "Z51", "=SUM(Odd[[Total $]])", False),
    ("odd_dollar_single", 1, "Z52", "=SUM(Odd[Total $])", False),
    ("odd_minus", 1, "Z53", "=SUM(Odd[[a-b]])", False),
    ("odd_minus_single", 1, "Z54", "=SUM(Odd[a-b])", False),
    ("odd_ampersand", 1, "Z55", "=SUM(Odd[[a&b]])", False),
    ("odd_colon", 1, "Z56", "=SUM(Odd[[a:b]])", False),
    ("odd_span", 1, "Z57", "=SUM(Odd[[Unit Price]:[Q'#]])", False),
    ("defined_name", 1, "Z58", "=SUM(Qtys)", False),
    ("defined_name_r1c1", 1, "Z59", "=SUM(QtysR1C1)", False),
    ("this_row_spelled", 1, "E2", "=Table1[[#This Row],[Qty]]*2", False),
    ("at", 1, "E3", "=Table1[@Qty]*2", False),
    ("implicit", 1, "F4", "=Table1[Qty]", False),
    ("at_outside", 1, "E9", "=Table1[@Qty]", False),
    ("at_header", 1, "E1", "=Table1[@Qty]", False),
    ("at_span", 1, "F2", "=SUM(Table1[@[Qty]:[Price]])", False),
    ("at_row", 1, "F3", "=COUNTA(Table1[@])", False),
    ("this_row_span_spelled", 1, "G2", "=SUM(Table1[[#This Row],[Qty]:[Price]])", False),
    ("this_row_alone_spelled", 1, "G3", "=COUNTA(Table1[#This Row])", False),
    ("at_bracketed", 1, "G4", "=Table1[@[Qty]]*1", False),
    ("unqualified_at_outside", 1, "E4", "=[@Qty]", False),
    ("odd_at_dollar", 1, "U2", "=Odd[@[Total $]]", False),
    ("odd_at_hash", 1, "V2", "=Odd[@[Q'#]]", False),
    ("odd_at_space", 1, "W2", "=Odd[@[Unit Price]]", False),
    ("at_with_reference", 1, "X3", "=Table1[@Qty]+A1", False),
    ("other_sheet", 2, "A1", "=SUM(Table1[Qty])", False),
    ("other_sheet_qualified", 2, "A2", "=SUM(Sheet1!Table1[Qty])", False),
    ("other_sheet_at", 2, "A3", "=Table1[@Qty]", False),
    ("inner_named", 1, "C21", "=Inner[@Qty]*2", False),
    ("inner_bare", 1, "C22", "=[@Qty]*3", False),
    ("inner_sum", 1, "C23", "=SUM(Inner[Qty])", False),
    ("inner_other", 1, "C24", "=SUM(Table1[Qty])", False),
    # Where one value is wanted: which of these Excel turns into this row's cell.
    ("scalar_data_column", 1, "Z61", "=Table1[[#Data],[Qty]]", False),
    ("scalar_data", 1, "Z62", "=Table1[#Data]", False),
    ("scalar_bare", 1, "Z63", "=Table1", False),
    ("scalar_span", 1, "Z64", "=Table1[[Qty]:[Price]]", False),
    ("scalar_all", 1, "Z65", "=Table1[#All]", False),
    ("scalar_all_column", 1, "Z66", "=Table1[[#All],[Qty]]", False),
    ("scalar_headers", 1, "Z67", "=Table1[#Headers]", False),
    ("scalar_headers_data_column", 1, "Z68", "=Table1[[#Headers],[#Data],[Qty]]", False),
    ("scalar_negated", 1, "Z69", "=-Table1[Qty]", False),
    ("scalar_sum", 1, "Z70", "=Table1[Qty]+1", False),
    ("scalar_abs", 1, "Z71", "=ABS(Table1[Qty])", False),
    ("scalar_if", 1, "Z72", "=IF(Table1[Qty]>4,1,0)", False),
    ("scalar_joined", 1, "Z73", '=Table1[Qty]&""', False),
    ("scalar_sumproduct", 1, "Z74", "=SUMPRODUCT(Table1[Qty])", False),
    ("scalar_rows", 1, "Z75", "=ROWS(Table1[Qty])", False),
    ("scalar_intersection", 1, "Z76", "=Table1[Qty] Table1[#Data]", False),
    ("scalar_grouped", 1, "Z77", "=(Table1[Qty])", False),
    ("scalar_odd_space", 1, "Z78", "=Odd[Unit Price]", False),
    ("scalar_odd_hash", 1, "Z79", "=Odd[Q'#]", False),
    ("scalar_header_cell", 1, "Z80", "=Table1[[#Headers],[Qty]]", False),
    ("scalar_mixed", 1, "Z81", "=Table1[@Qty]+Table1[Price]", False),
    ("scalar_isref", 1, "Z82", "=ISREF(Table1[Qty])", False),
    ("scalar_choose", 1, "Z83", "=CHOOSE(1,Table1[Qty])", False),
    ("scalar_index", 1, "Z84", "=INDEX(Table1,1,2)", False),
    ("scalar_this_row_span", 1, "Z85", "=Table1[[#This Row],[Qty]:[Price]]", False),
    ("scalar_iferror", 1, "Z86", "=IFERROR(Table1[Qty],0)", False),
    ("scalar_max", 1, "Z87", "=MAX(Table1[Qty])*Table1[Qty]", False),
    ("scalar_in_row", 1, "AG2", "=Table1[Qty]*1", False),
    ("scalar_data_in_row", 1, "AG3", "=Table1[[#Data],[Qty]]*1", False),
    ("scalar_bare_in_row", 1, "AG4", "=Table1*1", False),
    ("scalar_one_row_simple", 1, "Z88", "=One[a]", False),
    ("scalar_one_row_bracketed", 1, "Z89", "=One[b c]", False),
    ("scalar_many_rows_bracketed", 1, "Z90", "=Many[x y]", False),
    ("scalar_many_rows_simple", 1, "Z91", "=Many[z]", False),
    ("scalar_many_rows_in_row", 1, "AS3", "=Many[x y]*1", False),
    # Written into the Note column of the table named Big: how a formula spells its own table from inside it.
    ("own_all", 1, "C31", "=ROWS(Big[#All])", False),
    ("own_bare", 1, "C32", "=ROWS(Big)", False),
    ("own_scalar", 1, "C33", "=Big[Qty]", False),
    ("own_unqualified_scalar", 1, "C34", "=[Qty]*1", False),
    ("own_span", 1, "C35", "=SUM(Big[[Qty]:[Price]])", False),
    ("own_at_span", 1, "C36", "=SUM(Big[@[Qty]:[Price]])", False),
    ("own_at_row", 1, "C37", "=COUNTA(Big[@])", False),
    ("own_all_column", 1, "C38", "=SUM(Big[[#All],[Qty]])", False),
    ("own_header_cell", 1, "C39", "=Big[[#Headers],[Qty]]", False),
    ("own_data_column", 1, "C40", "=SUM(Big[[#Data],[Qty]])", False),
    ("own_data", 1, "C41", "=ROWS(Big[#Data])", False),
    ("own_unqualified_all", 1, "C42", "=ROWS([#All])", False),
]

#: Formulas written to a block, copied and filled; the cells whose formulas are read afterwards.
BLOCKS = "\n".join([
    'ws.Range("Y2:Y4").Formula = "=Table1[@Qty]*1"',
    'ws.Range("Z1").Copy ws.Range("AA1")',
    'ws.Range("E3").Copy ws.Range("E6")',
    'ws.Range("AB1").Formula = "=SUM(Table1[Name])"',
    'ws.Range("AB1:AE1").FillRight',
    'ws.Range("AB2").Formula = "=SUM(Table1[Qty])"',
    'ws.Range("AB2").AutoFill ws.Range("AB2:AD2")',
    'ws.Range("AB3").Formula = "=SUM(Table1[Qty])"',
    'ws.Range("AB3:AB5").FillDown',
    'ws.Range("AB7").Formula = "=SUM(Table1[[Qty]:[Qty]])"',
    'ws.Range("AB7:AC7").FillRight',
    'ws.Range("AB8").Formula = "=SUM(Table1[Qty])"',
    'ws.Range("AB8").Copy ws.Range("AC8:AD8")',
    'ws.Range("AB9").Formula = "=SUM(Table1[[#All],[Name]])"',
    'ws.Range("AB9").AutoFill ws.Range("AB9:AC9")',
    'ws.Range("AF12").Formula = "=SUM(Table1[Qty])"',
    'ws.Range("AF12").AutoFill ws.Range("AF12:AF14")',
    'ws.Range("AE24").Formula = "=SUM(Table1[Qty])"',
    'ws.Range("AE24").AutoFill ws.Range("AB24:AE24")',
    'ws.Range("AB25").Formula = "=SUM(Table1[Name])"',
    'ws.Range("AC25").Formula = "=SUM(Table1[Qty])"',
    'ws.Range("AB25:AC25").AutoFill ws.Range("AB25:AF25")',
    'ws.Range("AB26").Formula = "=SUM(Table1[Qty])"',
    'ws.Range("AB26").AutoFill ws.Range("AB26:AD26"), xlFillCopy',
])
BLOCK_CELLS = ["Y2", "Y3", "Y4", "AA1", "E6", "AC1", "AD1", "AE1", "AC2", "AD2", "AB4", "AB5", "AC7", "AC8", "AD8",
               "AC9", "AF13", "AF14", "AB24", "AC24", "AD24", "AD25", "AE25", "AF25", "AC26", "AD26"]

#: Formulas filled right by AutoFill from AB to AE, one row each: the row and the formula.
AUTOFILLS: list[tuple[int, str]] = [
    (12, "=SUM(Table1[Qty])"),
    (13, "=SUM(Table1[[Qty]:[Price]])"),
    (14, "=SUM(Table1[[Qty]:[Qty]])"),
    (15, "=SUM(Table1[[#Data],[Qty]])"),
    (16, "=SUM(Table1[[#Headers],[Qty]])"),
    (17, "=ROWS(Table1)"),
    (18, "=ROWS(Table1[#All])"),
    (19, "=SUM(Table1[Name])+SUM(Inner[Qty])"),
    (20, "=SUM(Odd[Unit Price])"),
    (21, "=SUM(Table1[Qty])+A1"),
    (22, "=SUM(Table1[[Price]:[Qty]])"),
    (23, "=SUM(Table1[@Qty])"),
    (3, "=Table1[@Qty]"),
    (4, "=SUM(Table1[@[Qty]:[Price]])"),
]

#: Every column name the character sweep tries, with a label for it.
CHARACTERS: list[tuple[str, str]] = [
    (f"c{code}", f"a{chr(code)}b") for code in [9, 10, 13, *range(32, 48), *range(58, 65), *range(91, 97),
                                                 *range(123, 127), 233]
] + [
    ("leading_space", " ab"),
    ("trailing_space", "ab "),
    ("leading_digit", "1ab"),
    ("cell_like", "A1"),
    ("r1c1_like", "R1C1"),
    ("boolean_like", "TRUE"),
    ("function_like", "Sum"),
    ("underscore", "_ab"),
    ("single_letter", "a"),
    ("leading_hash", "#x"),
    ("leading_at", "@x"),
    ("leading_quote", "'x"),
    ("digits", "123"),
]


def _spelled(text: str) -> str:
    """VBA that spells ``text``: literals for printable ASCII, ChrW for the rest."""
    pieces: list[str] = []
    run = ""
    for char in text:
        if " " <= char <= "~":
            run += char
            continue
        if run:
            pieces.append(_literal(run))
            run = ""
        pieces.append(f"ChrW({ord(char)})")
    if run or not pieces:
        pieces.append(_literal(run))
    return " & ".join(pieces)


def _escaped(name: str) -> str:
    """A column name as a formula spells it inside brackets."""
    return "".join("'" + char if char in "[]#'@" else char for char in name)

#: A table's headers overwritten one step after another: each step's name and its VBA.
REPEATS: list[tuple[str, str]] = [
    ("right_repeats_left", 'ws3.Range("C20").Value = "Two"'),
    ("other_case", 'ws3.Range("A20").Value = "two"'),
    ("cleared", 'ws3.Range("B20").ClearContents'),
    ("decimal", 'ws3.Range("A20").Value = 1.5'),
    ("boolean", 'ws3.Range("B20").Value = True'),
    ("date", 'ws3.Range("C20").Value = DateSerial(2020, 1, 2)'),
    ("setter_repeats", 'ws3.ListObjects("Rep").ListColumns(2).Name = "1.5"'),
    ("setter_unique", 'ws3.ListObjects("Rep").ListColumns(3).Name = "Cost"'),
    ("setter_blank", 'ws3.ListObjects("Rep").ListColumns(1).Name = ""'),
]

#: Questions: what each answers, or the error it raises.
QUESTIONS: list[tuple[str, str]] = [
    ("range_bare", 'ws.Range("Table1").Address'),
    ("range_column", 'ws.Range("Table1[Qty]").Address'),
    ("range_all", 'ws.Range("Table1[#All]").Address'),
    ("range_headers", 'ws.Range("Table1[#Headers]").Address'),
    ("range_data", 'ws.Range("Table1[#Data]").Address'),
    ("range_empty", 'ws.Range("Table1[]").Address'),
    ("range_all_column", 'ws.Range("Table1[[#All],[Qty]]").Address'),
    ("range_span", 'ws.Range("Table1[[Qty]:[Price]]").Address'),
    ("range_case", 'ws.Range("table1[qty]").Address'),
    ("range_at", 'ws.Range("Table1[@Qty]").Address'),
    ("range_totals", 'ws.Range("Table1[#Totals]").Address'),
    ("range_unknown", 'ws.Range("Table1[Nope]").Address'),
    ("range_odd", "ws.Range(\"Odd[[Q'#]]\").Address"),
    ("range_headers_data", 'ws.Range("Table1[[#Headers],[#Data]]").Address'),
    ("range_qualified", 'ws.Range("Sheet1!Table1[Qty]").Address'),
    ("range_other_sheet", 'ws2.Range("Table1[Qty]").Address'),
    ("range_application", 'Application.Range("Table1[Qty]").Address'),
    ("range_spaced", 'ws.Range("Table1[ Qty ]").Address'),
    ("evaluate_sum", 'Application.Evaluate("SUM(Table1[Qty])")'),
    ("evaluate_rows", 'ws.Evaluate("ROWS(Table1[#All])")'),
    ("evaluate_at", 'TypeName(Application.Evaluate("Table1[@Qty]"))'),
    ("evaluate_reference", 'Application.Evaluate("Table1[Qty]").Address'),
    ("evaluate_other_sheet", 'ws2.Evaluate("SUM(Table1[Qty])")'),
    ("evaluate_unqualified", 'TypeName(Application.Evaluate("[Qty]"))'),
    ("odd_columns", 'ColumnNames(ws.ListObjects("Odd"))'),
    ("name_refers_to", 'ws.Parent.Names("Qtys").RefersTo'),
]

#: The renames, one after another: each step's name and its VBA.
RENAMES: list[tuple[str, str]] = [
    ("table", 'ws.ListObjects("Table1").Name = "Sales"'),
    ("column", 'ws.Range("B1").Value = "Count"'),
    ("number", 'ws.Range("C1").Value = 5'),
    ("blank", 'ws.Range("A1").ClearContents'),
    ("repeat", 'ws.Range("A1").Value = "Count"'),
    ("list_column", 'ws.ListObjects(1).ListColumns(3).Name = "Cost"'),
]

HELPERS = """Private Function Rec(key As String, value As String) As String
    Rec = key & "~:~" & value & "~|~"
End Function

Private Function W(key As String, cell As Object, text As String, r1c1 As Boolean) As String
    On Error Resume Next
    If r1c1 Then cell.FormulaR1C1 = text Else cell.Formula = text
    If Err.Number <> 0 Then W = Rec("w:" & key, "!" & Err.Number) Else W = Rec("w:" & key, "ok")
End Function

Private Function Look(cell As Object) As String
    Dim v As Variant
    v = cell.Value
    Look = cell.Formula & "~;~" & cell.FormulaR1C1 & "~;~" & TypeName(v) & ":" & CStr(v)
End Function

Private Function AddName(key As String, ws As Object, name As String, text As String, r1c1 As Boolean) As String
    On Error Resume Next
    If r1c1 Then ws.Parent.Names.Add Name:=name, RefersToR1C1:=text Else ws.Parent.Names.Add Name:=name, RefersTo:=text
    If Err.Number <> 0 Then AddName = Rec("w:" & key, "!" & Err.Number) Else AddName = Rec("w:" & key, "ok")
End Function

Private Function NameText(ws As Object, name As String) As String
    On Error Resume Next
    NameText = "!"
    NameText = ws.Parent.Names(name).RefersTo
End Function

Private Function Try(cell As Object, text As String) As String
    On Error Resume Next
    cell.Formula = text
    If Err.Number <> 0 Then Try = "!" & Err.Number Else Try = cell.Formula
End Function

Private Function ColumnNames(t As Object) As String
    Dim c As Object, out As String
    For Each c In t.ListColumns
        out = out & c.Name & "=" & TypeName(c.Range.Cells(1).Value) & ":" & c.Range.Cells(1).Value & "/"
    Next
    ColumnNames = out
End Function
"""


def _sheet(sheet: int) -> str:
    return "ws" if sheet == 1 else "ws2"


def _literal(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def procedures() -> str:
    """The VBA the model runs too: Build, Writes, Reads, Blocks, Questions and Renames."""
    lines = ["Public Sub Build(ws As Object, ws2 As Object)", BUILD, "End Sub",
             "Public Function Writes(ws As Object, ws2 As Object) As String", "Dim out As String",
             'out = out & AddName("name_add", ws, "Qtys", "=Table1[Qty]", False)',
             'out = out & AddName("name_add_r1c1", ws, "QtysR1C1", "=Table1[Qty]", True)']
    lines += [f'out = out & W("{name}", {_sheet(sheet)}.Range("{cell}"), {_literal(formula)}, {r1c1})'
              for name, sheet, cell, formula, r1c1 in CASES]
    lines += ["Writes = out", "End Function",
              "Public Function Reads(ws As Object, ws2 As Object, step As String) As String", "Dim out As String"]
    lines += [f'out = out & Rec(step & ":{name}", Look({_sheet(sheet)}.Range("{cell}")))'
              for name, sheet, cell, _, _ in CASES]
    lines += ['out = out & Rec(step & ":columns", ColumnNames(ws.ListObjects(1)))',
              'out = out & Rec(step & ":refers_to", NameText(ws, "Qtys") & "/" & NameText(ws, "QtysR1C1"))',
              "Reads = out", "End Function",
              "Public Function Blocks(ws As Object) As String", "Dim out As String", "On Error Resume Next"]
    for index, line in enumerate(BLOCKS.splitlines()):
        lines += [line, f'If Err.Number <> 0 Then out = out & Rec("block_error:{index}", CStr(Err.Number)): Err.Clear']
    for row, formula in AUTOFILLS:
        lines += [f'ws.Range("AH{row}").Formula = {_literal(formula)}',
                  f'ws.Range("AH{row}").AutoFill ws.Range("AH{row}:AK{row}")',
                  f'If Err.Number <> 0 Then out = out & Rec("fill_error:{row}", CStr(Err.Number)): Err.Clear']
    lines += [f'out = out & Rec("block:{cell}", Look(ws.Range("{cell}")))' for cell in BLOCK_CELLS]
    lines += [f'out = out & Rec("fill:{row}", ws.Range("AI{row}").Formula & "~;~" & ws.Range("AJ{row}").Formula & '
              f'"~;~" & ws.Range("AK{row}").Formula)' for row, _ in AUTOFILLS]
    lines += ["Blocks = out", "End Function",
              "Public Function Questions(ws As Object, ws2 As Object) As String", "Dim out As String"]
    lines += [f'out = out & Rec("{name}", Q{index}(ws, ws2))' for index, (name, _) in enumerate(QUESTIONS)]
    lines += ["Questions = out", "End Function"]
    for index, (_, question) in enumerate(QUESTIONS):
        lines += [f"Private Function Q{index}(ws As Object, ws2 As Object) As String", "On Error GoTo Bad",
                  f"Q{index} = {question}", "Exit Function", "Bad:", f'Q{index} = "!" & Err.Number', "End Function"]
    lines += ["Public Function Renames(ws As Object, ws2 As Object) As String", "Dim out As String"]
    for step, vba in RENAMES:
        lines += ["Err.Clear", "On Error Resume Next", vba, f'out = out & Rec("{step}:error", CStr(Err.Number))',
                  "On Error GoTo 0", f'out = out & Reads(ws, ws2, "{step}")']
    lines += ["Renames = out", "End Function"]
    count = len(CHARACTERS)
    lines += ["Public Function Characters(ws3 As Object) As String", "Dim out As String, t As Object"]
    lines += [f'ws3.Cells(1, {index}).Value = "\'" & {_spelled(name)}' for index, (_, name) in
              enumerate(CHARACTERS, 1)]
    lines += [f"ws3.Cells(2, {index}).Value = {index}" for index in range(1, count + 1)]
    lines += [f"Set t = ws3.ListObjects.Add(xlSrcRange, ws3.Range(ws3.Cells(1, 1), ws3.Cells(2, {count})), , xlYes)",
              't.Name = "Chars"']
    for index, (label, name) in enumerate(CHARACTERS, 1):
        escaped = _spelled(_escaped(name))
        lines.append(f'out = out & Rec("char:{label}", t.ListColumns({index}).Name & "~;~" & '
                     f'Try(ws3.Cells(2, {count + 1 + index}), "=Chars[@[" & {escaped} & "]]") & "~;~" & '
                     f'Try(ws3.Cells(6, {index}), "=SUM(Chars[[" & {escaped} & "]])") & "~;~" & '
                     f'Try(ws3.Cells(7, {index}), "=SUM(Chars[" & {escaped} & "])"))')
    lines += ["Characters = out", "End Function",
              "Public Function Repeats(ws3 As Object) As String", "Dim out As String",
              'ws3.Range("A20:C20").Value = Array("One", "Two", "Three")',
              'ws3.Range("A21:C21").Value = Array(1, 2, 3)',
              'ws3.ListObjects.Add(xlSrcRange, ws3.Range("A20:C21"), , xlYes).Name = "Rep"',
              'ws3.Range("E20").Formula = "=SUM(Rep[One])"',
              'ws3.Range("E21").Formula = "=SUM(Rep[Two])"',
              'ws3.Range("E22").Formula = "=SUM(Rep[Three])"']
    for step, vba in REPEATS:
        lines += ["Err.Clear", "On Error Resume Next", vba, f'out = out & Rec("{step}:error", CStr(Err.Number))',
                  "On Error GoTo 0", f'out = out & Rec("{step}:columns", ColumnNames(ws3.ListObjects("Rep")))',
                  f'out = out & Rec("{step}:formulas", ws3.Range("E20").Formula & "~;~" & ws3.Range("E21").Formula & '
                  '"~;~" & ws3.Range("E22").Formula)']
    lines += ["Repeats = out", "End Function"]
    return HELPERS + "\n".join(lines) + "\n"


def formula2s() -> str:
    """What Formula2 reads for every case: Excel only, since the model has no Formula2."""
    lines = ["Public Function Formula2s(ws As Object, ws2 As Object) As String", "Dim out As String"]
    lines += [f'out = out & Rec("f2:{name}", {_sheet(sheet)}.Range("{cell}").Formula2)'
              for name, sheet, cell, _, _ in CASES]
    return "\n".join([*lines, "Formula2s = out", "End Function"]) + "\n"


def module() -> str:
    first, second = FOLDER / "structured.xlsx", FOLDER / "renamed.xlsx"
    probe = "\n".join([
        "Public Function Probe() As String",
        "Dim wb As Object, ws As Object, ws2 As Object, ws3 As Object, out As String",
        "Application.DisplayAlerts = False",
        "Set wb = Workbooks.Add(xlWBATWorksheet)",
        "Set ws = wb.Worksheets(1)",
        "Set ws2 = wb.Worksheets.Add(After:=ws)",
        "Set ws3 = wb.Worksheets.Add(After:=ws2)",
        "ws.Activate",
        'out = Rec("settings", Application.AutoCorrect.AutoFillFormulasInLists & "/" & '
        'Application.AutoCorrect.AutoExpandListRange)',
        "Build ws, ws2",
        "out = out & Writes(ws, ws2)",
        'out = out & Reads(ws, ws2, "r")',
        "out = out & Formula2s(ws, ws2)",
        "out = out & Blocks(ws)",
        "out = out & Questions(ws, ws2)",
        "out = out & Characters(ws3)",
        "out = out & Repeats(ws3)",
        f'wb.SaveAs Filename:="{first}", FileFormat:=51',
        "out = out & Renames(ws, ws2)",
        f'wb.SaveAs Filename:="{second}", FileFormat:=51',
        "wb.Close False",
        "Probe = out",
        "End Function",
    ])
    return procedures() + formula2s() + probe + "\n"


_CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"[^>]*?(?:/>|>(.*?)</c>)', re.DOTALL)
_F = re.compile(r"<f\b[^>]*?(?:/>|>.*?</f>)", re.DOTALL)


def formulas(path: Path) -> dict[str, dict[str, str]]:
    """The <f> element Excel wrote for every formula cell, per sheet part, and each table part."""
    out: dict[str, dict[str, str]] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            text = package.read(name).decode("utf-8")
            if name.startswith("xl/worksheets/sheet"):
                cells: dict[str, str] = {}
                for match in _CELL.finditer(text):
                    element = _F.search(match.group(2) or "")
                    if element is not None:
                        cells[match.group(1)] = element.group(0)
                out[name] = cells
            elif name.startswith("xl/tables/"):
                out[name] = {"xml": text}
    return out


def main() -> None:
    FOLDER.mkdir(parents=True, exist_ok=True)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers: dict[str, str] = {}
    for record in str(result.value).split("~|~"):
        if "~:~" in record:
            key, value = record.split("~:~", 1)
            answers[key] = value
    record = {
        "procedures": procedures(),
        "cases": [{"name": name, "sheet": sheet, "cell": cell, "formula": formula, "r1c1": r1c1}
                  for name, sheet, cell, formula, r1c1 in CASES],
        "block_cells": BLOCK_CELLS,
        "questions": [{"name": name, "question": question} for name, question in QUESTIONS],
        "renames": [{"step": step, "vba": vba} for step, vba in RENAMES],
        "answers": answers,
        "files": {"structured": formulas(FOLDER / "structured.xlsx"), "renamed": formulas(FOLDER / "renamed.xlsx")},
    }
    (FOLDER / "structured.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for key, value in answers.items():
        print(f"{key:40} {value}")


if __name__ == "__main__":
    main()
