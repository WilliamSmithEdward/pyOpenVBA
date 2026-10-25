"""Application.ConvertFormula in live Excel: each formula converted each way, and what Excel answers.

Each case is one call of ConvertFormula, or a search for the longest formula of one shape that it still converts,
and records what came back -- its TypeName and its text, an error value's number, an array's bounds and items -- or
the error raised. The cases run in order in one Excel, which remembers the spelling it first saw each name in.

The first cases stand in a workbook made from the worksheet template, named Sheet1, whose sheets are Sheet1,
"Sheet 2" and Data; the rest in a new workbook, Book2, with more sheets and names, and a few in a third saved as
"conv test probe.xlsx" in the folder the probe is given.

    python scripts/measure_convert_formula.py

writes tests/fixtures/convert_formula.json, which tests/test_excel_convert_formula.py replays.
"""

from __future__ import annotations

import json
import shutil
import tempfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "convert_formula.json"
A1, R1C1 = "xlA1", "xlR1C1"
ABSOLUTE = ("xlAbsolute", "xlAbsRowRelColumn", "xlRelRowAbsColumn", "xlRelative")

A1_FORMULAS = ["=A1+B2", "=$A$1+A$2+$B3", "=SUM(A1:B3)", "=SUM($A$1:B$3)", "=A:A", "=$1:$2", "=A1:A", "=Sheet2!A1",
               "='Sheet 2'!B2:C3", "=MyName+A1", '="A1"&B1', "=IF(A1>0,B1,C1)", "=XFD1048576", "=A1 B2",
               "=A1,B2", "=Sheet1:Sheet2!A1", "=[Book2.xlsx]Sheet1!A1", "=#REF!+A1", "={1,2;3,4}", "=1+2", "A1+B1",
               "", "=A1+", "=SUM(", "=LOG10(A1)", "=R1C1", "=RC", "=a1+b2", "=A1%", "=-A1", "=@A1:A3", "=A1#",
               "=Table1[Col]", "=LET(x,A1,x+B1)", "=A1:B2:C3"]
R1C1_FORMULAS = ["=R1C1+R[1]C[1]", "=RC", "=R[-1]C", "=SUM(R1C1:R3C2)", "=C1", "=R1", "=C[1]:C[2]", "=R2C2",
                 "=Sheet2!R1C1", "='Sheet 2'!RC[1]", "=R[-5]C", "=A1", "=R1C1:R2C2:R3C3", "=r1c1"]


def vba_text(text: str) -> str:
    """``text`` as a VBA string expression, a line feed spelled vbLf."""
    return '"' + text.replace('"', '""').replace("\n", '" & vbLf & "') + '"'


def call(formula: str, start: str, end: str | None = None, absolute: str | None = None,
         relative: str | None = None) -> str:
    args = [vba_text(formula), start, end or "", absolute or "", relative or ""]
    while args and not args[-1]:
        args.pop()
    return "Application.ConvertFormula(" + ", ".join(args) + ")"


#: Each case: a label, the workbook it stands in ("template", "book" or "saved"), the cell to stand on first as
#: Sheet!Cell or None, and the VBA expression whose value is recorded.
CASES: list[tuple[str, str, str | None, str]] = []


def case(label: str, expression: str, *, book: str = "book", stand: str | None = None) -> None:
    CASES.append((label, book, stand, expression))


for index, formula in enumerate(A1_FORMULAS):
    case(f"a1_{index}_same", call(formula, A1), book="template")
    for absolute in ABSOLUTE:
        case(f"a1_{index}_{absolute}", call(formula, A1, A1, absolute), book="template")
    case(f"a1_{index}_r1c1", call(formula, A1, R1C1), book="template")
    case(f"a1_{index}_r1c1_from_c3", call(formula, A1, R1C1, None, 'ws.Range("C3")'), book="template")
    case(f"a1_{index}_r1c1_absolute", call(formula, A1, R1C1, "xlAbsolute", 'ws.Range("C3")'), book="template")
    case(f"a1_{index}_r1c1_mixed", call(formula, A1, R1C1, "xlAbsRowRelColumn", 'ws.Range("C3")'), book="template")
for index, formula in enumerate(R1C1_FORMULAS):
    case(f"r1c1_{index}_a1", call(formula, R1C1, A1), book="template")
    case(f"r1c1_{index}_a1_from_b2", call(formula, R1C1, A1, None, 'ws.Range("B2")'), book="template")
    case(f"r1c1_{index}_a1_from_e5_absolute", call(formula, R1C1, A1, "xlAbsolute", 'ws.Range("E5")'),
         book="template")
    case(f"r1c1_{index}_same", call(formula, R1C1), book="template")
    case(f"r1c1_{index}_relative", call(formula, R1C1, R1C1, "xlRelative", 'ws.Range("B2")'), book="template")
for label, formula, start, end, absolute, relative, stand in [
    ("active_d4", "=A1", A1, R1C1, None, None, "Sheet1!D4"),
    ("active_other_sheet", "=A1", A1, R1C1, None, None, "Sheet 2!B2"),
    ("relative_to_area", "=A1", A1, R1C1, None, 'ws.Range("C3:D4")', None),
    ("relative_to_other_sheet", "=A1", A1, R1C1, None, 'wb.Worksheets(2).Range("B2")', None),
    ("long_255", "=" + "+".join(["A1"] * 85)[:254], A1, A1, "xlAbsolute", None, None),
    ("long_256", "=" + "+".join(["A1"] * 86)[:255], A1, A1, "xlAbsolute", None, None),
    ("long_300", "=" + "+".join(["A1"] * 100), A1, A1, "xlAbsolute", None, None),
    ("number_formula", "5", A1, A1, "xlAbsolute", None, None),
    ("absolute_number", "=A1", A1, A1, "1", None, None),
    ("absolute_invalid", "=A1", A1, A1, "7", None, None),
    ("from_invalid", "=A1", "3", A1, None, None, None),
    ("to_invalid", "=A1", A1, "3", None, None, None),
    ("relative_to_text", "=A1", A1, R1C1, None, '"C3"', None),
    ("quoted_active_quoted", "='Sheet 2'!A1", A1, A1, None, None, "Sheet 2!B2"),
    ("quoted_active_data", "='Sheet 2'!A1", A1, A1, None, None, "Data!A1"),
    ("quoted_missing", "='My Sheet'!A1", A1, A1, None, None, None),
    ("quoted_digit", "='1Sheet'!A1", A1, A1, None, None, None),
    ("quoted_plain", "='Sheet2'!A1", A1, A1, None, None, None),
    ("quoted_r1c1", "='Sheet 2'!A1", A1, R1C1, None, None, None),
    ("long_same_255", "=" + "1+" * 126 + "1", A1, A1, None, None, None),
    ("long_same_256", "=" + "1+" * 127 + "1", A1, A1, None, None, None),
    ("long_output_over", "=" + "+".join(["A1"] * 70), A1, A1, "xlAbsolute", None, None),
    ("long_output_under", "=" + "+".join(["A1"] * 63), A1, A1, "xlAbsolute", None, None),
    ("lower_function", "=sum(a1:b2)", A1, A1, None, None, None),
    ("spaces", "= A1 + B1", A1, A1, None, None, None),
    ("trailing_space", "=A1+B1 ", A1, A1, None, None, None),
    ("spaced_call", "=SUM( A1 , B1 )", A1, A1, "xlAbsolute", None, None),
    ("line_break", "=A1+\nB1", A1, A1, None, None, None),
    ("unknown_function", "=NOSUCH(A1)", A1, A1, "xlAbsolute", None, None),
    ("text_function_name", "=IF(TRUE,A1)", A1, A1, "xlAbsolute", None, None),
    ("number_spelling", "=1.50+1E3+.5", A1, A1, None, None, None),
    ("error_literal", "=#N/A", A1, A1, None, None, None),
    ("percent_text", "=50%", A1, A1, None, None, None),
]:
    case(label, call(formula, start, end, absolute, relative), book="template", stand=stand)

# The second workbook: what its sheets and names make of references, and how names are spelled.
for label, formula, *rest in [
    ("book_sheet", "=Data!A1"), ("book_first_sheet", "=Sheet1!A1"), ("book_missing", "=Nosuch!A1"),
    ("book_quoted_missing", "='No such'!A1"), ("book_sum", "=SUM(Data!A1:B2)"), ("book_pair", "=Data!A1:Data!B2"),
    ("book_three_d", "=Sheet1:Data!A1"), ("book_three_d_reversed", "=Data:Sheet1!A1"),
    ("book_three_d_same", "=Data:Data!A1"), ("book_three_d_quoted", "='Sheet 2:Data'!A1"),
    ("book_three_d_quoted_reversed", "='Data:Sheet 2'!A1"), ("book_name_range_ref", "=Sheet1:'Sheet 2'!A1"),
    ("book_lower_sheet", "=data!a1"), ("book_lower_quoted", "='sheet 2'!a1"), ("book_digit_sheet", "='1Sheet'!A1"),
    ("book_apostrophe_sheet", "='O''Brien'!A1"), ("book_cell_sheet", "='A1'!B2"), ("book_r1c1_sheet", "='R1C1'!B2"),
    ("book_error_ref", "=Data!#REF!"), ("book_self", "=[Book2]Data!A1"), ("book_self_quoted", "='[Book2]Sheet 2'!A1"),
    ("book_self_missing", "=[Book2]Nosuch!A1"), ("book_never_seen", "=Book8!A1"),
    ("book_sheet_r1c1", "=Data!A1:B2", A1, R1C1),
    ("names_first_lower", "=zzname1+1"), ("names_then_upper", "=ZZNAME1+1"), ("names_then_mixed", "=ZzName1"),
    ("names_func_lower", "=zzfunc1(1)"), ("names_func_upper", "=ZZFUNC1(1)"),
    ("names_shared_name", "=zzshared"), ("names_shared_call", "=ZZSHARED(1)"),
    ("names_shared_call_first", "=zzshared2(1)"), ("names_shared_name_after", "=ZZSHARED2"),
    ("names_defined_lower", "=defd"), ("names_defined_mixed", "=DEFD+Defd"),
    ("names_qualified_missing", "=Data!Nope"), ("names_qualified_book_level", "=Sheet1!TheName"),
    ("names_qualified_quoted", "='Sheet 2'!TheName"), ("names_qualified_local", "=Data!localname"),
    ("names_qualified_unknown_sheet", "=Book9!TheName"), ("names_external_book", "=[Book9]Data!A1"),
    ("names_plain_text", "hello"), ("names_plain_text_upper", "HELLO world"), ("names_let", "=LET(Zq,1,zq+ZQ)"),
    ("names_bound_twice", "=LET(zq,1,zq)"),
    ("space_colon", "=A1 :B2"), ("colon_space", "=A1: B2"), ("spaced_colon", "=A1 : B2"),
    ("spaced_colon_r1c1", "=A1 :B2", A1, R1C1), ("spaced_colon_names", "=zza :zzb"), ("spaced_area_chain", "=A1 :B2:C3"),
    ("space_in_sum", "=SUM(A1 :B2)"), ("space_before_paren", "=SUM (A1)"), ("name_space_paren", "=zzname1 (A1)"),
    ("bare_function", "=sum"), ("bare_function_operand", "=Len+1"), ("bare_function_range", "=SUM:A1"),
    ("bare_function_r1c1", "=sum", A1, R1C1),
    ("space_before_comma", "=SUM(A1 ,B2)"), ("space_after_comma", "=SUM(A1, B2)"), ("union_spaced", "=(A1 ,B2)"),
    ("space_inside_parens", "=( A1 )"), ("space_empty_call", "=PI( )"), ("space_after_equals", "=  A1"),
    ("space_after_equals_r1c1", "=  A1", A1, R1C1), ("leading_space", " =A1"), ("two_equals", "==A1"),
    ("plus_start", "+A1"), ("minus_start", "-A1", A1, R1C1), ("at_start", "@A1"), ("braces", "{=A1}"),
    ("plus_ref", "=+A1"), ("plus_number", "=+1"), ("minus_zero", "=-0"), ("minus_plus", "=-+1"),
    ("trailing_lf", "=A1\n"), ("trailing_lf_lf", "=A1\n\n"), ("lf_r1c1", "=A1+\nB1", A1, R1C1),
    ("array_spaces", "={1, 2 ;3}"), ("array_numbers", "={1.50,1E3,-0.5}"), ("array_text", '={"a",TRUE,#N/A}'),
    ("huge_number", "=1E308"), ("precise_number", "=123456789012345678"), ("small_number", "=1E-5"),
    ("sum_of_tenths", "=0.1+0.2"), ("big_number", "=1.23456789012345E+20"), ("quoted_string", '="a""b"'),
    ("lower_error", "=#div/0!"), ("lower_booleans", "=true+false"), ("true_call", "=TRUE()"),
    ("arity_too_many", "=LEN(1,2)"), ("arity_none", "=LEN()"), ("arity_if", "=IF()"), ("arity_now", "=NOW(1)"),
    ("cells_wanted", "=SUMIF(1,1)"), ("row_of_number", "=ROW(1)"), ("structured", "=Items[Qty]"),
    ("structured_r1c1", "=items[qty]", A1, R1C1), ("structured_all", "=Items[#All]"), ("let_cell", "=LET(x1,1,x1)"),
    ("quoted_name", "='A1'"), ("quoted_name_r1c1", "='A1'", A1, R1C1), ("spill_r1c1", "=A1#", A1, R1C1),
    ("r1c1_text_in_a1", "=SUM(R1C1:R2C2)", A1, A1), ("a1_text_in_r1c1", "=SUM(A1:B2)", R1C1, R1C1),
    ("a1_absolute_in_r1c1", "=$A$1", R1C1, A1),
    ("r1c1_offset_rows_max", "=R[1048575]C", R1C1, A1), ("r1c1_offset_rows_over", "=R[1048576]C", R1C1, A1),
    ("r1c1_offset_rows_min", "=R[-1048575]C", R1C1, A1), ("r1c1_offset_rows_under", "=R[-1048576]C", R1C1, A1),
    ("r1c1_offset_columns_max", "=RC[16383]", R1C1, A1), ("r1c1_offset_columns_over", "=RC[16384]", R1C1, A1),
    ("r1c1_offset_columns_under", "=RC[-16384]", R1C1, A1), ("r1c1_offset_far", "=R[2000000]C", R1C1, A1),
    ("r1c1_zero_offsets", "=R[0]C[0]", R1C1, R1C1), ("r1c1_corner", "=R1048576C16384", R1C1, A1),
    ("r1c1_column_past", "=R1C16385", R1C1, A1), ("r1c1_row_zero", "=R0C1", R1C1, A1),
    ("r1c1_row_past", "=R1048577C1", R1C1, A1), ("r1c1_lower_offsets", "=r[1]c[-1]", R1C1, R1C1),
    ("r1c1_current_row", "=R", R1C1, A1), ("r1c1_rows", "=R1:R[2]", R1C1, A1), ("r1c1_columns", "=C[-1]:C[1]", R1C1, A1),
    ("r1c1_row_relative", "=R[1]", R1C1, A1), ("r1c1_order", "=R2C2:R1C1", R1C1, None),
    ("r1c1_sheet", "=Data!R1C1", R1C1, A1), ("r1c1_quoted_sheet", "='Sheet 2'!RC", R1C1, R1C1),
    ("order", "=B2:A1+$B$2:A1"), ("order_absolute", "=B2:A1", A1, A1, "xlAbsolute"),
    ("order_r1c1", "=B2:A1", A1, R1C1), ("order_rows_columns", "=3:1+B:A"),
    ("mixed_area_relative", "=A$1:$B2", A1, R1C1, "xlRelRowAbsColumn", 'ws.Range("C3")'),
    ("r1c1_to_relative", "=R1C1:R2C[1]", R1C1, R1C1, "xlRelative", 'ws.Range("C3")'),
]:
    start = rest[0] if rest else A1
    end = rest[1] if len(rest) > 1 else None
    absolute = rest[2] if len(rest) > 2 else None
    relative = rest[3] if len(rest) > 3 else None
    case(label, call(formula, start, end, absolute, relative))

# The arguments: what each kind of value given for each does.
for label, expression in [
    ("formula_integer", "Application.ConvertFormula(5, xlA1)"),
    ("formula_double", "Application.ConvertFormula(1.5, xlA1)"),
    ("formula_boolean", "Application.ConvertFormula(True, xlA1)"),
    ("formula_empty", "Application.ConvertFormula(Empty, xlA1)"),
    ("formula_null", "Application.ConvertFormula(Null, xlA1)"),
    ("formula_error", "Application.ConvertFormula(CVErr(2007), xlA1)"),
    ("formula_date", "Application.ConvertFormula(#1/2/2020#, xlA1)"),
    ("formula_cell", 'Application.ConvertFormula(ws.Range("H1"), xlA1, xlR1C1)'),
    ("formula_column", 'Application.ConvertFormula(ws.Range("H1:H2"), xlA1, xlR1C1)'),
    ("formula_row", 'Application.ConvertFormula(ws.Range("H1:I1"), xlA1, xlR1C1)'),
    ("formula_with_blank", 'Application.ConvertFormula(ws.Range("H2:H4"), xlA1)'),
    ("formula_formula_cell", 'Application.ConvertFormula(ws.Range("H3"), xlA1, xlR1C1)'),
    ("from_zero", 'Application.ConvertFormula("=A1", 0, xlR1C1)'),
    ("from_minus_one", 'Application.ConvertFormula("=A1", -1, xlR1C1)'),
    ("from_two", 'Application.ConvertFormula("=A1", 2, xlR1C1)'),
    ("from_true", 'Application.ConvertFormula("=A1", True, xlR1C1)'),
    ("from_false", 'Application.ConvertFormula("=RC", False, xlA1)'),
    ("from_text", 'Application.ConvertFormula("=A1", "1", xlR1C1)'),
    ("from_text_r1c1", 'Application.ConvertFormula("=RC", "-4150", xlA1)'),
    ("from_text_bad", 'Application.ConvertFormula("=RC", "x", xlA1)'),
    ("from_empty", 'Application.ConvertFormula("=RC", Empty, xlA1)'),
    ("from_null", 'Application.ConvertFormula("=A1", Null, xlR1C1)'),
    ("from_big", 'Application.ConvertFormula("=A1", 1E+20, xlR1C1)'),
    ("to_zero", 'Application.ConvertFormula("=A1", xlA1, 0)'),
    ("to_minus_one", 'Application.ConvertFormula("=A1", xlA1, -1)'),
    ("to_false", 'Application.ConvertFormula("=A1", xlA1, False)'),
    ("to_true", 'Application.ConvertFormula("=RC", xlR1C1, True)'),
    ("to_text_number", 'Application.ConvertFormula("=A1", xlA1, "-4150")'),
    ("to_text_one", 'Application.ConvertFormula("=A1", xlA1, "1")'),
    ("to_text_true", 'Application.ConvertFormula("=RC", xlR1C1, "TRUE")'),
    ("to_text_false_lower", 'Application.ConvertFormula("=A1", xlA1, "false")'),
    ("to_null", 'Application.ConvertFormula("=A1", xlA1, Null)'),
    ("to_empty", 'Application.ConvertFormula("=A1", xlA1, Empty)'),
    ("to_error", 'Application.ConvertFormula("=A1", xlA1, CVErr(2042))'),
    ("absolute_zero", 'Application.ConvertFormula("=$A$1+A1", xlA1, xlA1, 0)'),
    ("absolute_four", 'Application.ConvertFormula("=$A$1", xlA1, xlA1, 4)'),
    ("absolute_five", 'Application.ConvertFormula("=A1", xlA1, xlA1, 5)'),
    ("absolute_minus_one", 'Application.ConvertFormula("=A1", xlA1, xlA1, -1)'),
    ("absolute_one_half", 'Application.ConvertFormula("=A1", xlA1, xlA1, 1.5)'),
    ("absolute_two_half", 'Application.ConvertFormula("=A1", xlA1, xlA1, 2.5)'),
    ("absolute_nine_tenths", 'Application.ConvertFormula("=$A$1", xlA1, xlA1, 0.9)'),
    ("absolute_four_nine", 'Application.ConvertFormula("=$A$1", xlA1, xlA1, 4.9)'),
    ("absolute_minus_half", 'Application.ConvertFormula("=$A$1", xlA1, xlA1, -0.5)'),
    ("absolute_text", 'Application.ConvertFormula("=A1", xlA1, xlA1, "2")'),
    ("absolute_text_half", 'Application.ConvertFormula("=A1", xlA1, xlA1, "1.5")'),
    ("absolute_text_bad", 'Application.ConvertFormula("=A1", xlA1, xlA1, "x")'),
    ("absolute_text_empty", 'Application.ConvertFormula("=$A$1", xlA1, xlA1, "")'),
    ("absolute_true", 'Application.ConvertFormula("=A1", xlA1, xlA1, True)'),
    ("absolute_false", 'Application.ConvertFormula("=$A$1", xlA1, xlA1, False)'),
    ("absolute_null", 'Application.ConvertFormula("=$A$1", xlA1, xlA1, Null)'),
    ("absolute_empty", 'Application.ConvertFormula("=$A$1", xlA1, xlA1, Empty)'),
    ("relative_text_cell", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , "R2C2")'),
    ("relative_text_lower", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , "r2c2")'),
    ("relative_text_offsets", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , "R[1]C[1]")'),
    ("relative_text_name", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , "TheName")'),
    ("relative_text_equals", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , "=R2C2")'),
    ("relative_text_column", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , "C2")'),
    ("relative_text_row", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , "R2")'),
    ("relative_text_a1", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , "B2")'),
    ("relative_text_a1_unused", 'Application.ConvertFormula("=A1", xlA1, xlA1, , "B2")'),
    ("relative_text_area", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , "R3C3:R4C4")'),
    ("relative_text_sheet", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , "Data!R3C3")'),
    ("relative_text_empty", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , "")'),
    ("relative_number", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , 5)'),
    ("relative_nothing", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , Nothing)'),
    ("relative_areas", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , ws.Range("C3,E5"))'),
    ("relative_other_book", 'Application.ConvertFormula("=A1", xlA1, xlR1C1, , ThisWorkbook.Worksheets(1).Range("D4"))'),
    ("relative_unused", 'Application.ConvertFormula("=A1", xlA1, xlA1, , ws.Range("D4"))'),
    ("relative_r1c1_same", 'Application.ConvertFormula("=R[1]C", xlR1C1, xlR1C1, , ws.Range("D4"))'),
    ("relative_r1c1_text", 'Application.ConvertFormula("=RC", xlR1C1, xlA1, , "R3C3")'),
]:
    case(label, expression)

# A saved workbook's name, which a sheet's reference and a name's carry.
for label, formula in [("saved_sheet", "=Sheet1!A1"), ("saved_name", "=Sheet1!SavedName"),
                       ("saved_two_sheets", "=Sheet1!A1+Sheet2!B1")]:
    case(label, call(formula, A1), book="saved")

#: The tails a search puts after a long text, as A1 reads them.
A1_TAILS = ["1", "65535", "65536", "1.5", "0", "1E3", "TRUE", "#N/A", "A1", "$A$1", "A1:B2", "A:A", "Data!A1",
            "Nosuch!A1", "-1", "-A1", "-TRUE", "-(1)", "{1}", '{"a"}', "PI()", "SUM(1)", "a", "1%", "1.0", "--1",
            "{1,2}", "{TRUE}", "A1#", "@A1", "XFD1048576", "A1 B2", "(A1,B2)", "NOW()", "TODAY()", 'LEN("a")',
            "ISBLANK(TRUE)", "IF(TRUE,TRUE)", "IF(TRUE,TRUE,TRUE)", 'CONCATENATE("a","b")', "SUM(TRUE,TRUE)",
            'LEN(LEN("a"))', 'NOSUCH("a")', "TRUE&(TRUE&(TRUE&A1))", " A1", "  A1", "TRUE& A1", "A1 &TRUE", "(A1)",
            "((A1))", "TRUE+(A1)", "SUM(A1)", "SUM(TRUE,A1)", 'LEN("a")&A1', "TRUE&TRUE&A1", '"b"&"c"&1.5',
            "IF(TRUE,A1)", "-(1.5)", "{1.5}", "{1;2}", "{TRUE,1}", '{"a",1}', 'CONCAT("a")', "XLOOKUP(TRUE,TRUE,TRUE)",
            "LET(x,TRUE,x)", "SUMIFS(A1,A1,TRUE)", "zzfunc9(TRUE)", "TRUE&zzname9", "Data!zzname9", "Sheet1!TheName",
            "Data!#REF!", "#REF!", "[Book9]Data!A1", "( A1)", " (A1)", "- A1", " -A1", "A1 & B1", "SUM( A1)",
            " SUM(A1)", " zzfunc9(A1)", "zzfunc9( A1)", "{-1}", "A1:B2:C3", "A1 :B2", "zzname9 A1", "B2:A1",
            "PI()&A1", "SUM(TRUE)&A1"]
#: The tails as R1C1 reads them.
R1C1_TAILS = ["RC", "R1C1", "R[1]C[1]", "1", "1.5", "C", "R1C1:R2C2", "Data!RC", "R[-1]C"]
#: Each mode a search converts in: From, To and ToAbsolute.
MODES = {"a1": "xlA1", "a1_to_r1c1": "xlA1, xlR1C1", "a1_absolute": "xlA1, xlA1, xlAbsolute",
         "r1c1_to_a1": "xlR1C1, xlA1", "r1c1": "xlR1C1"}
for mode, args in MODES.items():
    tails = R1C1_TAILS if mode.startswith("r1c1") else A1_TAILS if mode == "a1" else A1_TAILS[:16]
    for tail in tails:
        case(f"longest_{mode}_{tail}", f"Longest({vba_text(tail)}, {list(MODES).index(mode)})")
#: Operands a search chains with one operator, "=a+a+...+a", as far as it converts.
for operand, operator in [("1", "+"), ("1.5", "+"), ("TRUE", "+"), ("#N/A", "+"), ('""', "+"), ("a", "+"),
                          ("A1", "+"), ("$A$1", "+"), ("A:A", "+"), ("A1:B2", "+"), ("Data!A1", "+"),
                          ("[Book9]S!A1", "+"), ("(1)", "+"), ("-1", "+"), ("1%", "+"), ("PI()", "+"),
                          ("SUM(1)", "+"), ("{1}", "+"), ("1", "&"), ("A1", " "), ("10", "+")]:
    case(f"chain_{operand}_{operator}", f"Chain({vba_text(operand)}, {vba_text(operator)})")

HELPER = r'''Private Function Shown(value As Variant) As String
    Dim item As Variant, out As String
    If IsArray(value) Then
        For Each item In value
            out = out & Shown(item) & ";"
        Next
        Shown = "Array(" & Bounds(value) & "):" & out
    ElseIf IsError(value) Then
        Shown = "Error:" & CLng(value)
    ElseIf IsNull(value) Then
        Shown = "Null"
    ElseIf IsObject(value) Then
        Shown = "Object:" & TypeName(value)
    Else
        Shown = TypeName(value) & ":" & CStr(value)
    End If
End Function

Private Function Bounds(value As Variant) As String
    Dim dimension As Long, out As String, upper As Long
    On Error GoTo Done
    For dimension = 1 To 3
        upper = UBound(value, dimension)
        out = out & IIf(dimension > 1, ",", "") & LBound(value, dimension) & " To " & upper
    Next
Done:
    Bounds = out
End Function

Private Function Built(tail As String, k As Long) As String
    Built = "=""" & String(k, "a") & """&" & tail
End Function

Private Function Converted(f As String, mode As Long) As Variant
    Select Case mode
        Case 0: Converted = Application.ConvertFormula(f, xlA1)
        Case 1: Converted = Application.ConvertFormula(f, xlA1, xlR1C1)
        Case 2: Converted = Application.ConvertFormula(f, xlA1, xlA1, xlAbsolute)
        Case 3: Converted = Application.ConvertFormula(f, xlR1C1, xlA1)
        Case 4: Converted = Application.ConvertFormula(f, xlR1C1)
    End Select
End Function

Private Function Longest(tail As String, mode As Long) As String
    Dim low As Long, high As Long, middle As Long, got As Variant
    low = -1: high = 300
    Do While high - low > 1
        middle = (low + high) \ 2
        If IsError(Converted(Built(tail, middle), mode)) Then high = middle Else low = middle
    Loop
    If low < 0 Then
        Longest = "none"
    Else
        got = Converted(Built(tail, low), mode)
        Longest = "k=" & low & " len=" & Len(got) & " " & Right(got, 40)
    End If
End Function

Private Function Chain(operand As String, operator As String) As String
    Dim low As Long, high As Long, middle As Long, f As String
    low = 0: high = 300
    Do While high - low > 1
        middle = (low + high) \ 2
        f = "=" & operand & Replace(String(middle - 1, " "), " ", operator & operand)
        If IsError(Application.ConvertFormula(f, xlA1)) Then high = middle Else low = middle
    Loop
    f = "=" & operand & Replace(String(low - 1, " "), " ", operator & operand)
    Chain = "n=" & low & " len=" & Len(f)
End Function
'''


def case_code(index: int, entry: tuple[str, str, str | None, str]) -> str:
    label, _, stand, expression = entry
    lines = [f"Private Function Case{index}(wb As Workbook, ws As Worksheet) As String",
             "    Dim got As Variant", "    On Error Resume Next"]
    if stand:
        sheet, _, cell = stand.rpartition("!")
        lines += [f'    wb.Worksheets("{sheet}").Activate', f'    wb.Worksheets("{sheet}").Range("{cell}").Select']
    lines += [f"    got = {expression}",
              f'    If Err.Number <> 0 Then Case{index} = {vba_text(label + "^E")} & Err.Number & ":" & Err.Description '
              f'Else Case{index} = {vba_text(label + "^")} & Shown(got)',
              "    Err.Clear", "    ws.Activate", '    ws.Range("A1").Select', "End Function"]
    return "\n".join(lines) + "\n"


SETUP = '''    Set template = Workbooks.Add(xlWBATWorksheet)
    template.Worksheets.Add(After:=template.Worksheets(1)).Name = "Sheet 2"
    template.Worksheets.Add(After:=template.Worksheets(2)).Name = "Data"
    Set book = Workbooks.Add
    For Each sheetName In Array("Data", "Sheet 2", "1Sheet", "O'Brien", "A1", "R1C1")
        book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count)).Name = sheetName
    Next
    book.Names.Add Name:="TheName", RefersTo:="=Data!$A$1"
    book.Names.Add Name:="Defd", RefersTo:="=1"
    book.Worksheets("Data").Names.Add Name:="LocalName", RefersTo:="=Data!$B$1"
    book.Worksheets(1).ListObjects.Add(xlSrcRange, book.Worksheets(1).Range("F1:G3"), , xlYes).Name = "Items"
    book.Worksheets(1).Range("F1").Value = "Qty"
    book.Worksheets(1).Range("G1").Value = "Cost"
    book.Worksheets(1).Range("H1").Value = "'=B1"
    book.Worksheets(1).Range("H2").Value = "'=C1"
    book.Worksheets(1).Range("I1").Value = "'=D1"
    book.Worksheets(1).Range("H3").Formula = "=1+1"
    Set saved = Workbooks.Add
    saved.Worksheets.Add(After:=saved.Worksheets(1)).Name = "Sheet2"
    saved.Names.Add Name:="SavedName", RefersTo:="=1"
    saved.SaveAs folder & "\\conv test probe.xlsx", 51
'''


def probe_code() -> str:
    """The probe: each case in a function, the cases run in batches of 40 in the workbook each stands in."""
    code = HELPER + "".join(case_code(index, entry) for index, entry in enumerate(CASES))
    books = {"template": "template", "book": "book", "saved": "saved"}
    runs: list[tuple[str, list[int]]] = []
    for index, (_, book, _, _) in enumerate(CASES):
        if not runs or runs[-1][0] != book or len(runs[-1][1]) == 40:
            runs.append((book, []))
        runs[-1][1].append(index)
    for number, (book, batch) in enumerate(runs):
        calls = "".join(f'    out = out & Case{index}(wb, ws) & "~"\n' for index in batch)
        code += (f"Private Function Batch{number}(wb As Workbook, ws As Worksheet) As String\n    Dim out As String\n"
                 "    wb.Activate\n    ws.Activate\n    ws.Range(\"A1\").Select\n"
                 + calls + f"    Batch{number} = out\nEnd Function\n")
    code += ("Public Function Probe(folder As String) As String\n"
             "    Dim template As Workbook, book As Workbook, saved As Workbook, out As String, sheetName As Variant\n"
             + SETUP
             + "".join(f"    out = out & Batch{number}({books[book]}, {books[book]}.Worksheets(1))\n"
                       for number, (book, _) in enumerate(runs))
             + "    saved.Close False\n    book.Close False\n    template.Close False\n    Probe = out\nEnd Function\n")
    return code


def main() -> None:
    code = probe_code()
    folder = Path(tempfile.mkdtemp(prefix="convert_formula_"))
    try:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(code, "Probe", [str(folder)], timeout=1800.0)
            assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    found = dict(entry.split("^", 1) for entry in str(result.value).split("~") if entry)
    record = {"probe": code, "cases": [[*entry, found[entry[0]]] for entry in CASES]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for entry in record["cases"]:
        print(entry[0], ascii(entry[-1])[:150])


if __name__ == "__main__":
    main()
