"""Measure VBA Functions called from cells' formulas, user-defined functions: what
each kind of argument arrives as (a reference, a number, text, an array, one left
out), what a returned value, array, Range, Date or Error value puts in the cell, a
run-time error, Application.Caller, writing to a cell, recalculation, Evaluate,
and what the function leaves in the Err of the macro that wrote the formula.

Each probe is a module whose ``Function Probe() As String`` the fixture records
(see vba_probe_fixture.py); the functions the formulas call are in the same
module.
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_vba_udfs.json'

FUNCTIONS = '''
Public Function Twice(x As Variant) As Variant
    Twice = x * 2
End Function

Public Function CountCells(r As Range) As Long
    CountCells = r.Cells.Count
End Function

Public Function CallerAddress() As String
    CallerAddress = Application.Caller.Address
End Function

Public Function Joined(ParamArray xs() As Variant) As String
    Dim x As Variant
    For Each x In xs
        Joined = Joined & "[" & TypeName(x) & "]"
    Next
End Function

Public Function Fails() As Double
    Fails = 1 / 0
End Function

Public Function Returns2D() As Variant
    Dim a(1 To 2, 1 To 2) As Variant
    a(1, 1) = 1: a(1, 2) = 2: a(2, 1) = 3: a(2, 2) = 4
    Returns2D = a
End Function

Public Function ErrValue() As Variant
    ErrValue = CVErr(xlErrNA)
End Function

Public Function OptionalArg(Optional n As Variant) As String
    OptionalArg = IsMissing(n) & "|" & TypeName(n)
End Function

Public Function SumRange(r As Range) As Double
    Dim c As Range
    For Each c In r
        SumRange = SumRange + c.Value
    Next
End Function

Public Function WhatType(v As Variant) As String
    WhatType = TypeName(v)
End Function

Public Function Volatile1() As Double
    Application.Volatile
    Volatile1 = 1
End Function

Public Function TouchCell() As String
    On Error Resume Next
    Range("Z1").Value = 5
    TouchCell = "err " & Err.Number
End Function

Private Function Hidden() As Long
    Hidden = 1
End Function

Public Function ArrayShape(v As Variant) As String
    ArrayShape = LBound(v, 1) & UBound(v, 1) & "/" & LBound(v, 2) & UBound(v, 2)
End Function

Public Function ADate() As Date
    ADate = #1/2/2024#
End Function

Public Function Nothing1() As Variant
End Function

Public Function ARange() As Range
    Set ARange = Range("A2")
End Function

Public Function ReadsA2() As Double
    ReadsA2 = ActiveSheet.Range("A2").Value
End Function

Public Function UsesTwice(x As Variant) As Variant
    UsesTwice = Twice(x) + 1
End Function

Public Function NumText() As String
    NumText = "42"
End Function

Public Function LongResult() As Long
    LongResult = 7
End Function

Public Function BoolResult() As Boolean
    BoolResult = True
End Function

Public Function OneD() As Variant
    OneD = Array(5, 6, 7)
End Function

Public Function ByRefArg(ByRef x As Variant) As Variant
    x = 99
    ByRefArg = 1
End Function

Public Function MyCaller() As String
    MyCaller = TypeName(Application.Caller)
End Function

Public Function OneDimShape(v As Variant) As String
    OneDimShape = LBound(v) & UBound(v) & "|" & TypeName(v)
End Function

Public Function Handled() As Variant
    On Error Resume Next
    Err.Raise 13
    Handled = Err.Number
End Function
'''

CASES = {
    'udf_scalar': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Twice(A1)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_literal': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Twice(5)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_text_arg': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Twice(""x"")"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_range_arg': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=CountCells(A1:B3)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_caller': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=CallerAddress()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_paramarray': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Joined(1,""a"",A1,A1:A2,TRUE)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_error': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Fails()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_array_corner': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Returns2D()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_cverr': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=ErrValue()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_optional_missing': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=OptionalArg()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_optional_given': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=OptionalArg(1)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_sum_range': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=SumRange(A1:A2)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_arg_type_cell': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=WhatType(A1)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_arg_type_number': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=WhatType(1)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_arg_type_array': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=WhatType({1,2})"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_in_expression': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Twice(A1)+1"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_nested': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Twice(Twice(A2))"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_side_effect': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=TouchCell()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_private': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Hidden()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_unknown': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=NoSuchFunction(1)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_case': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=twice(A1)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_module_qualified': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Harness.Twice(A1)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_recalc_on_change': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Twice(A1)"\nws.Range("A1").Value = 10\nProbe = ws.Range("C3").Value',
    'udf_evaluate': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nProbe = Application.Evaluate("Twice(7)") & "|" & ws.Evaluate("Twice(A2)")',
    'udf_array_spill_formula2': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("E1").Formula2 = "=Returns2D()"\nProbe = ws.Range("F2").Value & "|" & ws.Range("E1").HasSpill',
    'array_arg_shape': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=ArrayShape({1,2,3})"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'array_arg_shape_2d': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=ArrayShape({1,2;3,4})"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'date_result': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=ADate()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'empty_result': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Nothing1()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'range_result': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=ARange()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'num_text_result': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=NumText()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'long_result': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=LongResult()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'bool_result': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=BoolResult()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'one_d_result': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=OneD()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'udf_calls_udf': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=UsesTwice(A1)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'byref_cell_arg': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=ByRefArg(A1)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'caller_typename': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=MyCaller()"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'error_arg': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=WhatType(1/0)"\nProbe = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Text',
    'caller_from_macro': 'On Error Resume Next\nDim v As Variant\nv = Application.Caller\nProbe = Err.Number & "|" & TypeName(v) & "|" & CStr(v)',
    'unreferenced_read_no_recalc': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Activate\nws.Range("C3").Formula = "=ReadsA2()"\nDim before As Double\nbefore = ws.Range("C3").Value\nws.Range("A2").Value = 40\nProbe = before & "|" & ws.Range("C3").Value\nws.Calculate\nProbe = Probe & "|" & ws.Range("C3").Value',
    'caller_err_kept': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nOn Error Resume Next\nErr.Raise 7\nws.Range("C3").Formula = "=Fails()"\nProbe = Err.Number & "|" & ws.Range("C3").Text & "|" & Err.Number',
    'one_d_spill': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("E1").Formula2 = "=OneD()"\nProbe = ws.Range("G1").Value & "|" & ws.Range("E1").SpillingToRange.Address',
    'date_result_numberformat': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=ADate()"\nProbe = ws.Range("C3").NumberFormat & "|" & TypeName(ws.Range("C3").Value) & "|" & ws.Range("C3").Value2',
    'one_row_array_shape': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=OneDimShape({1,2,3})"\nProbe = ws.Range("C3").Text',
    'one_col_array_shape': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=ArrayShape({1;2;3})"\nProbe = ws.Range("C3").Text',
    'fail_no_prior_error': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nOn Error Resume Next\nws.Range("C3").Formula = "=Fails()"\nProbe = Err.Number & "|" & ws.Range("C3").Text',
    'fail_without_resume_next': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Fails()"\nProbe = Err.Number & "|" & ws.Range("C3").Text',
    'ok_udf_err_kept': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nOn Error Resume Next\nErr.Raise 7\nws.Range("C3").Formula = "=Twice(A1)"\nProbe = Err.Number & "|" & ws.Range("C3").Text',
    'handled_udf_err': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nOn Error Resume Next\nErr.Raise 7\nws.Range("C3").Formula = "=Handled()"\nProbe = Err.Number & "|" & ws.Range("C3").Text',
    'fail_then_read_err': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nws.Range("C3").Formula = "=Fails()"\nDim t As String\nt = ws.Range("C3").Text\nProbe = Err.Number & "|" & Err.Description & "|" & t',
    'fail_value_read': 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3\nws.Range("A2").Value = 4\nOn Error Resume Next\nws.Range("C3").Formula = "=Fails()"\nErr.Clear\nDim v\nv = ws.Range("C3").Value\nProbe = Err.Number & "|" & CStr(v)',
}


def macro(name: str) -> str:
    return FUNCTIONS + '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT)


if __name__ == '__main__':
    main()
