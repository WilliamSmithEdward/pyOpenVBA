"""Calculation-client nesting and cleanup around the worksheet-UDF display boundary."""
from __future__ import annotations

import pytest

from pyopenvba._a1 import Area
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._model import Range
from pyopenvba.apps.excel._udf import evaluation_origin, running
from pyopenvba.exceptions import VBARuntimeError


def assert_clean(app: ExcelApplication) -> None:
    assert app.application.formula_origins == []
    assert app.application.udf_origins == []
    assert app.application.udf_callers == []


def test_origin_is_captured_for_invocation_and_restored_after_base_exception() -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    target = Range(book.sheets_[0], [Area(1, 1, 1, 1)])
    with evaluation_origin(app.application, 'worksheet'), running(app.application, target):
        with evaluation_origin(app.application, 'conditional'):
            # Changing the formula client alone cannot reclassify an already running VBA helper.
            with pytest.raises(VBARuntimeError) as failure:
                target.DisplayFormat()
            assert failure.value.number == 1004
            with running(app.application, target):
                assert target.DisplayFormat() is not None
        with pytest.raises(KeyboardInterrupt), evaluation_origin(app.application, 'evaluate'), running(app.application, target):
            assert target.DisplayFormat() is not None
            raise KeyboardInterrupt
        with pytest.raises(VBARuntimeError):
            target.DisplayFormat()
    assert target.DisplayFormat() is not None
    assert_clean(app)


@pytest.mark.parametrize('assignment,expected', [
    ('Range("B1").Formula="=' + 'SUM(' * 60 + 'ReadDisplay()' + ')' * 60 + '"', '1004'),
    ('Range("B1:B2").FormulaArray="=ReadDisplay()+ROW(B1:B2)*0"', '1004:1004'),
    ('Range("B1").Formula2="=ReadDisplay()+SEQUENCE(2)*0"', '1004:1004'),
], ids=['deep-scalar', 'legacy-array', 'dynamic-array'])
def test_actual_scalar_and_array_evaluation_captures_worksheet_origin(assignment: str, expected: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("""Function ReadDisplay() As Long
On Error GoTo Failed
ReadDisplay=Helper()
Exit Function
Failed:
ReadDisplay=Err.Number
End Function
Function Helper() As Long
Helper=Abs(CLng(Range("A1").DisplayFormat.Font.Bold))
End Function
Function Probe() As String
""" + assignment + """
Probe=CStr(Range("B1").Value)
If Not IsEmpty(Range("B2").Value) Then Probe=Probe & ":" & CStr(Range("B2").Value)
End Function
Function MacroRead() As Boolean
MacroRead=Range("A1").DisplayFormat.Font.Bold
End Function""", name='Harness')
    assert app.run('Probe') == expected
    assert app.run('MacroRead') is False
    assert_clean(app)


def test_referenced_worksheet_formula_overrides_conditional_origin_temporarily() -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module("""Function Leaf() As Long
On Error GoTo Failed
Leaf=CLng(Range("C1").DisplayFormat.Font.Bold)
Exit Function
Failed:
Leaf=Err.Number
End Function
Function ConditionalProbe() As Boolean
ConditionalProbe=(Range("B1").Value=1004 And Not Range("C1").DisplayFormat.Font.Bold)
End Function
Sub Setup()
Dim f As Object
Range("B1").Formula="=Leaf()"
Set f=Range("A1").FormatConditions.Add(xlExpression,Formula1:="=ConditionalProbe()")
f.Font.Bold=True
End Sub
Function Painted() As Boolean
Painted=Range("A1").DisplayFormat.Font.Bold
End Function""", name='Harness')
    app.run('Setup')
    book.sheets_[0].cells_[(1, 2)].stale = True
    assert app.run('Painted') is True
    assert book.sheets_[0].cells_[(1, 2)].value == 1004
    assert_clean(app)


def test_text_can_enter_conditional_udf_from_worksheet_udf() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("""Function ConditionalProbe() As Boolean
ConditionalProbe=Not Range("C1").DisplayFormat.Font.Bold
End Function
Function ReadText() As String
ReadText=Range("A1").Text
End Function
Function Probe() As String
Dim f As Object
Range("A1").Value=20
Set f=Range("A1").FormatConditions.Add(xlExpression,Formula1:="=ConditionalProbe()")
f.NumberFormat="0.00"
Range("B1").Formula="=ReadText()"
Probe=Range("B1").Value
End Function""", name='Harness')
    assert app.run('Probe') == '20.00'
    assert_clean(app)


@pytest.mark.parametrize('evaluate', ['Application.Evaluate', 'Worksheets(1).Evaluate'])
def test_evaluate_scope_ends_before_outer_worksheet_udf_continues(evaluate: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("""Function Leaf() As Long
Leaf=Abs(CLng(Range("C1").DisplayFormat.Font.Bold))+7
End Function
Function Outer() As String
Dim answer As Variant
On Error GoTo Failed
answer=""" + evaluate + """("Leaf()")
answer=CStr(answer) & ":"
answer=answer & CStr(Range("C1").DisplayFormat.Font.Bold)
Outer="UNEXPECTED:" & answer
Exit Function
Failed:
Outer=CStr(answer) & CStr(Err.Number)
End Function
Function Probe() As String
Range("B1").Formula="=Outer()"
Probe=Range("B1").Value
End Function""", name='Harness')
    assert app.run('Probe') == '7:1004'
    assert_clean(app)


def test_worksheet_origin_is_application_scoped_across_workbooks() -> None:
    app = ExcelApplication()
    host = app.add_workbook()
    app.add_workbook()
    app.activate_workbook(host)
    app.add_module("""Function OtherBookRead() As Long
On Error GoTo Failed
OtherBookRead=CLng(Workbooks(2).Worksheets(1).Range("A1").DisplayFormat.Font.Bold)
Exit Function
Failed:
OtherBookRead=Err.Number
End Function
Function Probe() As String
Range("A1").Formula="=OtherBookRead()"
Probe=CStr(Range("A1").Value)
End Function""", name='Harness', workbook=host)
    assert app.run('Probe') == '1004'
    assert_clean(app)


def test_unhandled_udf_error_and_deferred_dependency_retry_leave_no_origins() -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module("""Function Unhandled() As Boolean
Unhandled=Range("A1").DisplayFormat.Font.Bold
End Function
Sub Setup()
Dim i As Long
Range("B1").Formula="=Unhandled()"
For i=2 To 40
Cells(i,2).Formula="=B" & (i-1)
Next
End Sub
Function Probe() As String
Probe=CStr(Range("B40").Value)
End Function
Function MacroRead() As Boolean
MacroRead=Range("A1").DisplayFormat.Font.Bold
End Function""", name='Harness')
    app.run('Setup')
    for cell in book.sheets_[0].cells_.values():
        if cell.formula:
            cell.stale = True
    assert app.run('Probe') == 'Error 2015'
    assert app.run('MacroRead') is False
    assert_clean(app)


@pytest.mark.parametrize('evaluate', ['Application.Evaluate', 'Worksheets(1).Evaluate'])
def test_nested_evaluate_udf_cannot_write_and_macro_can_write_afterward(evaluate: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("""Function Writer() As Long
On Error GoTo Failed
Range("C1").Value=99
Writer=-1
Exit Function
Failed:
Writer=Err.Number
End Function
Function Outer() As Variant
Outer=""" + evaluate + """("Writer()")
End Function
Function Probe() As String
Range("C1").Value=17
Range("B1").Formula="=Outer()"
Probe=CStr(Range("B1").Value) & ":" & CStr(Range("C1").Value)
End Function
Function MacroWrite() As Long
Range("C1").Value=23
MacroWrite=Range("C1").Value
End Function""", name='Harness')
    assert app.run('Probe') == '1004:17'
    assert_clean(app)
    assert app.run('MacroWrite') == 23
    assert_clean(app)


def test_self_caller_text_uses_zero_then_actual_cache_without_marking_circular() -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module("""Public Detail As String
Function ReadSelf() As Double
Detail=Application.Caller.Text
ReadSelf=20
End Function
Sub Setup()
Range("A1").Formula="=ReadSelf()"
End Sub
Function Snapshot() As String
Snapshot=CStr(Range("A1").Value2) & ":" & Detail
End Function
Sub Recalculate()
Application.CalculateFull
End Sub""", name='Harness')
    app.run('Setup')
    assert app.run('Snapshot') == '20:0'
    assert book.calculator.circular == set()
    assert_clean(app)
    app.run('Recalculate')
    assert app.run('Snapshot') == '20:20'
    assert book.calculator.circular == set()
    assert_clean(app)


@pytest.mark.parametrize('read', ['Application.Caller.Value2', 'Range("A1").Value'])
def test_self_caller_value_reads_keep_existing_circular_behavior(read: str) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module("""Public Detail As String
Function ReadSelf() As Double
Dim v As Variant
Detail="before"
v=""" + read + """
Detail="after"
ReadSelf=20
End Function
Function Probe() As String
Range("A1").Formula="=ReadSelf()"
Probe=CStr(Range("A1").Value2) & ":" & Detail
End Function""", name='Harness')
    assert app.run('Probe') == '0:before'
    assert (book.sheets_[0].name.lower(), 1, 1) in book.calculator.circular
    assert_clean(app)
