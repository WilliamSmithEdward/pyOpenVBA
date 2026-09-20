"""The interpreter: statements, procedures, and the three ways to fail.

The expression semantics are held to live Excel in
``test_vba_semantics.py``.  This covers the parts a probe cannot reach:
control flow, how arguments are passed, what a class module does, and
which of the three errors comes out of which mistake.
"""

from __future__ import annotations

import pytest

from pyopenvba.exceptions import VBACompileError, VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter import Interpreter


def run(source: str, procedure: str = "Main", *args: object) -> Interpreter:
    vba = Interpreter()
    vba.add_module(source, name="Module1")
    vba.run(procedure, list(args))
    return vba


def printed(source: str, procedure: str = "Main") -> list[str]:
    return run(source, procedure).console


def value(source: str, procedure: str = "Main", *args: object) -> object:
    vba = Interpreter()
    vba.add_module(source, name="Module1")
    return vba.run(procedure, list(args))


# --- control flow ---------------------------------------------------------------------


def test_a_for_loop_counts_with_its_step() -> None:
    assert printed(
        "Sub Main()\n"
        "    Dim i As Long\n"
        "    For i = 10 To 1 Step -3\n"
        "        Debug.Print i\n"
        "    Next i\n"
        "End Sub\n"
    ) == [" 10 ", " 7 ", " 4 ", " 1 "]


def test_a_for_loop_that_never_runs_leaves_the_counter_past_the_limit() -> None:
    assert value(
        "Function Main() As Long\n"
        "    Dim i As Long\n"
        "    For i = 5 To 1\n"
        "    Next i\n"
        "    Main = i\n"
        "End Function\n"
    ) == 5


def test_nested_loops_share_one_next() -> None:
    assert printed(
        "Sub Main()\n"
        "    Dim i As Long, j As Long\n"
        "    For i = 1 To 2\n"
        "    For j = 1 To 2\n"
        "        Debug.Print i * 10 + j\n"
        "    Next j, i\n"
        "End Sub\n"
    ) == [" 11 ", " 12 ", " 21 ", " 22 "]


def test_exit_for_leaves_only_the_inner_loop() -> None:
    assert value(
        "Function Main() As Long\n"
        "    Dim i As Long, j As Long, n As Long\n"
        "    For i = 1 To 3\n"
        "        For j = 1 To 3\n"
        "            If j = 2 Then Exit For\n"
        "            n = n + 1\n"
        "        Next j\n"
        "    Next i\n"
        "    Main = n\n"
        "End Function\n"
    ) == 3


def test_do_loop_while_runs_its_body_once_before_testing() -> None:
    assert value(
        "Function Main() As Long\n"
        "    Dim n As Long\n"
        "    Do\n"
        "        n = n + 1\n"
        "    Loop While n < 0\n"
        "    Main = n\n"
        "End Function\n"
    ) == 1


def test_select_case_takes_the_first_match_only() -> None:
    assert printed(
        "Sub Main()\n"
        "    Dim n As Long\n"
        "    n = 3\n"
        "    Select Case n\n"
        "        Case 1, 2, 3\n"
        '            Debug.Print "list"\n'
        "        Case Is >= 3\n"
        '            Debug.Print "compare"\n'
        "    End Select\n"
        "End Sub\n"
    ) == ["list"]


def test_a_single_line_if_carries_its_else() -> None:
    assert printed(
        "Sub Main()\n"
        '    If 1 = 2 Then Debug.Print "yes" Else Debug.Print "no"\n'
        "End Sub\n"
    ) == ["no"]


def test_goto_jumps_to_a_label() -> None:
    assert printed(
        "Sub Main()\n"
        "    GoTo Skip\n"
        '    Debug.Print "never"\n'
        "Skip:\n"
        '    Debug.Print "here"\n'
        "End Sub\n"
    ) == ["here"]


def test_with_holds_the_object_for_a_leading_dot() -> None:
    vba = Interpreter()
    vba.add_module(
        "Sub Main()\n"
        "    Dim c As New Collection\n"
        "    With c\n"
        '        .Add "a"\n'
        '        .Add "b"\n'
        "        Debug.Print .Count\n"
        "    End With\n"
        "End Sub\n",
        name="Module1",
    )
    vba.run("Main")
    assert vba.console == [" 2 "]


# --- procedures ------------------------------------------------------------------------


def test_an_argument_is_by_reference_unless_it_says_otherwise() -> None:
    assert value(
        "Sub Bump(n As Long)\n"
        "    n = n + 1\n"
        "End Sub\n"
        "Function Main() As Long\n"
        "    Dim v As Long\n"
        "    v = 1\n"
        "    Bump v\n"
        "    Main = v\n"
        "End Function\n"
    ) == 2


def test_byval_keeps_the_caller_out_of_it() -> None:
    assert value(
        "Sub Bump(ByVal n As Long)\n"
        "    n = n + 1\n"
        "End Sub\n"
        "Function Main() As Long\n"
        "    Dim v As Long\n"
        "    v = 1\n"
        "    Bump v\n"
        "    Main = v\n"
        "End Function\n"
    ) == 1


def test_an_optional_argument_reports_itself_missing() -> None:
    assert printed(
        "Sub Show(Optional n As Variant)\n"
        "    Debug.Print IsMissing(n)\n"
        "End Sub\n"
        "Sub Main()\n"
        "    Show\n"
        "    Show 1\n"
        "End Sub\n"
    ) == ["True", "False"]


def test_an_optional_argument_takes_its_default() -> None:
    assert value(
        "Function Twice(Optional ByVal n As Long = 21) As Long\n"
        "    Twice = n * 2\n"
        "End Function\n"
        "Function Main() As Long\n"
        "    Main = Twice()\n"
        "End Function\n"
    ) == 42


def test_named_arguments_go_to_the_right_parameters() -> None:
    assert value(
        "Function Gap(ByVal a As Long, ByVal b As Long) As Long\n"
        "    Gap = a - b\n"
        "End Function\n"
        "Function Main() As Long\n"
        "    Main = Gap(b:=1, a:=10)\n"
        "End Function\n"
    ) == 9


def test_a_paramarray_collects_what_is_left() -> None:
    assert value(
        "Function Total(ParamArray parts() As Variant) As Long\n"
        "    Dim i As Long\n"
        "    For i = LBound(parts) To UBound(parts)\n"
        "        Total = Total + parts(i)\n"
        "    Next i\n"
        "End Function\n"
        "Function Main() As Long\n"
        "    Main = Total(1, 2, 3, 4)\n"
        "End Function\n"
    ) == 10


def test_recursion_works_and_is_bounded() -> None:
    assert value(
        "Function Fact(ByVal n As Long) As Double\n"
        "    If n <= 1 Then\n"
        "        Fact = 1\n"
        "    Else\n"
        "        Fact = n * Fact(n - 1)\n"
        "    End If\n"
        "End Function\n"
        "Function Main() As Double\n"
        "    Main = Fact(10)\n"
        "End Function\n"
    ) == 3628800.0


def test_a_run_away_recursion_is_error_28_rather_than_a_python_crash() -> None:
    with pytest.raises(VBARuntimeError) as raised:
        value(
            "Function Deeper(ByVal n As Long) As Long\n"
            "    Deeper = Deeper(n + 1)\n"
            "End Function\n"
            "Function Main() As Long\n"
            "    Main = Deeper(1)\n"
            "End Function\n"
        )
    assert raised.value.number == 28


def test_a_static_local_keeps_its_value_between_calls() -> None:
    assert printed(
        "Sub Count()\n"
        "    Static n As Long\n"
        "    n = n + 1\n"
        "    Debug.Print n\n"
        "End Sub\n"
        "Sub Main()\n"
        "    Count\n"
        "    Count\n"
        "    Count\n"
        "End Sub\n"
    ) == [" 1 ", " 2 ", " 3 "]


# --- arrays and types --------------------------------------------------------------------


def test_redim_preserve_keeps_what_was_there() -> None:
    assert printed(
        "Sub Main()\n"
        "    Dim a() As Long\n"
        "    ReDim a(1 To 2)\n"
        "    a(1) = 10\n"
        "    ReDim Preserve a(1 To 4)\n"
        "    a(4) = 40\n"
        "    Debug.Print a(1), a(4), UBound(a)\n"
        "End Sub\n"
    ) == [" 10            40            4 "]


def test_a_two_dimensional_array_indexes_both_ways() -> None:
    assert value(
        "Function Main() As Long\n"
        "    Dim grid(1 To 2, 1 To 3) As Long\n"
        "    grid(2, 3) = 6\n"
        "    Main = grid(2, 3) + UBound(grid, 2)\n"
        "End Function\n"
    ) == 9


def test_a_user_defined_type_holds_its_fields() -> None:
    assert printed(
        "Type Point\n"
        "    X As Long\n"
        "    Y As Long\n"
        "End Type\n"
        "Sub Main()\n"
        "    Dim p As Point\n"
        "    p.X = 3\n"
        "    p.Y = 4\n"
        "    Debug.Print p.X * p.Y\n"
        "End Sub\n"
    ) == [" 12 "]


def test_an_enum_names_its_numbers() -> None:
    assert printed(
        "Enum Colour\n"
        "    Red\n"
        "    Green\n"
        "    Blue = 10\n"
        "    Violet\n"
        "End Enum\n"
        "Sub Main()\n"
        "    Debug.Print Red, Green, Blue, Violet\n"
        "End Sub\n"
    ) == [" 0             1             10            11 "]


def test_a_class_module_holds_state_and_answers_its_properties() -> None:
    vba = Interpreter()
    vba.add_module(
        "Private mName As String\n"
        "Public Property Get Name() As String\n"
        "    Name = mName\n"
        "End Property\n"
        "Public Property Let Name(ByVal value As String)\n"
        "    mName = UCase(value)\n"
        "End Property\n"
        "Public Function Greet() As String\n"
        '    Greet = "hello " & mName\n'
        "End Function\n",
        name="Person",
        kind="class",
    )
    vba.add_module(
        "Sub Main()\n"
        "    Dim p As New Person\n"
        '    p.Name = "ada"\n'
        "    Debug.Print p.Name\n"
        "    Debug.Print p.Greet()\n"
        "End Sub\n",
        name="Module1",
    )
    vba.run("Main")
    assert vba.console == ["ADA", "hello ADA"]


def test_class_initialize_runs_when_the_instance_is_made() -> None:
    vba = Interpreter()
    vba.add_module(
        "Public Tag As String\n"
        "Private Sub Class_Initialize()\n"
        '    Tag = "ready"\n'
        "End Sub\n",
        name="Thing",
        kind="class",
    )
    vba.add_module(
        "Sub Main()\n    Dim t As New Thing\n    Debug.Print t.Tag\nEnd Sub\n",
        name="Module1",
    )
    vba.run("Main")
    assert vba.console == ["ready"]


# --- error handling ------------------------------------------------------------------------


def test_on_error_resume_next_carries_on_to_the_next_statement() -> None:
    assert printed(
        "Sub Main()\n"
        "    On Error Resume Next\n"
        "    Err.Raise 5\n"
        '    Debug.Print "after " & Err.Number\n'
        "End Sub\n"
    ) == ["after 5"]


def test_resume_next_inside_a_loop_carries_on_with_the_loop() -> None:
    """Resume Next goes to the statement after the failing one.

    That statement is inside the loop, so the loop keeps its counter and
    carries on: 2 is printed after the handler, because the Print is
    what follows the Raise.
    """
    assert printed(
        "Sub Main()\n"
        "    Dim i As Long\n"
        "    On Error GoTo Bad\n"
        "    For i = 1 To 3\n"
        "        If i = 2 Then Err.Raise 5\n"
        "        Debug.Print i\n"
        "    Next i\n"
        "    Exit Sub\n"
        "Bad:\n"
        '    Debug.Print "caught " & i\n'
        "    Resume Next\n"
        "End Sub\n"
    ) == [" 1 ", "caught 2", " 2 ", " 3 "]


def test_err_clears_after_the_handler_resumes() -> None:
    assert printed(
        "Sub Main()\n"
        "    On Error GoTo Bad\n"
        "    Err.Raise 9\n"
        "Done:\n"
        "    Debug.Print Err.Number\n"
        "    Exit Sub\n"
        "Bad:\n"
        "    Resume Next\n"
        "End Sub\n"
    ) == [" 0 "]


def test_an_error_with_no_handler_comes_out_as_a_runtime_error() -> None:
    with pytest.raises(VBARuntimeError) as raised:
        value("Sub Main()\n    Err.Raise 13\nEnd Sub\n")
    assert raised.value.number == 13
    assert "Module1.Main line 2" in raised.value.where


def test_err_raise_carries_its_description() -> None:
    with pytest.raises(VBARuntimeError) as raised:
        value('Sub Main()\n    Err.Raise 1001, "me", "no good"\nEnd Sub\n')
    assert (raised.value.number, raised.value.description) == (1001, "no good")


# --- the three kinds of failure --------------------------------------------------------------


def test_source_that_is_not_vba_is_a_compile_error() -> None:
    with pytest.raises(VBACompileError):
        Interpreter().add_module("Sub Main()\n    If Then\nEnd Sub\n", name="Module1")


def test_an_unclosed_block_is_a_compile_error_naming_the_block() -> None:
    with pytest.raises(VBACompileError) as raised:
        Interpreter().add_module("Sub Main()\n    If 1 = 1 Then\nEnd Sub\n", name="Module1")
    assert "End If" in str(raised.value)


def test_a_missing_variable_under_option_explicit_is_a_compile_error() -> None:
    with pytest.raises(VBACompileError):
        value("Option Explicit\nSub Main()\n    x = 1\nEnd Sub\n")


def test_a_statement_vba_has_and_this_lacks_is_unsupported_not_a_compile_error() -> None:
    with pytest.raises(VBAUnsupportedError) as raised:
        value('Sub Main()\n    Open "c:\\x.txt" For Output As #1\nEnd Sub\n')
    assert "file I/O" in str(raised.value)


def test_an_unsupported_statement_still_parses_so_the_module_loads() -> None:
    vba = Interpreter()
    vba.add_module(
        "Sub Touched()\n"
        '    Kill "c:\\x.txt"\n'
        "End Sub\n"
        "Function Main() As Long\n"
        "    Main = 7\n"
        "End Function\n",
        name="Module1",
    )
    assert int(vba.run("Main")) == 7  # type: ignore[arg-type]


def test_on_error_cannot_swallow_an_unsupported_error() -> None:
    """The gap is ours, so it has to reach the caller.

    Letting On Error Resume Next past it would leave the run carrying on
    over a step that never happened, and report a result computed from a
    state that never existed.
    """
    with pytest.raises(VBAUnsupportedError):
        value(
            "Sub Main()\n"
            "    On Error Resume Next\n"
            '    Kill "c:\\x.txt"\n'
            "End Sub\n"
        )


def test_a_declare_into_a_dll_is_unsupported_and_says_which_library() -> None:
    with pytest.raises(VBAUnsupportedError) as raised:
        value(
            'Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long\n'
            "Function Main() As Long\n"
            "    Main = GetTickCount()\n"
            "End Function\n"
        )
    assert "kernel32" in str(raised.value)


def test_createobject_is_unsupported_rather_than_a_silent_nothing() -> None:
    with pytest.raises(VBAUnsupportedError):
        value('Sub Main()\n    Dim o As Object\n    Set o = CreateObject("Scripting.Dictionary")\nEnd Sub\n')


# --- dialogs and the console ------------------------------------------------------------------


def test_msgbox_does_not_block_and_is_recorded() -> None:
    vba = Interpreter()
    vba.add_module('Sub Main()\n    MsgBox "hello", vbOKOnly, "Title"\nEnd Sub\n', name="Module1")
    vba.run("Main")
    assert vba.dialogs == [("MsgBox", "hello", "Title", 0)]


def test_a_queued_answer_is_what_msgbox_comes_back_with() -> None:
    vba = Interpreter()
    vba.answers.append(7)
    vba.add_module(
        "Sub Main()\n"
        '    If MsgBox("go on?", 4) = 7 Then Debug.Print "said no"\n'
        "End Sub\n",
        name="Module1",
    )
    vba.run("Main")
    assert vba.console == ["said no"]


def test_conditional_compilation_picks_one_branch() -> None:
    assert printed(
        "Sub Main()\n"
        "#If VBA7 Then\n"
        '    Debug.Print "seven"\n'
        "#Else\n"
        '    Debug.Print "six"\n'
        "#End If\n"
        "End Sub\n"
    ) == ["seven"]


def test_a_blanked_branch_keeps_the_line_numbers() -> None:
    with pytest.raises(VBARuntimeError) as raised:
        value(
            "Sub Main()\n"
            "#If Win16 Then\n"
            "    Dim old As Long\n"
            "#End If\n"
            "    Err.Raise 5\n"
            "End Sub\n"
        )
    assert "line 5" in raised.value.where
