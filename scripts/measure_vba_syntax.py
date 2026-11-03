"""Everyday VBA syntax, each form in a module of its own, recording its answer, the run-time error it stops with, or that it does not
compile: which forms compile is part of the language.

    python scripts/measure_vba_syntax.py

writes tests/fixtures/vba_syntax.json, which tests/test_vba_syntax.py replays.
"""
from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig
from pyvbaharness.results import HarnessError

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_syntax.json'
D = 'Dim s As String, i As Long, v As Variant, o As Object\n'
DICT = 'Set o = CreateObject("Scripting.Dictionary")\no("k") = 5\n'
#: name: (module-level text, body).
CASES: dict[str, tuple[str, str]] = {
    # Bang.
    'bang_dictionary': ('', DICT + 's = o!k'),
    'bang_assign': ('', DICT + 'o!k = 6\ns = o("k")'),
    'bang_bracket': ('', DICT + 'o("a b") = 7\ns = o![a b]'),
    'bang_after_call': ('', 's = Get1()!k'),
    'bang_space_before': ('', DICT + 's = o !k'),
    'bang_with': ('', DICT + 'With o\ns = !k\nEnd With'),
    'bang_range': ('', 's = TypeName(ActiveSheet!A1)'),
    'bang_suffix_still': ('', 'Dim x!\nx! = 1.5\ns = x! & TypeName(x!)'),
    'bang_suffix_then_op': ('', 'Dim x!\nx! = 2\ns = x! * 2'),
    'bang_literal': ('', 's = TypeName(1!) & 2!'),
    'bang_function_suffix': ('Function Half!(a)\nHalf! = a / 2\nEnd Function', 's = Half!(3)'),
    # Suffixes.
    'suffix_all': ('', 'Dim a%, b&, c!, d#, e@, f$\ns = TypeName(a) & TypeName(b) & TypeName(c) & TypeName(d) & TypeName(e) & TypeName(f)'),
    'suffix_function': ('', 's = Left$("abc", 1) & Mid$("abc", 2) & Right$("abc", 1) & UCase$("a") & LCase$("A") & Trim$(" a ") & TypeName(Left$("a", 1))'),
    'suffix_chr': ('', 's = Chr$(65) & ChrW$(66) & Str$(5) & Space$(1) & String$(2, "x") & Format$(1, "0") & Hex$(255) & Oct$(8)'),
    'suffix_conflict': ('', 'Dim a%\na& = 1'),
    'suffix_literal_hash': ('', 's = TypeName(1#) & TypeName(1@) & TypeName(1&) & TypeName(1%) & TypeName(&HFF) & TypeName(&HFF&) & TypeName(&O7) & TypeName(32768)'),
    'hex_negative': ('', 's = &HFFFF & "|" & &HFFFF& & "|" & &H8000 & "|" & &O177777'),
    'date_literal': ('', 's = #1/2/2020# & "|" & #3:04 PM# & "|" & TypeName(#1/2/2020#) & "|" & #January 2, 2020#'),
    'exponent_d': ('', 's = 1D2 & "|" & 1E2 & "|" & TypeName(1E2) & TypeName(1D2)'),
    # Continuation and separators.
    'continuation': ('', 's = "a" & _\n"b" & _\n"c"'),
    'continuation_in_call': ('', 's = Left( _\n"abc", _\n1)'),
    'colon_many': ('', 'i = 1: i = i + 1: s = i'),
    'colon_label': ('', 'GoTo L1\ni = 9\nL1: i = i + 1: s = i'),
    'number_label': ('', 'GoTo 10\ni = 9\n10 i = i + 1\ns = i'),
    'number_lines_erl': ('', 'On Error GoTo B\n10 i = 1\n20 i = 1 / 0\n30 s = "no"\nExit Function\nB:\ns = Erl()'),
    'comment_rem': ('', "s = \"a\" ' c\nRem c\ns = s & \"b\": Rem c"),
    'comment_continuation': ('', "s = \"a\" ' c _\n' still comment\ns = s & \"b\""),
    'if_line_colon': ('', 'If True Then i = 1: i = 2\nIf False Then i = 5: i = 6\ns = i'),
    'if_line_else_colon': ('', 'If False Then i = 1 Else i = 2: i = i + 1\ns = i'),
    'if_line_nested': ('', 'If True Then If False Then i = 1 Else i = 2\ns = i'),
    'if_then_newline_colon': ('', 'If True Then: i = 1: End If\ns = i'),
    'else_if_two_words': ('', 'If False Then\ni = 1\nElse If True Then\ni = 2\nEnd If\nEnd If\ns = i'),
    # Statements.
    'let_keyword': ('', 'Let i = 5\ns = i'),
    'call_keyword': ('Sub P(ByRef a)\na = a + 1\nEnd Sub', 'i = 1\nCall P(i)\nP i\nP (i)\ns = i'),
    'gosub': ('', 'GoSub A1\ns = s & "b"\nExit Function\nA1:\ns = s & "a"\nReturn'),
    'return_no_gosub': ('', 'Return'),
    'on_gosub': ('', 'i = 2\nOn i GoSub A1, A2\ns = s & "e"\nExit Function\nA1:\ns = "1"\nReturn\nA2:\ns = "2"\nReturn'),
    'on_goto': ('', 'i = 2\nOn i GoTo A1, A2\nExit Function\nA1:\ns = "1"\nExit Function\nA2:\ns = "2"'),
    'select_is_to': ('', 'i = 5\nSelect Case i\nCase Is < 3, 9\ns = "a"\nCase 4 To 6\ns = "b"\nCase Else\ns = "c"\nEnd Select'),
    'select_text_like': ('', 'Select Case "abc"\nCase "a" To "b"\ns = "x"\nCase Else\ns = "y"\nEnd Select'),
    'select_compare_first': ('', 'Select Case True\nCase 1 > 2\ns = "a"\nCase 2 > 1\ns = "b"\nEnd Select'),
    'with_nested_dot': ('', 'With ActiveSheet\nWith .Range("A1")\n.Value = 3\ns = .Value + .Parent.Range("A1").Value\nEnd With\nEnd With'),
    'with_new': ('', 'With CreateObject("Scripting.Dictionary")\n.Add "a", 1\ns = .Count\nEnd With'),
    'for_step_negative': ('', 'For i = 3 To 1 Step -1\ns = s & i\nNext i\ns = s & i'),
    'for_float': ('', 'Dim d As Double\nFor d = 0 To 1 Step 0.25\ns = s & d & ","\nNext'),
    'for_next_two': ('', 'Dim j As Long\nFor i = 1 To 2\nFor j = 1 To 2\ns = s & i & j\nNext j, i'),
    'for_wrong_next': ('', 'Dim j As Long\nFor i = 1 To 2\nFor j = 1 To 2\nNext i\nNext j'),
    'for_exit_do_inside': ('', 'For i = 1 To 2\nExit Do\nNext'),
    'for_change_bound': ('', 'Dim n As Long\nn = 2\nFor i = 1 To n\nn = 5\ns = s & i\nNext'),
    'while_wend': ('', 'While i < 2\ni = i + 1\nWend\ns = i'),
    'do_exit': ('', 'Do\ni = i + 1\nIf i = 3 Then Exit Do\nLoop\ns = i'),
    'erase_fixed': ('', 'Dim a(2) As Long\na(1) = 5\nErase a\ns = a(1) & UBound(a)'),
    'erase_dynamic': ('', 'Dim a() As Long\nReDim a(2)\nErase a\ns = UBound(a)'),
    'redim_preserve': ('', 'Dim a() As Long\nReDim a(1)\na(1) = 7\nReDim Preserve a(3)\ns = a(1) & UBound(a)'),
    'redim_preserve_first_dim': ('', 'Dim a() As Long\nReDim a(1, 1)\nReDim Preserve a(2, 1)'),
    'redim_undeclared': ('', 'ReDim zz(2)\nzz(1) = 3\ns = zz(1)'),
    'redim_to_syntax': ('', 'Dim a() As Long\nReDim a(2 To 4)\ns = LBound(a) & UBound(a)'),
    'dim_to_negative': ('', 'Dim a(-1 To 1) As Long\ns = LBound(a) & UBound(a)'),
    'mid_statement': ('', 's = "abcde"\nMid(s, 2, 2) = "XY"\nMid$(s, 4) = "Z"'),
    'lset_rset': ('', 'Dim a As String * 4\nLSet a = "x"\ns = "[" & a & "]"\nRSet a = "y"\ns = s & "[" & a & "]"'),
    'fixed_string': ('', 'Dim a As String * 3\na = "abcdef"\ns = a & Len(a)'),
    'static_var': ('Function C()\nStatic n As Long\nn = n + 1\nC = n\nEnd Function', 'i = C()\ns = C()'),
    'static_function': ('Static Function C()\nDim n As Long\nn = n + 1\nC = n\nEnd Function', 'i = C()\ns = C()'),
    'const_expression': ('Const A1 = 2, B1 = A1 * 3\nPrivate Const C1 As String = "x" & "y"', 's = B1 & C1'),
    'const_local_assign': ('', 'Const K = 1\nK = 2'),
    'enum_implicit': ('Enum E1\nA1\nB1 = 5\nC1\nEnd Enum', 's = A1 & B1 & C1 & E1.C1'),
    'type_nested': ('Type T1\nX As Long\nY(1) As String\nEnd Type\nType T2\nI As T1\nEnd Type', 'Dim q As T2\nq.I.X = 5\nq.I.Y(1) = "a"\ns = q.I.X & q.I.Y(1)'),
    'type_copy': ('Type T1\nX As Long\nEnd Type', 'Dim a As T1, b As T1\na.X = 1\nb = a\nb.X = 2\ns = a.X & b.X'),
    'type_in_variant': ('Type T1\nX As Long\nEnd Type', 'Dim a As T1\nv = a'),
    'option_base': ('Option Base 1', 'Dim a(2) As Long\ns = LBound(a) & UBound(a) & LBound(Array(1))'),
    'option_compare_text': ('Option Compare Text', 's = ("a" = "A") & ("a" < "B") & ("a" Like "A") & InStr("xA", "a") & StrComp("a", "A")'),
    'option_compare_binary': ('Option Compare Binary', 's = ("a" = "A") & ("a" < "B") & ("a" Like "A") & InStr("xA", "a")'),
    'option_explicit': ('Option Explicit', 'zz = 1'),
    'option_private': ('Option Private Module', 's = "ok"'),
    'deftype': ('DefLng A-C\nDefStr S', 'Dim a1, b1, st\ns = TypeName(a1) & TypeName(st)'),
    'declare_ptrsafe': ('#If VBA7 Then\nDeclare PtrSafe Function GetTickCount Lib "kernel32" () As Long\n#Else\nDeclare Function GetTickCount Lib "kernel32" () As Long\n#End If', 's = "ok"'),
    'cond_compile': ('#Const X1 = 2', '#If X1 = 2 Then\ns = "a"\n#ElseIf X1 = 3 Then\ns = "b"\n#Else\ns = "c"\n#End If\n#If Win64 Then\ns = s & "w"\n#End If\n#If VBA7 Then\ns = s & "7"\n#End If\n#If Mac Then\ns = s & "m"\n#End If'),
    'cond_undefined': ('', '#If Nope Then\ns = "a"\n#Else\ns = "b"\n#End If'),
    'longptr': ('', 'Dim p As LongPtr, q As LongLong\ns = TypeName(p) & TypeName(q) & LenB(p)'),
    'typeof_is': ('', 'Set o = ActiveSheet\ns = (TypeOf o Is Worksheet) & (TypeOf o Is Range) & (TypeOf o Is Object)'),
    'typeof_not': ('', 'Set o = ActiveSheet\nIf Not TypeOf o Is Range Then s = "n"'),
    'is_nothing_not': ('', 'If Not o Is Nothing Then s = "a" Else s = "b"'),
    'addressof': ('Sub Cb()\nEnd Sub\nFunction Ptr(ByVal p As LongPtr)\nPtr = p <> 0\nEnd Function', 's = Ptr(AddressOf Cb)'),
    'varptr': ('', 'Dim q As Long\ns = (VarPtr(q) <> 0) & (StrPtr("a") <> 0) & (ObjPtr(ActiveSheet) <> 0)'),
    # Operators.
    'operators_int': ('', 's = 7 \\ 2 & "|" & 7 Mod 2 & "|" & -7 Mod 2 & "|" & 7.5 Mod 2 & "|" & 2 ^ 3 ^ 2 & "|" & -2 ^ 2 & "|" & 7.5 \\ 2'),
    'operators_logic': ('', 's = (True Xor False) & (True Eqv False) & (True Imp False) & (5 And 3) & (5 Or 3) & (Not 5) & (5 Xor 3) & (5 Eqv 3) & (5 Imp 3)'),
    'operators_compare_null': ('', 'v = Null\ns = IsNull(v + 1) & IsNull(v & "a") & IsNull(v = 1) & (v & "a") & IsNull(Not v) & (True Or v) & IsNull(False Or v) & (False And v)'),
    'operators_concat_number': ('', 's = 1 & 2 & TypeName(1 & 2) & "1" + 2 & TypeName("1" + 2) & ("a" & 1)'),
    'plus_text_number': ('', 's = "a" + 1'),
    'like_patterns': ('', 's = ("abc" Like "a*") & ("a1" Like "a#") & ("ab" Like "a?") & ("b" Like "[a-c]") & ("d" Like "[!a-c]") & ("a*" Like "a[*]") & ("" Like "*")'),
    'comparison_chain': ('', 's = 1 < 2 < 3 & "|" & (3 > 2 > 1) & "|" & (1 = 1 = True)'),
    'integer_overflow': ('', 's = 200 * 200'),
    'long_mixed': ('', 's = 200& * 200 & TypeName(200 * 2) & TypeName(40000) & TypeName(2.5) & TypeName(1 / 2) & TypeName(2 \\ 1) & TypeName(True + 1) & TypeName(CByte(1) + CByte(1))'),
    'rounding_bankers': ('', 's = CInt(2.5) & CInt(3.5) & Round(2.5) & Round(2.675, 2) & CLng(-1.5) & Int(-1.5) & Fix(-1.5)'),
    'implicit_date_math': ('', 's = #1/2/2020# + 1 & "|" & TypeName(#1/2/2020# + 1) & "|" & #1/3/2020# - #1/2/2020# & "|" & TypeName(#1/3/2020# - #1/2/2020#)'),
    # Procedures.
    'named_args': ('Function F(a, Optional b = 2, Optional c = 3)\nF = a & b & c\nEnd Function', 's = F(1, c:=9) & F(b:=5, a:=4)'),
    'named_dup': ('Function F(a)\nF = a\nEnd Function', 's = F(a:=1, a:=2)'),
    'missing_optional': ('Function F(Optional a)\nF = IsMissing(a) & TypeName(a)\nEnd Function', 's = F() & F(1)'),
    'skipped_comma': ('Function F(Optional a = 1, Optional b = 2)\nF = a & b\nEnd Function', 's = F(, 5)'),
    'paramarray_empty': ('Function F(ParamArray p())\nF = UBound(p) & LBound(p) & TypeName(p)\nEnd Function', 's = F() & "|" & F(1, 2)'),
    'paramarray_byref': ('Sub F(ParamArray p())\np(0) = 9\nEnd Sub', 'i = 1\nF i\ns = i'),
    'function_as_variable': ('Function F()\nF = 1\nF = F + 1\nEnd Function', 's = F'),
    'recursive_name': ('Function F(n)\nIf n = 0 Then F = 0 Else F = n + F(n - 1)\nEnd Function', 's = F(3)'),
    'exit_property': ('Property Get P()\nP = 1\nExit Property\nP = 2\nEnd Property', 's = P'),
    'property_let_get': ('Dim m\nProperty Let P(a)\nm = a * 2\nEnd Property\nProperty Get P()\nP = m\nEnd Property', 'P = 2\ns = P'),
    'property_set': ('Dim m As Object\nProperty Set P(a)\nSet m = a\nEnd Property\nProperty Get P()\nSet P = m\nEnd Property', 'Set P = ActiveSheet\ns = TypeName(P)'),
    'byval_call_parens': ('Sub B(a)\na = 9\nEnd Sub', 'i = 1\nB (i)\ns = i\nB i\ns = s & i'),
    'array_literal_func': ('Function F()\nF = Array(1, 2)\nEnd Function', 's = F()(1) & UBound(F)'),
    'return_array_typed': ('Function F() As Long()\nDim a(1) As Long\na(1) = 4\nF = a\nEnd Function', 's = F()(1)'),
    'array_assign_copy': ('', 'Dim a(1) As Long, b() As Long\na(1) = 1\nb = a\nb(1) = 2\ns = a(1) & b(1)'),
    'array_assign_fixed': ('', 'Dim a(1) As Long, b(1) As Long\nb = a'),
    'variant_array_copy': ('', 'v = Array(1, 2)\nDim w\nw = v\nw(0) = 9\ns = v(0) & w(0)'),
    'for_each_array': ('', 'For Each v In Array(1, 2)\ns = s & v\nNext'),
    'for_each_long_var': ('', 'For Each i In Array(1)\nNext'),
    'split_join_filter': ('', 's = Join(Split("a,b", ","), "-") & UBound(Split("")) & Join(Filter(Array("ab", "c"), "a"), "") & TypeName(Split("a"))'),
    'set_non_object': ('', 'Set v = 5'),
    'let_object_default': ('', 'ActiveSheet.Range("A1").Value = 4\nv = ActiveSheet.Range("A1")\ns = TypeName(v) & v'),
    'error_statement': ('', 'On Error Resume Next\nError 11\ns = Err.Number & Err.Description'),
    'err_raise_custom': ('', 'On Error Resume Next\nErr.Raise vbObjectError + 5, "src", "desc"\ns = Err.Number - vbObjectError & Err.Source & Err.Description'),
    'resume_label': ('', 'On Error GoTo B\ni = 1 / 0\nA1:\ns = s & "a"\nExit Function\nB:\ns = "b"\nResume A1'),
    'resume_same': ('', 'On Error GoTo B\ni = 1 / v\ns = i\nExit Function\nB:\nv = 1\nResume'),
    'on_error_minus_one': ('', 'On Error GoTo B\ni = 1 / 0\nExit Function\nB:\nOn Error GoTo -1\nOn Error Resume Next\ni = 1 / 0\ns = Err.Number'),
    'error_in_handler': ('', 'On Error GoTo B\ni = 1 / 0\nExit Function\nB:\ni = 1 / 0'),
    'stop_statement_skipped': ('', 'If False Then Stop\ns = "a"'),
    'debug_print': ('', 'Debug.Print "a"; 1, "b"\nDebug.Assert True\ns = "ok"'),
    'beep_doevents': ('', 'i = DoEvents\ns = TypeName(i)'),
    'keyword_names': ('', 'Dim Name As String, Value As Long, Left As Long\nName = "a"\nValue = 1\nLeft = 2\ns = Name & Value & Left'),
    'keyword_named_arg_range': ('', 'Dim Range As Long\nRange = 1\ns = Range & TypeName(ActiveSheet.Range("A1"))'),
    'bracket_name': ('', 'Dim [a b] As Long\n[a b] = 2\ns = [a b]'),
    'bracket_evaluate': ('', 's = [1+1] & TypeName([A1])'),
    'unicode_name': ('', 'Dim café As Long\ncafé = 2\ns = café'),
    'underscore_name': ('', 'Dim a_b As Long, _c As Long'),
    'line_too_many_continuations': ('', 's = "a"' + ' & _\n"b"' * 25),
    'case_insensitive': ('', 'DIM A AS LONG\nA = 1\nIF a = 1 THEN S = "x"'),
    'empty_function_call': ('Function F()\nEnd Function', 's = TypeName(F()) & TypeName(F)'),
    'sub_value': ('Sub P()\nEnd Sub', 'i = P()'),
    'sub_call_undefined': ('', 'Nope1'),
    'wrong_arg_count': ('Function F(a)\nEnd Function', 'v = F(1, 2)'),
    'implicit_variant_type': ('', 'zz = 1\ns = TypeName(zz)\nzz = "a"\ns = s & TypeName(zz)'),
    'dim_multi_only_last': ('', 'Dim a, b As Long\ns = TypeName(a) & TypeName(b)'),
    'dim_new_auto': ('Class1Free', 's = "ok"'),
    'date_function_no_parens': ('', 'v = Date\ns = TypeName(v) & TypeName(Now) & TypeName(Timer) & TypeName(Time)'),
    'date_assign': ('', 'Dim d As Date\nd = "1/2/2020"\ns = d & TypeName(d) & CDbl(d)'),
    'string_compare_numbers': ('', 's = ("10" < "9") & (10 < 9) & ("10" < 9) & (v = 0) & (v = "") & IsEmpty(v)'),
    'boolean_text': ('', 's = CStr(True) & CBool("True") & CBool(1) & CInt(True) & (True = -1)'),
    'val_cdbl': ('', 's = Val("12abc") & Val("&H10") & Val(" 1 2") & CDbl("1.5") & CStr(1.5) & Str(1.5) & "|" & Str(-1)'),
    'iif_both_evaluated': ('Function F()\nF = 1\ni9 = i9 + 1\nEnd Function', 's = IIf(True, 1, 1 / 0)'),
    'choose_switch': ('', 's = Choose(2, "a", "b") & Switch(False, "x", True, "y") & TypeName(Choose(9, 1)) & TypeName(Switch(False, 1))'),
}
SAFE = re.compile(r'^\w+$')
PROBE = """{top}

Private Function Get1() As Object
    Set Get1 = CreateObject("Scripting.Dictionary")
    Get1("k") = 5
End Function

Function Probe() As String
    {d}{body}
    Probe = s
End Function
"""


def module(top: str, body: str) -> str:
    # A module-level statement has to come before the helper, and Option before everything.
    return PROBE.format(top=top, d=D.replace('\n', '\n    '), body=body.replace('\n', '\n    '))


def main() -> None:
    probes: list[dict[str, object]] = []
    excel = None
    try:
        for name, (top, body) in CASES.items():
            if top == 'Class1Free':
                continue
            source = module(top, body)
            if excel is None:
                excel = ExcelSession(HarnessConfig(lock_wait_s=1200.0))
                excel.__enter__()
                excel.new_document()
            try:
                result = excel.run_vba(source, 'Probe', timeout=30.0, line_numbers=False)
            except HarnessError as refused:
                # The editor would not take the module at all, as with too many line continuations.
                probes.append({'name': name, 'macro': source, 'compile': True})
                print(name, 'refused', str(refused)[:60], flush=True)
                continue
            if result.ok and not result.dialogs:
                answer: dict[str, object] = {'result': result.value}
            elif any(dialog.classification == 'compile-error' for dialog in result.dialogs):
                answer = {'compile': True}
            else:
                found = re.search(r'VBA error (\d+)', str(result.error))
                assert found, (name, result)
                answer = {'error': int(found.group(1))}
            if not result.ok:
                excel.__exit__(None, None, None)
                excel = None
            probes.append({'name': name, 'macro': source, **answer})
            print(name, answer, flush=True)
    finally:
        if excel is not None:
            excel.__exit__(None, None, None)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
