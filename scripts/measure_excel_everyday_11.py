"""An eleventh sweep of everyday macros: what On Error Resume Next does when the header of a loop or a branch fails. VBA goes on to
the next statement, which is inside the block.

    python scripts/measure_excel_everyday_11.py

writes tests/fixtures/excel_everyday_11.json, which tests/test_excel_everyday_7.py replays with the others.
"""
from __future__ import annotations

from pathlib import Path

from measure_excel_everyday_7 import PRE
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_11.json'
ON = 'Dim w As Variant\nOn Error Resume Next\nd = 0\na = 5\nw = "abc"\n'
TAIL = '\ns = s & "|" & Err.Number'
CASES = {
    'for_each_nothing': ON + 'Set o = Nothing\nFor Each k In o\ns = s & "in"\nNext\ns = s & "out"' + TAIL,
    'for_each_number': ON + 'For Each k In a\ns = s & "in"\nNext\ns = s & "out"' + TAIL,
    'for_each_text': ON + 'For Each k In w\ns = s & "in"\nNext\ns = s & "out"' + TAIL,
    'for_each_empty_variant': ON + 'For Each k In v\ns = s & "in"\nNext\ns = s & "out"' + TAIL,
    'for_each_uninitialised_array': ON + 'Dim u() As Long\nFor Each k In u\ns = s & "in"\nNext\ns = s & "out"' + TAIL,
    'for_each_gone_range': ON + 'Set r = ws.Range("B3")\nws.Rows(3).Delete\nFor Each c In r\ns = s & "in"\nNext\ns = s & "out"' + TAIL,
    'for_each_gone_range_var': ON + 'Set r = ws.Range("B3")\nws.Rows(3).Delete\nFor Each k In r\ns = s & TypeName(k)\nNext\ns = s & "out"' + TAIL,
    'for_each_clears_after': ON + 'For Each k In a\nErr.Clear\ns = s & Err.Number\nNext\ns = s & "out"' + TAIL,
    'for_each_twice_body': ON + 'For Each k In a\ns = s & "a"\ns = s & "b"\nNext\ns = s & "c"' + TAIL,
    'for_each_error_stops_without_resume': 'Set o = Nothing\nFor Each k In o\ns = s & "in"\nNext',
    'for_bad_bound': ON + 'For i = 1 To "x"\ns = s & "in" & i\nIf Len(s) > 20 Then Exit For\nNext\ns = s & "out"' + TAIL,
    'for_bad_start': ON + 'For i = "x" To 2\ns = s & "in" & i\nIf Len(s) > 20 Then Exit For\nNext\ns = s & "out"' + TAIL,
    'for_bad_step': ON + 'For i = 1 To 2 Step "x"\ns = s & "in" & i\nIf Len(s) > 20 Then Exit For\nNext\ns = s & "out"' + TAIL,
    'for_division_bound': ON + 'For i = 1 To 1 / d\ns = s & "in" & i\nIf Len(s) > 20 Then Exit For\nNext\ns = s & "out"' + TAIL,
    'do_while_top': ON + 'Do While 1 / d\nn = n + 1\ns = s & "in"\nIf n > 3 Then Exit Do\nLoop\ns = s & "out"' + TAIL,
    'do_until_top': ON + 'Do Until 1 / d\nn = n + 1\ns = s & "in"\nIf n > 3 Then Exit Do\nLoop\ns = s & "out"' + TAIL,
    'loop_while_bottom': ON + 'Do\nn = n + 1\ns = s & "in"\nIf n > 3 Then Exit Do\nLoop While 1 / d\ns = s & "out"' + TAIL,
    'loop_until_bottom': ON + 'Do\nn = n + 1\ns = s & "in"\nIf n > 3 Then Exit Do\nLoop Until 1 / d\ns = s & "out"' + TAIL,
    'while_wend': ON + 'While 1 / d\nn = n + 1\ns = s & "in"\nIf n > 3 Then GoTo Done\nWend\nDone:\ns = s & "out"' + TAIL,
    'if_block': ON + 'If 1 / d Then\ns = "then"\nElse\ns = "else"\nEnd If\ns = s & "out"' + TAIL,
    'if_block_no_else': ON + 'If 1 / d Then\ns = "then"\nEnd If\ns = s & "out"' + TAIL,
    'if_line': ON + 'If 1 / d Then s = "then" Else s = "else"\ns = s & "out"' + TAIL,
    'elseif': ON + 'If False Then\ns = "one"\nElseIf 1 / d Then\ns = "two"\nElse\ns = "three"\nEnd If\ns = s & "out"' + TAIL,
    'elseif_after_true': ON + 'If True Then\ns = "one"\nElseIf 1 / d Then\ns = "two"\nElse\ns = "three"\nEnd If\ns = s & "out"' + TAIL,
    'select_header': ON + 'Select Case 1 / d\nCase 1\ns = "one"\nCase Else\ns = "else"\nEnd Select\ns = s & "out"' + TAIL,
    'select_case_label': ON + 'Select Case 1\nCase 1 / d\ns = "one"\nCase Else\ns = "else"\nEnd Select\ns = s & "out"' + TAIL,
    'select_second_label': ON + 'Select Case 2\nCase 1\ns = "one"\nCase 1 / d\ns = "two"\nCase Else\ns = "else"\nEnd Select\ns = s & "out"' + TAIL,
    'with_header': ON + 'Set o = Nothing\nWith o\ns = s & "in"\ns = s & .Name\nEnd With\ns = s & "out"' + TAIL,
    'with_number': ON + 'With a\ns = s & "in"\nEnd With\ns = s & "out"' + TAIL,
    'exit_for_after_failed_header': ON + 'For Each k In a\ns = s & "in"\nExit For\nNext\ns = s & "out"' + TAIL,
    'exit_do_after_failed_header': ON + 'Do While 1 / d\ns = s & "in"\nExit Do\nLoop\ns = s & "out"' + TAIL,
    'nested_for_each_failed': ON + 'For i = 1 To 2\nFor Each k In a\ns = s & "in" & i\nNext\ns = s & "mid"\nNext\ns = s & "out"' + TAIL,
    'goto_handler_for_each': 'a = 5\nOn Error GoTo Bad\nFor Each k In a\ns = s & "in"\nNext\ns = s & "out"\nGoTo Done\nBad:\ns = s & "bad" & Err.Number\nResume Next\nDone:',
}


def macro(name: str) -> str:
    cleanup = "\nOn Error GoTo 0\nApplication.DisplayAlerts = False\nws.Delete\nApplication.DisplayAlerts = True"
    return "\nFunction Probe() As String\n" + PRE + CASES[name] + cleanup + "\nProbe = s\nEnd Function\n"


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
