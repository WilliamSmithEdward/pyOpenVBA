"""A held-out round of TextRange2, the text of a shape: the members the model has, on texts and runs no earlier probe used, and a first
measurement of the text methods it lacks (ChangeCase, TrimText, AddPeriods, RemovePeriods, Find, Replace, InsertSymbol, LtrRun, RtlRun,
MathZones, Select). The model's answers are written down first, each case run alone, so a refusal answers that case only.

Each case makes a workbook with a rectangle R (100, 50, 240 by 120) holding one of four texts through TextFrame2.TextRange.Text, with two
of them given a bold part to make runs, runs its statement on tr, its TextRange, and reads what the statement returned and the text after.
A paragraph end reads as <CR>, a line break as <VT>.

    python scripts/measure_text_range2_round.py predict
    python scripts/measure_text_range2_round.py measure
    python scripts/measure_text_range2_round.py predict 2
    python scripts/measure_text_range2_round.py measure 2

writes tests/fixtures/text_range2_round.json, and text_range2_round_2.json for round 2, which tests/test_excel_text_range2_round.py replays.
"""
from __future__ import annotations

import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/text_range2_round.json'
SEPARATOR = '%%'
#: The texts, as VBA expressions, and the part made bold, if any.
TEXTS = [('"alpha beta. Gamma delta!" & vbCr & "second line here?"', (7, 4)), ('"  padded text  "', None),
         ('"one" & vbCr & "two" & vbCr & "three"', (5, 3)), ('"Mixed CASE words, here"', None)]
#: What every case reads after its statement.
AFTER = ['tr.Text', 'tr.Length', 'tr.Runs.Count', 'tr.Paragraphs.Count']
#: Reads of the members the model has, each on its own.
READS = ['tr.Count', 'tr.Characters(3, 5).Text', 'tr.Characters(3, 5).Start', 'tr.Characters(3, 5).Length', 'tr.Words.Count',
         'tr.Words(2).Text', 'tr.Words(2).Start', 'tr.Paragraphs(1).Text', 'tr.Paragraphs(1).Length', 'tr.Sentences.Count',
         'tr.Sentences(1).Text', 'tr.Runs(1).Text', 'tr.Runs(1).Font.Bold', 'tr.Paragraphs.Item(1).Start', 'tr.LanguageID',
         'tr.ParagraphFormat.Alignment', 'tr.Characters(2, 3).Font.Bold', 'tr.Characters.Count', 'tr.Words(1).Length']
#: Statements, each answering what it returns where it returns something.
STATEMENTS = ['v = tr.InsertAfter("X").Text', 'v = tr.InsertBefore("Y").Start', 'tr.Characters(1, 3).Delete',
              'v = tr.Paragraphs(1).InsertAfter("Z").Start', 'tr.ChangeCase msoCaseSentence', 'tr.ChangeCase msoCaseLower',
              'tr.ChangeCase msoCaseUpper', 'tr.ChangeCase msoCaseTitle', 'tr.ChangeCase msoCaseToggle', 'v = tr.TrimText.Text',
              'v = tr.TrimText.Start', 'tr.AddPeriods', 'tr.RemovePeriods', 'v = tr.Find("e").Start', 'v = tr.Find("E", MatchCase:=msoTrue).Start',
              'v = tr.Find("e", After:=5).Start', 'v = tr.Find("zzz").Length', 'v = tr.Find("t", WholeWords:=msoTrue).Start',
              'v = tr.Replace("e", "EE").Start', 'v = tr.Replace("e", "EE", MatchCase:=msoTrue).Text', 'v = tr.Replace("zzz", "q") Is Nothing',
              'v = tr.InsertSymbol("Arial", 65).Text', 'v = tr.InsertSymbol("Arial", 945, msoTrue).Start', 'tr.LtrRun', 'tr.RtlRun',
              'v = tr.MathZones.Count', 'tr.Select']
#: Round 2, on new texts: what round 1 left to the model's own rule (a Find with no match, a part's Find and Replace, case on a part,
#: sentence case after a stop, periods already there, trailing tabs and paragraph ends, the symbol's font, what Select selects, and the
#: direction RtlRun gives).
TEXTS_2 = [('"first one. second two? third" & vbCr & "next para"', None), ('"ab cd" & vbCr & "ef."', (1, 2)),
           ('"x  y" & vbTab', None), ('"Hello world hello"', (7, 5)), ('"don\'t stop-now (x)"', None)]
READS_2 = ['tr.Words.Count', 'tr.Words(1).Text', 'tr.Sentences.Count', 'tr.Sentences(2).Text']
STATEMENTS_2 = ['v = tr.Find("zzz") Is Nothing', 'v = tr.Find("zzz").Text', 'v = tr.Find("zzz").Count', 'v = tr.Characters(3, 6).Find("o").Start',
                'v = tr.Find("HELLO").Start', 'v = tr.Find("o", After:=100).Start', 'v = tr.Find("o").Count',
                'v = tr.Replace("hello", "bye", WholeWords:=msoTrue).Start', 'v = tr.Characters(1, 5).Replace("o", "0").Start',
                'tr.Characters(2, 4).ChangeCase msoCaseUpper', 'tr.Characters(2, 4).ChangeCase msoCaseTitle', 'tr.ChangeCase msoCaseSentence',
                'tr.ChangeCase msoCaseTitle', 'v = tr.TrimText.Length', 'v = tr.TrimText.Count', 'tr.AddPeriods', 'tr.RemovePeriods',
                'v = tr.InsertSymbol("Arial", 65).Font.Name', 'v = tr.Characters(2, 2).InsertSymbol("Symbol", 66).Start',
                'v = tr.InsertSymbol("Arial", 200).Text', 'tr.Select: v = TypeName(Selection)', 'tr.RtlRun: v = tr.ParagraphFormat.TextDirection',
                'tr.LtrRun: v = tr.ParagraphFormat.TextDirection', 'v = tr.MathZones(1).Text']
#: Round 3, held out on the code rounds 1 and 2 built, on new texts and arguments.
TEXTS_3 = [('"The cat sat. on the mat" & vbCr & "a-b c\'d"', (5, 3)), ('"  hi" & vbCr & "there."', None), ('"ABC def GHI"', (5, 3)),
           ('"q? r! s."', None)]
READS_3 = ['tr.Words.Count', 'tr.Words(1).Text', 'tr.Words(2).Start', 'tr.Sentences.Count', 'tr.Paragraphs(2).Text', 'tr.Runs.Count']
STATEMENTS_3 = ['tr.ChangeCase msoCaseSentence', 'tr.ChangeCase msoCaseLower', 'tr.ChangeCase msoCaseUpper', 'tr.ChangeCase msoCaseTitle',
                'tr.ChangeCase msoCaseToggle', 'tr.Characters(3, 5).ChangeCase msoCaseTitle', 'tr.Characters(1, 3).ChangeCase msoCaseLower',
                'v = tr.TrimText.Text', 'v = tr.Characters(1, 4).TrimText.Length', 'tr.AddPeriods', 'tr.RemovePeriods', 'v = tr.Find("a").Start',
                'v = tr.Find("A", MatchCase:=msoTrue).Start', 'v = tr.Find("e", After:=3).Start', 'v = tr.Find("cat", WholeWords:=msoTrue).Start',
                'v = tr.Characters(5, 3).Find("t").Start', 'v = tr.Find("q").Length', 'v = tr.Replace("a", "@").Start',
                'v = tr.Replace("the", "THE", MatchCase:=msoTrue) Is Nothing', 'v = tr.Characters(2, 6).Replace("e", "3").Text',
                'v = tr.InsertSymbol("Arial", 169).Text', 'v = tr.InsertSymbol("Calibri", 8364, msoTrue).Text',
                'v = tr.Characters(1, 1).InsertSymbol("Arial", 66).Font.Name', 'v = tr.MathZones.Count', 'tr.Select: v = TypeName(Selection)',
                'tr.RtlRun: v = tr.ParagraphFormat.TextDirection']
ROUNDS = {1: (TEXTS, READS, STATEMENTS, OUT), 2: (TEXTS_2, READS_2, STATEMENTS_2, OUT.with_name('text_range2_round_2.json')),
          3: (TEXTS_3, READS_3, STATEMENTS_3, OUT.with_name('text_range2_round_3.json'))}
ROUND = 1


def cases() -> list[list[Any]]:
    texts, reads, statements, _out = ROUNDS[ROUND]
    every: list[list[Any]] = []
    for text, bold in texts:
        every += [[text, bold, 'read', read] for read in reads]
        every += [[text, bold, 'do', statement] for statement in statements]
    return every


def _case(case: list[Any]) -> str:
    text, bold, how, what = case
    lines = ['    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set sh = ws.Shapes.AddShape(msoShapeRectangle, 100, 50, 240, 120)', '    Set tr = sh.TextFrame2.TextRange',
             f'    tr.Text = {text}']
    if bold:
        lines.append(f'    tr.Characters({bold[0]}, {bold[1]}).Font.Bold = msoTrue')
    lines += ['    s = ""', '    Err.Clear: v = Empty']
    if how == 'read':
        lines += [f'    v = {what}', '    s = Show(v)']
    else:
        lines += [f'    {what}', '    s = "e" & Err.Number & "=" & Show(v)']
        for read in AFTER:
            lines += ['    Err.Clear: v = Empty', f'    v = {read}', '    s = s & "|" & Show(v)']
    lines += ['    Set tr = Nothing: Set sh = Nothing', '    wb.Close SaveChanges:=False', f'    out = out & "{SEPARATOR}" & s']
    return '\n'.join(lines) + '\n'


HELPERS = '''Private Function Show(ByVal v As Variant) As String
    If Err.Number Then Show = "e" & Err.Number: Exit Function
    Show = TypeName(v) & ":" & Replace(Replace(Replace(CStr(v), vbCr, "<CR>"), vbLf, "<LF>"), Chr(11), "<VT>")
End Function
'''
DECLARE = '    Dim wb As Workbook, ws As Worksheet, sh As Object, tr As Object, v As Variant, s As String\n    On Error Resume Next\n'
PART = 12


def macro(chosen: list[list[Any]]) -> str:
    parts = [chosen[at:at + PART] for at in range(0, len(chosen), PART)]
    procedures = ''.join(f'Private Sub Part{index}(ByRef out As String)\n{DECLARE}{"".join(_case(case) for case in part)}End Sub\n'
                         for index, part in enumerate(parts))
    calls = ''.join(f'    Part{index} out\n' for index in range(len(parts)))
    return (f'{HELPERS}{procedures}Function Probe() As String\n    Dim out As String\n    Application.DisplayAlerts = False\n{calls}'
            f'    Application.DisplayAlerts = True\n    Probe = Mid(out, {len(SEPARATOR) + 1})\nEnd Function\n')


def alone(case: list[Any]) -> str:
    """The model's answer to one case run by itself."""
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBACompileError, VBAUnsupportedError

    app = ExcelApplication()
    app.add_workbook()
    try:
        app.add_module(macro([case]), name='H')
        return str(app.run('H.Probe'))
    except VBACompileError:
        return 'does not compile'
    except VBAUnsupportedError:
        return 'refused'


def predict() -> None:
    sys.path.insert(0, str(ROOT / 'src'))
    every = cases()
    text = json.dumps({'cases': every, 'model': [alone(case) for case in every]}, indent=1) + '\n'
    ROUNDS[ROUND][3].write_text(text, encoding='utf-8', newline='\n')
    print(len(every), 'cases; predictions sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())


def measure() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, Any] = json.loads(ROUNDS[ROUND][3].read_text(encoding='utf-8'))
    every: list[list[Any]] = record['cases']
    answers: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for at in range(0, len(every), 60):
            result = excel.run_vba(macro(every[at:at + 60]), 'Probe', (), timeout=600, line_numbers=False)
            assert result.ok, (at, result.error, [one.message for one in result.dialogs])
            answers += str(result.value).split(SEPARATOR)
            print(at + 60, 'cases', flush=True)
    record.update({'measured_at': datetime.now(timezone.utc).isoformat(), 'excel': answers})
    ROUNDS[ROUND][3].write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    wrong = sum(model != excel and model != 'refused' for model, excel in zip(record['model'], answers))
    print('model wrong', wrong, 'refused', record['model'].count('refused'), 'of', len(answers))


if __name__ == '__main__':
    ROUND = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
