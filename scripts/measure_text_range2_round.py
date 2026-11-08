"""A held-out round of TextRange2, the text of a shape: the members the model has, on texts and runs no earlier probe used, and a first
measurement of the text methods it lacks (ChangeCase, TrimText, AddPeriods, RemovePeriods, Find, Replace, InsertSymbol, LtrRun, RtlRun,
MathZones, Select). The model's answers are written down first, each case run alone, so a refusal answers that case only.

Each case makes a workbook with a rectangle R (100, 50, 240 by 120) holding one of four texts through TextFrame2.TextRange.Text, with two
of them given a bold part to make runs, runs its statement on tr, its TextRange, and reads what the statement returned and the text after.
A paragraph end reads as <CR>, a line break as <VT>.

    python scripts/measure_text_range2_round.py predict
    python scripts/measure_text_range2_round.py measure

writes tests/fixtures/text_range2_round.json, which tests/test_excel_text_range2_round.py replays.
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


def cases() -> list[list[Any]]:
    every: list[list[Any]] = []
    for text, bold in TEXTS:
        every += [[text, bold, 'read', read] for read in READS]
        every += [[text, bold, 'do', statement] for statement in STATEMENTS]
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
    OUT.write_text(text, encoding='utf-8', newline='\n')
    print(len(every), 'cases; predictions sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())


def measure() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, Any] = json.loads(OUT.read_text(encoding='utf-8'))
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
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    wrong = sum(model != excel and model != 'refused' for model, excel in zip(record['model'], answers))
    print('model wrong', wrong, 'refused', record['model'].count('refused'), 'of', len(answers))


if __name__ == '__main__':
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
