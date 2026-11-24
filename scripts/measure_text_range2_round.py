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
#: Rounds 4 and 5, held out on the code after round 3's one miss, with what the statements write to the file.
TEXTS_4 = [('"Rain falls. sun shines! (yes) no" & vbCr & "end-of-day it\'s late"', (1, 4)), ('"  two  words  "', None),
           ('"tab" & vbTab & vbTab', None), ('"AbC dEf" & vbCr & "x." & vbCr & "y!"', (3, 3))]
STATEMENTS_4 = ['tr.ChangeCase msoCaseSentence', 'tr.ChangeCase msoCaseTitle', 'tr.ChangeCase msoCaseToggle', 'tr.Characters(4, 6).ChangeCase msoCaseTitle',
                'tr.Characters(2, 5).ChangeCase msoCaseUpper', 'v = tr.TrimText.Length', 'v = tr.Characters(2, 3).TrimText.Text', 'tr.AddPeriods',
                'tr.RemovePeriods', 'v = tr.Find("s").Start', 'v = tr.Find("S", MatchCase:=msoTrue).Start', 'v = tr.Find("zz", After:=2).Start',
                'v = tr.Find("zz", After:=50).Start', 'v = tr.Find("no", WholeWords:=msoTrue).Start', 'v = tr.Characters(3, 2).Find("a").Start',
                'v = tr.Replace("s", "$").Text', 'v = tr.Characters(1, 3).Replace("x", "y") Is Nothing', 'v = tr.Replace("t", "T", After:=4).Start',
                'v = tr.InsertSymbol("Arial", 128).Text', 'v = tr.InsertSymbol("Arial", 937, msoTrue).Start', 'v = tr.MathZones.Count',
                'tr.Select: v = TypeName(Selection)', 'tr.LtrRun: v = tr.ParagraphFormat.TextDirection']
TEXTS_5 = [('"one two. three four? five" & vbCr & "six-seven eight\'s nine"', (5, 3)), ('"  ab"', None), ('"Zed  " & vbTab', (1, 1)),
           ('"p." & vbCr & "Q" & vbCr & "r?"', None)]
ROUNDS = {1: (TEXTS, READS, STATEMENTS, OUT), 2: (TEXTS_2, READS_2, STATEMENTS_2, OUT.with_name('text_range2_round_2.json')),
          3: (TEXTS_3, READS_3, STATEMENTS_3, OUT.with_name('text_range2_round_3.json')),
          4: (TEXTS_4, READS_3, STATEMENTS_4, OUT.with_name('text_range2_round_4.json')),
          5: (TEXTS_5, READS_3, STATEMENTS_4, OUT.with_name('text_range2_round_5.json')),
          # Rounds 6 and 7, held out on the code after round 5's two misses.
          6: ([('"big dog. small CAT ran" & vbCr & "it is so"', (1, 3)), ('"  lone"', None), ('"ends with tab" & vbTab', (6, 4)),
               ('"Fine?" & vbCr & "maybe." & vbCr & "sure"', None)], READS_3, STATEMENTS_4, OUT.with_name('text_range2_round_6.json')),
          7: ([('"wet sand. dry rock! cold sea" & vbCr & "warm wind"', (6, 4)), ('"  xy  z"', None), ('"NASA team" & vbTab', None),
               ('"a." & vbCr & "bb?" & vbCr & "cc"', (2, 2))], READS_3, STATEMENTS_4, OUT.with_name('text_range2_round_7.json'))}
#: Round 8, a first measurement of the members the model lacks: For Each over a range and its collections, the bounds, which follow
#: the display's DPI, the clipboard methods, each Paste after a Copy or Cut of the probe's own text so nothing of the user's clipboard
#: is read, and InsertChartField on a shape's text.
STATEMENTS_8 = ['n = 0: For Each x In tr: n = n + 1: v = v & x.Text & "/": Next: v = n & ":" & v',
                'n = 0: For Each x In tr.Characters: n = n + 1: v = v & x.Text & "/": Next: v = n & ":" & v',
                'n = 0: For Each x In tr.Words: n = n + 1: v = v & x.Text & "/": Next: v = n & ":" & v',
                'n = 0: For Each x In tr.Runs: n = n + 1: v = v & x.Text & "/": Next: v = n & ":" & v',
                'n = 0: For Each x In tr.Paragraphs: n = n + 1: v = v & TypeName(x) & "/": Next: v = n & ":" & v',
                'v = tr.BoundLeft', 'v = tr.BoundTop', 'v = tr.BoundWidth', 'v = tr.BoundHeight', 'v = tr.Characters(1, 3).BoundWidth',
                'v = tr.Paragraphs(1).BoundHeight',
                'tr.Characters(1, 3).RotatedBounds x1, y1, x2, y2, x3, y3, x4, y4: v = x1 & "," & y1 & "," & x2 & "," & y2 & "," & x3 & "," & y3 & "," & x4 & "," & y4',
                'tr.Characters(1, 3).Copy: v = tr.Characters(5, 0).Paste.Text', 'tr.Characters(1, 3).Copy: v = tr.Characters(5, 2).Paste.Start',
                'tr.Copy: v = tr.Characters(2, 0).Paste.Length', 'tr.Characters(1, 4).Cut: v = tr.Characters(2, 0).Paste.Text',
                'tr.Characters(1, 4).Cut', 'tr.Characters(5, 3).Copy: v = tr.Characters(1, 0).PasteSpecial(msoClipboardFormatPlainText).Text',
                'tr.Characters(5, 3).Copy: v = tr.Characters(1, 0).PasteSpecial(msoClipboardFormatHTML).Text',
                'ws.Range("A1").Value = "cell": ws.Range("A1").Copy: v = tr.Characters(1, 0).Paste.Text',
                'v = tr.InsertChartField(msoChartFieldValue).Text', 'v = tr.InsertChartField(msoChartFieldFormula, "=Sheet1!A1").Text']
ROUNDS[8] = ([('"red fox. blue jay" & vbCr & "green owl"', (5, 3)), ('"solo"', None)], [], STATEMENTS_8,
             OUT.with_name('text_range2_round_8.json'))
#: Round 9, the text clipboard and the cells' together: where a paste of text lands and what it brings, PasteSpecial, Cut, what
#: CutCopyMode and ClipboardFormats say, and each way a copy of cells and a copy of text meet. Every case starts from a clipboard it
#: set itself, a cell copied and let go, so nothing of the user's clipboard, or an earlier case's, is read.
END = 'tr.Characters(tr.Length + 1, 0)'
STATEMENTS_9 = ['tr.Characters(1, 3).Copy: v = tr.Characters(1, 0).Paste.Text', 'tr.Characters(1, 3).Copy: v = tr.Characters(3, 0).Paste.Text',
                f'tr.Characters(1, 3).Copy: v = {END}.Paste.Text', 'tr.Characters(1, 3).Copy: v = tr.Characters(2, 2).Paste.Start',
                'tr.Characters(1, 3).Copy: v = tr.Paste.Text', f'tr.Copy: v = {END}.Paste.Length', f'tr.Characters(5, 3).Copy: v = {END}.Paste.Font.Bold',
                'tr.Characters(1, 3).Cut: v = tr.Characters(2, 0).Paste.Text', 'tr.Characters(2, 2).Cut: v = tr.Text',
                f'tr.Characters(1, 3).Copy: v = {END}.PasteSpecial(msoClipboardFormatPlainText).Text',
                f'tr.Characters(1, 3).Copy: v = {END}.PasteSpecial(msoClipboardFormatHTML).Text',
                f'tr.Characters(1, 3).Copy: v = {END}.PasteSpecial(msoClipboardFormatRTF).Text',
                'tr.Characters(1, 3).Copy: v = tr.Characters(1, 0).PasteSpecial(msoClipboardFormatPlainText).Text',
                f'v = {END}.Paste.Text', 'tr.Copy: v = Application.CutCopyMode', 'tr.Cut: v = Application.CutCopyMode',
                f'tr.Copy: Application.CutCopyMode = False: v = {END}.Paste.Text',
                'tr.Characters(1, 3).Copy: f = Application.ClipboardFormats: For k = LBound(f) To UBound(f): v = v & f(k) & ",": Next',
                f'ws.Range("A1").Value = "cell": ws.Range("A1").Copy: v = {END}.Paste.Text',
                'ws.Range("A1").Value = "cell": ws.Range("A1").Copy: v = tr.Characters(1, 0).Paste.Text',
                'ws.Range("A1").Value = "cell": ws.Range("A1").Copy: tr.Characters(1, 3).Copy: v = Application.CutCopyMode',
                'tr.Characters(1, 3).Copy: ws.Range("C3").Select: ws.Paste: v = ws.Range("C3").Formula & "|" & ws.Shapes.Count',
                'tr.Characters(1, 3).Copy: ws.Paste ws.Range("C3"): v = ws.Range("C3").Formula & "|" & ws.Shapes.Count',
                'tr.Characters(1, 3).Copy: ws.Range("C3").PasteSpecial xlPasteValues: v = ws.Range("C3").Formula',
                'tr.Characters(1, 3).Copy: ws.Range("C3").PasteSpecial xlPasteAll: v = ws.Range("C3").Formula & "|" & ws.Shapes.Count',
                'tr.Characters(1, 3).Copy: ws.Range("C3").Insert: v = ws.Range("C3").Formula & "|" & ws.Range("C4").Formula',
                f'tr.Characters(1, 3).Copy: ws.Range("A1").Value = "x": ws.Range("A1").Copy ws.Range("B1"): v = {END}.Paste.Text',
                'ws.Range("A1").Value = "cell": ws.Range("A1").Copy: tr.Characters(1, 3).Copy: ws.Range("C3").Select: ws.Paste: v = ws.Range("C3").Formula']
ROUNDS[9] = ([('"red fox. blue jay" & vbCr & "green owl"', (5, 3)), ('"solo"', None)], [], STATEMENTS_9,
             OUT.with_name('text_range2_round_9.json'))
#: What round 9 reads after each statement besides AFTER.
AFTER_9 = ['tr.Characters(1, 3).Font.Bold', 'Application.CutCopyMode']
#: Round 10, what round 9 left open: what a pasted run carries, the characters a copied cell brings, and what a sheet does with copied
#: text that reads as a number, a formula or more than one line. The pasted part is the range Paste answers, p.
STATEMENTS_10 = ['tr.Characters(1, 3).Copy: Set p = tr.Characters(3, 0).Paste: v = p.Font.Bold & "|" & p.Font.Name & "|" & p.Font.Size & "|" & p.Font.Italic',
                 'tr.Characters(1, 3).Copy: Set p = tr.Characters(tr.Length + 1, 0).Paste: v = p.Font.Bold & "|" & p.Font.Name & "|" & p.Font.Size',
                 'tr.Characters(1, 3).Copy: Set p = tr.Characters(tr.Length + 1, 0).PasteSpecial(msoClipboardFormatPlainText): v = p.Font.Bold & "|" & p.Start & "|" & p.Length',
                 'tr.Characters(1, 3).Font.Italic = msoTrue: tr.Characters(1, 3).Copy: Set p = tr.Characters(tr.Length + 1, 0).Paste: v = p.Font.Italic & "|" & p.Font.Bold',
                 'tr.Characters(2, 2).Font.Size = 20: tr.Characters(1, 3).Copy: Set p = tr.Characters(tr.Length + 1, 0).Paste: v = p.Font.Size & "|" & p.Runs.Count',
                 'ws.Range("A1").Value = "cell": ws.Range("A1").Copy: Set p = tr.Characters(tr.Length + 1, 0).Paste: For k = 1 To p.Length: v = v & AscW(Mid(p.Text, k, 1)) & ",": Next',
                 'ws.Range("A1").Value = "cell": ws.Range("A2").Value = 5: ws.Range("A1:A2").Copy: Set p = tr.Characters(tr.Length + 1, 0).Paste: For k = 1 To p.Length: v = v & AscW(Mid(p.Text, k, 1)) & ",": Next',
                 'ws.Range("A1").Value = "a": ws.Range("B1").Value = "b": ws.Range("A1:B1").Copy: Set p = tr.Characters(tr.Length + 1, 0).Paste: For k = 1 To p.Length: v = v & AscW(Mid(p.Text, k, 1)) & ",": Next',
                 'ws.Range("A1").Value = "cell": ws.Range("A1").Copy: Set p = tr.Characters(tr.Length + 1, 0).Paste: v = p.Font.Bold & "|" & p.Font.Name & "|" & p.Font.Size',
                 'tr.Text = "12": tr.Copy: ws.Paste ws.Range("C3"): v = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Formula & "|" & ws.Range("C4").Formula',
                 'tr.Text = "=1+1": tr.Characters(1, 4).Copy: ws.Paste ws.Range("C3"): v = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Formula',
                 'tr.Text = "007": tr.Characters(1, 3).Copy: ws.Paste ws.Range("C3"): v = TypeName(ws.Range("C3").Value) & ":" & ws.Range("C3").Formula',
                 'tr.Characters(1, 3).Copy: ws.Range("C3").Value = "old": ws.Paste ws.Range("C3:D4"): v = ws.Range("C3").Formula & "|" & ws.Range("D4").Formula & "|" & ws.Range("C3").Font.Bold',
                 'tr.Copy: ws.Paste ws.Range("C3"): v = ws.Range("C3").Formula & "|" & ws.Range("C4").Formula & "|" & ws.Range("C5").Formula',
                 'tr.Characters(5, 3).Copy: ws.Paste ws.Range("C3"): v = ws.Range("C3").Formula & "|" & ws.Range("C3").Font.Bold & "|" & ws.Range("C3").Font.Name',
                 'tr.Characters(1, 3).Copy: ws.Range("C3").Select: ws.Paste: v = TypeName(Selection) & ":" & Selection.Address',
                 'tr.Characters(1, 3).Copy: ws.Paste ws.Range("C3"): v = Application.CutCopyMode & "|" & ws.Range("C3").Formula: Set p = tr.Characters(tr.Length + 1, 0).Paste: v = v & "|" & p.Text',
                 'tr.Characters(1, 3).Copy: ws.Range("C3").Value = 1: ws.Range("C3").Insert: v = ws.Range("C3").Formula & "|" & ws.Range("C4").Formula',
                 'tr.Characters(1, 3).Copy: ws.Range("C3").Value = 1: ws.Range("C3").PasteSpecial xlPasteValues: v = ws.Range("C3").Formula']
ROUNDS[10] = ([('"red fox. blue jay" & vbCr & "green owl"', (5, 3)), ('"solo"', None)], [], STATEMENTS_10,
              OUT.with_name('text_range2_round_10.json'))
#: Round 11, held out on the text clipboard rounds 9 and 10 built: new texts and places, and round 8's paste just before a bold run,
#: which answered nothing there with no clipboard set first.
STATEMENTS_11 = ['tr.Characters(1, 2).Copy: Set p = tr.Characters(4, 0).Paste: v = p.Text & "|" & p.Start & "|" & p.Font.Bold',
                 'tr.Characters(2, 2).Copy: Set p = tr.Characters(5, 0).Paste: v = p.Text & "|" & p.Start',
                 'tr.Characters(4, 2).Copy: Set p = tr.Characters(2, 0).Paste: v = p.Font.Bold & "|" & p.Length',
                 'tr.Characters(1, 1).Copy: Set p = tr.Characters(2, 3).Paste: v = p.Text & "|" & p.Start',
                 'tr.Characters(3, 2).Cut: Set p = tr.Characters(tr.Length + 1, 0).Paste: v = p.Text & "|" & p.Start',
                 'tr.Characters(2, 2).Copy: Set p = tr.Characters(3, 0).PasteSpecial(msoClipboardFormatPlainText): v = p.Text & "|" & p.Font.Bold',
                 'tr.Characters(1, 2).Copy: Application.CutCopyMode = False: v = tr.Characters(tr.Length + 1, 0).Paste.Text',
                 'tr.Characters(1, 2).Copy: ws.Range("A1").Value = 1: ws.Range("A1").Copy ws.Range("B1"): v = tr.Characters(2, 0).Paste.Text',
                 'tr.Characters(1, 2).Copy: f = Application.ClipboardFormats: v = f(1) & "," & f(UBound(f)) & "," & UBound(f)',
                 'tr.Characters(1, 2).Cut: v = Application.CutCopyMode & "|" & tr.Text',
                 'tr.Characters(1, 2).Copy: ws.Paste ws.Range("B2"): v = ws.Range("B2").Formula & "|" & ws.Range("B2").Font.Bold',
                 'tr.Characters(4, 2).Copy: ws.Range("B2").Select: ws.Paste: v = ws.Range("B2").Formula & "|" & ws.Range("B2").Font.Bold',
                 'tr.Characters(1, 2).Copy: Set p = tr.Characters(3, 0).Paste: v = tr.Characters(1, 4).Font.Bold',
                 'ws.Range("A1").Value = 1: ws.Range("A1").Copy: tr.Characters(1, 2).Copy: v = Application.CutCopyMode']
ROUNDS[11] = ([('"ab cdef gh" & vbCr & "ij kl"', (4, 2)), ('"wxyz"', (1, 1)), ('"red fox. blue jay" & vbCr & "green owl"', (5, 3))], [],
              STATEMENTS_11, OUT.with_name('text_range2_round_11.json'))
#: Round 12, held out again on the build after round 11's two misses, on new texts; round 11's third text met a clipboard held by
#: another program, which round 12's measure runs again.
ROUNDS[12] = ([('"pq rs tuv" & vbCr & "wx"', (4, 2)), ('"mnop qr"', (6, 2)), ('"big cat. tiny ant" & vbCr & "slow"', (1, 3))], [],
              STATEMENTS_11, OUT.with_name('text_range2_round_12.json'))
#: Round 13, what round 11's case 16 left open: Characters running past the text's end, read, copied and cut, late and through a
#: variable typed TextRange2 (t).
_PAST_END = ['v = {r}.Characters(4, 2).Text', 'v = {r}.Characters(4, 2).Length', 'v = {r}.Characters(4, 2).Start',
             'v = {r}.Characters(5, 1).Text', 'v = {r}.Characters(5, 1).Start', 'v = {r}.Characters(6, 1).Length',
             'v = {r}.Characters(3, 10).Text', 'v = {r}.Characters(0, 2).Text',
             '{r}.Characters(4, 2).Copy: v = {r}.Characters({r}.Length + 1, 0).Paste.Text',
             '{r}.Characters(3, 10).Copy: v = {r}.Characters(2, 0).Paste.Text',
             '{r}.Characters(5, 1).Copy: v = {r}.Characters(2, 0).Paste.Text', '{r}.Characters(4, 2).Cut: v = {r}.Text']
STATEMENTS_13 = [*(one.format(r='tr') for one in _PAST_END), *('Set t = tr: ' + one.format(r='t') for one in _PAST_END)]
ROUNDS[13] = ([('"wxyz"', None), ('"ab cd"', (1, 2))], [], STATEMENTS_13, OUT.with_name('text_range2_round_13.json'))
#: Round 11's markup.
MARKUP_11 = ['tr.Characters(2, 2).Copy: tr.Characters(6, 0).Paste', 'tr.Characters(6, 2).Copy: tr.Characters(2, 0).Paste',
             'tr.Characters(1, 3).Copy: tr.Characters(5, 3).Paste', 'tr.Characters(5, 2).Copy: tr.Characters(tr.Length + 1, 0).Paste',
             'tr.Characters(2, 3).Copy: tr.Characters(3, 0).PasteSpecial msoClipboardFormatPlainText', 'tr.Characters(1, 2).Cut']
#: Round 10's markup: what a paste writes to the file.
MARKUP_10 = ['tr.Characters(1, 3).Copy: tr.Characters(5, 0).Paste', 'tr.Characters(5, 3).Copy: tr.Characters(1, 3).Paste',
             'tr.Characters(1, 3).Copy: tr.Characters(tr.Length + 1, 0).Paste',
             'tr.Characters(1, 3).Copy: tr.Characters(tr.Length + 1, 0).PasteSpecial msoClipboardFormatPlainText',
             'tr.Characters(1, 3).Copy: tr.Characters(tr.Length + 1, 0).PasteSpecial msoClipboardFormatHTML', 'tr.Characters(5, 3).Cut',
             'ws.Range("A1").Value = "cell": ws.Range("A1").Copy: tr.Characters(tr.Length + 1, 0).Paste', 'tr.Copy: tr.Characters(4, 0).Paste']
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
    if ROUND >= 9:
        # A clipboard the case set itself: a cell copied, then let go.
        lines += ['    ws.Range("Z1").Value = "z": ws.Range("Z1").Copy: Application.CutCopyMode = False', '    ws.Range("Z1").ClearContents']
    lines += ['    s = ""', '    Err.Clear: v = Empty']
    if how == 'read':
        lines += [f'    v = {what}', '    s = Show(v)']
    else:
        lines += [f'    {what}', '    s = "e" & Err.Number & "=" & Show(v)']
        for read in AFTER + (AFTER_9 if ROUND >= 9 else []):
            lines += ['    Err.Clear: v = Empty', f'    v = {read}', '    s = s & "|" & Show(v)']
    lines += ['    Set tr = Nothing: Set sh = Nothing', '    wb.Close SaveChanges:=False', f'    out = out & "{SEPARATOR}" & s']
    return '\n'.join(lines) + '\n'


HELPERS = '''Private Function Show(ByVal v As Variant) As String
    If Err.Number Then Show = "e" & Err.Number: Exit Function
    Show = TypeName(v) & ":" & Replace(Replace(Replace(CStr(v), vbCr, "<CR>"), vbLf, "<LF>"), Chr(11), "<VT>")
End Function
'''
DECLARE = ('    Dim wb As Workbook, ws As Worksheet, sh As Object, tr As Object, v As Variant, s As String, x As Object, n As Long\n'
           '    Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single, x3 As Single, y3 As Single, x4 As Single, y4 As Single\n'
           '    Dim f As Variant, k As Long, p As Object\n'
           '    On Error Resume Next\n')
#: From round 13, a variable typed TextRange2, so early bound calls are measured too.
TYPED = '    Dim t As TextRange2\n'
PART = 12


def macro(chosen: list[list[Any]]) -> str:
    parts = [chosen[at:at + PART] for at in range(0, len(chosen), PART)]
    procedures = ''.join(f'Private Sub Part{index}(ByRef out As String)\n{DECLARE}{TYPED if ROUND >= 13 else ""}{"".join(_case(case) for case in part)}End Sub\n'
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


#: From round 4, what a statement writes to the file: each saved with a text of two runs, and the shape's txBody read back.
MARKUP = ['', 'tr.RtlRun', 'tr.LtrRun', 'tr.Paragraphs(1).RtlRun', 'tr.ChangeCase msoCaseUpper', 'tr.AddPeriods',
          'tr.Characters(5, 3).InsertSymbol "Arial", 66', 'tr.Replace "def", "xyz"',
          # From round 5.
          'tr.InsertSymbol "Arial", 66', 'tr.Characters(2, 2).InsertSymbol "Arial", 66', 'tr.ChangeCase msoCaseTitle',
          'tr.InsertAfter ".": tr.RemovePeriods']


def markups() -> list[str]:
    """The statements whose markup the round records: rounds 10 and 11 have their own."""
    return MARKUP_11 if ROUND >= 11 else MARKUP_10 if ROUND >= 10 else MARKUP


def markup_macro() -> str:
    lines = ['Function Marks(ByVal folder As String) As String', '    Dim wb As Workbook, ws As Worksheet, sh As Object, tr As Object',
             '    Application.DisplayAlerts = False']
    for index, statement in enumerate(markups()):
        lines += ['    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
                  '    Set sh = ws.Shapes.AddShape(msoShapeRectangle, 100, 50, 240, 120)', '    Set tr = sh.TextFrame2.TextRange',
                  '    tr.Text = "abc def"', '    tr.Characters(1, 3).Font.Bold = msoTrue', f'    {statement}',
                  f'    wb.SaveAs folder & Application.PathSeparator & "m{index}.xlsx", xlOpenXMLWorkbook', '    wb.Close SaveChanges:=False']
    return '\n'.join([*lines, '    Application.DisplayAlerts = True', 'End Function', ''])


def text_bodies(folder: Path) -> list[list[str]]:
    import re
    import zipfile

    bodies: list[list[str]] = []
    for index in range(len(markups())):
        with zipfile.ZipFile(folder / f'm{index}.xlsx') as package:
            bodies.append(re.findall(r'<xdr:txBody>.*?</xdr:txBody>', package.read('xl/drawings/drawing1.xml').decode('utf-8'), re.S))
    return bodies


def model_markup_each(folder: Path) -> list[list[str] | str]:
    """The model's markup for each of the round's statements run by itself: its txBodies, or "refused" where it refuses."""
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    global MARKUP, MARKUP_10, MARKUP_11
    kept = (MARKUP, MARKUP_10, MARKUP_11)
    found: list[list[str] | str] = []
    try:
        for statement in list(markups()):
            MARKUP = MARKUP_10 = MARKUP_11 = [statement]
            app = ExcelApplication(sandbox=False)
            app.add_workbook()
            app.add_module(markup_macro(), name='M')
            try:
                app.run('M.Marks', str(folder))
            except VBAUnsupportedError:
                found.append('refused')
                continue
            found.append(text_bodies(folder)[0])
    finally:
        MARKUP, MARKUP_10, MARKUP_11 = kept
    return found


def model_markup(folder: Path) -> list[list[str]] | str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(markup_macro(), name='M')
    try:
        app.run('M.Marks', str(folder))
    except VBAUnsupportedError as refused:
        return f'refused: {refused}'
    return text_bodies(folder)


def predict() -> None:
    import tempfile

    sys.path.insert(0, str(ROOT / 'src'))
    every = cases()
    record: dict[str, Any] = {'cases': every, 'model': [alone(case) for case in every]}
    if ROUND >= 4:
        with tempfile.TemporaryDirectory() as folder:
            # From round 11, each statement alone, so one refused leaves the others' markup to compare.
            record['model_markup'] = model_markup_each(Path(folder)) if ROUND >= 11 else model_markup(Path(folder))
    text = json.dumps(record, indent=1) + '\n'
    ROUNDS[ROUND][3].write_text(text, encoding='utf-8', newline='\n')
    print(len(every), 'cases; predictions sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())


def measure() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, Any] = json.loads(ROUNDS[ROUND][3].read_text(encoding='utf-8'))
    every: list[list[Any]] = record['cases']
    answers: list[str] = []
    height = 'Function H() As String\n    H = CStr(ActiveSheet.StandardHeight)\nEnd Function\n'
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        # From round 8, the display's scale at both ends of the run, which the bounds follow (StandardHeight 15 is 96 DPI).
        heights = [str(excel.run_vba(height, 'H', line_numbers=False).value)] if ROUND >= 8 else []
        # From round 12, a text's cases at a time, run again, up to three times, while Excel could not open the clipboard
        # (CLIPBRD_E_CANT_OPEN, -2147221040): another program held it, as Remote Desktop's clipboard does after each copy.
        step = len(ROUNDS[ROUND][2]) if ROUND >= 12 else 60
        tries: list[int] = []
        for at in range(0, len(every), step):
            for attempt in range(1, 4):
                result = excel.run_vba(macro(every[at:at + step]), 'Probe', (), timeout=600, line_numbers=False)
                assert result.ok, (at, result.error, [one.message for one in result.dialogs])
                got = str(result.value).split(SEPARATOR)
                if ROUND < 12 or not any('-2147221040' in one for one in got):
                    break
            tries.append(attempt)
            answers += got
            print(at + step, 'cases', flush=True)
        if ROUND >= 12:
            record['tries'] = tries
        if ROUND >= 4:
            import tempfile

            with tempfile.TemporaryDirectory() as folder:
                result = excel.run_vba(markup_macro(), 'Marks', (folder,), timeout=300, line_numbers=False)
                assert result.ok, (result.error, [one.message for one in result.dialogs])
                record['excel_markup'] = text_bodies(Path(folder))
            print('markup same as the model:', record['excel_markup'] == record['model_markup'])
        if heights:
            heights.append(str(excel.run_vba(height, 'H', line_numbers=False).value))
            record['standard_height'] = heights
            print('StandardHeight', heights, flush=True)
    record.update({'measured_at': datetime.now(timezone.utc).isoformat(), 'excel': answers})
    ROUNDS[ROUND][3].write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    wrong = sum(model != excel and model != 'refused' for model, excel in zip(record['model'], answers))
    print('model wrong', wrong, 'refused', record['model'].count('refused'), 'of', len(answers))


if __name__ == '__main__':
    ROUND = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
