"""Characters(start, length) and TextFrame2.TextRange.Characters on a shape's text: what part of the text each covers, its
Font, the errors, and the text body each saved shape has.

    python scripts/measure_shape_runs.py

writes tests/fixtures/shape_runs.json, which tests/test_excel_shape_runs.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_runs.json'
RECT = 'ws.Shapes.AddShape(1, 10, 10, 100, 40)'
TEXT = 'ws.Shapes.AddTextbox(1, 10, 10, 100, 40)'
SET = 'sh.TextFrame.Characters.Text = "Hello World": '
TWO = 'sh.TextFrame.Characters.Text = "ab" & vbLf & "cd": '
C = 'sh.TextFrame.Characters'
R = 'sh.TextFrame2.TextRange.Characters'
CASES = [
    (RECT, SET.rstrip(': ')),
    *[(RECT, f'{SET}s2 = {C}({args}).Text') for args in ('1, 5', '7, 5', '7', '1', '2, 0', '12, 1', '20, 1', '0, 1', '-1, 1', '1, 100', '1, -1')],
    (RECT, f'{SET}s2 = {C}(7, 5).Count'), (RECT, f'{SET}s2 = {C}(1, 100).Count'),
    (RECT, f'{SET}s2 = {C}(, 5).Text'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True'), (RECT, f'{SET}{C}(7, 5).Font.Bold = True'), (RECT, f'{SET}{C}(3, 2).Font.Bold = True'),
    (RECT, f'{SET}{C}(1, 11).Font.Bold = True'), (RECT, f'{SET}{C}(1, 100).Font.Bold = True'), (RECT, f'{SET}{C}(7).Font.Bold = True'),
    (RECT, f'{SET}{C}(20, 1).Font.Bold = True'), (RECT, f'{SET}{C}(0, 1).Font.Bold = True'), (RECT, f'{SET}{C}(2, 0).Font.Bold = True'),
    (RECT, f'{SET}{C}(1, 5).Font.Size = 14'), (RECT, f'{SET}{C}(1, 5).Font.Italic = True'), (RECT, f'{SET}{C}(1, 5).Font.Color = 255'),
    (RECT, f'{SET}{C}(1, 5).Font.Name = "Arial"'), (RECT, f'{SET}{C}(1, 5).Font.Underline = xlUnderlineStyleSingle'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(7, 5).Font.Italic = True'),
    (RECT, f'{SET}{C}(7, 5).Font.Italic = True: {C}(1, 5).Font.Bold = True'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(3, 5).Font.Italic = True'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(6, 1).Font.Bold = True'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(1, 5).Font.Bold = False'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}.Font.Size = 14'), (RECT, f'{SET}{C}.Font.Size = 14: {C}(1, 5).Font.Bold = True'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}.Font.Bold = False'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}.Text = "New"'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(1, 5).Text = "Bye"'),
    (RECT, f'{SET}{C}(1, 5).Text = "Bye"'), (RECT, f'{SET}{C}(7, 5).Text = ""'), (RECT, f'{SET}{C}(3, 0).Text = "XY"'),
    (RECT, f'{SET}{C}(1, 5).Delete'), (RECT, f'{SET}{C}(1, 5).Insert "Bye"'), (RECT, f'{SET}{C}(6, 0).Insert "!"'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(1, 5).Delete'),
    (RECT, f'{TWO}{C}(1, 1).Font.Bold = True'), (RECT, f'{TWO}{C}(4, 1).Font.Bold = True'), (RECT, f'{TWO}{C}(2, 3).Font.Bold = True'),
    (RECT, f'{TWO}{C}(3, 1).Font.Bold = True'), (RECT, f'{TWO}s2 = {C}(3, 1).Text & "|" & {C}(4, 2).Text'),
    (TEXT, f'{SET}{C}(1, 5).Font.Bold = True'), (TEXT, f'{SET}{C}(7, 5).Font.Color = 255'),
    (RECT, f'{SET}s2 = {R}(7, 5).Text'), (RECT, f'{SET}{R}(1, 5).Font.Bold = msoTrue'), (RECT, f'{SET}{R}(7, 5).Font.Size = 14'),
    (RECT, f'{SET}{R}(1, 5).Font.Bold = msoTrue: {R}(7, 5).Font.Italic = msoTrue'), (RECT, f'{SET}{R}(1, 5).Text = "Bye"'),
    (RECT, f'{SET}{R}(3, 2).Font.UnderlineStyle = msoUnderlineSingleLine'), (RECT, f'{SET}{R}(1, 5).Font.Fill.ForeColor.RGB = 255'),
    (RECT, f'{SET}s2 = sh.TextFrame2.TextRange.Runs.Count'), (RECT, f'{SET}{C}(1, 5).Font.Bold = True: s2 = sh.TextFrame2.TextRange.Runs.Count'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: s2 = sh.TextFrame2.TextRange.Runs(1).Text & "|" & sh.TextFrame2.TextRange.Runs(2).Text'),
    # What a replaced text takes from the runs before it, and the empty paragraph mark a whole replacement can leave.
    (RECT, f'{C}.Text = "A": {C}.Text = "New"'), (RECT, f'{SET}{C}.Text = "New"'), (RECT, f'{SET}{C}(7, 5).Font.Bold = True: {C}.Text = "New"'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(7, 5).Font.Italic = True: {C}.Text = "New"'), (RECT, f'{SET}{C}(1, 11).Font.Bold = True: {C}.Text = "New"'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(1, 11).Text = "New"'), (RECT, f'{SET}{C}(7, 5).Font.Bold = True: {C}(1, 11).Text = "New"'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}.Text = "New": {C}.Text = "Again"'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(3, 5).Text = "XY"'), (RECT, f'{SET}{C}(3, 2).Font.Bold = True: {C}(1, 5).Text = "XY"'),
    (RECT, f'{SET}{C}(3, 2).Font.Bold = True: {C}(3, 2).Text = ""'), (RECT, f'{SET}{C}(3, 2).Font.Bold = True: {C}(4, 1).Insert "XY"'),
    (RECT, f'{SET}{C}(3, 2).Font.Bold = True: {C}(1, 2).Insert "XY"'), (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(6, 1).Text = "XY"'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(1, 5).Text = ""'), (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(1, 11).Delete'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}.Delete'), (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}.Text = ""'),
    (RECT, f'{SET}{C}.Delete'), (RECT, f'{SET}{C}(1, 5).Font.Bold = True: sh.TextFrame2.DeleteText'),
    (RECT, f'{SET}{C}.Font.Bold = True: {C}.Text = "New"'), (RECT, f'{SET}{C}.Font.Bold = True: {C}(1, 5).Text = "Bye"'),
    # No text at all.
    (RECT, f's2 = {C}(1, 1).Text'), (RECT, f'{C}(1, 1).Font.Bold = True'), (RECT, f's2 = {R}(1, 1).Text'), (RECT, f'{R}(1, 1).Font.Bold = msoTrue'), (RECT, f's2 = {C}(1, 1).Count'),
    # Whole settings over part settings, and part settings going back to where they began.
    (RECT, f'{SET}{C}.Font.Color = 255: {C}(1, 5).Font.Color = 65280'), (RECT, f'{SET}{C}(1, 5).Font.Color = 255: {C}.Font.Color = 65280'),
    (RECT, f'{SET}{C}(1, 5).Font.Size = 14: {C}(1, 2).Font.Size = 11'), (RECT, f'{SET}{C}.Font.Bold = True: {C}(1, 5).Font.Bold = False'),
    # Each other setting on part of the text, for what a mixed range reads.
    *[(RECT, f'{SET}{C}(1, 5).Font.{one}') for one in (
        'Strikethrough = True', 'Superscript = True', 'Subscript = True', 'FontStyle = "Bold"', 'ColorIndex = 3', 'TintAndShade = 0.5',
        'Shadow = True', 'Underline = xlUnderlineStyleDouble', 'Size = 14.5', 'Name = ""')],
    *[(RECT, f'{SET}{R}(1, 5).Font.{one}') for one in (
        'Italic = msoTrue', 'Size = 14', 'Name = "Arial"', 'Strike = msoSingleStrike', 'StrikeThrough = msoTrue', 'DoubleStrikeThrough = msoTrue', 'Superscript = msoTrue',
        'Subscript = msoTrue', 'BaselineOffset = 0.3', 'Caps = msoAllCaps', 'Allcaps = msoTrue', 'Smallcaps = msoTrue', 'Spacing = 2', 'Kerning = 12', 'Equalize = msoTrue')],
    # Line breaks.
    *[(RECT, f'{one}: s2 = {C}.Count & "|" & sh.TextFrame2.TextRange.Paragraphs.Count & "|" & sh.TextFrame2.TextRange.Paragraphs(1).Text') for one in (
        f'{C}.Text = "ab" & vbCr & "cd"', f'{C}.Text = "ab" & vbCrLf & "cd"', f'{C}.Text = "ab" & vbLf & "cd"',
        'sh.TextFrame2.TextRange.Text = "ab" & vbCr & "cd"', 'sh.TextFrame2.TextRange.Text = "ab" & vbLf & "cd"', 'sh.TextFrame2.TextRange.Text = "ab" & vbCrLf & "cd"', f'{C}.Text = "ab" & vbCr'
    )],
    # TextRange2's own ranges.
    *[(RECT, f'{SET}s2 = {R}({args}).Text') for args in ('0, 1', '12, 1', '20, 1', '1, 100', '1, 0', '7', '1, -1', ', 5', '12, 0', '11, 5', '-1, 1', '3, 0')],
    (RECT, f'{SET}s2 = {R}(7, 5).Start & "|" & {R}(7, 5).Length & "|" & {R}(7, 5).Count'),
    (RECT, f'{SET}s2 = sh.TextFrame2.TextRange.Start & "|" & sh.TextFrame2.TextRange.Length & "|" & sh.TextFrame2.TextRange.Count'),
    (RECT, f'{SET}s2 = {R}.Text & "|" & {R}.Count & "|" & {R}.Start & "|" & {R}.Length'),
    (RECT, f'{SET}s2 = {R}(1, 5).Characters(2, 2).Text & "|" & {R}(7, 5).Characters(2, 2).Start'), (RECT, f'{SET}s2 = {R}(7, 5).Characters(9, 1).Text'),
    (RECT, f'{SET}{R}(7, 5).Characters(2, 2).Font.Bold = msoTrue'),
    (RECT, f'{SET}s2 = sh.TextFrame2.TextRange.Item(1).Text'), (RECT, f'{SET}s2 = sh.TextFrame2.TextRange(1).Text'), (RECT, f'{SET}s2 = sh.TextFrame2.TextRange.Item(2).Text'),
    (RECT, f'{SET}s2 = {R}(1, 5).Count & "|" & {R}(1, 5).Item(2).Text & "|" & {R}(1, 5)(1).Text'),
    (RECT, f'{SET}s2 = TypeName({C}(1, 5)) & "|" & TypeName({C}(1, 5).Parent) & "|" & TypeName({C}(1, 5).Font.Parent) & "|" & TypeName({R}(1, 5)) & "|" & TypeName({R}(1, 5).Parent) & "|" & TypeName({R}(1, 5).Font.Parent)'),
    *[(RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(7, 5).Font.Italic = True: s2 = sh.TextFrame2.TextRange.Runs{tail}') for tail in (
        '.Count', '.Text', '(1).Text', '(2).Text', '(3).Text', '(4).Text', '(0).Text', '(1, 2).Text', '(2, 2).Text', '(3, 5).Text', '(2).Start', '(2).Length', '(2).Count',
        '(3).Font.Italic', '.Item(3).Text', '().Text', '(1).Runs.Count')],
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: s2 = {R}(3, 5).Runs.Count & "|" & {R}(3, 5).Runs(2).Text & "|" & {R}(3, 5).Runs(2).Start'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: sh.TextFrame2.TextRange.Runs(2).Font.Italic = msoTrue'),
    (RECT, f's2 = sh.TextFrame2.TextRange.Runs.Count'), (RECT, f's2 = sh.TextFrame2.TextRange.Runs(1).Text'),
    (RECT, f'{TWO}{C}(1, 1).Font.Bold = True: s2 = sh.TextFrame2.TextRange.Runs.Count'),
    # Paragraphs and words.
    *[(RECT, f'{SET}s2 = sh.TextFrame2.TextRange.{tail}') for tail in (
        'Paragraphs.Count', 'Paragraphs(1).Text', 'Paragraphs(2).Text', 'Words.Count', 'Words(1).Text', 'Words(2).Text', 'Words(3).Text', 'Words(1, 2).Text', 'Words(0).Text', 'Words.Text')],
    (RECT, f'{SET}sh.TextFrame2.TextRange.Words(2).Font.Bold = msoTrue'), (RECT, f'{SET}sh.TextFrame2.TextRange.Words(1).Font.Bold = msoTrue'),
    (RECT, f'{SET}sh.TextFrame2.TextRange.Paragraphs(1).Font.Bold = msoTrue'),
    (RECT, f'{C}.Text = "ab cd" & vbCr & "ef": s2 = sh.TextFrame2.TextRange.Paragraphs.Count & "|" & sh.TextFrame2.TextRange.Paragraphs(1).Text & "|" & sh.TextFrame2.TextRange.Paragraphs(2).Text & "|" & sh.TextFrame2.TextRange.Paragraphs(2).Start'),
    (RECT, f'{C}.Text = "ab cd" & vbCr & "ef": sh.TextFrame2.TextRange.Paragraphs(2).Font.Bold = msoTrue'),
    (RECT, f'{C}.Text = "ab cd" & vbCr & "ef": sh.TextFrame2.TextRange.Paragraphs(1).Font.Bold = msoTrue'),
    (RECT, f'{C}.Text = "ab cd" & vbCr & "ef": {C}(4, 3).Font.Bold = True'),
    (RECT, f'{C}.Text = "ab cd" & vbLf & "ef": s2 = sh.TextFrame2.TextRange.Words.Count & "|" & sh.TextFrame2.TextRange.Words(2).Text & "|" & sh.TextFrame2.TextRange.Words(3).Text'),
    (RECT, f'{C}.Text = "a, b.  c": s2 = sh.TextFrame2.TextRange.Words.Count & "|" & sh.TextFrame2.TextRange.Words(1).Text & "|" & sh.TextFrame2.TextRange.Words(2).Text & "|" & sh.TextFrame2.TextRange.Words(3).Text'),
    # Adding and taking away.
    *[(RECT, f'{SET}s2 = sh.TextFrame2.TextRange.{tail}') for tail in ('InsertAfter("X").Text', 'InsertBefore("X").Text', 'InsertAfter("X").Start', 'InsertBefore("X").Start')],
    *[(RECT, f'{SET}{C}(1, 5).Font.Bold = True: {R}({args}).{tail}') for args, tail in (
        ('1, 5', 'InsertAfter "X"'), ('1, 5', 'InsertBefore "X"'), ('7, 5', 'InsertAfter "X"'), ('7, 5', 'InsertBefore "X"'), ('3, 2', 'InsertAfter "X"'), ('3, 2', 'InsertBefore "X"'),
        ('1, 5', 'Delete'), ('3, 5', 'Delete'), ('7, 5', 'Delete'), ('1, 11', 'Delete'))],
    (RECT, f'{SET}sh.TextFrame2.TextRange.Delete'), (RECT, f'sh.TextFrame2.TextRange.InsertAfter "X"'), (RECT, f'sh.TextFrame2.TextRange.InsertBefore "X"'),
    # A range kept while the text changes.
    (RECT, f'{SET}Set o = {C}(7, 5): {C}.Text = "A new longer text": s2 = o.Text'), (RECT, f'{SET}Set o = {R}(7, 5): {C}.Text = "A new longer text": s2 = o.Text'),
    (RECT, f'{SET}Set o = {C}(7, 5): {C}.Text = "Hi": s2 = o.Text'), (RECT, f'{SET}Set o = {R}(7, 5): {C}.Text = "Hi": s2 = o.Text'),
    (RECT, f'{SET}Set o = {R}(7, 5): o.Font.Bold = msoTrue: {C}(1, 1).Text = "XYZ": s2 = o.Text: o.Font.Italic = msoTrue'),
    # Emptying a text whose last run has a setting; one run made of two settings; a replacement with paragraphs.
    (RECT, f'{SET}{C}(7, 5).Font.Bold = True: {C}.Text = ""'), (RECT, f'{SET}{C}(7, 5).Font.Bold = True: {C}.Delete'),
    (RECT, f'{SET}{C}(7, 5).Font.Bold = True: sh.TextFrame2.TextRange.Delete'), (RECT, f'{SET}{C}(7, 5).Font.Bold = True: sh.TextFrame2.TextRange.Text = ""'),
    (RECT, f'{SET}{C}(7, 5).Font.Bold = True: {C}.Text = "": {C}.Text = "New"'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(6, 6).Font.Bold = True: {C}.Text = "New"'),
    (RECT, f'{C}.Text = "ab" & vbCr & "cd": {C}(1, 1).Font.Bold = True: {C}.Text = "New"'), (RECT, f'{C}.Text = "ab" & vbCr & "cd": {C}(4, 1).Font.Bold = True: {C}.Text = "New"'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}.Text = "x" & vbCr & "y"'), (RECT, f'{SET}{C}(7, 5).Font.Bold = True: {C}.Text = "x" & vbCr & "y"'),
    (RECT, f'{SET}{C}(7, 5).Font.Bold = True: {C}.Text = "x" & vbLf & "y"'),
    # A range's edge on a paragraph break, and breaks in a partial replacement.
    (RECT, f'{C}.Text = "ab cd" & vbCr & "ef": {C}(6, 1).Font.Bold = True: s2 = sh.TextFrame2.TextRange.Runs.Count'),
    (RECT, f'{C}.Text = "ab cd" & vbCr & "ef": {C}(5, 3).Font.Bold = True: s2 = sh.TextFrame2.TextRange.Runs.Count'),
    (RECT, f'{SET}{C}(1, 5).Text = "a" & vbCr & "b"'), (RECT, f'{SET}{C}(1, 5).Text = "a" & vbLf & "b"'), (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {C}(3, 1).Text = "a" & vbCr & "b"'),
    # The empty paragraph mark a replacement left, under later edits.
    (RECT, f'{SET}{C}(7, 5).Font.Bold = True: {C}.Text = "New": {C}(1, 1).Text = "XY"'), (RECT, f'{SET}{C}(7, 5).Font.Bold = True: {C}.Text = "New": {C}(2, 1).Font.Italic = True'),
    (RECT, f'{SET}{C}(7, 5).Font.Bold = True: {C}.Text = "New": {C}.Font.Size = 14'), (RECT, f'{SET}{C}(7, 5).Font.Bold = True: {C}.Text = "New": {C}.Font.Bold = False'),
    (RECT, f'{SET}{C}(7, 5).Font.Bold = True: {C}.Text = "New": {C}.Text = "Again"'),
    # What a mixed range reads when its first character has no setting.
    *[(RECT, f'{SET}{C}(7, 5).Font.{one}') for one in (
        'Underline = xlUnderlineStyleSingle', 'Strikethrough = True', 'Superscript = True', 'Subscript = True', 'Shadow = True', 'Size = 14', 'Color = 255', 'Name = "Arial"')],
    *[(RECT, f'{SET}{R}(7, 5).Font.{one}') for one in (
        'Strike = msoSingleStrike', 'DoubleStrikeThrough = msoTrue', 'Superscript = msoTrue', 'Subscript = msoTrue', 'BaselineOffset = 0.3', 'Caps = msoAllCaps', 'Smallcaps = msoTrue',
        'Spacing = 2', 'Kerning = 12', 'Equalize = msoTrue', 'UnderlineStyle = msoUnderlineSingleLine')],
    # More of TextRange2's ranges, the empty one above all.
    *[(RECT, f'{SET}s2 = {R}({args}).Text') for args in ('-1, 5', '1, -5', '3, -1', '0, 0', '0, 5', '-5, 0', '13, 0')],
    (RECT, f'{SET}s2 = {R}(3, 0).Start & "|" & {R}(3, 0).Length & "|" & {R}(3, 0).Count'),
    (RECT, f'{SET}{R}(3, 0).Font.Bold = msoTrue'), (RECT, f'{SET}s2 = {R}(3, 0).Font.Bold'), (RECT, f'{SET}{C}(1, 5).Font.Bold = True: s2 = {R}(3, 0).Font.Bold'),
    (RECT, f'{SET}{R}(3, 0).Text = "X"'), (RECT, f'{SET}{R}(3, 0).Delete'), (RECT, f'{SET}{R}(3, 0).InsertAfter "X"'), (RECT, f'{SET}{R}(3, 0).InsertBefore "X"'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {R}(6, 0).InsertAfter "X"'), (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {R}(6, 0).InsertBefore "X"'), (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {R}(1, 0).InsertBefore "X"'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {R}(6, 0).Text = "X"'),
    (RECT, f'{SET}s2 = {R}(1, 5).Count & "|" & {R}(1, 5).Item(2).Text'), (RECT, f'{SET}s2 = {R}.Item(2).Text'), (RECT, f'{SET}s2 = {R}(1, 5).Item(9).Text'), (RECT, f'{SET}s2 = {R}(1, 5).Item(1).Count'),
    # What each range counts.
    *[(RECT, f'{SET}{C}(1, 5).Font.Bold = True: s2 = sh.TextFrame2.TextRange.{tail}') for tail in (
        'Runs(1, 2).Count', 'Runs(1, 2).Runs.Count', 'Runs(1, 2).Item(1).Text', 'Words.Item(1).Text', 'Words(1).Count', 'Words(1, 2).Count', 'Words(1).Words.Count', 'Paragraphs(1).Count', 'Paragraphs(1).Length', 'Paragraphs.Item(1).Text')],
    *[(RECT, f'{SET}s2 = {R}{tail}') for tail in ('(1, 5).Words.Count', '(3, 5).Words.Count', '(3, 5).Words(1).Text', '(3, 5).Words(2).Start', '(3, 5).Paragraphs.Count', '(3, 5).Paragraphs(1).Text')],
    *[(RECT, f's2 = sh.TextFrame2.TextRange.{tail}') for tail in (
        'Words.Count', 'Paragraphs.Count', 'Characters.Count', 'Count', 'Item(1).Text', 'Start & "|" & sh.TextFrame2.TextRange.Length', 'Words(1).Text', 'Paragraphs(1).Text', 'Characters(1, 0).Text', 'Characters.Text')],
    # Walking each kind of range.
    *[(RECT, f'{SET}{C}(1, 5).Font.Bold = True: For Each o In sh.TextFrame2.TextRange{tail}: s2 = s2 & o.Text & "/" & o.Start & ",": Next') for tail in (
        '.Runs', '.Words', '.Paragraphs', '.Characters', '.Characters(1, 3)', '', '.Runs(1)', '.Words(1)', '.Runs(1, 2)')],
    (RECT, f'{C}.Text = "ab cd" & vbCr & "ef": For Each o In sh.TextFrame2.TextRange.Paragraphs: o.Font.Size = 14: s2 = s2 & o.Text & ",": Next'),
    (RECT, f'For Each o In sh.TextFrame2.TextRange.Runs: s2 = s2 & "x": Next'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: For Each o In sh.TextFrame2.TextRange.Runs: o.Font.Size = 14: Next'),
    # Characters' other members.
    (RECT, f'{SET}s2 = {C}(7, 5).Caption'), (RECT, f'{SET}{C}(1, 5).Caption = "Bye"'), (RECT, f'{SET}s2 = {C}(7, 5)'), (RECT, f'{SET}{C}(1, 5) = "Bye"'),
    (RECT, f'{SET}s2 = {C}(7, 5).Start'), (RECT, f'{SET}s2 = {C}(7, 5).Length'),
    (TEXT, f'{SET}{C}(7, 5).Font.Bold = True: {C}.Text = "New"'), (TEXT, f'{SET}s2 = sh.TextFrame2.TextRange.Runs.Count'),
    # A run with a setting that stops short of its paragraph's break, and which side InsertBefore looks to.
    (RECT, f'{C}.Text = "ab cd" & vbCr & "ef": {C}(4, 2).Font.Bold = True: s2 = sh.TextFrame2.TextRange.Runs.Count'),
    (RECT, f'{SET}{C}(1, 5).Font.Bold = True: {R}(6, 1).InsertBefore "X"'), (RECT, f'{SET}{C}(7, 1).Font.Bold = True: {R}(6, 1).InsertAfter "X"'),
]
READS = [*[f'{C}(1, 5).Font.{one}' for one in ('Underline', 'Strikethrough', 'Superscript', 'Subscript', 'FontStyle', 'ColorIndex', 'TintAndShade', 'Shadow')],
         *[f'{C}(1, 7).Font.{one}' for one in ('Bold', 'Italic', 'Size', 'Color', 'Name', 'Underline', 'Strikethrough', 'Superscript', 'Subscript', 'FontStyle', 'ColorIndex', 'TintAndShade', 'Shadow')],
         *[f'{R}(1, 7).Font.{one}' for one in ('Bold', 'Italic', 'Size', 'Name', 'UnderlineStyle', 'Strike', 'StrikeThrough', 'DoubleStrikeThrough', 'Superscript', 'Subscript', 'BaselineOffset', 'Caps', 'Allcaps', 'Smallcaps', 'Spacing', 'Kerning', 'Equalize')],
         *[f'{R}(1, 5).Font.{one}' for one in ('Italic', 'Size', 'Name')],
         'sh.TextFrame2.TextRange.Runs.Count', f'{C}(1, 5).Font.Bold', f'{C}(7, 5).Font.Bold', f'{C}.Font.Bold', f'{C}(1, 5).Font.Size', f'{C}(7, 5).Font.Size', f'{C}.Font.Size',
         f'{C}(1, 5).Font.Italic', f'{C}(3, 5).Font.Italic', f'{C}.Font.Italic', f'{C}(1, 5).Font.Color', f'{C}.Font.Color', f'{C}(1, 5).Font.Name', f'{C}.Font.Name',
         f'{C}.Text', f'{C}.Count', f'{R}(1, 5).Font.Bold', f'{R}.Font.Bold', 'sh.TextFrame2.TextRange.Font.Bold', 's2']


def _probe() -> str:
    lines: list[str] = []
    chunks = [range(at, min(at + 80, len(CASES))) for at in range(0, len(CASES), 80)]
    # In procedures of their own, since one holding them all is too large to compile. Each takes its own error: passed up
    # through a call, Excel's -2147024809 turns into VBA's 5.
    for number, chunk in enumerate(chunks):
        lines += [f'Private Sub Act{number}(ByVal which As Long, ByVal ws As Worksheet, ByVal sh As Object, o As Object, s2 As Variant, failure As Long)', '    On Error GoTo Failed', '    Select Case which']
        for index in chunk:
            lines += [f'    Case {index}', f'        {CASES[index][1]}']
        lines += ['    End Select', '    Exit Sub', 'Failed:', '    failure = Err.Number', 'End Sub', '']
    lines += ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
              '    Dim wb As Workbook, ws As Worksheet, sh As Object, o As Object, s As String, v As String, s2 As Variant, failure As Long',
              '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)']
    boxes = ', '.join(str(index) for index, (shape, _) in enumerate(CASES) if shape == TEXT)
    lines += ['    Select Case which', f'    Case {boxes}: Set sh = {TEXT}', f'    Case Else: Set sh = {RECT}', '    End Select',
              '    Select Case which \\ 80']
    for number in range(len(chunks)):
        lines += [f'    Case {number}: Act{number} which, ws, sh, o, s2, failure']
    lines += ['    End Select', '    s = IIf(failure = 0, "ok", "Err " & failure)', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "~" & Replace(v, vbLf, "<LF>")']
    lines += ['    On Error GoTo 0', '    Tried = s',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              'End Function',
              '', 'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=560.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = []
        for index in range(len(CASES)):
            with zipfile.ZipFile(Path(folder) / f'c{index}.xlsx') as package:
                xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
            parts.append(re.findall(r'<xdr:txBody>.*?</xdr:txBody>', xml, re.DOTALL)[0])
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{shape}: {statements}', answer, part] for (shape, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
