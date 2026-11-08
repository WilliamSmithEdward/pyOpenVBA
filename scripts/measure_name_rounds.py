"""Held-out rounds over the PARTIAL rows of Name and Names whose code reaches no reported gap: a name's Name and NameLocal,
Index, Visible, Comment, Category and CategoryLocal, MacroType, Delete and ValidWorkbookParameter, and Names.Count. Each is
read, set and read back where it sets.

The book is the table rounds' (scripts/measure_table_pivot_rounds.py): a new book with one to three workbook names and
perhaps a sheet's own, the same saved and reopened, or the names Excel saved in tests/fixtures/object_roundtrip.xlsx.

    python scripts/measure_name_rounds.py predict <seed>
    python scripts/measure_name_rounds.py measure <seed>

tests/test_excel_name_rounds.py replays the rounds.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_table_pivot_rounds as base  # noqa: E402

OUT = base.FIXTURES / 'name_rounds'
REACHES = {'Name': base.TABLES.reaches['Name'], 'Names': base.TABLES.reaches['Names']}
READ = 'x = Key(o.Name, book) & " " & o.Index & " " & wbL.Names.Count'
MEMBERS = {
    'Name': [
        'x = Key(o.Name, book)', 'x = Key(o.NameLocal, book)', 'x = Key(o.Index, book)',
        f'o.Name = "Zeta_1": {READ}', f'o.Name = "aaa_first": {READ}', f'o.Name = UCase(o.Name): {READ}',
        f'o.Name = "1bad": {READ}', f'o.Name = "A1": {READ}', f'o.Name = "R1C1": {READ}', f'o.Name = "has space": {READ}',
        f'o.Name = wbL.Names(wbL.Names.Count).Name: {READ}', f'o.Name = "": {READ}', f'o.Name = 12: {READ}',
        f'o.Name = "_under.dot": {READ}', f'o.Name = "Имя_" & ChrW(1488): {READ}', f'o.Name = String(255, "n"): {READ}',
        f'o.Name = String(256, "n"): {READ}', f'o.NameLocal = "Local_set": {READ}', f'o.NameLocal = "bad name": {READ}',
        'wbL.Names.Add "Aaa", "=1": x = Key(o.Index, book) & " " & wbL.Names("Aaa").Index',
        'wbL.Names.Add "zzz", "=2": x = Key(o.Index, book) & " " & wbL.Names("zzz").Index',
        'x = Key(o.Visible, book)', 'o.Visible = False: x = Key(o.Visible, book) & " " & wbL.Names.Count',
        'o.Visible = False: o.Visible = True: x = Key(o.Visible, book)', 'o.Visible = 0: x = Key(o.Visible, book)',
        'o.Visible = "no": x = Key(o.Visible, book)', 'o.Visible = Null: x = Key(o.Visible, book)',
        'x = Key(o.Comment, book)', 'o.Comment = "c" & vbLf & "d": x = Key(o.Comment, book)', 'o.Comment = String(255, "z"): x = Len(o.Comment)',
        'o.Comment = String(256, "z"): x = Len(o.Comment)', 'o.Comment = 3.5: x = Key(o.Comment, book)', 'o.Comment = "": x = Key(o.Comment, book)',
        'x = Key(o.Category, book)', 'x = Key(o.CategoryLocal, book)', 'o.Category = "Financial": x = Key(o.Category, book)',
        'o.MacroType = xlFunction: o.Category = "financial": x = Key(o.Category, book) & " " & o.CategoryLocal',
        'o.MacroType = xlFunction: o.Category = "My Own": x = Key(o.Category, book) & " " & o.CategoryLocal',
        'o.MacroType = xlCommand: o.CategoryLocal = "Text": x = Key(o.Category, book) & " " & o.CategoryLocal',
        'o.MacroType = xlFunction: o.Category = 4: x = Key(o.Category, book)', 'o.MacroType = xlFunction: x = Key(o.Category, book)',
        'x = Key(o.MacroType, book)', 'o.MacroType = xlFunction: x = Key(o.MacroType, book)', 'o.MacroType = xlCommand: x = Key(o.MacroType, book)',
        'o.MacroType = xlCommand: o.MacroType = xlNone: x = Key(o.MacroType, book)', 'o.MacroType = xlNotXLM: x = Key(o.MacroType, book)',
        'o.MacroType = 0: x = Key(o.MacroType, book)', 'o.MacroType = "1": x = Key(o.MacroType, book)',
        'o.Delete: x = wbL.Names.Count', 'o.Delete: x = Key(wbL.Names(1), book)', 'x = Key(o.ValidWorkbookParameter, book)',
        'o.MacroType = xlFunction: x = Key(o.ValidWorkbookParameter, book)',
    ],
    'Names': [
        'x = Key(o.Count, book)', 'o.Add "Fresh", "=Main!$C$1": x = Key(o.Count, book)', 'o.Add "Fresh", "=3": o.Add "Fresh", "=4": x = Key(o.Count, book)',
        'o(1).Delete: x = Key(o.Count, book)', 'o.Add "שם", "=5": o.Add "ชื่อ", "=6": x = Key(o.Count, book)',
        'wsL.Names.Add "Own", "=7": x = Key(o.Count, book) & " " & wsL.Names.Count', 'o(1).Visible = False: x = Key(o.Count, book)',
    ],
}
LISTS = base.Plan(REACHES, MEMBERS, (), {}, base.setup, base.books, base.DIMS, OUT, base.SAMPLE)


if __name__ == '__main__':
    base.main(LISTS)
