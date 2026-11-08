"""Held-out rounds over ListObject's PARTIAL rows whose code reaches no reported gap: its table style options, AutoFilter and
ShowAutoFilter, Comment, AlternativeText, Summary, TableStyle, Active, DisplayRightToLeft, Slicers, Sort, the members a table of
other sources answers (QueryTable, Refresh, SharePointURL, TableObject, Unlink, XmlMap) and its default member. Each is read, set
and read back where it sets. ExportToVisio and Publish are not asked: one starts Visio, the other sends the table to a server.

The book is the table rounds' (scripts/measure_table_pivot_rounds.py), with a seeded set of the table's options set before any
save, so a reopened book reads them from the file; the Excel-saved book gets them in the session half the time.

    python scripts/measure_list_object_rounds.py predict <seed>
    python scripts/measure_list_object_rounds.py measure <seed>

tests/test_excel_list_object_rounds.py replays the rounds.
"""
from __future__ import annotations

import random
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_table_pivot_rounds as base  # noqa: E402

OUT = base.FIXTURES / 'list_object_rounds'
REACHES = {'ListObject': base.TABLES.reaches['ListObject']}
FLAGS = ('ShowTableStyleColumnStripes', 'ShowTableStyleFirstColumn', 'ShowTableStyleLastColumn', 'ShowTableStyleRowStripes', 'ShowAutoFilter')


def _flag_calls(flag: str) -> list[str]:
    return [f'x = Key(o.{flag}, book)', f'o.{flag} = Not o.{flag}: x = Key(o.{flag}, book)', f'o.{flag} = True: x = Key(o.{flag}, book)',
            f'o.{flag} = 0: x = Key(o.{flag}, book)', f'o.{flag} = "yes": x = Key(o.{flag}, book)']


MEMBERS = {'ListObject': [
    *[call for flag in FLAGS for call in _flag_calls(flag)],
    'x = o', 'x = Key(o.Active, book)', 'ws.Activate: ws.Range(tcell).Select: x = Key(o.Active, book)', 'ws2.Activate: x = Key(o.Active, book)',
    'x = Key(o.AlternativeText, book)', 'o.AlternativeText = "alt " & String(300, "a"): x = Len(o.AlternativeText)', 'o.AlternativeText = "": x = Key(o.AlternativeText, book)',
    'x = Key(o.Summary, book)', 'o.Summary = "sum" & vbLf & "two": x = Key(o.Summary, book)', 'o.Summary = 12: x = Key(o.Summary, book)',
    'x = Key(o.Comment, book)', 'o.Comment = "c" & vbLf & "d": x = Key(o.Comment, book)', 'o.Comment = String(255, "z"): x = Len(o.Comment)',
    'o.Comment = String(256, "z"): x = Len(o.Comment)', 'o.Comment = 3.5: x = Key(o.Comment, book)',
    'x = Key(o.AutoFilter, book)', 'x = Key(o.AutoFilter.FilterMode, book)', 'x = Key(o.AutoFilter.Range, book)',
    'o.ShowAutoFilter = False: x = Key(o.AutoFilter, book)', 'o.ShowAutoFilter = False: o.ShowAutoFilter = True: x = Key(o.AutoFilter.Filters.Count, book)',
    'x = Key(o.DisplayName, book)', 'x = Key(o.DisplayRightToLeft, book)', 'ws.DisplayRightToLeft = True: x = Key(o.DisplayRightToLeft, book)',
    'x = Key(o.TableStyle, book)', 'o.TableStyle = "TableStyleLight9": x = Key(o.TableStyle, book)', 'o.TableStyle = "": x = Key(o.TableStyle, book)',
    'o.TableStyle = "NoSuchStyle": x = Key(o.TableStyle, book)', 'o.TableStyle = "TableStyleDark2": x = Key(o.TableStyle, book)',
    'x = Key(o.Slicers, book) & o.Slicers.Count', 'x = TypeName(o.Sort) & " " & o.Sort.SortFields.Count & " " & o.Sort.Header',
    'x = Key(o.QueryTable, book)', 'o.Refresh: x = "refreshed"', 'x = Key(o.SharePointURL, book)', 'x = Key(o.TableObject, book)',
    'o.Unlink: x = Key(o.Range, book)', 'x = Key(o.XmlMap, book)',
]}


def _options(seed: int, kind: str) -> list[str]:
    """Seeded table options set on every table the sheet holds."""
    rng = random.Random(f'{seed}:{kind}:options')
    lines = []
    for flag in FLAGS[:4]:
        if rng.random() < 0.4:
            lines.append(f'    ws.ListObjects(1).{flag} = {rng.choice(("True", "False"))}')
    if rng.random() < 0.3:
        lines.append('    ws.ListObjects(1).ShowAutoFilter = False')
    if rng.random() < 0.4:
        lines.append(f'    ws.ListObjects(1).Comment = "note {seed % 13}"')
    if rng.random() < 0.4:
        lines.append(f'    ws.ListObjects(1).AlternativeText = "alt {seed % 11}": ws.ListObjects(1).Summary = "sum {seed % 7}"')
    if rng.random() < 0.4:
        lines.append(f'    ws.ListObjects(1).TableStyle = "{rng.choice(("TableStyleLight1", "TableStyleMedium9", "TableStyleDark3", ""))}"')
    return lines


def setup(kind: str, seed: int) -> str:
    """The table rounds' book, with the options set before a reopened book is saved."""
    built = base.setup('new' if kind == 'reopened' else kind, seed)
    options = _options(seed, kind)
    if kind == 'file':
        return built + ('\n'.join(options) + '\n' if options and random.Random(f'{seed}:file').random() < 0.5 else '')
    built += '\n'.join(options) + '\n' if options else ''
    if kind == 'reopened':
        built += '\n'.join(['    saved = path & "\\r" & Format(Now, "hhnnss") & (Int(Timer * 1000) Mod 1000) & ".xlsx"', '    wbT.SaveAs saved, xlOpenXMLWorkbook',
                            '    wbT.Close False', '    Set wbT = Workbooks.Open(saved): Set wbL = wbT',
                            '    Set ws = wbT.Worksheets("Main"): Set wsL = ws: Set ws2 = wbT.Worksheets("Second")']) + '\n'
    return built


LISTS = base.Plan(REACHES, MEMBERS, (), {}, setup, base.books, base.DIMS, OUT, base.SAMPLE)


if __name__ == '__main__':
    base.main(LISTS)
