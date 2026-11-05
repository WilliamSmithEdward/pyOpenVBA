"""A sweep any object can use: set a member, read it, save, reopen, read it again. A value that changes across the save is one the file
lost, which no read-back alone can show.

    TARGETS = [Target('rect', 'Set s = ws.Shapes.AddShape(1, 10, 10, 100, 40): s.Name = "T"', 'Set s = ws.Shapes("T")')]
    SETTINGS = {'s.Rotation': ['30']}
    ACCEPTED = {'rect: s.Rotation = 30': 'why the file does not keep it yet'}
    EXCEL_LOSES = {'sheet: s.ScrollArea = "A1:B2"': 'the fixture that measured Excel losing it too'}

    @pytest.mark.parametrize('case', cases(TARGETS, SETTINGS), ids=str)
    def test_each_setting_comes_back(case, tmp_path):
        check(case, tmp_path, ACCEPTED, EXCEL_LOSES)

A set the model refuses, or that raises as a macro would see, passes: nothing was claimed. A lost value fails unless ACCEPTED names it,
and a named one that now comes back fails too, so the list only shrinks. EXCEL_LOSES is what Excel's own save drops as well, by a fixture
that measured it: the model is right to lose those, and fails if it ever keeps one. A value naming its workbook, as OnAction's Book1!M does,
is compared without the name. It says nothing of whether a kept value matches Excel; that takes
a measurement. No case may use a member that sends mail or follows a link.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError

#: These could send real mail from the machine's account, in Excel and in any probe written from a sweep.
_NEVER = re.compile(r'\.Follow\b|mailto|FollowHyperlink|\b(?:SendMail|SendForReview|ReplyWithChanges|EndReview|Route|RouteWorkbook|RoutingSlip|HasRoutingSlip|HasMailer|Mailer|MailEnvelope|SendFaxOverInternet'
                    r'|MailLogon|MailLogoff|MailSession|MailSystem|xlDialogSendMail|Reply|ReplyAll)\b', re.IGNORECASE)


@dataclass(frozen=True)
class Target:
    """What is set. ``make`` gets it into ``s`` with the sheet in ``ws``, making it if need be; ``find`` gets it back after the reopen. With a
    ``book``, the file to open first."""

    name: str
    make: str
    find: str
    book: Path | None = None
    #: What sort of thing it is, for a setting to apply to.
    kind: str = ''


@dataclass(frozen=True)
class Case:
    target: Target
    path: str
    value: str

    def __str__(self) -> str:
        return f'{self.target.name}: {self.path} = {self.value}'


def cases(targets: list[Target], settings: dict[str, list[str]], kinds: dict[str, tuple[str, ...]] | None = None) -> list[Case]:
    """Each setting on each target. ``kinds`` narrows a setting, by its path, to targets of those kinds."""
    found: list[Case] = []
    for path, values in settings.items():
        for target in targets:
            if kinds is None or path not in kinds or target.kind in kinds[path]:
                found += [Case(target, path, value) for value in values]
    return found


def _reader(make: str, statement: str, read: str) -> str:
    return (f'Public Function Go() As String\n    Dim ws As Object, s As Object, v As Variant\n    Set ws = ActiveWorkbook.Worksheets(1)\n    {make}\n    {statement}\n'
            f'    v = {read}\n    If IsObject(v) Then Go = "object" Else Go = CStr(v)\nEnd Function\n'
            'Public Function Book() As String\n    Book = ActiveWorkbook.Name\nEnd Function\n')


def without_book(value: str, book: str) -> str:
    """A value with the name of its workbook taken from any reference to it, quoted or not, since a save renames the book."""
    return value.replace(f"'{book}'!", '').replace(f'{book}!', '')


def outcome(case: Case, folder: Path) -> tuple[str, str]:
    """``kept``, ``no claim`` or ``lost``, and what was seen."""
    assert not _NEVER.search(f'{case.path} {case.value} {case.target.make} {case.target.find}'), 'no mail or link-following members in a sweep'
    target = case.target
    if target.book is not None:
        app = ExcelApplication.open(target.book, with_vba=False)
    else:
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
    app.add_module(_reader(target.make, f'{case.path} = {case.value}', case.path), name='Sweep')
    try:
        before = without_book(str(app.run('Sweep.Go')), str(app.run('Sweep.Book')))
    except (VBAUnsupportedError, VBARuntimeError) as refused:
        return 'no claim', str(refused)[:100]
    saved = folder / 'saved.xlsx'
    app.save(saved)
    again = ExcelApplication.open(saved, with_vba=False)
    again.add_module(_reader(target.find, '', case.path), name='Sweep')
    try:
        after = without_book(str(again.run('Sweep.Go')), str(again.run('Sweep.Book')))
    except VBAUnsupportedError as refused:
        return 'no claim', f'the set stood at {before}; the reopened file reports {str(refused)[:80]}'
    if after == before:
        return 'kept', after
    return 'lost', f'{before} before the save, {after} after it'


def check(case: Case, folder: Path, accepted: dict[str, str], excel_loses: dict[str, str] | None = None) -> None:
    """A loss must be on one of the lists, and what the lists name must still be lost."""
    result, seen = outcome(case, folder)
    name = str(case)
    if excel_loses is not None and name in excel_loses:
        assert result == 'lost', f'{name} is lost by Excel too ({excel_loses[name]}), and the model does not lose it ({result}: {seen})'
    elif name in accepted:
        assert result == 'lost', f'{name} no longer loses its value ({result}: {seen}): take it off the accepted list'
    else:
        assert result != 'lost', f'{name} is lost across a save: {seen}'
