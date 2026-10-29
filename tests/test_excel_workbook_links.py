"""Workbook.LinkSources, LinkInfo, BreakLink, UpdateLinks and SaveLinkValues, on a workbook whose links' sources are
closed.

Measured by scripts/measure_workbook_links.py, whose macros the model runs
here as Excel ran them, over the files Excel saved: Dest.xlsx reads Beta.xlsx
and Alpha.xlsx, and is reopened alone, its links not updated. What the
model saves of each link is compared with what Excel saved.

Two things Excel does are left out, each held by a strict xfail below: it
refuses BreakLink with 1004 on a workbook a macro has just opened until
something lists its links, and it adds a hidden name, _xlfn.SINGLE, on
reading a formula Formula2 shows with an @ (the model does not make that
name; tests/test_excel_worksheet_rename.py notes the same). The name lists
are compared without it.
"""
from __future__ import annotations

import base64
import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/workbook_links.json').read_text(encoding='utf-8'))
FILES: dict[str, dict[str, str]] = RECORD['files']
#: The hidden name Excel makes for Formula2's @, as the name lists show it.
SINGLE = '_xlfn.SINGLE==#NAME?|'
#: Of a file Excel saved again, the parts the fixture keeps, and the file they were saved again from.
RESAVED = {'PlainAgain.xlsx': 'Plain.xlsx', 'NamedAgain.xlsx': 'Named.xlsx', 'NamedThird.xlsx': 'NamedAgain.xlsx',
           'LocalAgain.xlsx': 'Local.xlsx', 'MixedAgain.xlsx': 'Mixed.xlsx', 'BlankAgain.xlsx': 'Blank.xlsx'}


def _part(file: str, part: str) -> str:
    """A part of a file Excel saved, as recorded: its text, or same:<file> for the same part of an earlier file."""
    text = FILES[file][part]
    return _part(text[len('same:'):], part) if text.startswith('same:') else text


def _kept_of_resaved(part: str) -> bool:
    """Whether the fixture keeps a part of a file saved again: those that name its links, all of them."""
    return 'externalLink' in part or part in ('xl/workbook.xml', 'xl/worksheets/sheet1.xml')


def _parts(file: str) -> dict[str, str]:
    """Every part of a file Excel saved; one the fixture keeps only some parts of takes the rest from the file it
    was saved again from."""
    base = {part: text for part, text in _parts(RESAVED[file]).items() if not _kept_of_resaved(part)} \
        if file in RESAVED else {}
    return base | {part: _part(file, part) for part in FILES[file]}


def _write(folder: Path, file: str) -> Path:
    """The file Excel saved, in ``folder``, beside where its links' sources were."""
    path = folder / file
    with zipfile.ZipFile(path, 'w') as package:
        for part, text in _parts(file).items():
            data = base64.b64decode(text[len('base64:'):]) if text.startswith('base64:') \
                else text.replace('<folder>', str(folder)).encode('utf-8')
            package.writestr(part, data)
    return path


def _app(folder: Path, module: str = RECORD['module']) -> ExcelApplication:
    """A workbook of the measuring macros, with Dest.xlsx, its two sources and the files the status case opens in
    ``folder``."""
    for file in ('Dest.xlsx', 'Alpha.xlsx', 'Beta.xlsx', 'Plain.xlsx', 'Blank.xlsx', 'NamedAgain.xlsx'):
        _write(folder, file)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['sink'], name='LinkSink', kind='class')
    app.add_module(module, name='LinksProbe')
    return app


def _run(app: ExcelApplication, folder: Path, proc: str, *args: str) -> str:
    return str(app.run(f'LinksProbe.{proc}', str(folder), *args)).replace(str(folder), '<folder>')


def _excels(key: str) -> list[str]:
    """Excel's answer, a piece at a time, without the hidden name the model does not make."""
    return RECORD[key].replace(SINGLE, '').split('|')


def _saved(path: Path, folder: Path) -> dict[str, str]:
    """The text parts of a file the model saved, its folder written <folder>."""
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name).decode('utf-8').replace(str(folder), '<folder>')
                for name in package.namelist() if not name.endswith('.bin')}


def _cells(parts: dict[str, str], sheet: str) -> dict[str, tuple[str, str, str]]:
    """Each cell of a saved sheet part: its type, its formula and its value, text as its text, whether shared as
    Excel keeps it or inline as the model writes it."""
    strings = re.findall(r'<si>(.*?)</si>', parts.get('xl/sharedStrings.xml', ''), re.DOTALL)
    out: dict[str, tuple[str, str, str]] = {}
    for found in re.finditer(r'<c r="([A-Z]+[0-9]+)"([^>]*?)(?:/>|>(.*?)</c>)', parts[sheet], re.DOTALL):
        kind = re.search(r'\bt="([A-Za-z]+)"', found.group(2))
        body = found.group(3) or ''
        formula = re.search(r'<f\b[^>]*>(.*?)</f>', body, re.DOTALL)
        value = re.search(r'<v>(.*?)</v>', body, re.DOTALL)
        text = value.group(1) if value else ''
        typed = kind.group(1) if kind else ''
        if typed == 's':
            text, typed = re.sub(r'<[^>]+>', '', strings[int(text)]), 'text'
        elif typed == 'inlineStr':
            text, typed = re.sub(r'<[^>]+>', '', body), 'text'
        out[found.group(1)] = (typed, formula.group(1) if formula else '', text)
    return out


def _references(parts: dict[str, str]) -> list[str]:
    """The parts a saved workbook's externalReferences name, in order, whatever relationship ids they take."""
    targets = dict(re.findall(r'<Relationship\b[^>]*?\bId="([^"]+)"[^>]*?\bTarget="([^"]+)"',
                              parts['xl/_rels/workbook.xml.rels']))
    return [targets[identity] for identity in re.findall(r'<externalReference r:id="([^"]+)"/>',
                                                         parts['xl/workbook.xml'])]


def _link_parts(parts: dict[str, str]) -> dict[str, str]:
    return {name: text for name, text in parts.items() if name.startswith('xl/externalLinks/')}


def test_link_sources_and_link_info(tmp_path: Path) -> None:
    """LinkSources lists the links' workbooks by their full names in the file's order, Beta's first as it was made
    first, and Empty for any other Type; LinkInfo answers xlLinkInfoStatus with 3, the status of a link whose
    workbook is closed, and #N/A or #VALUE! for anything else; UpdateLinks and SaveLinkValues read the file's."""
    got = _run(_app(tmp_path), tmp_path, 'SurveyLinks')
    assert got.split('|') == _excels('survey')


@pytest.mark.xfail(strict=True, reason="Excel adds the hidden name _xlfn.SINGLE when it reads a formula Formula2 "
                                       "shows with an @; the model does not")
def test_the_hidden_name_excel_adds_for_formula2s_at(tmp_path: Path) -> None:
    assert _run(_app(tmp_path), tmp_path, 'SurveyLinks') == RECORD['survey']


def test_breaking_a_link_raises_no_event(tmp_path: Path) -> None:
    """BreakLink raises neither SheetChange nor SheetCalculate; CalculateFull and an edit after it do."""
    assert _run(_app(tmp_path), tmp_path, 'BreakEvents') == RECORD['events']


@pytest.mark.parametrize('which', ['Alpha', 'Beta', 'All'])
def test_breaking_a_link(which: str, tmp_path: Path) -> None:
    """Each formula that reads the link, itself or through names, is the value it showed, an array's and a spill's
    cells too; a name that reads it is #REF!, saved as #REF! and read back as #REF!#REF!; the link is gone from
    LinkSources and from the file, the other renumbered."""
    app = _app(tmp_path)
    got = _run(app, tmp_path, 'BreakCase', which, f'Broke{which}.xlsx')
    assert got.split('|') == _excels(f'break_{which.lower()}')
    mine, excels = _saved(tmp_path / f'Broke{which}.xlsx', tmp_path), _parts(f'Broke{which}.xlsx')
    assert _link_parts(mine) == _link_parts(excels)
    for sheet in ('xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'):
        assert _cells(mine, sheet) == _cells(excels, sheet)
    pattern = r'<definedNames>.*?</definedNames>'
    assert re.findall(pattern, mine['xl/workbook.xml']) == re.findall(pattern, excels['xl/workbook.xml'])
    # Excel numbers the workbook's relationships again on every save; the model keeps the file's.
    assert _references(mine) == _references(excels)
    links = sorted(name for name in mine if name.startswith('xl/externalLinks/externalLink'))
    assert [name for name in re.findall(r'PartName="/(xl/externalLinks/[^"]+)"', mine['[Content_Types].xml'])] \
        == links
    targets = re.findall(r'Target="(externalLinks/[^"]+)"', mine['xl/_rels/workbook.xml.rels'])
    assert sorted('xl/' + target for target in targets) == links


def test_break_link_errors(tmp_path: Path) -> None:
    """A path to no link, an empty name, an OLE link's Type 1004; another Type error 5, text no number 13, Null 94,
    and a missing Type 450; the file's name alone, either case and / for \\ break the link; a link broken already is
    left as it is."""
    got = _run(_app(tmp_path), tmp_path, 'ErrorCases')
    listed = RECORD['errors'].split('#', 1)[1]
    assert got.split('#', 1)[1] == listed


@pytest.mark.xfail(strict=True, reason="Excel refuses BreakLink with 1004 on a workbook a macro has just opened "
                                       "until its links are listed; the model breaks the link at once")
def test_break_link_before_the_links_are_listed(tmp_path: Path) -> None:
    got = _run(_app(tmp_path), tmp_path, 'ErrorCases')
    assert got.split('#', 1)[0] == RECORD['errors'].split('#', 1)[0]


def test_link_settings(tmp_path: Path) -> None:
    """UpdateLinks takes 1 to 3, rounded as a Long, else 1004, and leaves the workbook saved; SaveLinkValues takes
    a Boolean and leaves it unsaved. Saved, they are the workbookPr's updateLinks and saveExternalLinkValues, the
    links' parts without their values, whose formulas are then #REF! each worked out in full; the workbook keeps
    the values, which a save with SaveLinkValues True again writes."""
    app = _app(tmp_path)
    assert _run(app, tmp_path, 'LinkSettings').split('|') == _excels('settings')
    assert _run(app, tmp_path, 'MoreCases').split('|') == _excels('more')
    for file in ('Never.xlsx', 'ValuesAgain.xlsx', 'Always.xlsx', 'NoLinks.xlsx'):
        mine, excels = _saved(tmp_path / file, tmp_path), _parts(file)
        pattern = r'<workbookPr\b[^>]*>'
        assert re.findall(pattern, mine['xl/workbook.xml']) == re.findall(pattern, excels['xl/workbook.xml'])
        assert _link_parts(mine) == _link_parts(excels)


#: The status case's changes to Alpha.xlsx, as its macro makes them.
CHANGES = '''Public Sub ChangeAlpha(ByVal folder As String)
    Dim a As Workbook
    Set a = Workbooks.Open(folder & "\\Alpha.xlsx")
    a.Worksheets("Sheet1").Range("A1").Value = 99
    a.Save
    a.Close False
End Sub

Public Sub RenameOther(ByVal folder As String)
    Dim a As Workbook
    Set a = Workbooks.Open(folder & "\\Alpha.xlsx")
    a.Worksheets("Other").Name = "Renamed"
    a.Save
    a.Close False
End Sub
'''


def test_link_statuses(tmp_path: Path) -> None:
    """A link whose workbook's file is there is 3, changed or not, 1 once the file is moved away, and 2 once a sheet
    a formula reads of it is renamed, a blank cell's sheet too, but not a sheet only listed, nor the one Excel adds
    for the workbook's names. The macro's own StatusOf reads each; the model taking no file system statement, the
    move away and back of Beta.xlsx, which the macro does with Name, is done here."""
    app = _app(tmp_path, RECORD['module'].replace('Private Function StatusOf', 'Public Function StatusOf'))
    app.add_module(CHANGES, name='Changes')

    def status(file: str = 'Dest.xlsx') -> str:
        return _run(app, tmp_path, 'StatusOf', file)

    got = f'as saved:{status()}#phantom:{status("NamedAgain.xlsx")}'
    app.run('Changes.ChangeAlpha', str(tmp_path))
    got += f'#alpha changed:{status()}'
    (tmp_path / 'Beta.xlsx').rename(tmp_path / 'Beta.moved')
    got += f'#beta gone:{status()}'
    (tmp_path / 'Beta.moved').rename(tmp_path / 'Beta.xlsx')
    app.run('Changes.RenameOther', str(tmp_path))
    got += f'#sheet gone:{status()}#unread sheet gone:{status("Plain.xlsx")}#blank cell sheet gone:{status("Blank.xlsx")}'
    assert got == RECORD['statuses']


@pytest.mark.parametrize('file', sorted(RESAVED))
def test_a_save_writes_the_links_as_excel(file: str, tmp_path: Path) -> None:
    """Saved again with its links' workbooks closed, a link one of whose workbook's names a formula reads gains a
    sheet named after that workbook, once; and after a break the other link keeps only the cells still read."""
    source = _write(tmp_path, RESAVED[file])
    app = ExcelApplication.open(source)
    if file == 'MixedAgain.xlsx':
        app.add_module(BREAK, name='Breaking')
        app.run('Breaking.Break')
    saved = app.save(tmp_path / file)
    mine, excels = _saved(saved, tmp_path), _parts(file)
    assert _link_parts(mine) == _link_parts(excels)
    assert _cells(mine, 'xl/worksheets/sheet1.xml') == _cells(excels, 'xl/worksheets/sheet1.xml')


def test_a_name_a_deleted_sheet_left(tmp_path: Path) -> None:
    """A name reading a deleted sheet is saved with #REF! for each reference it lost, and read back with #REF!#REF!
    there, as a broken link's name is."""
    app = _app(tmp_path)
    assert _run(app, tmp_path, 'DeletedSheet') == RECORD['deleted_sheet']
    pattern = r'<definedNames>.*?</definedNames>'
    assert re.findall(pattern, _saved(tmp_path / 'Gone.xlsx', tmp_path)['xl/workbook.xml']) \
        == re.findall(pattern, _parts('Gone.xlsx')['xl/workbook.xml'])


#: Breaks Alpha's link in the workbook ActiveWorkbook is, its links listed first.
BREAK = ('Sub Break()\nDim v As Variant\nv = ActiveWorkbook.LinkSources()\n'
         'ActiveWorkbook.BreakLink ActiveWorkbook.Path & "\\Alpha.xlsx", xlLinkTypeExcelLinks\nEnd Sub\n')


def test_a_protected_sheets_formulas_are_refused(tmp_path: Path) -> None:
    """Excel stops a break over a protected sheet's formulas with an alert DisplayAlerts does not hold back."""
    app = ExcelApplication.open(_write(tmp_path, 'Dest.xlsx'))
    app.add_module(BREAK.replace('v = ActiveWorkbook', 'ActiveWorkbook.Worksheets("Second").Protect\n'
                                                       'v = ActiveWorkbook', 1), name='Breaking')
    with pytest.raises(VBAUnsupportedError, match='protected'):
        app.run('Breaking.Break')
    assert app.workbook.sheets_[0].cells_[(2, 1)].formula


def test_a_link_a_validation_reads_is_refused(tmp_path: Path) -> None:
    from pyopenvba._a1 import Area
    from pyopenvba.apps.excel._validation import Rule

    app = ExcelApplication.open(_write(tmp_path, 'Dest.xlsx'))
    app.workbook.sheets_[0].validations.append(Rule([Area(1, 6, 1, 6)], kind=3, formula1='[2]Sheet1!$B$1:$B$5'))
    app.add_module(BREAK, name='Breaking')
    with pytest.raises(VBAUnsupportedError, match='validation'):
        app.run('Breaking.Break')


def test_link_info_of_an_edition_is_refused(tmp_path: Path) -> None:
    app = ExcelApplication.open(_write(tmp_path, 'Dest.xlsx'))
    app.add_module('Function Edition() As Variant\nEdition = ActiveWorkbook.LinkInfo("Alpha.xlsx", 3, 1, 1)\n'
                   'End Function\n', name='Asking')
    with pytest.raises(VBAUnsupportedError, match='EditionRef'):
        app.run('Asking.Edition')


def test_a_workbook_with_a_dde_link(tmp_path: Path) -> None:
    """A link to a DDE source the model leaves out: LinkSources still lists the workbooks, and asking for the other
    kinds of link, or breaking one, reports itself."""
    path = _write(tmp_path, 'Dest.xlsx')
    with zipfile.ZipFile(path) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    parts['xl/externalLinks/externalLink3.xml'] = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<externalLink '
        b'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><ddeLink ddeService="Excel" '
        b'ddeTopic="Book1"/></externalLink>')
    parts['xl/workbook.xml'] = parts['xl/workbook.xml'].replace(
        b'</externalReferences>', b'<externalReference r:id="rId99"/></externalReferences>')
    parts['xl/_rels/workbook.xml.rels'] = parts['xl/_rels/workbook.xml.rels'].replace(
        b'</Relationships>', b'<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        b'relationships/externalLink" Target="externalLinks/externalLink3.xml"/></Relationships>')
    with zipfile.ZipFile(path, 'w') as package:
        for name, data in parts.items():
            package.writestr(name, data)
    app = ExcelApplication.open(path)
    app.add_module('Function Listed() As Long\nListed = UBound(ActiveWorkbook.LinkSources(xlExcelLinks))\n'
                   'End Function\nFunction Others() As Variant\nOthers = ActiveWorkbook.LinkSources(xlOLELinks)\n'
                   'End Function\n' + BREAK, name='Asking')
    assert app.run('Asking.Listed') == 2
    for macro in ('Asking.Others', 'Asking.Break'):
        with pytest.raises(VBAUnsupportedError, match='OLE or DDE'):
            app.run(macro)


def test_a_link_to_a_workbook_the_application_has_open_is_refused(tmp_path: Path) -> None:
    app = _app(tmp_path)
    app.add_module('Function Open2() As String\nDim d As Workbook\n'
                   'Set d = Workbooks.Open(Filename:=ThisWorkbook.Path & "\\Dest.xlsx", UpdateLinks:=0)\n'
                   'Workbooks.Open ThisWorkbook.Path & "\\Alpha.xlsx"\n'
                   'd.BreakLink d.Path & "\\Alpha.xlsx", xlLinkTypeExcelLinks\nEnd Function\n', name='OpenSource')
    app.workbook.path = str(tmp_path)
    with pytest.raises(VBAUnsupportedError, match='open'):
        app.run('OpenSource.Open2')
