"""FillFormat as live Excel answered six held-out rounds (scripts/measure_fill_format_round.py): the members the model has, the texture
members on fills with no texture, PictureEffects, and pictures and textures from files. Each case runs alone in the model,
typed and late, read by read, then saved, opened again and read once more, and its saved spPr is compared. A read or case the model
refuses is allowed; one it answers must be Excel's. Round 5 was measured at StandardHeight 14.5 (144 DPI) and its spPr
matched the model's 96-DPI answer; the rounds before it did not record StandardHeight. Round 6 compares the reads of the book opened again only."""
from __future__ import annotations

import importlib.util
import json
import tempfile
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBASandboxError, VBAUnsupportedError
from pyopenvba.shapes._read_look import read_look

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_fill_format_round', ROOT / 'scripts/measure_fill_format_round.py')
assert _SPEC is not None and _SPEC.loader is not None
PROBE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(PROBE)


@pytest.mark.parametrize('number', [1, 2, 3, 4, 5])
def test_each_case_answers_as_excel_did(number: int, monkeypatch: pytest.MonkeyPatch) -> None:
    record: dict[str, Any] = json.loads((ROOT / f'tests/fixtures/fill_format_round_{number}.json').read_text(encoding='utf-8'))
    monkeypatch.setattr(PROBE, 'ROUND', number)
    monkeypatch.setattr(PROBE, 'READS', record['reads'])
    names = ['set', *record['reads'], *(f'reopened {one}' for one in record['reads'])]
    differ: list[tuple[Any, ...]] = []
    with tempfile.TemporaryDirectory() as folder:
        PROBE.write_files(Path(folder))
        for index, (case, excel) in enumerate(zip(record['cases'], record['excel'])):
            got = PROBE.alone(case, index, Path(folder))
            if got[0] == 'refused':
                continue
            differ += [(case, name, mine, theirs) for name, mine, theirs in zip(names, got[0].split('|'), excel[0].split('|'))
                       if mine not in (theirs, 'refused')]
            if got[1] != excel[1]:
                differ.append((case, 'spPr', got[1], excel[1]))
    assert differ == []


def test_a_picture_fill_opened_again_answers_as_excel_did(monkeypatch: pytest.MonkeyPatch) -> None:
    """Round 6: the reads of the book opened again only, which hold nothing the display's DPI changes."""
    record: dict[str, Any] = json.loads((ROOT / 'tests/fixtures/fill_format_round_6.json').read_text(encoding='utf-8'))
    monkeypatch.setattr(PROBE, 'ROUND', 6)
    monkeypatch.setattr(PROBE, 'READS', record['reads'])
    differ: list[tuple[Any, ...]] = []
    with tempfile.TemporaryDirectory() as folder:
        PROBE.write_files(Path(folder))
        for index, (case, excel) in enumerate(zip(record['cases'], record['excel'])):
            got = PROBE.alone(case, index, Path(folder))
            if got[0] == 'refused':
                continue
            differ += [(case, name, mine, theirs) for name, mine, theirs in zip(record['reads'], PROBE.reopened(got[0]), PROBE.reopened(excel[0]))
                       if mine not in (theirs, 'refused')]
    assert differ == []


def _picture_filled_book(member: str, folder: Path) -> ExcelApplication:
    """A book saved with a shape S whose fill is p.png by ``member``, and a module H whose Q(folder, change) opens it again, makes
    ``change`` to S's fill, if any, and answers its Type, and whose R saves it again as g.xlsx."""
    (folder / 'p.png').write_bytes(PROBE.files()['p.png'])
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Sub P(ByVal folder As String)\n    Dim wb As Object\n    Set wb = Workbooks.Add\n'
                   '    wb.Worksheets(1).Shapes.AddShape(msoShapeOval, 20, 20, 120, 60).Name = "S"\n'
                   f'    wb.Worksheets(1).Shapes("S").Fill.{member} folder & Application.PathSeparator & "p.png"\n'
                   '    wb.SaveAs folder & Application.PathSeparator & "f.xlsx", xlOpenXMLWorkbook\n    wb.Close False\nEnd Sub\n'
                   'Function Q(ByVal folder As String, ByVal change As String) As String\n    Dim wb As Object\n'
                   '    Set wb = Workbooks.Open(folder & Application.PathSeparator & "f.xlsx")\n'
                   '    If change = "Solid" Then wb.Worksheets(1).Shapes("S").Fill.Solid\n'
                   '    If change = "Line" Then wb.Worksheets(1).Shapes("S").Line.Weight = 3\n'
                   '    Q = CStr(wb.Worksheets(1).Shapes("S").Fill.Type)\nEnd Function\n'
                   'Sub R(ByVal folder As String)\n    ActiveWorkbook.SaveAs folder & Application.PathSeparator & "g.xlsx", xlOpenXMLWorkbook\nEnd Sub\n',
                   name='H')
    app.run('H.P', str(folder))
    return app


@pytest.mark.parametrize(('member', 'kind'), [('UserPicture', '6'), ('UserTextured', '4')])
def test_a_picture_fill_in_a_book_opened_again_reads_as_before(member: str, kind: str, tmp_path: Path) -> None:
    """Excel answers a picture fill opened again as it did before the save (rounds 2 to 4); a save of the book unchanged keeps it."""
    app = _picture_filled_book(member, tmp_path)
    assert app.run('H.Q', str(tmp_path), '') == kind
    app.run('H.R', str(tmp_path))
    with zipfile.ZipFile(tmp_path / 'f.xlsx') as first, zipfile.ZipFile(tmp_path / 'g.xlsx') as second:
        assert second.read('xl/drawings/drawing1.xml') == first.read('xl/drawings/drawing1.xml')
        assert second.read('xl/media/image1.png') == first.read('xl/media/image1.png') == PROBE.files()['p.png']


@pytest.mark.parametrize('change', ['Solid', 'Line'])
def test_a_change_to_a_picture_fill_from_a_file_is_refused(change: str, tmp_path: Path) -> None:
    """Only reads of a picture fill opened again are measured, so a change to its fill or line reports itself."""
    app = _picture_filled_book('UserPicture', tmp_path)
    with pytest.raises(VBAUnsupportedError, match='with a picture fill from a file'):
        app.run('H.Q', str(tmp_path), change)


@pytest.mark.parametrize('blip', ['<a:blip r:embed="rId1"><a:alphaModFix amt="70000"/></a:blip><a:stretch><a:fillRect/></a:stretch>',
                                  '<a:blip r:embed="rId1"/><a:srcRect l="10"/><a:stretch><a:fillRect/></a:stretch>',
                                  '<a:blip r:link="rId1"/><a:stretch><a:fillRect/></a:stretch>'])
def test_a_picture_fill_excel_s_user_picture_does_not_write_is_refused(blip: str) -> None:
    """A picture fill with effects, a crop or a linked image is not read, as nothing has measured what Excel answers of it."""
    with pytest.raises(VBAUnsupportedError, match="blipFill Excel's UserPicture or UserTextured does not write"):
        read_look(f'<xdr:spPr><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:blipFill>{blip}</a:blipFill></xdr:spPr>')


def test_a_picture_from_a_file_is_refused_under_the_sandbox(tmp_path: Path) -> None:
    (tmp_path / 'p.png').write_bytes(PROBE.files()['p.png'])
    for member in ('UserPicture', 'UserTextured'):
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(f'Sub P(ByVal path As String)\n    ActiveSheet.Shapes.AddShape(msoShapeOval, 20, 20, 120, 60).Fill.{member} path\nEnd Sub\n',
                       name='H')
        with pytest.raises(VBASandboxError):
            app.run('H.P', str(tmp_path / 'p.png'))
