"""Chart.Export and CopyPicture, replayed from live Excel (scripts/measure_output_pictures.py, tests/fixtures/output_pictures.json) with a small
fake for the function each hands its work to, and the hook's own rules: no function, the sandbox, and the clipboard after."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._hooks import HookRequest
from pyopenvba.exceptions import VBASandboxError, VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'output_pictures.json').read_text(encoding='utf-8'))
CASES: list[dict[str, Any]] = RECORD['cases']
_spec = importlib.util.spec_from_file_location('measure_output_pictures', Path(__file__).parents[1] / 'scripts/measure_output_pictures.py')
assert _spec is not None and _spec.loader is not None
_script = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_script)
OWNERS: dict[str, tuple[str, str]] = getattr(_script, 'OWNERS')
STATES: dict[str, tuple[str, str, str]] = getattr(_script, 'STATES')
row_source: Any = getattr(_script, 'row_source')
#: What the fake function writes for each kind Excel draws, so a file's kind shows as Excel's did.
MAGIC = {'PNG': b'\x89PNG', 'GIF': b'GIF8', 'JPG': b'\xff\xd8', 'JPEG': b'\xff\xd8', 'BMP': b'BM'}
KINDS = ((b'\x89PNG', 'png'), (b'GIF8', 'gif'), (b'\xff\xd8', 'jpeg'), (b'BM', 'bmp'))


def _drawn(request: HookRequest) -> None:
    if request.member == 'Export' and request.path is not None:
        kind = str(request.arguments.get('FilterName', request.path.suffix[1:])).upper()
        request.path.write_bytes(MAGIC[kind] + b'....')


def _kind(data: bytes) -> str:
    return next((kind for magic, kind in KINDS if data.startswith(magic)), 'before' if data == b'before' else f'other:{data[:4].hex()}')


def _folder(tmp_path: Path) -> tuple[Path, Path]:
    book = tmp_path / 'book.xlsx'
    book.write_bytes((FIXTURES / 'output_pictures.xlsx').read_bytes())
    folder = tmp_path / 'run'
    (folder / 'sub').mkdir(parents=True)
    (folder / 'there.png').write_bytes(b'before')
    (folder / 'ro.png').write_bytes(b'before')
    (folder / 'ro.png').chmod(0o444)
    return book, folder


def _row(case: dict[str, Any]) -> str:
    get, declared = OWNERS[case['owner']] if case['member'] != 'state' else (STATES[case['row']][1], case['owner'])
    return str(row_source(case['call'], declared if case['binding'] == 'Typed' else 'Object', get, live=False))


@pytest.mark.parametrize('case', CASES, ids=[f'{one["member"]} {one["owner"]} {one["binding"]} {one["row"]}' for one in CASES])
def test_each_call_answers_as_excel_did(case: dict[str, Any], tmp_path: Path) -> None:
    """Excel's answer to the paste after it, which the model does not make, is left out; what was written is compared by name, and by kind for a chart sheet.
    A chart in a frame Excel has not drawn since its workbook opened exports an empty file, which the model, drawing nothing, does not follow."""
    book, folder = _folder(tmp_path)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.set_hook('Export', _drawn)
    app.set_hook('CopyPicture', _drawn)
    app.add_module(_row(case), name='Probe')
    try:
        got = str(app.run('Probe.Row', str(book), str(folder))).replace(str(folder), '<DIR>')
    finally:
        (folder / 'ro.png').chmod(0o666)
    assert got == case['excel'].split('|paste=')[0]
    written = sorted(f'{str(one.relative_to(folder)).replace(chr(92), "/")} {_kind(one.read_bytes())}' for one in folder.rglob('*')
                     if one.is_file() and one.name != 'ro.png' and not (one.name == 'there.png' and one.read_bytes() == b'before'))
    if 'Frame' in case['owner'] or 'ChartObjects' in case['call'] or case['row'] == 'export_frame_hidden_sheet':
        assert [one.split(' ')[0] for one in written] == [one.split(' ')[0] for one in case['files']]
    else:
        assert written == case['files']


def _book(tmp_path: Path, *, sandbox: bool = False) -> tuple[ExcelApplication, Path]:
    book, folder = _folder(tmp_path)
    (folder / 'ro.png').chmod(0o666)
    app = ExcelApplication(sandbox=sandbox)
    app.add_workbook()
    app.open_workbook(book, with_vba=False)
    return app, folder


def _macro(app: ExcelApplication, body: str, folder: Path) -> object:
    app.add_module(f'Public Function Go(ByVal folder As String) As Variant\n    Dim wb As Workbook\n    Set wb = Workbooks("book.xlsx")\n{body}\nEnd Function\n', name='Go')
    return app.run('Go.Go', str(folder))


def test_with_no_function_a_drawing_reports_itself_and_nothing_is_written(tmp_path: Path) -> None:
    app, folder = _book(tmp_path)
    for body in ('    wb.Charts("Plot").Export folder & "\\a.png"', '    wb.Worksheets("One").Range("A1:B3").CopyPicture'):
        with pytest.raises(VBAUnsupportedError, match='set_hook'):
            _macro(app, body, folder)
    assert not (folder / 'a.png').exists()


def test_a_kind_excel_does_not_draw_is_the_empty_file_excel_writes_with_no_function(tmp_path: Path) -> None:
    app, folder = _book(tmp_path)
    assert _macro(app, '    Go = wb.Charts("Plot").Export(folder & "\\a.tif")', folder) is True
    assert (folder / 'a.tif').read_bytes() == b''


def test_the_function_hears_what_excel_was_given(tmp_path: Path) -> None:
    app, folder = _book(tmp_path)
    heard: list[HookRequest] = []
    app.set_hook('Export', heard.append)
    app.set_hook('CopyPicture', heard.append)
    _macro(app, '    wb.Charts("Plot").Export folder & "\\a.gif", FilterName:="GIF"\n    wb.Worksheets("One").Shapes("Box").CopyPicture xlPrinter, xlPicture', folder)
    export, picture = heard
    assert (export.member, export.owner, export.path, dict(export.arguments)) == ('Export', 'Chart', folder / 'a.gif', {'Filename': str(folder / 'a.gif'), 'FilterName': 'GIF'})
    assert (picture.member, picture.owner, picture.path, dict(picture.arguments)) == ('CopyPicture', 'Shape', None, {'Appearance': 2, 'Format': -4147})


def test_under_the_sandbox_an_export_outside_the_folders_never_reaches_the_function(tmp_path: Path) -> None:
    app, folder = _book(tmp_path, sandbox=True)
    heard: list[HookRequest] = []
    app.set_hook('Export', heard.append, folders=[tmp_path / 'elsewhere'])
    with pytest.raises(VBASandboxError):
        _macro(app, '    wb.Charts("Plot").Export folder & "\\a.png"', folder)
    assert not heard
    app.set_hook('Export', heard.append, folders=[folder])
    _macro(app, '    wb.Charts("Plot").Export folder & "\\b.png"', folder)
    assert heard[0].path == folder / 'b.png' or heard[0].path == (folder / 'b.png').resolve()


def test_under_the_sandbox_the_empty_file_is_gated_too(tmp_path: Path) -> None:
    app, folder = _book(tmp_path, sandbox=True)
    with pytest.raises(VBASandboxError):
        _macro(app, '    wb.Charts("Plot").Export folder & "\\a.tif"', folder)
    assert not (folder / 'a.tif').exists()


def test_a_picture_runs_its_function_under_the_sandbox_as_it_writes_no_file(tmp_path: Path) -> None:
    app, folder = _book(tmp_path, sandbox=True)
    heard: list[HookRequest] = []
    app.set_hook('CopyPicture', heard.append)
    _macro(app, '    wb.Worksheets("One").Range("A1:B3").CopyPicture', folder)
    assert heard[0].owner == 'Range'


def test_a_hidden_shape_copies_nothing_and_its_function_is_not_called(tmp_path: Path) -> None:
    app, folder = _book(tmp_path)
    heard: list[HookRequest] = []
    app.set_hook('CopyPicture', heard.append)
    # Cells copied first are gone from the clipboard after, so a paste fails as Excel's did (picture_shape_hidden).
    got = _macro(app, '    wb.Worksheets("One").Range("A1").Copy\n    wb.Worksheets("One").Shapes("Box").Visible = msoFalse\n'
                      '    wb.Worksheets("One").Shapes("Box").CopyPicture\n    wb.Worksheets("Two").Activate\n    wb.Worksheets("Two").Range("A1").Select\n'
                      '    On Error Resume Next\n    wb.Worksheets("Two").Paste\n    Go = Err.Number & "|" & Application.CutCopyMode', folder)
    assert str(got) == '1004|0'
    assert not heard


def test_a_paste_of_the_picture_reports_itself_until_cells_are_copied(tmp_path: Path) -> None:
    app, folder = _book(tmp_path)
    app.set_hook('CopyPicture', _drawn)
    with pytest.raises(VBAUnsupportedError, match='picture'):
        _macro(app, '    wb.Worksheets("One").Range("A1:B3").CopyPicture\n    wb.Worksheets("Two").Activate\n    wb.Worksheets("Two").Paste wb.Worksheets("Two").Range("D1")', folder)
    got = _macro(app, '    wb.Worksheets("One").Range("A1:B3").CopyPicture\n    wb.Worksheets("One").Range("A2").Copy\n'
                      '    wb.Worksheets("Two").Paste wb.Worksheets("Two").Range("D1")\n    Go = wb.Worksheets("Two").Range("D1").Value & "|" & Application.CutCopyMode', folder)
    assert str(got) == '2|1'
