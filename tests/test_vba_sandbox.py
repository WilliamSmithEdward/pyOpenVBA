"""The sandbox: with it on, running VBA cannot reach outside the engine, and Python's own calls still can."""
from __future__ import annotations

import ast
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.powerpoint import PowerPointApplication
from pyopenvba.apps.word import WordApplication
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter import Interpreter, VBASandboxError

SOURCE = Path(__file__).parent.parent / 'src/pyopenvba'


def excel(statements: str, folder: Path, *, sandbox: bool = True) -> ExcelApplication:
    app = ExcelApplication(sandbox=sandbox)
    app.add_workbook()
    app.add_module(f'Public Sub Go(folder)\n{statements}\nEnd Sub\n', name='Probe')
    return app


#: Each reaches outside the engine; the last group has no shim yet, so with the sandbox off it is still refused.
WRITES_OR_READS = [
    'ActiveWorkbook.SaveAs folder & "/out.xlsx"',
    'ActiveWorkbook.SaveCopyAs folder & "/out.xlsx"',
    'Workbooks.Open folder & "/there.xlsx"',
    'Workbooks.Open folder & "/not there.xlsx"',
    # A backslash or another case reaches nothing the plain path could not: off Windows the path is turned into the disk's own only past the gate.
    'Workbooks.Open folder & "\\there.xlsx"',
    'Workbooks.Open UCase(folder) & "/THERE.XLSX"',
    'Workbooks.Open "..\\" & folder & "\\there.xlsx"',
    'ActiveWorkbook.SaveAs folder & "\\out.xlsx"',
    'ActiveWorkbook.SaveCopyAs folder & "\\out.xlsx"',
    'Workbooks.OpenText UCase(folder) & "\\THERE.CSV"',
    'Workbooks.OpenText folder & "/there.csv"',
    'ActiveSheet.SetBackgroundPicture folder & "/there.png"',
    'ActiveWorkbook.ApplyTheme folder & "/there.png"',
    'ActiveSheet.PageSetup.LeftHeaderPicture.Filename = folder & "/there.png"',
    'ActiveWorkbook.UpdateLink folder & "/there.xlsx"',
    'ActiveWorkbook.ChangeLink "there.xlsx", folder & "/there.xlsx"',
    'ActiveWorkbook.OpenLinks folder & "/there.xlsx"',
    'ActiveWorkbook.Connections.AddFromFile folder & "/there.png"',
    'Workbooks.CheckOut folder & "/there.xlsx"',
    'Workbooks.OpenXML folder & "/there.xlsx"',
    'Workbooks.OpenDatabase folder & "/there.xlsx"',
    'b = ActiveWorkbook.CustomXMLParts.Add().Load(folder & "/there.csv")',
]
NOT_YET = [
    'x = Shell("cmd")',
    'x = GetSetting("a", "b", "c")',
    'Set x = GetObject("", "Excel.Application")',
    'Set x = CreateObject("Shell.Application")',
]


@pytest.fixture
def folder(tmp_path: Path) -> Path:
    for name in ('there.xlsx', 'there.csv', 'there.png'):
        (tmp_path / name).write_bytes(b'')
    return tmp_path


@pytest.mark.parametrize('statement', WRITES_OR_READS + NOT_YET)
def test_the_sandbox_refuses_it_and_touches_nothing(statement: str, folder: Path) -> None:
    before = sorted(path.name for path in folder.iterdir())
    app = excel(f'    {statement}', folder)
    with pytest.raises(VBASandboxError, match='not supported with sandbox on'):
        app.run('Probe.Go', str(folder))
    assert sorted(path.name for path in folder.iterdir()) == before


@pytest.mark.parametrize('statement', NOT_YET)
def test_without_the_sandbox_a_gap_is_still_only_a_gap(statement: str, folder: Path) -> None:
    app = excel(f'    {statement}', folder, sandbox=False)
    with pytest.raises(VBAUnsupportedError) as raised:
        app.run('Probe.Go', str(folder))
    assert not isinstance(raised.value, VBASandboxError)


def test_without_the_sandbox_vba_saves_a_file(folder: Path) -> None:
    excel('    ActiveWorkbook.SaveAs folder & "/out.xlsx"', folder, sandbox=False).run('Probe.Go', str(folder))
    assert (folder / 'out.xlsx').stat().st_size > 0


def test_an_open_saved_workbook_cannot_save_or_read_its_neighbours(folder: Path) -> None:
    saved = ExcelApplication()
    saved.add_workbook()
    saved.save(folder / 'book.xlsx')
    app = ExcelApplication.open(folder / 'book.xlsx')
    app.add_module('Public Sub Go()\n    ActiveWorkbook.Save\nEnd Sub\n', name='Probe')
    stamp = (folder / 'book.xlsx').read_bytes()
    with pytest.raises(VBASandboxError):
        app.run('Probe.Go')
    assert (folder / 'book.xlsx').read_bytes() == stamp


def test_vba_cannot_trap_it(folder: Path) -> None:
    app = excel('    On Error Resume Next\n    ActiveWorkbook.SaveAs folder & "/out.xlsx"\n    Range("A1").Value = 1',
                folder)
    with pytest.raises(VBASandboxError):
        app.run('Probe.Go', str(folder))
    assert not (folder / 'out.xlsx').exists()


def test_it_cannot_be_reached_through_an_event(folder: Path) -> None:
    """A macro the file itself starts is running VBA too."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go(folder)\n    ActiveWorkbook.SaveCopyAs folder & "/out.xlsx"\nEnd Sub\n', name='Probe')
    app.add_module('Public Sub Outer(folder)\n    Application.Run "Probe.Go", folder\nEnd Sub\n', name='Starter')
    with pytest.raises(VBASandboxError):
        app.run('Starter.Outer', str(folder))
    assert not (folder / 'out.xlsx').exists()


def test_python_is_not_vba(folder: Path) -> None:
    """Python's own calls work with the sandbox on, also once a macro has been refused."""
    app = excel('    ActiveWorkbook.SaveAs folder & "/out.xlsx"', folder)
    app.save(folder / 'before.xlsx')
    with pytest.raises(VBASandboxError):
        app.run('Probe.Go', str(folder))
    app.save(folder / 'after.xlsx')
    assert ExcelApplication.open(folder / 'after.xlsx').workbook is not None


def test_word_keeps_its_vba_in() -> None:
    app = WordApplication()
    app.add_document()
    app.add_module('Public Sub Go()\n    ActiveDocument.SaveAs "out.docx"\nEnd Sub\n', name='Probe')
    with pytest.raises(VBASandboxError):
        app.run('Probe.Go')
    app = WordApplication()
    app.add_document()
    app.add_module('Public Sub Go()\n    Documents.Open "not there.docx"\nEnd Sub\n', name='Probe')
    with pytest.raises(VBASandboxError):
        app.run('Probe.Go')


def test_powerpoint_keeps_its_vba_in() -> None:
    app = PowerPointApplication()
    app.add_presentation()
    app.add_module('Public Sub Go()\n    ActivePresentation.SaveAs "out.pptx"\nEnd Sub\n', name='Probe')
    with pytest.raises(VBASandboxError):
        app.run('Probe.Go')
    app = PowerPointApplication()
    app.add_presentation()
    app.add_module('Public Sub Go()\n    Presentations.Open "not there.pptx"\nEnd Sub\n', name='Probe')
    with pytest.raises(VBASandboxError):
        app.run('Probe.Go')


def test_the_flag_defaults_on() -> None:
    assert Interpreter().sandbox
    assert ExcelApplication().interpreter.sandbox
    assert WordApplication().interpreter.sandbox
    assert PowerPointApplication().interpreter.sandbox
    assert not ExcelApplication(sandbox=False).interpreter.sandbox


def test_a_declare_is_refused() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long\n'
                   'Public Sub Go()\n    x = GetTickCount()\nEnd Sub\n', name='Probe')
    with pytest.raises(VBASandboxError):
        app.run('Probe.Go')


#: The calls that reach the file system.
_REACHING = {'create_connection', 'open', 'read_bytes', 'write_bytes', 'read_text', 'write_text', 'is_file', 'is_dir',
             'exists', 'mkdir', 'unlink', 'rmdir', 'stat', 'ZipFile', 'urlopen', 'socket', 'Popen', 'system', 'getenv'}
#: Functions that reach out behind a gate in every caller VBA can get to.
_BEHIND_A_GATE = {
    ('apps/excel/_vbide.py', 'read_source_file'),  # VBComponents.Import, CodeModule.AddFromFile
    ('apps/excel/_workbook_links.py', '_sheet_names'),  # status
    ('interpreter/_files.py', 'flush'),  # a file that Open let through
    ('interpreter/_files.py', 'spelled'),  # full, which every caller reaches through its gate
    ('interpreter/_fso.py', '_flush'),  # a stream that Open let through
    ('interpreter/_fso.py', 'DriveExists'),  # _table, which runs it first
    ('interpreter/_fso.py', '_actual'),  # GetFile, GetFolder, GetSpecialFolder
    ('interpreter/_fso.py', '_listed'),  # members that have passed it
    ('interpreter/_fso.py', '_exists'),  # _check, which runs it first
    ('interpreter/_files.py', 'attributes_of'),  # Open, Kill, FileCopy, GetAttr
    ('interpreter/_files.py', 'read_only'),  # Open, Kill, FileCopy
    ('apps/excel/_vba_save.py', '_open_zip'),  # the bytes already in memory
    ('apps/excel/_output.py', '_check_writable'),  # _hooks.gate, which every caller runs first
    ('apps/excel/_output.py', '_check_export_path'),  # _hooks.gate or Interpreter.outside, which export_chart runs first
    ('apps/excel/_hooks.py', 'saved_bytes'),  # a temporary file of its own, for the caller's function
}


def _reaches(function: ast.AST) -> list[str]:
    found: list[str] = []
    for node in ast.walk(function):
        if isinstance(node, ast.Call):
            name = node.func.attr if isinstance(node.func, ast.Attribute) else getattr(node.func, 'id', '')
            if name in _REACHING:
                found.append(name)
    return found


def _gated(function: ast.AST) -> bool:
    """Whether the function goes through the gate: Interpreter.outside, or _hooks.link_gate, which is that gate unless the host
    granted folders for linked files and then checks the path against them."""
    return any(isinstance(node, ast.Attribute) and node.attr == 'outside'
               or isinstance(node, ast.Name) and node.id == 'link_gate' for node in ast.walk(function))


def test_nothing_reaches_outside_past_the_gate() -> None:
    """A new way out of the engine has to go through Interpreter.outside: the applications' __init__ modules are the
    Python API, and everything else a macro could get to is checked here."""
    ungated: list[str] = []
    for folder in ('apps/excel', 'apps/word', 'apps/powerpoint', 'interpreter'):
        for path in sorted((SOURCE / folder).glob('*.py')):
            if path.name == '__init__.py':
                continue
            relative = path.relative_to(SOURCE).as_posix()
            tree = ast.parse(path.read_text(encoding='utf-8'))
            functions = [node for node in ast.walk(tree) if isinstance(node, ast.FunctionDef)]
            for function in functions:
                if _reaches(function) and not _gated(function) and (relative, function.name) not in _BEHIND_A_GATE:
                    ungated.append(f'{relative}: {function.name} calls {_reaches(function)}')
    assert not ungated
