"""set_com_class: CreateObject and GetObject make what a host's factory makes. What the model answers itself is replayed from live Excel
(scripts/measure_com_objects.py); the rest holds what the design William approved says."""
from __future__ import annotations

import datetime as _dt
import importlib.util
import json
from decimal import Decimal
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._hooks import HookRequest
from pyopenvba.apps.powerpoint import PowerPointApplication
from pyopenvba.apps.word import WordApplication
from pyopenvba.exceptions import VBARuntimeError, VBASandboxError, VBAStepLimitError, VBAUnsupportedError
from pyopenvba.interpreter._com_hook import ComRequest
from pyopenvba.interpreter._values import VBAArray, error

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_com_objects', ROOT / 'scripts/measure_com_objects.py')
assert _SPEC is not None and _SPEC.loader is not None
MEASURE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(MEASURE)
RECORD = json.loads((ROOT / 'tests/fixtures/com_objects.json').read_text(encoding='utf-8'))
#: Names Excel answered 429 for, where the model cannot know what a machine has: here a host says, through its factory, that there is none.
NONE_HERE = ('No.Such.Class', 'NoDots', '5')
#: A known gap of the binding model, not of the hook: a Set of an object of another class into a typed variable is 13 in Excel, and the model takes
#: it, for its own classes as for this one.
GAPS = ('Set c = CreateObject("Scripting.Dictionary")',)


def _app(*, sandbox: bool = True) -> ExcelApplication:
    app = ExcelApplication(sandbox=sandbox)
    app.add_workbook()
    return app


@pytest.mark.parametrize('declared', ['Object', 'Variant'])
def test_what_the_model_answers_around_the_hook_is_what_excel_did(declared: str) -> None:
    app = _app(sandbox=False)
    for name in NONE_HERE:
        app.set_com_class(name, lambda request: None)
    app.add_module(MEASURE.macro(declared), name='H')
    found = str(app.run('H.Probe')).split(MEASURE.SEPARATOR)
    wanted = RECORD['answers'][declared].split(MEASURE.SEPARATOR)
    assert len(found) == len(wanted)
    assert not [(line[0], excel, model) for line, excel, model in zip(MEASURE.LINES, wanted, found) if excel != model and line[0] not in GAPS]
    assert [(excel, model) for line, excel, model in zip(MEASURE.LINES, wanted, found) if line[0] in GAPS] == [('13=Nothing', '0=Dictionary')]


def test_with_nothing_registered_an_unknown_class_is_refused_as_before() -> None:
    app = _app()
    app.add_module('Public Function Make()\n    On Error Resume Next\n    Set o = CreateObject("No.Such.Class")\nEnd Function\n', name='H')
    with pytest.raises(VBASandboxError):
        app.run('H.Make')
    off = _app(sandbox=False)
    off.add_module('Public Function Make()\n    On Error Resume Next\n    Set o = CreateObject("No.Such.Class")\nEnd Function\n', name='H')
    with pytest.raises(VBAUnsupportedError):
        off.run('H.Make')


class _File:
    def __init__(self, name: str, size: int) -> None:
        self.Name = name
        self.Size = size


class _Files:
    """A collection, as a real FileSystemObject's Files is: For Each walks it into an Object."""

    def __init__(self, *files: _File) -> None:
        self.files = files
        self.Count = len(files)

    def __iter__(self) -> Any:
        return iter(self.files)


class _Folder:
    def __init__(self, path: str) -> None:
        self.Path = path
        self.Files = _Files(_File('a.txt', 3), _File('b.txt', 5))
        self.Names = ['a.txt', 'b.txt']


class _FileSystem:
    """A host's own stand-in, which reaches nothing."""

    vba_type_name = 'FileSystemObject'

    def __init__(self) -> None:
        self.Label = 'start'
        self.calls: list[object] = []

    def GetFolder(self, path: str) -> _Folder:
        return _Folder(path)

    def Echo(self, value: object) -> object:
        self.calls.append(value)
        return value

    def Fail(self) -> None:
        raise OSError('the disk is not there')

    def Refuse(self) -> None:
        raise error(53, 'File not found')

    def __getitem__(self, key: object) -> object:
        return f'item {key}'


def _with_fso(app: ExcelApplication) -> _FileSystem:
    made = _FileSystem()
    app.set_com_class('Scripting.FileSystemObject', lambda request: made)
    return made


def test_objects_a_host_object_answers_are_adapted_at_every_depth() -> None:
    app = _app()
    _with_fso(app)
    app.add_module('''Public Function Walk() As String
    Dim fso As Object, f As Object, s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    s = TypeName(fso) & ";" & fso.GetFolder("C:\\data").Path & ";"
    For Each f In fso.GetFolder("C:\\data").Files
        s = s & f.Name & "=" & f.Size & ","
    Next
    Walk = s & fso.GetFolder("x").Files.Count & ";" & fso("k") & ";" & fso.label
End Function
''', name='H')
    assert app.run('H.Walk') == 'FileSystemObject;C:\\data;a.txt=3,b.txt=5,2;item k;start'


def test_a_list_is_an_array_which_for_each_walks_into_a_variant_only() -> None:
    """VBA's own rule for an array, which a host's list is: a real collection would be an object, as _Files is."""
    app = _app()
    _with_fso(app)
    app.add_module('''Public Function Names() As String
    Dim fso As Object, one As Variant, s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each one In fso.GetFolder("x").Names
        s = s & one & ","
    Next
    Names = s
End Function
''', name='H')
    assert app.run('H.Names') == 'a.txt,b.txt,'


def test_values_cross_as_the_table_says_both_ways() -> None:
    app = _app()
    made = _with_fso(app)
    app.add_module('''Public Function Cross() As String
    Dim fso As Object, a(2 To 3) As Long, back As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    a(2) = 7: a(3) = 8
    back = fso.Echo(CCur(1.5))
    Cross = TypeName(back) & ";" & TypeName(fso.Echo(#1/2/2020 3:04:05 PM#)) & ";" & TypeName(fso.Echo(CLng(4))) & ";" & TypeName(fso.Echo(Empty))
    back = fso.Echo(a)
    Cross = Cross & ";" & LBound(back) & ";" & TypeName(fso.Echo(Null)) & ";" & TypeName(fso.Echo(2.5)) & ";" & TypeName(fso.Echo(True))
End Function
''', name='H')
    assert app.run('H.Cross') == 'Currency;Date;Long;Empty;2;Null;Double;Boolean'
    assert made.calls[0] == Decimal('1.5')
    assert made.calls[1] == _dt.datetime(2020, 1, 2, 15, 4, 5)
    assert made.calls[2] == 4 and type(made.calls[2]) is int
    assert made.calls[3] is None
    assert isinstance(made.calls[4], VBAArray) and made.calls[4].bounds == [(2, 3)]


def test_a_list_comes_back_as_a_zero_based_array_and_a_large_int_as_a_double() -> None:
    app = _app()
    made = _FileSystem()
    made.Echo = lambda value: [1, 'two', 2 ** 40] if value == 'list' else value  # type: ignore[method-assign]
    app.set_com_class('Scripting.FileSystemObject', lambda request: made)
    app.add_module('''Public Function Back() As String
    Dim v As Variant
    v = CreateObject("Scripting.FileSystemObject").Echo("list")
    Back = LBound(v) & ";" & UBound(v) & ";" & TypeName(v(0)) & ";" & TypeName(v(2))
End Function
''', name='H')
    assert app.run('H.Back') == '0;2;Long;Double'


def test_a_member_is_found_case_insensitively_and_a_missing_one_is_438() -> None:
    app = _app()
    made = _with_fso(app)
    app.add_module('''Public Function Members() As String
    Dim fso As Object, e As Long
    Set fso = CreateObject("scripting.filesystemobject")
    fso.LABEL = "changed"
    On Error Resume Next
    fso.NoSuchMember
    e = Err.Number
    Err.Clear
    fso.calls
    Members = fso.label & ";" & e
End Function
''', name='H')
    assert app.run('H.Members') == 'changed;438'
    assert made.Label == 'changed'


def test_nothing_with_an_underscore_is_a_member() -> None:
    app = _app()
    _with_fso(app)
    app.add_module('Public Function Hidden()\n    On Error Resume Next\n    Dim x As Variant\n    x = CreateObject("Scripting.FileSystemObject").[__class__]\n'
                   '    Hidden = Err.Number\nEnd Function\n', name='H')
    assert app.run('H.Hidden') == 438


def test_what_a_host_object_raises_is_440_or_its_own_vba_error() -> None:
    app = _app()
    _with_fso(app)
    app.add_module('''Public Function Errors() As String
    Dim fso As Object, s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    fso.Fail
    s = Err.Number & ";"
    Err.Clear
    fso.Refuse
    Errors = s & Err.Number
End Function
''', name='H')
    assert app.run('H.Errors') == '440;53'


def test_a_factory_that_raises_follows_the_same_rule() -> None:
    app = _app()

    def broken(request: ComRequest) -> object:
        raise RuntimeError('no licence')

    app.set_com_class('Host.Broken', broken)
    app.set_com_class('Host.Refusing', lambda request: (_ for _ in ()).throw(error(70, 'Permission denied')))
    app.set_com_class('Host.Absent', lambda request: None)
    app.add_module('''Public Function Made() As String
    Dim o As Object, s As String
    On Error Resume Next
    Set o = CreateObject("Host.Broken")
    s = Err.Number & ";"
    Err.Clear
    Set o = CreateObject("Host.Refusing")
    s = s & Err.Number & ";"
    Err.Clear
    Set o = CreateObject("Host.Absent")
    Made = s & Err.Number
End Function
''', name='H')
    assert app.run('H.Made') == '440;70;429'
    with pytest.raises(VBARuntimeError) as raised:
        app.add_module('Public Sub Bare()\n    Set o = CreateObject("Host.Broken")\nEnd Sub\n', name='B')
        app.run('B.Bare')
    assert isinstance(raised.value.__cause__, RuntimeError)


def test_a_registered_name_wins_over_our_own_and_none_takes_it_away() -> None:
    app = _app()
    app.set_com_class('Scripting.Dictionary', lambda request: _FileSystem())
    app.add_module('Public Function Kind() As String\n    Kind = TypeName(CreateObject("Scripting.Dictionary"))\nEnd Function\n', name='H')
    assert app.run('H.Kind') == 'FileSystemObject'
    app.set_com_class('Scripting.Dictionary', None)
    assert app.run('H.Kind') == 'Dictionary'


def test_the_request_carries_where_the_call_came_from() -> None:
    app = _app()
    heard: list[ComRequest] = []
    app.set_com_class('Host.Thing', lambda request: heard.append(request) or _FileSystem())
    app.add_module('Public Sub Main()\n    Helper\nEnd Sub\nPublic Sub Helper()\n    Set o = CreateObject("Host.Thing", "")\n    Set o = GetObject(, "host.thing")\nEnd Sub\n', name='M')
    app.run('M.Main')
    first, second = heard
    assert (first.prog_id, first.kind, first.server, first.path, first.caller, first.entry) == ('Host.Thing', 'create', '', None, 'M.Helper', 'M.Main')
    assert (second.prog_id, second.kind) == ('host.thing', 'get')
    assert first.document is app.application.vba_get('ThisWorkbook')


def test_set_hook_requests_carry_caller_and_entry_too() -> None:
    app = _app()
    heard: list[HookRequest] = []
    app.set_hook('PrintOut', heard.append)
    app.add_module('Public Sub Main()\n    Helper\nEnd Sub\nPublic Sub Helper()\n    ActiveWorkbook.Worksheets(1).PrintOut Copies:=1\nEnd Sub\n', name='M')
    app.run('M.Main')
    assert (heard[0].caller, heard[0].entry) == ('M.Helper', 'M.Main')


def test_a_host_object_that_runs_vba_draws_on_the_run_budget() -> None:
    app = ExcelApplication(step_limit=1000)
    app.add_workbook()
    app.add_module('Public Sub Spin()\n    Dim i As Long\n    For i = 1 To 600\n    Next\nEnd Sub\n'
                   'Public Sub Main()\n    Spin\n    CreateObject("Host.Calls").Back\nEnd Sub\n', name='H')

    class _Calls:
        def Back(self) -> None:
            app.run('H.Spin')

    app.set_com_class('Host.Calls', lambda request: _Calls())
    with pytest.raises(VBAStepLimitError):
        app.run('H.Main')


def test_withevents_on_a_host_object_is_refused() -> None:
    app = _app()
    _with_fso(app)
    app.add_module('Private WithEvents fso As Scripting.FileSystemObject\nPublic Sub Hook()\n    Set fso = CreateObject("Scripting.FileSystemObject")\nEnd Sub\n',
                   name='Listener', kind='class')
    app.add_module('Public Sub Main()\n    Dim l As New Listener\n    l.Hook\nEnd Sub\n', name='H')
    with pytest.raises(VBAUnsupportedError, match='WithEvents'):
        app.run('H.Main')


def test_the_sandbox_lets_a_registered_class_through_and_gates_getobject_files(tmp_path: Path) -> None:
    app = _app()
    heard: list[ComRequest] = []
    app.set_com_class('Host.Doc', lambda request: heard.append(request) or _FileSystem())
    app.add_module('Public Sub Open(ByVal path As String)\n    Set o = GetObject(path, "Host.Doc")\nEnd Sub\n', name='H')
    with pytest.raises(VBASandboxError):
        app.run('H.Open', str(tmp_path / 'a.doc'))
    assert not heard
    allowed = tmp_path / 'in'
    allowed.mkdir()
    app.set_com_class('Host.Doc', lambda request: heard.append(request) or _FileSystem(), folders=[allowed])
    with pytest.raises(VBASandboxError, match='outside the folders'):
        app.run('H.Open', str(tmp_path / 'a.doc'))
    app.run('H.Open', str(allowed / 'a.doc'))
    assert heard[0].path == (allowed / 'a.doc').resolve() and heard[0].kind == 'get'


def test_a_backslash_cannot_carry_a_getobject_path_out_of_its_grant(tmp_path: Path) -> None:
    """Off Windows a backslash is a character of a name, so folder/..\\away\\x.xlsx checked as given is one file in the folder, where opened it is
    outside it. The gate checks the path as it will be opened, on every system (Agent 03's finding on #574)."""
    app = _app()
    heard: list[ComRequest] = []
    allowed = tmp_path / 'grant'
    allowed.mkdir()
    (tmp_path / 'away').mkdir()
    app.set_com_class('Host.Doc', lambda request: heard.append(request) or _FileSystem(), folders=[allowed])
    app.add_module('Public Sub Open(ByVal path As String)\n    Set o = GetObject(path, "Host.Doc")\nEnd Sub\n', name='H')
    with pytest.raises(VBASandboxError):
        app.run('H.Open', str(allowed) + '/..\\away\\x.xlsx')
    assert not heard
    app.run('H.Open', str(allowed) + '\\inner\\..\\x.xlsx')
    assert heard[0].path is not None and heard[0].path.parent == allowed.resolve()


def test_getobject_of_a_file_for_a_class_no_host_supplies_is_refused(tmp_path: Path) -> None:
    app = _app(sandbox=False)
    app.add_module('Public Sub Open(ByVal path As String)\n    Set o = GetObject(path, "No.Host.Class")\nEnd Sub\n', name='H')
    with pytest.raises(VBAUnsupportedError):
        app.run('H.Open', str(tmp_path / 'a.doc'))


@pytest.mark.parametrize('name', ['', ' x', 'x ', 5, None])
def test_a_progid_must_be_a_plain_name(name: object) -> None:
    with pytest.raises(ValueError, match='ProgID'):
        _app().set_com_class(name, lambda request: None)  # type: ignore[arg-type]


def test_a_factory_must_be_callable() -> None:
    with pytest.raises(TypeError):
        _app().set_com_class('Host.Thing', 'not a function')  # pyright: ignore[reportArgumentType]


@pytest.mark.parametrize('make', [WordApplication, PowerPointApplication])
def test_word_and_powerpoint_take_classes_too(make: type) -> None:
    app = make()
    app.set_com_class('Host.Thing', lambda request: _FileSystem())  # pyright: ignore[reportUnknownLambdaType]
    app.add_module('Public Function Kind() As String\n    Kind = TypeName(CreateObject("Host.Thing"))\nEnd Function\n', name='H')
    assert app.run('H.Kind') == 'FileSystemObject'
