"""The host's step limit: a run of VBA that takes more steps than allowed stops with VBAStepLimitError, which On Error cannot trap, whatever the run passed through."""
from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from pyopenvba import ExcelFile
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.powerpoint import PowerPointApplication
from pyopenvba.apps.word import WordApplication
from pyopenvba.exceptions import VBAStepLimitError

SPIN = '''Public Done As Long

Public Function Spin(ByVal times As Long) As Long
    Dim i As Long
    For i = 1 To times
        Done = Done + 1
    Next
    Spin = Done
End Function

Public Sub Forever()
    Range("A1").Value = "before"
    On Error Resume Next
    Do
    Loop
End Sub

Public Function HowMany() As Long
    HowMany = Done
End Function

Public Function FirstCell() As Variant
    FirstCell = Range("A1").Value
End Function

Public Function ErrNumber() As Long
    ErrNumber = Err.Number
End Function
'''


def _app(limit: int | None) -> ExcelApplication:
    app = ExcelApplication(step_limit=limit)
    app.add_workbook()
    app.add_module(SPIN, name='H')
    return app


def test_there_is_no_limit_unless_the_host_sets_one() -> None:
    app = _app(None)
    assert app.interpreter.step_limit is None
    assert app.run('H.Spin', 100_000) == 100_000


def test_an_endless_loop_stops_and_on_error_cannot_trap_it() -> None:
    app = _app(1000)
    with pytest.raises(VBAStepLimitError, match='step_limit=1000'):
        app.run('H.Forever')
    # What ran before it is kept, not rolled back, and it is no run-time error, so Err was not set.
    assert app.run('H.FirstCell') == 'before'
    assert app.run('H.ErrNumber') == 0


def test_the_next_run_starts_with_a_fresh_budget() -> None:
    app = _app(1000)
    with pytest.raises(VBAStepLimitError):
        app.run('H.Forever')
    assert app.run('H.Spin', 100) == 100
    assert app.run('H.Spin', 100) == 200


def test_a_run_inside_the_limit_is_untouched_and_one_past_it_is_not() -> None:
    app = _app(1000)
    assert app.run('H.Spin', 200) == 200
    with pytest.raises(VBAStepLimitError):
        app.run('H.Spin', 2000)


def test_class_code_and_evaluate_draw_on_the_run_budget() -> None:
    app = _app(1000)
    app.add_module('Public Function Work(ByVal times As Long) As Long\n    Work = Spin(times)\nEnd Function\n', name='Worker', kind='class')
    app.add_module('''Public Function ByClass() As Long
    Dim w As New Worker
    ByClass = w.Work(300) + w.Work(300) + w.Work(300)
End Function
Public Function ByEvaluate() As Long
    ByEvaluate = Spin(300) + Application.Evaluate("Spin(300)") + Spin(300)
End Function
''', name='Callers')
    with pytest.raises(VBAStepLimitError):
        app.run('Callers.ByClass')
    with pytest.raises(VBAStepLimitError):
        app.run('Callers.ByEvaluate')
    # Each alone was well inside it.
    assert app.run('H.Spin', 300) == app.run('H.HowMany')


def test_a_python_callback_that_runs_vba_inside_a_run_draws_on_the_same_budget() -> None:
    """What a host hands VBA to call back, here a callback DoEvents pumps, is part of the run it was called in."""
    app = _app(1000)
    app.add_module('Public Sub Pumping()\n    Spin 300\n    DoEvents\n    Spin 300\nEnd Sub\n', name='P')

    def callback() -> None:
        app.run('H.Spin', 600)

    app.interpreter.pumped.append(callback)
    with pytest.raises(VBAStepLimitError):
        app.run('P.Pumping')


def test_an_event_macro_draws_on_the_run_that_raised_it(tmp_path: Path) -> None:
    path = tmp_path / 'events.xlsm'
    shutil.copy(Path(__file__).parent / 'fixtures' / 'shapes' / 'excel_shapes.xlsm', path)
    with ExcelFile(path) as host:
        host.set_module('Sheet1', 'Private Sub Worksheet_Change(ByVal Target As Range)\n    Dim i As Long\n    For i = 1 To 600\n    Next\nEnd Sub\n')
        host.set_module('Module1', 'Public Sub Change()\n    Dim i As Long\n    For i = 1 To 600\n    Next\n    Sheet1.Range("B2").Value = 1\nEnd Sub\n'
                                   'Public Sub Quiet()\n    Dim i As Long\n    For i = 1 To 600\n    Next\nEnd Sub\n')
        host.save()
    app = ExcelApplication.open(path, step_limit=1000)
    app.run('Module1.Quiet')
    with pytest.raises(VBAStepLimitError):
        app.run('Module1.Change')
    assert ExcelApplication.open(path, step_limit=5000).run('Module1.Change') is None


@pytest.mark.parametrize('limit', [0, -1, 1.5, True, '10'])
def test_a_limit_must_be_a_whole_number_of_at_least_one(limit: object) -> None:
    with pytest.raises(ValueError, match='step_limit'):
        ExcelApplication(step_limit=limit)  # type: ignore[arg-type]
    with pytest.raises(ValueError, match='step_limit'):
        WordApplication(step_limit=limit)  # type: ignore[arg-type]


@pytest.mark.parametrize('make', [WordApplication, PowerPointApplication])
def test_word_and_powerpoint_take_the_limit_too(make: type) -> None:
    app = make(step_limit=500)
    app.add_module('Public Sub Forever()\n    Do\n    Loop\nEnd Sub\n', name='H')
    with pytest.raises(VBAStepLimitError):
        app.run('H.Forever')
