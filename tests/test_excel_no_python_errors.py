"""Whatever a macro passes, what comes out is a VBA error or an answer. A Python exception reaching a macro is one On Error cannot treat, and no comparison with Excel looks for it.

This needs no Excel, only that no other kind of exception escapes. It calls members directly from Python, where the sandbox does not apply (it gates what VBA is running), so a
SaveAs given "x" really writes "x". So every test here starts in an empty temporary folder and refuses to run anywhere else, and no value can name a path that leaves it.
"""
from __future__ import annotations

import faulthandler
from collections.abc import Callable, Iterator
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAError
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import EMPTY, MISSING, NOTHING, NULL

#: Mail and routing are never touched, and these wait, end the session or put off work to a time.
LEFT_OUT = frozenset({
    'sendmail', 'sendforreview', 'replywithchanges', 'endreview', 'route', 'routeworkbook', 'routingslip', 'hasroutingslip', 'hasmailer', 'mailer', 'mailenvelope',
    'sendfaxoverinternet', 'maillogon', 'maillogoff', 'mailsession', 'mailsystem', 'followhyperlink', 'follow',
    'wait', 'quit', 'ontime', 'onkey', 'onrepeat', 'onundo', 'volatile', 'stop',
})
TEXTS = ('x', '', '2')
SETUP = '''Sub Setup()
    Dim ws As Worksheet, wb As Workbook, pc As PivotCache
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("Name", "Team", "Qty")
    ws.Range("A2:C2").Value = Array("Ann", "x", 5)
    ws.Range("A3:C3").Value = Array("Bob", "y", 7)
    ws.Range("E1").Formula = "=SUM(C2:C3)"
    wb.Names.Add "Total", "=Sheet1!$E$1"
    ws.Range("A1").AddComment "note"
    ws.Range("C2:C3").FormatConditions.Add xlCellValue, xlGreater, "1"
    ws.Shapes.AddShape 1, 10, 10, 50, 50
    ws.ChartObjects.Add(10, 100, 200, 100).Chart.SetSourceData ws.Range("A1:C3")
    ws.ListObjects.Add xlSrcRange, ws.Range("H1:I3"), , xlYes
    Set pc = wb.PivotCaches.Create(xlDatabase, "Sheet1!R1C1:R3C3")
    pc.CreatePivotTable ws.Range("L3"), "P1"
End Sub
Function Awkward(ByVal which As Long) As Variant
    Dim a() As Long
    If which = 1 Then Awkward = Array(1, "a")
    If which = 2 Then Awkward = a
    If which = 3 Then Awkward = #1/2/2020#
End Function
'''
#: Each a way from Application to an object, reading only properties.
PATHS: dict[str, list[tuple[str, list[object]]]] = {
    'Application': [],
    'Workbook': [('ActiveWorkbook', [])],
    'Worksheet': [('ActiveSheet', [])],
    'Worksheets': [('ActiveWorkbook', []), ('Worksheets', [])],
    'Range': [('ActiveSheet', []), ('Range', ['A1:B2'])],
    'Font': [('ActiveSheet', []), ('Range', ['A1']), ('Font', [])],
    'Interior': [('ActiveSheet', []), ('Range', ['A1']), ('Interior', [])],
    'Borders': [('ActiveSheet', []), ('Range', ['A1']), ('Borders', [])],
    'Name': [('ActiveWorkbook', []), ('Names', [1])],
    'Comment': [('ActiveSheet', []), ('Range', ['A1']), ('Comment', [])],
    'Shapes': [('ActiveSheet', []), ('Shapes', [])],
    'Shape': [('ActiveSheet', []), ('Shapes', [1])],
    'ChartObject': [('ActiveSheet', []), ('ChartObjects', [1])],
    'Chart': [('ActiveSheet', []), ('ChartObjects', [1]), ('Chart', [])],
    'ListObject': [('ActiveSheet', []), ('ListObjects', [1])],
    'PivotTable': [('ActiveSheet', []), ('PivotTables', [1])],
    'PageSetup': [('ActiveSheet', []), ('PageSetup', [])],
    'Window': [('ActiveWindow', [])],
}


@pytest.fixture(autouse=True)
def in_an_empty_folder(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Iterator[None]:
    monkeypatch.chdir(tmp_path)
    # Nothing here may run on long enough to hide a member that never returns.
    faulthandler.dump_traceback_later(400, exit=True)
    yield
    faulthandler.cancel_dump_traceback_later()


def test_no_text_given_can_name_a_path_out_of_the_folder() -> None:
    for text in TEXTS:
        assert not set(text) & set('/\\:.~%$'), text


def _fresh() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SETUP, name='WalkSetup')
    app.run('WalkSetup.Setup')
    return app


def _values(app: ExcelApplication) -> dict[str, object]:
    made: dict[str, object] = {text or 'blank': text for text in TEXTS}
    made.update(null=NULL, empty=EMPTY, negative=-1, huge=1e308, wide=2.0 ** 40, nothing=NOTHING)
    made['array'] = app.interpreter.run('WalkSetup.Awkward', [1])
    made['unallocated'] = app.interpreter.run('WalkSetup.Awkward', [2])
    made['date'] = app.interpreter.run('WalkSetup.Awkward', [3])
    made['object'] = _follow(app, [('ActiveSheet', []), ('Range', ['A1'])])
    return made


def _follow(app: ExcelApplication, path: list[tuple[str, list[object]]]) -> Any:
    subject: Any = app.application
    for name, args in path:
        subject = subject.vba_get(name, args)
    return subject


def _escapes(call: Callable[[], object]) -> str | None:
    try:
        call()
    except VBAError:
        return None
    except (KeyboardInterrupt, SystemExit):
        raise
    except BaseException as failed:  # noqa: BLE001 - this is what is looked for
        return f'{type(failed).__name__}: {str(failed)[:80]}'
    return None


@pytest.mark.parametrize('name', sorted(PATHS))
def test_no_python_error_escapes_a_member(name: str) -> None:
    path = PATHS[name]
    sample = _follow(_fresh(), path)
    assert isinstance(sample, VBAObject)
    members = {key: spec for key, spec in type(sample)._vba_members.items() if key not in LEFT_OUT}  # pyright: ignore[reportPrivateUsage]
    found: list[str] = []
    for spec in members.values():
        app = _fresh()
        for value_name in list(_values(app)):
            def again() -> tuple[Any, dict[str, object]]:
                nonlocal app
                try:
                    return _follow(app, path), _values(app)
                except Exception:  # noqa: BLE001 - an earlier call took the object away: begin again
                    app = _fresh()
                    return _follow(app, path), _values(app)

            subject, made = again()
            value = made[value_name]
            for position in range(max(len(spec.parameters), 1 if spec.varargs else 0)):
                escape = _escapes(lambda: subject.vba_get(spec.name, [MISSING] * position + [value]))
                if escape:
                    found.append(f'{name}.{spec.name} argument {position} {value_name}: {escape}')
                subject, made = again()
                value = made[value_name]
            if spec.setter is not None:
                escape = _escapes(lambda: subject.vba_set(spec.name, value))
                if escape:
                    found.append(f'{name}.{spec.name} set {value_name}: {escape}')
    assert not found, found[:30]
